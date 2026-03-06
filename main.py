"""Main CLI entry point for the accounting mailbox reader."""

import json
import logging
import os
import sys
import traceback
from datetime import datetime
from pathlib import Path

import click

from src.email_reader import EmailReader
from src.output_formatter import OutputFormatter
from src.config import config
from src.email_classifier import EmailClassifier
from src.confluence_sync import ConfluenceSyncer
from src.pass2_processor import Pass2Processor, reconcile_priority, _PRIORITY_RANK
from src.pass2_inv_processor import Pass2InvProcessor
from src.keyword_triage import KeywordTriage
from src.email_translator import EmailTranslator
from src.conversation_matcher import match_conversations, update_superseded_jsons
from src.teams_notifier import TeamsNotifier
from src.graph_client import GraphAPIClient
from src.daily_stats import DailyStats
from src.jira_client import JiraClient
from src.correction_logger import scan_corrections, get_pending_corrections_count

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
)

# Add file handler to SharePoint-synced folder if configured
_log_folder = os.getenv("LOCAL_FOLDER_PATH", "")
if _log_folder:
    from logging.handlers import RotatingFileHandler
    _log_file = os.path.join(_log_folder, "process.log")
    _file_handler = RotatingFileHandler(_log_file, maxBytes=1_000_000, backupCount=3, encoding="utf-8")
    _file_handler.setLevel(logging.INFO)
    _file_handler.setFormatter(logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s"))
    logging.getLogger().addHandler(_file_handler)

# Suppress verbose Azure SDK logging
logging.getLogger("azure.core.pipeline.policies.http_logging_policy").setLevel(logging.WARNING)
logging.getLogger("azure.identity").setLevel(logging.WARNING)
logging.getLogger("msal").setLevel(logging.WARNING)

logger = logging.getLogger(__name__)


def send_error_notification(error_message: str, error_details: str):
    """Send email notification to admin when process fails."""
    try:
        recipient_email = "matthias.dibianco@eatplanted.com"
        graph = GraphAPIClient()

        subject = f"[ALERT] Accounting Mailbox Reader Process Failed - {datetime.now().strftime('%Y-%m-%d %H:%M')}"
        body = f"""
Process failed with the following error:

ERROR: {error_message}

DETAILS:
{error_details}

Timestamp: {datetime.now().isoformat()}
Mailbox: {config.accounting_mailbox}

Please check the logs at:
C:\\Users\\MatthiasDiBianco\\.accounting_mailbox_reader\\process.log
"""

        graph.send_mail(
            mailbox=config.accounting_mailbox,
            to_recipients=[recipient_email],
            subject=subject,
            body=body,
            is_html=False
        )
        logger.info(f"Error notification sent to {recipient_email}")
    except Exception as e:
        logger.error(f"Failed to send error notification: {e}")


def send_daily_summary():
    """Send daily summary email aggregating all runs since last summary."""
    try:
        recipient_email = "matthias.dibianco@eatplanted.com"
        agg = DailyStats.aggregate_runs_since_last_summary()

        if agg["run_count"] == 0:
            logger.info("No runs since last summary — skipping email")
            return

        # Format period
        since_dt = datetime.fromisoformat(agg["since"].replace("Z", "+00:00"))
        since_str = since_dt.strftime("%Y-%m-%d %H:%M")
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M")

        # LLM breakdown
        llm_breakdown = ""
        if agg["llm_calls_by_model"]:
            llm_breakdown = "\n  LLM Calls by Model:"
            for model, count in sorted(agg["llm_calls_by_model"].items()):
                llm_breakdown += f"\n    - {model}: {count}"

        # Corrections warning
        pending_corrections = get_pending_corrections_count()
        corrections_section = ""
        if pending_corrections > 0:
            corrections_section = f"""

CLASSIFICATION CORRECTIONS

*** {pending_corrections} correction(s) queued ***
Tell Claude Code: "update from corrections"
File: config/corrections.yaml"""

        body = f"""Accounting Mailbox Reader - Summary
Period: {since_str} to {now_str}

Runs: {agg['run_count']}

{agg['total_emails_processed']} emails processed
  - {agg['categories_by_keywords']} categorized by keywords
  - {agg['categories_by_llm']} categorized by LLM
{agg['total_ven_rem_analyzed']} VEN-REM analyzed
{agg['total_ven_followup_analyzed']} VEN-FOLLOWUP analyzed
{agg['total_ven_inv_processed']} VEN-INV processed
{agg['total_emails_archived']} emails archived
{agg['total_human_completed_moved']} human-completed moved
{agg['total_jsons_saved']} JSONs saved{llm_breakdown}{corrections_section}
"""

        graph = GraphAPIClient()
        graph.send_mail(
            mailbox=config.accounting_mailbox,
            to_recipients=[recipient_email],
            subject=f"Daily Summary - Accounting Mailbox - {now_str}",
            body=body,
            is_html=False
        )
        logger.info(f"Daily summary sent to {recipient_email}")

        DailyStats.set_last_summary_sent()
    except Exception as e:
        logger.error(f"Failed to send daily summary: {e}")


@click.group()
def cli():
    """Accounting Mailbox Reader - Extract and analyze emails from the accounting mailbox."""
    pass


@cli.command()
@click.option(
    "--mailbox",
    default=config.accounting_mailbox,
    help="Target mailbox email address",
)
@click.option(
    "--days",
    default=config.days_back,
    type=int,
    help="How many days back to read emails",
)
@click.option(
    "--max",
    default=config.max_emails,
    type=int,
    help="Maximum number of emails to read",
)
@click.option(
    "--format",
    type=click.Choice(["json", "table", "detailed"]),
    default="table",
    help="Output format",
)
@click.option(
    "--output",
    type=click.Path(),
    help="Save output to file (optional)",
)
@click.option(
    "--no-attachments",
    is_flag=True,
    help="Skip attachment extraction",
)
@click.option(
    "--no-body",
    is_flag=True,
    help="Skip full body content",
)
@click.option(
    "--search",
    help="Optional search query (OData syntax)",
)
@click.option(
    "--upload-sharepoint",
    is_flag=True,
    help="Upload individual email JSONs to SharePoint",
)
@click.option(
    "--classify",
    is_flag=True,
    help="Run LLM classification on low-confidence keyword results",
)
@click.option(
    "--deep",
    is_flag=True,
    help="Run Pass 2 deep analysis on VEN-REM emails (implies --classify)",
)
@click.option(
    "--force-llm",
    is_flag=True,
    help="Force LLM classification on ALL emails (reset avenue, overrides keywords)",
)
@click.option(
    "--write-back",
    is_flag=True,
    help="Write category tags back to Outlook emails",
)
def read(mailbox, days, max, format, output, no_attachments, no_body, search, upload_sharepoint, classify, deep, force_llm, write_back):
    """Read emails from the accounting mailbox."""
    try:
        logger.info("Initializing email reader...")
        reader = EmailReader()

        logger.info("Fetching emails...")
        emails = reader.read_emails(
            max_results=max,
            days_back=days,
            search_query=search,
            extract_attachments=not no_attachments,
        )

        if not emails:
            click.echo("No emails found.")
            return

        click.echo(f"\n[OK] Found {len(emails)} emails\n")

        # --deep implies --classify
        if deep:
            classify = True

        # ── Pass 0: Keyword triage (always runs, zero API calls) ──
        click.echo("Running keyword triage (Pass 0)...")
        triage = KeywordTriage()
        for idx, email in enumerate(emails, 1):
            keyword_result = triage.classify(email.to_dict())
            email.classification = keyword_result
            cat = keyword_result["primary_category"]["id"]
            conf = keyword_result["confidence_level"]
            prio = keyword_result["priority"]
            click.echo(f"  [{idx}/{len(emails)}] {cat} ({conf}) {prio} - {email.subject[:50]}...")

        high_count = sum(1 for e in emails if e.classification.get("confidence_level") == "HIGH")
        med_count = sum(1 for e in emails if e.classification.get("confidence_level") == "MEDIUM")
        low_count = sum(1 for e in emails if e.classification.get("confidence_level") == "LOW")
        click.echo(f"[OK] Keyword triage: {high_count} HIGH, {med_count} MEDIUM, {low_count} LOW\n")

        # ── Conversation matching (group RE/FW chains) ──
        if config.local_folder_path:
            click.echo("Matching conversations...")
            conv_results = match_conversations(emails, config.local_folder_path)
            linked = 0
            for email in emails:
                info = conv_results.get(email.id)
                if info:
                    email.conversation_id = info["conversation_id"]
                    email.conversation_position = info["position"]
                    email.is_latest_in_conversation = info["is_latest"]
                    email.related_emails = info["related_emails"]
                    email.is_chain = info.get("is_chain", False)
                    if info["position"] > 1:
                        linked += 1
            click.echo(
                f"[OK] Conversations: {linked} linked to existing, "
                f"{sum(1 for e in emails if not e.is_latest_in_conversation)} superseded\n"
            )

        # ── LLM classification (only when needed) ──
        if classify or force_llm:
            if force_llm:
                emails_for_llm = list(emails)
                click.echo(f"Force-LLM: classifying ALL {len(emails_for_llm)} emails with LLM...")
            else:
                emails_for_llm = [
                    e for e in emails
                    if e.classification.get("confidence_level") == "LOW"
                ]

            if emails_for_llm:
                if not force_llm:
                    click.echo(f"Running LLM classification on {len(emails_for_llm)} low-confidence emails...")
                try:
                    classifier = EmailClassifier()
                    for idx, email in enumerate(emails_for_llm, 1):
                        click.echo(f"  LLM [{idx}/{len(emails_for_llm)}]: {email.subject[:50]}...")
                        keyword_result = email.classification
                        llm_result = classifier.classify(
                            email.to_dict(),
                            keyword_classification=keyword_result,
                        )
                        llm_result["classification_method"] = "llm"
                        llm_result["keyword_classification"] = keyword_result
                        # Ensure LLM cannot downgrade keyword priority (take max)
                        kw_prio = keyword_result.get("priority", "PRIO_MEDIUM")
                        llm_prio = llm_result.get("priority", "PRIO_MEDIUM")
                        if _PRIORITY_RANK.get(kw_prio, 1) > _PRIORITY_RANK.get(llm_prio, 1):
                            llm_result["priority"] = kw_prio
                        email.classification = llm_result
                    click.echo(f"[OK] LLM classified {len(emails_for_llm)} emails\n")
                except Exception as e:
                    logger.error(f"LLM classification error: {e}", exc_info=True)
                    click.echo(f"[ERR] LLM classification failed: {e}", err=True)
            else:
                click.echo("All emails HIGH/MEDIUM confidence — no LLM classification needed.\n")

        # ── Pass 2: Deep analysis on VEN-REM emails ──
        if deep:
            ven_rem_emails = [
                e for e in emails
                if e.classification
                and e.classification.get("primary_category", {}).get("id") == "VEN-REM"
            ]

            if ven_rem_emails:
                click.echo(f"Running Pass 2 deep analysis on {len(ven_rem_emails)} VEN-REM emails...")
                try:
                    processor = Pass2Processor()
                    if not processor.invoice_lookup.is_configured:
                        click.echo(
                            "[WARN] Fabric SQL not configured. "
                            "LLM extraction will run but BC lookups will be skipped. "
                            "Set FABRIC_SQL_ENDPOINT, FABRIC_CLIENT_ID, FABRIC_CLIENT_SECRET in .env"
                        )
                    for idx, email in enumerate(ven_rem_emails, 1):
                        click.echo(
                            f"  Pass 2 [{idx}/{len(ven_rem_emails)}]: "
                            f"{email.subject[:50]}..."
                        )
                        pass2_result = processor.process_email(email.to_dict())
                        if pass2_result:
                            email.pass2_results = pass2_result
                            # Reconcile priority: max(keyword/LLM prio, reminder-level prio)
                            urgency = pass2_result.get("urgency_level")
                            if urgency is not None:
                                old_prio = email.classification.get("priority", "PRIO_MEDIUM")
                                new_prio = reconcile_priority(old_prio, urgency)
                                if new_prio != old_prio:
                                    logger.info(f"Priority reconciled: {old_prio} -> {new_prio} (urgency_level={urgency})")
                                email.classification["priority"] = new_prio
                            # Handle reclassification
                            if pass2_result.get("reclassified"):
                                new_cat = pass2_result["reclassified_to"]
                                click.echo(
                                    f"    [RECLASSIFIED] VEN-REM -> {new_cat}: "
                                    f"{pass2_result.get('verification_reasoning', '')[:60]}..."
                                )
                                email.classification["primary_category"] = {
                                    "id": new_cat,
                                    "name": KeywordTriage.CATEGORY_NAMES.get(new_cat, new_cat),
                                }
                                email.classification["classification_method"] = "llm_reclassified"
                    processor.close()
                    click.echo(f"[OK] Pass 2 complete for {len(ven_rem_emails)} emails\n")
                except Exception as e:
                    logger.error(f"Pass 2 error: {e}", exc_info=True)
                    click.echo(f"[ERR] Pass 2 failed: {e}", err=True)
            else:
                click.echo("No VEN-REM emails found for Pass 2 analysis.\n")

        # ── Translation: English summary + body for non-English emails ──
        # Translate emails that: have no pass2_results, confidence >= MEDIUM, not already English
        emails_for_translation = [
            e for e in emails
            if e.classification
            and not e.pass2_results
            and e.classification.get("confidence_level") in ("HIGH", "MEDIUM")
            and not EmailTranslator.is_likely_english(e.to_dict())
        ]

        if emails_for_translation:
            click.echo(f"Translating {len(emails_for_translation)} non-English emails...")
            try:
                translator = EmailTranslator()
                for idx, email in enumerate(emails_for_translation, 1):
                    click.echo(f"  Translate [{idx}/{len(emails_for_translation)}]: {email.subject[:50]}...")
                    result = translator.translate(email.to_dict())
                    if result:
                        if result.get("summary"):
                            email.classification["summary"] = result["summary"]
                        if result.get("body_english"):
                            email.body_english = result["body_english"]
                        if result.get("model_used"):
                            email.classification["translation_model"] = result["model_used"]
                click.echo(f"[OK] Translated {len(emails_for_translation)} emails\n")
            except Exception as e:
                logger.error(f"Translation error: {e}", exc_info=True)
                click.echo(f"[ERR] Translation failed: {e}", err=True)

        # ── Outlook write-back ──
        if write_back:
            click.echo("Writing categories + extended properties back to Outlook...")
            write_count = 0
            ext_prop_id = "String {00020329-0000-0000-C000-000000000046} Name EmailClassifierData"
            for email in emails:
                if not email.classification:
                    continue
                cat_id = email.classification.get("primary_category", {}).get("id", "")
                prio = email.classification.get("priority", "")
                categories = [c for c in [cat_id, prio] if c]
                if categories:
                    # Build extended property payload (classification + pass2 + translation)
                    ext_data = dict(email.classification)
                    if hasattr(email, "body_english") and email.body_english:
                        ext_data["body_english"] = email.body_english[:2000]
                    if hasattr(email, "pass2_results") and email.pass2_results:
                        ext_data["pass2_results"] = email.pass2_results
                    updates = {
                        "categories": categories,
                        "isRead": True,
                        "singleValueExtendedProperties": [{
                            "id": ext_prop_id,
                            "value": json.dumps(ext_data, ensure_ascii=False, default=str),
                        }],
                    }
                    result = reader.graph_client.update_message(
                        mailbox, email.id, updates
                    )
                    if result:
                        write_count += 1
                    else:
                        click.echo(f"  [WARN] Failed to write back: {email.subject[:40]}")
            click.echo(f"[OK] Wrote categories + properties to {write_count}/{len(emails)} emails\n")

        # Upload to SharePoint if requested
        if upload_sharepoint:
            # Priority: Local folder > Power Automate > SharePoint direct
            if config.local_folder_path:
                click.echo(f"Saving emails to local folder: {config.local_folder_path}")
                stats = OutputFormatter.save_to_local_folder(
                    emails,
                    folder_path=config.local_folder_path
                )
                result_msg = "Local Save Complete"
            elif config.power_automate_flow_url:
                click.echo("Uploading emails via Power Automate...")
                stats = OutputFormatter.save_via_power_automate(
                    emails,
                    flow_url=config.power_automate_flow_url
                )
                result_msg = "Power Automate Upload Complete"
            else:
                click.echo("Uploading emails to SharePoint (direct)...")
                # Reuse the same authenticated graph_client from the reader
                stats = OutputFormatter.save_to_sharepoint(emails, graph_client=reader.graph_client)
                result_msg = "SharePoint Upload Complete"

            click.echo(
                f"\n[OK] {result_msg}:\n"
                f"  - Total: {stats['total']}\n"
                f"  - Successful: {stats['successful']}\n"
                f"  - Failed: {stats['failed']}"
            )
            if stats['errors']:
                click.echo("\nErrors:")
                for error in stats['errors']:
                    click.echo(f"  - {error}")
            click.echo()

        # Format output
        if format == "json":
            result = OutputFormatter.to_json(emails, pretty=True)
        elif format == "detailed":
            result = OutputFormatter.to_detailed_text(
                emails,
                include_body=not no_body,
                include_extracted_text=True,
            )
        else:  # table
            result = OutputFormatter.to_console_table(emails, include_attachments=True)

        # Output (replace Unicode chars that Windows cp1252 can't encode)
        try:
            click.echo(result)
        except UnicodeEncodeError:
            safe = result.encode("cp1252", errors="replace").decode("cp1252")
            click.echo(safe)

        # Save to file if requested
        if output:
            if OutputFormatter.save_to_file(result, output):
                click.echo(f"\n[OK] Output saved to {output}")

    except Exception as e:
        logger.error(f"Error: {e}", exc_info=True)
        click.echo(f"Error: {e}", err=True)
        sys.exit(1)


@cli.command()
@click.option(
    "--no-attachments",
    is_flag=True,
    help="Skip attachment extraction",
)
@click.option(
    "--search",
    help="Optional search query (OData syntax)",
)
def preview(no_attachments, search):
    """Preview emails without saving to file."""
    try:
        logger.info("Initializing email reader...")
        reader = EmailReader()

        logger.info("Fetching emails (limited to 5 for preview)...")
        emails = reader.read_emails(
            max_results=5,
            days_back=1,
            search_query=search,
            extract_attachments=not no_attachments,
        )

        if not emails:
            click.echo("No emails found.")
            return

        click.echo(f"\n[OK] Preview of {len(emails)} most recent emails\n")
        result = OutputFormatter.to_console_table(emails, include_attachments=True)
        click.echo(result)

    except Exception as e:
        logger.error(f"Error: {e}", exc_info=True)
        click.echo(f"Error: {e}", err=True)
        sys.exit(1)


@cli.command()
def config_show():
    """Show current configuration."""
    click.echo("\n" + "="*60)
    click.echo("CURRENT CONFIGURATION")
    click.echo("="*60)

    click.echo(f"\nMailbox:           {config.accounting_mailbox}")
    click.echo(f"Dry Run Mode:      {config.dry_run}")
    click.echo(f"Max Emails:        {config.max_emails}")
    click.echo(f"Days Back:         {config.days_back}")
    click.echo(f"Max Attachment Size: {config.max_attachment_size_mb} MB")
    click.echo(f"Attachment Formats: {', '.join(config.attachment_formats)}")

    click.echo("\nAzure Setup:")
    click.echo(f"  Client ID:     {'[OK] Configured' if config.azure_client_id else '[ERR] Missing'}")
    click.echo(f"  Tenant ID:     {'[OK] Configured' if config.azure_tenant_id else '[ERR] Missing'}")
    click.echo(f"  Client Secret: {'[OK] Configured' if config.azure_client_secret else '[ERR] Missing'}")

    click.echo("\nFabric / Business Central (Pass 2):")
    click.echo(f"  SQL Endpoint:  {'[OK] ' + config.fabric_sql_endpoint[:40] + '...' if config.fabric_sql_endpoint else '[  ] Not configured'}")
    click.echo(f"  Database:      {'[OK] ' + config.fabric_database if config.fabric_database else '[  ] Not configured'}")
    click.echo(f"  Client ID:     {'[OK] Configured' if config.fabric_client_id else '[  ] Not configured'}")
    click.echo(f"  Client Secret: {'[OK] Configured' if config.fabric_client_secret else '[  ] Not configured'}")
    tc = config.bc_table_config
    click.echo(f"  Table:         [{tc['database']}].[{tc['schema']}].[{tc['table']}]")

    click.echo("\n" + "="*60 + "\n")


@cli.command()
def init():
    """Initialize .env file with required variables."""
    env_file = Path(".env")

    if env_file.exists():
        click.echo("[OK] .env file already exists")
        return

    env_example = Path(".env.example")
    if env_example.exists():
        with open(env_example) as f:
            content = f.read()
        with open(env_file, "w") as f:
            f.write(content)
        click.echo("[OK] Created .env file from .env.example")
        click.echo("\nNext steps:")
        click.echo("1. Edit .env and fill in your Azure credentials")
        click.echo("2. Run: python main.py config-show")
        click.echo("3. Run: python main.py preview")
    else:
        click.echo("[ERR] .env.example not found")
        sys.exit(1)


@cli.command()
def sync_categories():
    """Sync category table from Confluence."""
    try:
        click.echo("Syncing categories from Confluence...")
        syncer = ConfluenceSyncer()

        if syncer.sync_categories():
            click.echo("\n[OK] Categories synced successfully")
        else:
            click.echo("\n[ERR] Category sync failed", err=True)
            sys.exit(1)

    except Exception as e:
        logger.error(f"Error syncing categories: {e}", exc_info=True)
        click.echo(f"[ERR] Error: {e}", err=True)
        sys.exit(1)


@cli.command()
@click.option(
    "--email-id",
    help="Test classification on a specific email ID",
)
def test_classify(email_id):
    """Test email classification on a single email."""
    try:
        logger.info("Initializing classifier...")
        classifier = EmailClassifier()

        logger.info("Initializing email reader...")
        reader = EmailReader()

        if email_id:
            # Fetch specific email
            click.echo(f"Fetching email ID: {email_id}")
            # Note: Would need to add a method to fetch single email by ID
            click.echo("[ERR] Fetching by ID not yet implemented. Use --days and --max to test.")
            sys.exit(1)
        else:
            # Fetch most recent email
            click.echo("Fetching most recent email for testing...")
            emails = reader.read_emails(
                max_results=1,
                days_back=7,
                extract_attachments=False,
            )

            if not emails:
                click.echo("[ERR] No emails found to classify")
                sys.exit(1)

            email = emails[0]

        click.echo(f"\n{'='*80}")
        click.echo(f"Testing classification on:")
        click.echo(f"Subject: {email.subject}")
        click.echo(f"From: {email.from_name} <{email.from_email}>")
        click.echo(f"Date: {email.received_datetime}")
        click.echo(f"{'='*80}\n")

        # Classify
        click.echo("Classifying...")
        classification = classifier.classify(email.to_dict())

        # Display results
        click.echo("\n" + "="*80)
        click.echo("CLASSIFICATION RESULT")
        click.echo("="*80)
        click.echo(f"\nCategory: {classification['primary_category']['id']} - {classification['primary_category']['name']}")
        if classification.get('secondary_category'):
            click.echo(f"Secondary: {classification['secondary_category']['id']} - {classification['secondary_category']['name']}")
        click.echo(f"\nPriority: {classification['priority']}")
        click.echo(f"Confidence: {classification['confidence_level']}")
        click.echo(f"Manual Review: {classification['requires_manual_review']}")

        if classification.get('extracted_entities'):
            click.echo(f"\nExtracted Entities:")
            for key, value in classification['extracted_entities'].items():
                if value:
                    click.echo(f"  {key}: {value}")

        click.echo(f"\nReasoning:")
        click.echo(f"{classification['reasoning']}")
        click.echo("="*80)

    except Exception as e:
        logger.error(f"Error testing classification: {e}", exc_info=True)
        click.echo(f"[ERR] Error: {e}", err=True)
        sys.exit(1)


@cli.command()
@click.option(
    "--mailbox",
    default=config.accounting_mailbox,
    help="Target mailbox email address",
)
@click.option(
    "--upload-sharepoint",
    is_flag=True,
    help="Upload/save email JSONs",
)
@click.option(
    "--dry-run",
    is_flag=True,
    help="Show what would happen without making changes in Outlook",
)
def process(mailbox, upload_sharepoint, dry_run):
    """Automated processing run: classify new emails, deep-analyze VEN-REM, archive processed.

    This is the command meant for scheduled execution. It:
    1. Fetches only NEW emails (since last run watermark)
    2. Classifies all (keywords + LLM on low-confidence)
    3. Flags them as OPEN in Outlook
    4. Runs Pass 2 on VEN-REM emails
    5. Translates non-English emails
    6. Archives agent-processed emails (Pass 2) to ARCHIVE/PROCESSED BY AGENT
    7. Scans inbox for human-completed flags → moves to ARCHIVE/PROCESSED BY HUMANS
    8. Saves JSONs and updates watermark
    """
    import hashlib
    from datetime import datetime, timedelta, timezone
    from src.watermark import get_watermark, update_watermark
    from src import api_counter

    try:
        click.echo("=" * 60)
        click.echo("AUTOMATED PROCESSING RUN")
        click.echo("=" * 60)

        # ── Health check: fail fast if Graph API is unreachable ──
        graph_check = GraphAPIClient()
        ok, msg = graph_check.check_health(mailbox)
        if not ok:
            click.echo(f"[ERR] Graph API health check failed: {msg}", err=True)
            logger.error(f"Health check failed: {msg}")
            sys.exit(1)
        click.echo(f"Graph API: OK")

        # Show API budget and snapshot for tracking
        api_snapshot_before = api_counter.get_today_breakdown()
        used_today = api_counter.get_today_total()
        remaining = api_counter.get_remaining()
        click.echo(f"API calls today: {used_today} | Remaining: {remaining}")

        # ── 1. Watermark: only fetch new emails ──
        watermark = get_watermark()
        click.echo(f"\nWatermark: {watermark}")

        watermark_dt = datetime.fromisoformat(watermark.replace("Z", "+00:00"))
        days_since = max((datetime.now(timezone.utc) - watermark_dt).days + 1, 1)

        reader = EmailReader()
        emails = reader.read_emails(
            max_results=250,
            days_back=days_since,
            extract_attachments=True,
        )

        # Filter to only emails AFTER watermark
        if emails:
            emails = [
                e for e in emails
                if e.received_datetime > watermark
            ]

        if not emails:
            click.echo("No new emails since last run.")
        else:
            click.echo(f"[OK] {len(emails)} new email(s) to process\n")

            # ── 2. Keyword triage (Pass 0) ──
            click.echo("Pass 0: Keyword triage...")
            triage = KeywordTriage()
            for email in emails:
                keyword_result = triage.classify(email.to_dict())
                email.classification = keyword_result
                email.processing_status = "OPEN"

            high = sum(1 for e in emails if e.classification.get("confidence_level") == "HIGH")
            med = sum(1 for e in emails if e.classification.get("confidence_level") == "MEDIUM")
            low = sum(1 for e in emails if e.classification.get("confidence_level") == "LOW")
            click.echo(f"  Keywords: {high} HIGH, {med} MEDIUM, {low} LOW")

            # ── 2b. Conversation matching ──
            if config.local_folder_path:
                click.echo("  Matching conversations...")
                conv_results = match_conversations(emails, config.local_folder_path)
                linked = 0
                for email in emails:
                    info = conv_results.get(email.id)
                    if info:
                        email.conversation_id = info["conversation_id"]
                        email.conversation_position = info["position"]
                        email.is_latest_in_conversation = info["is_latest"]
                        email.related_emails = info["related_emails"]
                        email.is_chain = info.get("is_chain", False)
                        if info["position"] > 1:
                            linked += 1
                superseded = sum(1 for e in emails if not e.is_latest_in_conversation)
                chains = sum(1 for e in emails if e.is_chain)
                click.echo(f"  Conversations: {linked} linked, {superseded} superseded, {chains} in chains")

            # ── 3. LLM classification on LOW confidence ──
            low_conf = [e for e in emails if e.classification.get("confidence_level") == "LOW"]
            if low_conf:
                click.echo(f"  LLM classifying {len(low_conf)} low-confidence...")
                try:
                    classifier = EmailClassifier()
                    for email in low_conf:
                        keyword_result = email.classification
                        llm_result = classifier.classify(
                            email.to_dict(), keyword_classification=keyword_result
                        )
                        llm_result["classification_method"] = "llm"
                        llm_result["keyword_classification"] = keyword_result
                        # Ensure LLM cannot downgrade keyword priority (take max)
                        kw_prio = keyword_result.get("priority", "PRIO_MEDIUM")
                        llm_prio = llm_result.get("priority", "PRIO_MEDIUM")
                        if _PRIORITY_RANK.get(kw_prio, 1) > _PRIORITY_RANK.get(llm_prio, 1):
                            llm_result["priority"] = kw_prio
                        email.classification = llm_result
                    click.echo(f"  [OK] LLM classified {len(low_conf)} emails")
                except Exception as e:
                    click.echo(f"  [ERR] LLM classification: {e}", err=True)

            # ── 4. Flag all as "flagged" + write categories to Outlook ──
            click.echo("\nOutlook: flagging + writing categories...")
            ext_prop_id = "String {00020329-0000-0000-C000-000000000046} Name EmailClassifierData"
            for email in emails:
                if dry_run:
                    continue
                cat_id = email.classification.get("primary_category", {}).get("id", "")
                prio = email.classification.get("priority", "")
                categories = [c for c in [cat_id, prio] if c]
                if email.is_chain:
                    categories.append("CHAIN")

                ext_data = dict(email.classification)
                updates = {
                    "categories": categories,
                    "isRead": True,
                    "flag": {"flagStatus": "flagged"},
                    "singleValueExtendedProperties": [{
                        "id": ext_prop_id,
                        "value": json.dumps(ext_data, ensure_ascii=False, default=str),
                    }],
                }
                reader.graph_client.update_message(mailbox, email.id, updates)
            click.echo(f"  [OK] {len(emails)} emails flagged + categorized + marked read")

            # ── 5. Pass 2 on VEN-REM/VEN-FOLLOWUP emails ──
            ven_rem = [
                e for e in emails
                if e.classification.get("primary_category", {}).get("id") in {"VEN-REM", "VEN-FOLLOWUP"}
            ]

            if ven_rem:
                click.echo(f"\nPass 2: deep analysis on {len(ven_rem)} VEN-REM/VEN-FOLLOWUP...")
                try:
                    processor = Pass2Processor()
                    for email in ven_rem:
                        click.echo(f"  Pass 2: {email.subject[:50]}...")
                        pass2_result = processor.process_email(email.to_dict())
                        if pass2_result:
                            email.pass2_results = pass2_result
                            # Reconcile priority: max(keyword/LLM prio, reminder-level prio)
                            urgency = pass2_result.get("urgency_level")
                            if urgency is not None:
                                old_prio = email.classification.get("priority", "PRIO_MEDIUM")
                                new_prio = reconcile_priority(old_prio, urgency)
                                if new_prio != old_prio:
                                    logger.info(f"Priority reconciled: {old_prio} -> {new_prio} (urgency_level={urgency})")
                                email.classification["priority"] = new_prio
                            primary_cat = email.classification.get("primary_category", {}).get("id", "VEN-REM")
                            if pass2_result.get("reclassified"):
                                new_cat = pass2_result["reclassified_to"]
                                click.echo(f"    [RECLASSIFIED] {primary_cat} -> {new_cat}")
                                email.classification["primary_category"] = {
                                    "id": new_cat,
                                    "name": KeywordTriage.CATEGORY_NAMES.get(new_cat, new_cat),
                                }
                                email.classification["classification_method"] = "llm_reclassified"
                                # VEN-FOLLOWUP requires active response — do NOT auto-archive
                                if new_cat == "VEN-FOLLOWUP":
                                    email.processing_status = "NEEDS REVIEW > VENDOR QUERY"
                            else:
                                # Pass 2 confirmed classification
                                verified = pass2_result.get("verified_category", primary_cat)
                                if verified == "VEN-REM":
                                    # Confirmed VEN-REM reminder → auto-archive
                                    email.processing_status = "ARCHIVE > PROCESSED BY AGENT"
                                elif verified == "VEN-FOLLOWUP":
                                    # Confirmed VEN-FOLLOWUP → flag for active response
                                    email.processing_status = "NEEDS REVIEW > VENDOR QUERY"
                    processor.close()
                    click.echo(f"  [OK] Pass 2 complete")
                except Exception as e:
                    click.echo(f"  [ERR] Pass 2: {e}", err=True)

            # ── 5b. Pass 2 on VEN-INV emails ──
            ven_inv = [
                e for e in emails
                if e.classification.get("primary_category", {}).get("id") == "VEN-INV"
            ]

            if ven_inv:
                click.echo(f"\nPass 2: processing {len(ven_inv)} VEN-INV...")
                notifier = TeamsNotifier()
                processed_count = 0
                unknown_count = 0
                try:
                    processor = Pass2InvProcessor()
                    for email in ven_inv:
                        click.echo(f"  VEN-INV: {email.subject[:50]}...")
                        pass2_result = processor.process_email(
                            email.to_dict(), reader.graph_client, dry_run=dry_run
                        )
                        if pass2_result:
                            email.pass2_results = pass2_result
                            # VEN-INV stays in inbox for manual review (no auto-archive)
                            email.processing_status = "NEEDS REVIEW > NEW INVOICE"
                            action = pass2_result.get("action_taken", "UNKNOWN")
                            click.echo(f"    [{action}]")

                            # Send Teams notification
                            if action == "UNKNOWN_ENTITY":
                                notifier.notify_ven_inv_unknown_entity(email.to_dict())
                                unknown_count += 1
                            else:
                                notifier.notify_ven_inv_processed(email.to_dict(), pass2_result)
                                processed_count += 1

                    processor.close()

                    # Send summary notification
                    if processed_count > 0 or unknown_count > 0:
                        notifier.notify_run_summary(processed_count, unknown_count)

                    click.echo(f"  [OK] VEN-INV processing complete ({processed_count} processed, {unknown_count} unknown)")
                except Exception as e:
                    click.echo(f"  [ERR] VEN-INV: {e}", err=True)

            # ── 5c. Mark NO_ACTION_NEEDED emails for archiving ──
            no_action = [
                e for e in emails
                if e.classification.get("primary_category", {}).get("id") == "NO_ACTION_NEEDED"
            ]
            if no_action:
                click.echo(f"\nMarking {len(no_action)} NO_ACTION_NEEDED emails for archiving...")
                for email in no_action:
                    email.processing_status = "ARCHIVE > PROCESSED BY AGENT"
                click.echo(f"  [OK] {len(no_action)} emails marked for archiving")

            # ── 5d. Jira ticket creation ──
            jira = JiraClient()
            if jira.enabled:
                jira_candidates = []
                for email in emails:
                    cat_id = email.classification.get("primary_category", {}).get("id", "")
                    prio = email.classification.get("priority", "")
                    trigger = None
                    if cat_id == "VEN-FOLLOWUP":
                        trigger = "VEN-FOLLOWUP"
                    elif prio == "PRIO_HIGHEST":
                        trigger = "PRIO_HIGHEST"
                    elif cat_id == "OTHER" and prio in ("PRIO_HIGH", "PRIO_HIGHEST"):
                        trigger = "OTHER"
                    if trigger and jira.should_create_ticket(email.to_dict()):
                        jira_candidates.append((email, trigger))

                if jira_candidates:
                    click.echo(f"\nJira: creating tickets for {len(jira_candidates)} email(s)...")
                    jira_created = 0
                    for email, trigger in jira_candidates:
                        if dry_run:
                            click.echo(f"  [DRY-RUN] Would create ticket: {trigger} - {email.subject[:50]}")
                            continue
                        # Check for existing ticket (duplicate prevention)
                        existing = jira.find_existing_ticket(email.id)
                        if existing:
                            email.jira_issue_key = existing
                            click.echo(f"  [EXISTS] {existing}: {email.subject[:50]}")
                            continue
                        key = jira.create_ticket(email.to_dict(), trigger)
                        if key:
                            email.jira_issue_key = key
                            email.processing_status = "ARCHIVE > JIRA_TICKET_CREATED"
                            jira_created += 1
                            safe_subj = email.subject[:50].encode('utf-8', errors='replace').decode('utf-8', errors='ignore')
                            click.echo(f"  [CREATED] {key} ({trigger}): {safe_subj}")
                    click.echo(f"  [OK] Jira: {jira_created} ticket(s) created")

            # ── 6. Translation for non-English emails ──
            to_translate = [
                e for e in emails
                if e.classification
                and not e.pass2_results
                and e.classification.get("confidence_level") in ("HIGH", "MEDIUM")
                and not EmailTranslator.is_likely_english(e.to_dict())
            ]
            if to_translate:
                click.echo(f"\nTranslating {len(to_translate)} non-English emails...")
                try:
                    translator = EmailTranslator()
                    for email in to_translate:
                        result = translator.translate(email.to_dict())
                        if result:
                            if result.get("summary"):
                                email.classification["summary"] = result["summary"]
                            if result.get("body_english"):
                                email.body_english = result["body_english"]
                            if result.get("model_used"):
                                email.classification["translation_model"] = result["model_used"]
                    click.echo(f"  [OK] Translated {len(to_translate)} emails")
                except Exception as e:
                    click.echo(f"  [ERR] Translation: {e}", err=True)

            # ── 7a. Mark superseded emails in this batch for archiving ──
            superseded_in_batch = [e for e in emails if not e.is_latest_in_conversation]
            if superseded_in_batch:
                click.echo(f"\nMarking {len(superseded_in_batch)} superseded emails for archiving...")
                for email in superseded_in_batch:
                    email.processing_status = "ARCHIVE > SUPERSEDED"

            # ── 7b. Move agent-processed + superseded + Jira-ticketed to archive folder ──
            to_archive = [
                e for e in emails
                if e.processing_status in (
                    "ARCHIVE > PROCESSED BY AGENT",
                    "ARCHIVE > SUPERSEDED",
                    "ARCHIVE > JIRA_TICKET_CREATED",
                )
            ]
            if to_archive and not dry_run:
                click.echo(f"\nArchiving {len(to_archive)} emails...")
                folder_id_agent = reader.graph_client.get_or_create_folder(
                    mailbox, "ARCHIVE/PROCESSED BY AGENT"
                )
                folder_id_superseded = reader.graph_client.get_or_create_folder(
                    mailbox, "ARCHIVE/SUPERSEDED"
                )
                folder_id_jira = reader.graph_client.get_or_create_folder(
                    mailbox, "ARCHIVE/JIRA_TICKET_CREATED"
                )
                for email in to_archive:
                    # Complete the flag
                    reader.graph_client.flag_message(mailbox, email.id, "complete")
                    # Update extended property with final status
                    ext_data = dict(email.classification) if email.classification else {}
                    ext_data["processing_status"] = email.processing_status
                    if email.pass2_results:
                        ext_data["pass2_results"] = email.pass2_results
                    if email.jira_issue_key:
                        ext_data["jira_issue_key"] = email.jira_issue_key
                    reader.graph_client.update_message(mailbox, email.id, {
                        "singleValueExtendedProperties": [{
                            "id": ext_prop_id,
                            "value": json.dumps(ext_data, ensure_ascii=False, default=str),
                        }],
                    })
                    # Move to appropriate archive folder
                    if email.processing_status == "ARCHIVE > SUPERSEDED":
                        dest_folder = folder_id_superseded
                        label = "SUPERSEDED"
                    elif email.processing_status == "ARCHIVE > JIRA_TICKET_CREATED":
                        dest_folder = folder_id_jira
                        label = "JIRA"
                    else:
                        dest_folder = folder_id_agent
                        label = "ARCHIVED"
                    if dest_folder:
                        moved = reader.graph_client.move_message(mailbox, email.id, dest_folder)
                        safe_subject = email.subject[:50].encode('utf-8', errors='replace').decode('utf-8', errors='ignore')
                        if moved:
                            click.echo(f"  [{label}] {safe_subject}...")
                        else:
                            click.echo(f"  [WARN] Failed to move: {safe_subject}")
                    else:
                        click.echo("  [ERR] Could not create archive folder")

            # ── 8. Save JSONs + update superseded conversations ──
            if upload_sharepoint:
                if config.local_folder_path:
                    stats = OutputFormatter.save_to_local_folder(emails, folder_path=config.local_folder_path)
                    click.echo(f"\n[OK] Saved {stats['successful']}/{stats['total']} JSONs")
                    # Mark older emails in conversations as superseded (JSON + Outlook)
                    try:
                        needs_archive = update_superseded_jsons(conv_results, config.local_folder_path)
                        # Archive old emails that were in active inbox states
                        if needs_archive and not dry_run:
                            click.echo(f"  Archiving {len(needs_archive)} superseded emails from prior runs...")
                            sup_folder_id = reader.graph_client.get_or_create_folder(
                                mailbox, "ARCHIVE/SUPERSEDED"
                            )
                            if sup_folder_id:
                                for old_email in needs_archive:
                                    try:
                                        # Add CHAIN category before archiving
                                        reader.graph_client.update_message(mailbox, old_email["message_id"], {
                                            "categories": ["CHAIN"],
                                        })
                                        reader.graph_client.flag_message(mailbox, old_email["message_id"], "complete")
                                        reader.graph_client.move_message(mailbox, old_email["message_id"], sup_folder_id)
                                        safe_subj = old_email["subject"][:50].encode('utf-8', errors='replace').decode('utf-8', errors='ignore')
                                        click.echo(f"    [SUPERSEDED] {safe_subj}")
                                    except Exception as e:
                                        logger.warning(f"Could not archive superseded email: {e}")
                    except NameError:
                        pass  # conv_results not set (no local_folder_path during matching)

            # ── 9. Update watermark ──
            if emails and not dry_run:
                latest = max(e.received_datetime for e in emails)
                update_watermark(latest)
                click.echo(f"Watermark updated to: {latest}")
            elif dry_run:
                click.echo("(dry run — watermark not updated)")

        # ── 10. Scan inbox for human-completed flags → archive ──
        click.echo("\nScanning inbox for human-completed emails...")
        completed = reader.graph_client.get_inbox_messages_by_flag(mailbox, "complete")

        if completed:
            click.echo(f"  Found {len(completed)} human-completed email(s)")
            if not dry_run:
                folder_id = reader.graph_client.get_or_create_folder(
                    mailbox, "ARCHIVE/PROCESSED BY HUMANS"
                )
                if folder_id:
                    for msg in completed:
                        msg_id = msg["id"]
                        subj = msg.get("subject", "")[:50]
                        safe_subj = subj.encode('utf-8', errors='replace').decode('utf-8', errors='ignore')

                        # Update existing JSON if it exists
                        if config.local_folder_path:
                            ts = msg.get("receivedDateTime", "").replace(":", "-").replace("T", "_").split(".")[0]
                            email_hash = hashlib.md5(msg_id.encode()).hexdigest()[:12]
                            json_file = Path(config.local_folder_path) / f"{ts}_{email_hash}.json"
                            if json_file.exists():
                                try:
                                    with open(json_file, "r", encoding="utf-8") as f:
                                        data = json.load(f)
                                    data["processing_status"] = "ARCHIVE > PROCESSED BY HUMANS"
                                    with open(json_file, "w", encoding="utf-8") as f:
                                        json.dump(data, f, indent=2, ensure_ascii=False)
                                except Exception as e:
                                    logger.warning(f"Could not update JSON for {subj}: {e}")

                        # Move to archive
                        moved = reader.graph_client.move_message(mailbox, msg_id, folder_id)
                        if moved:
                            click.echo(f"  [HUMAN-ARCHIVED] {safe_subj}...")
                        else:
                            click.echo(f"  [WARN] Failed to move: {safe_subj}")
                else:
                    click.echo("  [ERR] Could not create archive folder")
        else:
            click.echo("  No human-completed emails found.")

        # ── 11. Scan WRONG_CLASSIFICATION folder for corrections ──
        click.echo("\nScanning for classification corrections...")
        try:
            corrections_found = scan_corrections(
                graph=reader.graph_client,
                mailbox=mailbox,
                local_folder=config.local_folder_path,
                dry_run=dry_run,
            )
            if corrections_found:
                click.echo(f"  Logged {corrections_found} new correction(s) to config/corrections.yaml")
            else:
                click.echo("  No new corrections found.")
        except Exception as e:
            logger.warning(f"Correction scan failed (non-fatal): {e}")
            click.echo(f"  [WARN] Correction scan failed: {e}")

        # ── 11b. Warn about pending corrections ──
        pending_corrections = get_pending_corrections_count()
        if pending_corrections > 0:
            click.echo(
                f"\n*** {pending_corrections} correction(s) queued in config/corrections.yaml ***"
                f"\n    Tell Claude Code: \"update from corrections\""
            )

        # API usage summary
        calls_this_run = api_counter.get_today_total() - used_today
        click.echo(f"\nAPI calls this run: {calls_this_run}")
        click.echo(f"API calls remaining today: {api_counter.get_remaining()}")
        breakdown = api_counter.get_today_breakdown()
        if breakdown:
            for model, count in sorted(breakdown.items()):
                click.echo(f"  {model}: {count}")

        click.echo(f"\n{'=' * 60}")
        click.echo("PROCESSING COMPLETE")
        click.echo("=" * 60)

        # ── Record daily stats ──
        ven_rem_count = sum(1 for e in emails if e.classification.get("primary_category", {}).get("id") == "VEN-REM")
        ven_followup_count = sum(1 for e in emails if e.classification.get("primary_category", {}).get("id") == "VEN-FOLLOWUP")
        ven_inv_count = sum(1 for e in emails if e.classification.get("primary_category", {}).get("id") == "VEN-INV")
        archived_count = sum(1 for e in emails if e.processing_status == "ARCHIVE > PROCESSED BY AGENT")
        human_completed_count = len(completed) if 'completed' in locals() and completed is not None else 0
        jsons_saved_count = len(emails) if upload_sharepoint and config.local_folder_path else 0

        # Count categorization methods
        keywords_count = sum(1 for e in emails if e.classification.get("confidence_level") == "HIGH")
        llm_count = sum(1 for e in emails if e.classification.get("confidence_level") in ("LOW", "MEDIUM") and e.classification.get("model_used"))

        # Get LLM breakdown: delta between before/after snapshots (captures ALL calls: classification, pass2, translation)
        api_snapshot_after = api_counter.get_today_breakdown()
        llm_breakdown = {}
        for model, count in api_snapshot_after.items():
            before = api_snapshot_before.get(model, 0)
            if count - before > 0:
                llm_breakdown[model] = count - before

        DailyStats.record_process_run(
            emails_processed=len(emails),
            categories_by_keywords=keywords_count,
            categories_by_llm=llm_count,
            ven_rem_analyzed=ven_rem_count,
            ven_followup_analyzed=ven_followup_count,
            ven_inv_processed=ven_inv_count,
            emails_archived=archived_count,
            human_completed_moved=human_completed_count,
            jsons_saved=jsons_saved_count,
            llm_calls_by_model=llm_breakdown,
        )

        # ── Send daily summary at 10:00 ──
        current_hour = datetime.now().hour
        if current_hour == 10:
            click.echo("\nSending daily summary email...")
            send_daily_summary()

    except Exception as e:
        error_msg = str(e)
        error_details = traceback.format_exc()
        logger.error(f"Process error: {error_msg}", exc_info=True)
        click.echo(f"[ERR] {error_msg}", err=True)

        # Send error notification email
        send_error_notification(error_msg, error_details)

        sys.exit(1)


@cli.command()
@click.option("--folder", default="Reminders", help="Outlook folder name/path to process")
@click.option("--budget", default=0, type=int, help="Max API calls (0 = use remaining daily budget)")
@click.option("--days", default=60, type=int, help="Only process emails from last N days")
@click.option("--mailbox", default=config.accounting_mailbox, help="Target mailbox")
@click.option("--dry-run", is_flag=True, help="Show what would happen without changes")
@click.option("--upload-sharepoint", is_flag=True, help="Save email JSONs locally")
def cleanup(folder, budget, days, mailbox, dry_run, upload_sharepoint):
    """Batch-process emails from an Outlook folder (e.g. Reminders cleanup).

    Designed to run after the normal process command, using remaining daily API budget.
    Each VEN-REM email costs 1 API call (Pass 2). No translation to save budget.

    Examples:
        python main.py cleanup                              # Use remaining budget
        python main.py cleanup --budget 19                  # Use exactly 19 calls
        python main.py cleanup --folder "Reminders"         # Specify folder
        python main.py cleanup --dry-run                    # Preview only
    """
    from datetime import datetime, timezone
    from src import api_counter

    try:
        click.echo("=" * 60)
        click.echo(f"CLEANUP: Processing folder '{folder}'")
        click.echo(f"Time: {datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M UTC')}")
        click.echo("=" * 60)

        # ── Budget calculation ──
        used_before = api_counter.get_today_total()
        breakdown = api_counter.get_today_breakdown()
        if budget <= 0:
            budget = api_counter.get_remaining()

        click.echo(f"\nAPI calls today: {used_before}")
        if breakdown:
            for model, count in sorted(breakdown.items()):
                click.echo(f"  {model}: {count}")
        click.echo(f"Budget for cleanup: {budget}")

        if budget <= 0:
            click.echo("\nNo API budget remaining today. Exiting.")
            return

        # ── Initialize ──
        reader = EmailReader()
        graph = reader.graph_client

        # ── Find the folder ──
        click.echo(f"\nLooking up folder: {folder}...")
        folder_id = graph.get_folder_id(mailbox, folder)
        if not folder_id:
            click.echo(f"[ERR] Folder '{folder}' not found in {mailbox}.")
            sys.exit(1)
        click.echo(f"  [OK] Found folder")

        # ── Fetch messages ──
        click.echo(f"Fetching emails (last {days} days)...")
        messages = graph.get_folder_messages(
            mailbox, folder_id, max_results=1000, days_back=days
        )

        if not messages:
            click.echo("No emails found in folder.")
            return

        click.echo(f"[OK] Found {len(messages)} emails\n")

        # ── Parse into Email objects (body + attachments) ──
        click.echo("Loading email bodies and attachments...")
        emails = []
        for idx, msg in enumerate(messages, 1):
            try:
                email_obj = reader._parse_message(msg)
                body = graph.get_message_body(mailbox, email_obj.id)
                if body:
                    email_obj.body = reader._extract_text_from_html(body)
                if email_obj.has_attachments:
                    attachments = graph.get_message_attachments(
                        mailbox, email_obj.id
                    )
                    if attachments:
                        email_obj.attachments = reader._process_attachments(
                            attachments, email_obj.id
                        )
                emails.append(email_obj)
            except Exception as e:
                logger.error(f"Error parsing message {idx}: {e}")

            if idx % 50 == 0:
                click.echo(f"  Loaded {idx}/{len(messages)} emails...")

        click.echo(f"[OK] Loaded {len(emails)} emails\n")

        # ── Pass 0: Keyword triage (free) ──
        click.echo("Pass 0: Keyword triage...")
        triage = KeywordTriage()
        ven_rem_emails = []
        for email_obj in emails:
            keyword_result = triage.classify(email_obj.to_dict())
            email_obj.classification = keyword_result
            if keyword_result["primary_category"]["id"] == "VEN-REM":
                ven_rem_emails.append(email_obj)

        cats = {}
        for e in emails:
            cat = e.classification["primary_category"]["id"]
            cats[cat] = cats.get(cat, 0) + 1
        for cat, count in sorted(cats.items(), key=lambda x: -x[1]):
            click.echo(f"  {cat}: {count}")

        # ── Conversation matching ──
        conv_results = {}
        if config.local_folder_path:
            click.echo("\nMatching conversations...")
            conv_results = match_conversations(emails, config.local_folder_path)
            linked = 0
            for email_obj in emails:
                info = conv_results.get(email_obj.id)
                if info:
                    email_obj.conversation_id = info["conversation_id"]
                    email_obj.conversation_position = info["position"]
                    email_obj.is_latest_in_conversation = info["is_latest"]
                    email_obj.related_emails = info["related_emails"]
                    email_obj.is_chain = info.get("is_chain", False)
                    if info["position"] > 1:
                        linked += 1
            superseded = sum(1 for e in emails if not e.is_latest_in_conversation)
            click.echo(f"  Conversations: {linked} linked, {superseded} superseded")

        click.echo(f"\n  VEN-REM for Pass 2: {len(ven_rem_emails)}")

        # ── Pass 2 on VEN-REM (respect budget) ──
        remaining = min(budget, api_counter.get_remaining())
        pass2_count = min(len(ven_rem_emails), remaining)

        if pass2_count > 0:
            click.echo(
                f"\nPass 2: processing {pass2_count}/{len(ven_rem_emails)} "
                f"VEN-REM emails (budget: {remaining})..."
            )
            try:
                processor = Pass2Processor()
                processed = 0
                for idx, email_obj in enumerate(ven_rem_emails[:pass2_count], 1):
                    if api_counter.get_remaining() <= 0:
                        click.echo(f"  Budget exhausted after {processed} emails.")
                        break

                    click.echo(
                        f"  Pass 2 [{idx}/{pass2_count}]: "
                        f"{email_obj.subject[:50]}..."
                    )
                    pass2_result = processor.process_email(email_obj.to_dict())
                    if pass2_result:
                        email_obj.pass2_results = pass2_result
                        primary_cat = email_obj.classification.get("primary_category", {}).get("id", "VEN-REM")
                        if pass2_result.get("reclassified"):
                            new_cat = pass2_result["reclassified_to"]
                            click.echo(f"    [RECLASSIFIED] {primary_cat} -> {new_cat}")
                            email_obj.classification["primary_category"] = {
                                "id": new_cat,
                                "name": KeywordTriage.CATEGORY_NAMES.get(
                                    new_cat, new_cat
                                ),
                            }
                            email_obj.classification["classification_method"] = "llm_reclassified"
                            if new_cat == "VEN-FOLLOWUP":
                                email_obj.processing_status = "NEEDS REVIEW > VENDOR QUERY"
                        else:
                            verified = pass2_result.get("verified_category", primary_cat)
                            if verified == "VEN-REM":
                                email_obj.processing_status = "ARCHIVE > PROCESSED BY AGENT"
                            elif verified == "VEN-FOLLOWUP":
                                email_obj.processing_status = "NEEDS REVIEW > VENDOR QUERY"
                        processed += 1

                processor.close()
                click.echo(f"  [OK] Pass 2 complete for {processed} emails")
            except Exception as e:
                click.echo(f"  [ERR] Pass 2: {e}", err=True)
                logger.error(f"Pass 2 error: {e}", exc_info=True)
        else:
            click.echo("\nNo budget for Pass 2 or no VEN-REM emails.")

        # ── Outlook write-back ──
        if not dry_run:
            ext_prop_id = (
                "String {00020329-0000-0000-C000-000000000046} "
                "Name EmailClassifierData"
            )

            click.echo("\nOutlook: writing categories...")
            write_count = 0
            for email_obj in emails:
                cat_id = email_obj.classification.get(
                    "primary_category", {}
                ).get("id", "")
                prio = email_obj.classification.get("priority", "")
                categories = [c for c in [cat_id, prio] if c]

                ext_data = dict(email_obj.classification)
                if email_obj.pass2_results:
                    ext_data["pass2_results"] = email_obj.pass2_results

                updates = {
                    "categories": categories,
                    "isRead": True,
                    "singleValueExtendedProperties": [{
                        "id": ext_prop_id,
                        "value": json.dumps(
                            ext_data, ensure_ascii=False, default=str
                        ),
                    }],
                }
                result = graph.update_message(mailbox, email_obj.id, updates)
                if result:
                    write_count += 1

            click.echo(
                f"  [OK] Categories written for {write_count}/{len(emails)} emails"
            )

            # Move agent-processed to PROCESSED BY AGENT
            agent_processed = [
                e for e in emails
                if e.processing_status == "ARCHIVE > PROCESSED BY AGENT"
            ]
            if agent_processed:
                click.echo(
                    f"\nMoving {len(agent_processed)} processed emails "
                    f"to ARCHIVE/PROCESSED BY AGENT..."
                )
                archive_id = graph.get_or_create_folder(
                    mailbox, "ARCHIVE/PROCESSED BY AGENT"
                )
                if archive_id:
                    moved_count = 0
                    for email_obj in agent_processed:
                        graph.flag_message(mailbox, email_obj.id, "complete")
                        moved = graph.move_message(
                            mailbox, email_obj.id, archive_id
                        )
                        if moved:
                            moved_count += 1
                            click.echo(f"  [MOVED] {email_obj.subject[:50]}...")
                        else:
                            click.echo(
                                f"  [WARN] Failed to move: {email_obj.subject[:40]}"
                            )
                    click.echo(f"  [OK] Moved {moved_count} emails")
                else:
                    click.echo("  [ERR] Could not create archive folder")
        else:
            click.echo("\n(dry run — no Outlook changes)")

        # ── Save JSONs ──
        if upload_sharepoint and config.local_folder_path:
            click.echo(f"\nSaving JSONs to {config.local_folder_path}...")
            stats = OutputFormatter.save_to_local_folder(
                emails, folder_path=config.local_folder_path
            )
            click.echo(
                f"[OK] Saved {stats['successful']}/{stats['total']} JSONs"
            )
            # Mark older emails in conversations as superseded
            if conv_results:
                needs_archive = update_superseded_jsons(conv_results, config.local_folder_path)
                if needs_archive:
                    logger.info(f"{len(needs_archive)} superseded emails need Outlook archiving (skipped in cleanup)")

        # ── Summary ──
        calls_used = api_counter.get_today_total() - used_before
        agent_count = sum(
            1 for e in emails
            if e.processing_status == "ARCHIVE > PROCESSED BY AGENT"
        )
        pass2_done = sum(1 for e in emails if e.pass2_results)

        click.echo(f"\n{'=' * 60}")
        click.echo("CLEANUP COMPLETE")
        click.echo(f"  Emails in folder:            {len(messages)}")
        click.echo(f"  Keyword-classified:          {len(emails)}")
        click.echo(f"  VEN-REM with Pass 2:         {pass2_done}")
        click.echo(f"  Moved to PROCESSED BY AGENT: {agent_count}")
        click.echo(f"  API calls used this run:     {calls_used}")
        click.echo(f"  API calls remaining today:   {api_counter.get_remaining()}")
        click.echo(
            f"  VEN-REM still pending:       "
            f"{len(ven_rem_emails) - pass2_done}"
        )
        click.echo("=" * 60)

        # ── Record cleanup stats and send daily summary ──
        DailyStats.record_cleanup_run(
            emails_in_folder=len(messages),
            emails_classified=len(emails),
            emails_with_pass2=pass2_done,
            emails_moved_to_archive=agent_count,
            api_calls_used=calls_used,
        )

        # Send daily summary email
        click.echo("\nSending daily summary email...")
        send_daily_summary()

    except Exception as e:
        error_msg = str(e)
        error_details = traceback.format_exc()
        logger.error(f"Cleanup error: {error_msg}", exc_info=True)
        click.echo(f"[ERR] {error_msg}", err=True)

        # Send error notification email
        send_error_notification(error_msg, error_details)

        sys.exit(1)


@cli.command()
def build_conversation_index():
    """Build/rebuild the conversation index from existing email JSONs.

    Scans all JSON files in the local email folder and creates the
    conversation_index.json used for matching email chains.
    """
    from src.conversation_matcher import build_index_from_existing

    if not config.local_folder_path:
        click.echo("[ERR] LOCAL_FOLDER_PATH not set in .env")
        sys.exit(1)

    click.echo(f"Building conversation index from: {config.local_folder_path}")
    count = build_index_from_existing(config.local_folder_path)
    click.echo(f"[OK] Indexed {count} emails (last 30 days)")


if __name__ == "__main__":
    cli()
