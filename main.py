"""Main CLI entry point for the accounting mailbox reader."""

import json
import logging
import sys
from pathlib import Path

import click

from src.email_reader import EmailReader
from src.output_formatter import OutputFormatter
from src.config import config
from src.email_classifier import EmailClassifier
from src.confluence_sync import ConfluenceSyncer
from src.pass2_processor import Pass2Processor
from src.keyword_triage import KeywordTriage
from src.email_translator import EmailTranslator

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
)

# Suppress verbose Azure SDK logging
logging.getLogger("azure.core.pipeline.policies.http_logging_policy").setLevel(logging.WARNING)
logging.getLogger("azure.identity").setLevel(logging.WARNING)
logging.getLogger("msal").setLevel(logging.WARNING)

logger = logging.getLogger(__name__)


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

    try:
        click.echo("=" * 60)
        click.echo("AUTOMATED PROCESSING RUN")
        click.echo("=" * 60)

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

            # ── 5. Pass 2 on VEN-REM emails ──
            ven_rem = [
                e for e in emails
                if e.classification.get("primary_category", {}).get("id") == "VEN-REM"
            ]

            if ven_rem:
                click.echo(f"\nPass 2: deep analysis on {len(ven_rem)} VEN-REM...")
                try:
                    processor = Pass2Processor()
                    for email in ven_rem:
                        click.echo(f"  Pass 2: {email.subject[:50]}...")
                        pass2_result = processor.process_email(email.to_dict())
                        if pass2_result:
                            email.pass2_results = pass2_result
                            if pass2_result.get("reclassified"):
                                new_cat = pass2_result["reclassified_to"]
                                click.echo(f"    [RECLASSIFIED] VEN-REM -> {new_cat}")
                                email.classification["primary_category"] = {
                                    "id": new_cat,
                                    "name": KeywordTriage.CATEGORY_NAMES.get(new_cat, new_cat),
                                }
                                email.classification["classification_method"] = "llm_reclassified"
                            else:
                                # Pass 2 confirmed VEN-REM → mark as agent-processed
                                email.processing_status = "ARCHIVE > PROCESSED BY AGENT"
                    processor.close()
                    click.echo(f"  [OK] Pass 2 complete")
                except Exception as e:
                    click.echo(f"  [ERR] Pass 2: {e}", err=True)

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

            # ── 7. Move agent-processed to archive folder ──
            agent_processed = [e for e in emails if e.processing_status == "ARCHIVE > PROCESSED BY AGENT"]
            if agent_processed and not dry_run:
                click.echo(f"\nArchiving {len(agent_processed)} agent-processed emails...")
                folder_id = reader.graph_client.get_or_create_folder(
                    mailbox, "ARCHIVE/PROCESSED BY AGENT"
                )
                if folder_id:
                    for email in agent_processed:
                        # Complete the flag
                        reader.graph_client.flag_message(mailbox, email.id, "complete")
                        # Update extended property with final status
                        ext_data = dict(email.classification)
                        ext_data["processing_status"] = email.processing_status
                        if email.pass2_results:
                            ext_data["pass2_results"] = email.pass2_results
                        reader.graph_client.update_message(mailbox, email.id, {
                            "singleValueExtendedProperties": [{
                                "id": ext_prop_id,
                                "value": json.dumps(ext_data, ensure_ascii=False, default=str),
                            }],
                        })
                        # Move to archive folder
                        moved = reader.graph_client.move_message(mailbox, email.id, folder_id)
                        if moved:
                            click.echo(f"  [ARCHIVED] {email.subject[:50]}...")
                        else:
                            click.echo(f"  [WARN] Failed to move: {email.subject[:40]}")
                else:
                    click.echo("  [ERR] Could not create archive folder")

            # ── 8. Save JSONs ──
            if upload_sharepoint:
                if config.local_folder_path:
                    stats = OutputFormatter.save_to_local_folder(emails, folder_path=config.local_folder_path)
                    click.echo(f"\n[OK] Saved {stats['successful']}/{stats['total']} JSONs")

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
                            click.echo(f"  [HUMAN-ARCHIVED] {subj}...")
                        else:
                            click.echo(f"  [WARN] Failed to move: {subj}")
                else:
                    click.echo("  [ERR] Could not create archive folder")
        else:
            click.echo("  No human-completed emails found.")

        click.echo(f"\n{'=' * 60}")
        click.echo("PROCESSING COMPLETE")
        click.echo("=" * 60)

    except Exception as e:
        logger.error(f"Process error: {e}", exc_info=True)
        click.echo(f"[ERR] {e}", err=True)
        sys.exit(1)


if __name__ == "__main__":
    cli()
