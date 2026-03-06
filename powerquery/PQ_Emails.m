let
    // ── Source: all JSON files in the local email folder ──
    FolderPath = "C:\Users\MatthiasDiBianco\planted foods AG\Circle Finance - Documents\Accounting\Mailbox\emails",
    Source = Folder.Files(FolderPath),
    FilterJSON = Table.SelectRows(Source, each Text.EndsWith([Name], ".json")),

    // ── Parse each JSON file ──
    AddContent = Table.AddColumn(FilterJSON, "JsonContent", each
        try Json.Document([Content]) otherwise null
    ),
    RemoveNulls = Table.SelectRows(AddContent, each [JsonContent] <> null),

    // ── Drop heavy Folder.Files columns (Content binary etc.) and buffer ──
    Trimmed = Table.SelectColumns(RemoveNulls, {"JsonContent"}),
    Buffered = Table.Buffer(Trimmed),

    // ── Extract flat email fields ──
    AddFields = Table.AddColumn(Buffered, "Parsed", each
        let
            j = [JsonContent],
            cls = try Record.Field(j, "classification") otherwise [],
            cat = try Record.Field(cls, "primary_category") otherwise [],
            ent = try Record.Field(cls, "extracted_entities") otherwise [],
            p2  = try Record.Field(j, "pass2_results") otherwise [],
            pe  = try Record.Field(p2, "planted_entity") otherwise [],
            frm = try Record.Field(j, "from") otherwise []
        in [
            email_id            = try j[id] otherwise null,
            outlook_link        = try j[outlook_link] otherwise null,
            processing_status   = try j[processing_status] otherwise null,
            from_email          = try frm[email] otherwise null,
            from_name           = try frm[name] otherwise null,
            subject             = try j[subject] otherwise null,
            received_datetime   = try DateTimeZone.FromText(j[received_datetime]) otherwise null,
            body_preview        = try j[body_preview] otherwise null,
            has_attachments     = try j[has_attachments] otherwise null,
            is_read             = try j[is_read] otherwise null,
            importance          = try j[importance] otherwise null,

            // Classification
            category_id         = try cat[id] otherwise null,
            category_name       = try cat[name] otherwise null,
            confidence_level    = try cls[confidence_level] otherwise null,
            priority            = try cls[priority] otherwise null,
            classification_method = try cls[classification_method] otherwise null,
            keyword_confidence  = try cls[keyword_confidence] otherwise null,
            requires_manual_review = try cls[requires_manual_review] otherwise null,
            summary             = try cls[summary] otherwise null,
            reasoning           = try cls[reasoning] otherwise null,

            // Extracted entities (from LLM classification)
            entity_vendor       = try ent[vendor] otherwise null,
            entity_customer     = try ent[customer] otherwise null,
            entity_invoice_nr   = try ent[invoice_number] otherwise null,
            entity_amount       = try ent[amount] otherwise null,
            entity_currency     = try ent[currency] otherwise null,
            entity_due_date     = try ent[due_date] otherwise null,

            // Pass 2
            has_pass2           = try (p2[pass2_timestamp] <> null) otherwise false,
            urgency_level       = try p2[urgency_level] otherwise null,
            urgency_reasoning   = try p2[urgency_reasoning] otherwise null,
            planted_entity_code = try pe[code] otherwise null,
            planted_entity_name = try pe[name] otherwise null,
            verified_category   = try p2[verified_category] otherwise null,
            invoice_count       = try List.Count(p2[invoices]) otherwise 0,

            // Pass 2 VEN-INV specific
            inv_action_taken    = try p2[action_taken] otherwise null,
            inv_forwarded_to    = try p2[forwarded_to] otherwise null,
            inv_invoices_address = try p2[invoices_address] otherwise null,
            inv_draft_reply_created = try p2[draft_reply_created] otherwise null,

            // Translation
            body_english        = try j[body_english] otherwise null,

            // Conversation threading
            is_latest_in_conversation = try j[is_latest_in_conversation] otherwise true,
            is_chain   = try j[is_chain] otherwise false,
            conversation_id           = try j[conversation_id] otherwise null,
            conversation_position     = try j[conversation_position] otherwise null,

            // Jira
            jira_issue_key      = try j[jira_issue_key] otherwise null,

            // Metadata
            json_filename       = try j[id] otherwise null
        ]
    ),

    // ── Expand parsed record into columns ──
    Expanded = Table.ExpandRecordColumn(AddFields, "Parsed", {
        "email_id", "outlook_link", "processing_status",
        "from_email", "from_name", "subject", "received_datetime",
        "body_preview", "has_attachments", "is_read", "importance",
        "category_id", "category_name", "confidence_level", "priority",
        "classification_method", "keyword_confidence", "requires_manual_review",
        "summary", "reasoning",
        "entity_vendor", "entity_customer", "entity_invoice_nr",
        "entity_amount", "entity_currency", "entity_due_date",
        "has_pass2", "urgency_level", "urgency_reasoning",
        "planted_entity_code", "planted_entity_name", "verified_category",
        "invoice_count",
        "inv_action_taken", "inv_forwarded_to", "inv_invoices_address", "inv_draft_reply_created",
        "body_english",
        "is_latest_in_conversation", "is_chain", "conversation_id", "conversation_position",
        "jira_issue_key",
        "json_filename"
    }),

    // ── Clean up: drop helper columns, keep only parsed fields ──
    Cleaned = Table.SelectColumns(Expanded, {
        "email_id", "outlook_link", "processing_status",
        "from_email", "from_name", "subject", "received_datetime",
        "body_preview", "has_attachments", "is_read", "importance",
        "category_id", "category_name", "confidence_level", "priority",
        "classification_method", "keyword_confidence", "requires_manual_review",
        "summary", "reasoning",
        "entity_vendor", "entity_customer", "entity_invoice_nr",
        "entity_amount", "entity_currency", "entity_due_date",
        "has_pass2", "urgency_level", "urgency_reasoning",
        "planted_entity_code", "planted_entity_name", "verified_category",
        "invoice_count",
        "inv_action_taken", "inv_forwarded_to", "inv_invoices_address", "inv_draft_reply_created",
        "body_english",
        "is_latest_in_conversation", "is_chain", "conversation_id", "conversation_position",
        "jira_issue_key"
    }),

    // ── Prefix priority for sort order (0 = low … 3 = highest) ──
    PrioMap = [PRIO_LOW = "0_PRIO_LOW", PRIO_MEDIUM = "1_PRIO_MEDIUM",
               PRIO_HIGH = "2_PRIO_HIGH", PRIO_HIGHEST = "3_PRIO_HIGHEST"],
    WithPrio = Table.TransformColumns(Cleaned, {
        {"priority", each try Record.Field(PrioMap, _) otherwise _, type text}
    }),

    // ── Set types ──
    Typed = Table.TransformColumnTypes(WithPrio, {
        {"received_datetime", type datetimezone},
        {"has_attachments", type logical},
        {"is_read", type logical},
        {"requires_manual_review", type logical},
        {"has_pass2", type logical},
        {"keyword_confidence", type number},
        {"entity_amount", type number},
        {"urgency_level", Int64.Type},
        {"invoice_count", Int64.Type},
        {"is_latest_in_conversation", type logical},
        {"is_chain", type logical},
        {"conversation_position", Int64.Type}
    })
in
    Typed
