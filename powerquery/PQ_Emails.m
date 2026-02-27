let
    // ── Source: all JSON files in the local email folder ──
    FolderPath = "C:\Users\MatthiasDiBianco\emails\emails",
    Source = Folder.Files(FolderPath),
    FilterJSON = Table.SelectRows(Source, each Text.EndsWith([Name], ".json")),

    // ── Parse each JSON file ──
    AddContent = Table.AddColumn(FilterJSON, "JsonContent", each
        try Json.Document([Content]) otherwise null
    ),
    RemoveNulls = Table.SelectRows(AddContent, each [JsonContent] <> null),

    // ── Extract flat email fields ──
    AddFields = Table.AddColumn(RemoveNulls, "Parsed", each
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

            // Translation
            body_english        = try j[body_english] otherwise null,

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
        "invoice_count", "body_english", "json_filename"
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
        "invoice_count", "body_english"
    }),

    // ── Set types ──
    Typed = Table.TransformColumnTypes(Cleaned, {
        {"received_datetime", type datetimezone},
        {"has_attachments", type logical},
        {"is_read", type logical},
        {"requires_manual_review", type logical},
        {"has_pass2", type logical},
        {"keyword_confidence", type number},
        {"entity_amount", type number},
        {"urgency_level", Int64.Type},
        {"invoice_count", Int64.Type}
    })
in
    Typed
