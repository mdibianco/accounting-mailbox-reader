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

    // ── Filter to only emails with pass2_results containing invoices ──
    AddPass2 = Table.AddColumn(RemoveNulls, "pass2", each
        try Record.Field([JsonContent], "pass2_results") otherwise null
    ),
    HasPass2 = Table.SelectRows(AddPass2, each [pass2] <> null),

    AddInvoiceList = Table.AddColumn(HasPass2, "invoices_list", each
        try Record.Field([pass2], "invoices") otherwise null
    ),
    HasInvoices = Table.SelectRows(AddInvoiceList, each [invoices_list] <> null),

    // ── Add email context columns for joining ──
    AddEmailContext = Table.AddColumn(HasInvoices, "EmailData", each
        let
            j = [JsonContent],
            frm = try Record.Field(j, "from") otherwise [],
            cls = try Record.Field(j, "classification") otherwise [],
            p2  = [pass2],
            pe  = try Record.Field(p2, "planted_entity") otherwise []
        in [
            email_id            = try j[id] otherwise null,
            outlook_link        = try j[outlook_link] otherwise null,
            subject             = try j[subject] otherwise null,
            received_datetime   = try j[received_datetime] otherwise null,
            from_email          = try frm[email] otherwise null,
            from_name           = try frm[name] otherwise null,
            category_id         = try Record.Field(try Record.Field(cls, "primary_category") otherwise [], "id") otherwise null,
            priority            = try cls[priority] otherwise null,
            urgency_level       = try p2[urgency_level] otherwise null,
            planted_entity_code = try pe[code] otherwise null,
            planted_entity_name = try pe[name] otherwise null
        ]
    ),

    // ── Expand invoices list into rows (one row per invoice) ──
    ExpandedList = Table.ExpandListColumn(AddEmailContext, "invoices_list"),

    // ── Expand each invoice record ──
    AddInvoiceFields = Table.AddColumn(ExpandedList, "InvoiceData", each
        let
            inv = [invoices_list],
            bc  = try Record.Field(inv, "bc_lookup") otherwise []
        in [
            vendor_name     = try inv[vendor_name] otherwise null,
            invoice_number  = try inv[invoice_number] otherwise null,
            invoice_date    = try Date.FromText(inv[invoice_date]) otherwise null,
            due_date        = try Date.FromText(inv[due_date]) otherwise null,
            amount          = try inv[amount] otherwise null,
            currency        = try inv[currency] otherwise null,
            bc_status       = try bc[status] otherwise null,
            bc_found        = try bc[found] otherwise null,
            bc_error        = try bc[error] otherwise null
        ]
    ),

    // ── Expand email context ──
    ExpandEmail = Table.ExpandRecordColumn(AddInvoiceFields, "EmailData", {
        "email_id", "outlook_link", "subject", "received_datetime",
        "from_email", "from_name", "category_id", "priority",
        "urgency_level", "planted_entity_code", "planted_entity_name"
    }),

    // ── Expand invoice fields ──
    ExpandInvoice = Table.ExpandRecordColumn(ExpandEmail, "InvoiceData", {
        "vendor_name", "invoice_number", "invoice_date", "due_date",
        "amount", "currency", "bc_status", "bc_found", "bc_error"
    }),

    // ── Clean up: keep only relevant columns ──
    Cleaned = Table.SelectColumns(ExpandInvoice, {
        // Email context (for joining / filtering)
        "email_id", "outlook_link", "subject", "received_datetime",
        "from_email", "from_name", "category_id", "priority",
        "urgency_level", "planted_entity_code", "planted_entity_name",
        // Invoice fields
        "vendor_name", "invoice_number", "invoice_date", "due_date",
        "amount", "currency", "bc_status", "bc_found", "bc_error"
    }),

    // ── Set types ──
    Typed = Table.TransformColumnTypes(Cleaned, {
        {"received_datetime", type text},
        {"invoice_date", type date},
        {"due_date", type date},
        {"amount", type number},
        {"urgency_level", Int64.Type},
        {"bc_found", type logical}
    })
in
    Typed
