let
    // ══════════════════════════════════════════════════════════════
    // CONFIG
    // ══════════════════════════════════════════════════════════════
    FolderPath = "C:\Users\MatthiasDiBianco\planted foods AG\Circle Finance - Documents\Accounting\Mailbox\emails",
    FuzzyThreshold = 0.80,

    // Levenshtein similarity: 1 = identical, 0 = completely different
    LevenshteinSimilarity = (s1 as text, s2 as text) as number =>
        let
            len1 = Text.Length(s1),
            len2 = Text.Length(s2),
            maxLen = List.Max({len1, len2}),
            sim = if maxLen = 0 then 1.0
                // Quick reject: if lengths differ by more than 20%, can't reach 80%
                else if Number.Abs(len1 - len2) / maxLen > 0.2 then 0
                else let
                    initRow = List.Numbers(0, len2 + 1),
                    finalRow = List.Accumulate(
                        {0..len1-1}, initRow,
                        (prev, i) =>
                            let c1 = Text.At(s1, i) in
                            List.Accumulate(
                                {0..len2-1},
                                {i + 1},
                                (row, j) =>
                                    let
                                        cost = if c1 = Text.At(s2, j) then 0 else 1,
                                        val = List.Min({
                                            row{j} + 1,
                                            prev{j + 1} + 1,
                                            prev{j} + cost
                                        })
                                    in row & {val}
                            )
                    )
                in 1 - finalRow{len2} / maxLen
        in sim,

    // ══════════════════════════════════════════════════════════════
    // PART 1: Load email invoices from local JSON files
    // ══════════════════════════════════════════════════════════════
    Source = Folder.Files(FolderPath),
    FilterJSON = Table.SelectRows(Source, each Text.EndsWith([Name], ".json")),

    AddContent = Table.AddColumn(FilterJSON, "JsonContent", each
        try Json.Document([Content]) otherwise null
    ),
    RemoveNulls = Table.SelectRows(AddContent, each [JsonContent] <> null),

    AddPass2 = Table.AddColumn(RemoveNulls, "pass2", each
        try Record.Field([JsonContent], "pass2_results") otherwise null
    ),
    HasPass2 = Table.SelectRows(AddPass2, each [pass2] <> null),

    AddInvoiceList = Table.AddColumn(HasPass2, "invoices_list", each
        try Record.Field([pass2], "invoices") otherwise null
    ),
    HasInvoices = Table.SelectRows(AddInvoiceList, each [invoices_list] <> null),

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

    ExpandedList = Table.ExpandListColumn(AddEmailContext, "invoices_list"),

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

    ExpandEmail = Table.ExpandRecordColumn(AddInvoiceFields, "EmailData", {
        "email_id", "outlook_link", "subject", "received_datetime",
        "from_email", "from_name", "category_id", "priority",
        "urgency_level", "planted_entity_code", "planted_entity_name"
    }),

    ExpandInvoice = Table.ExpandRecordColumn(ExpandEmail, "InvoiceData", {
        "vendor_name", "invoice_number", "invoice_date", "due_date",
        "amount", "currency", "bc_status", "bc_found", "bc_error"
    }),

    Invoices = Table.SelectColumns(ExpandInvoice, {
        "email_id", "outlook_link", "subject", "received_datetime",
        "from_email", "from_name", "category_id", "priority",
        "urgency_level", "planted_entity_code", "planted_entity_name",
        "vendor_name", "invoice_number", "invoice_date", "due_date",
        "amount", "currency", "bc_status", "bc_found", "bc_error"
    }),

    // ══════════════════════════════════════════════════════════════
    // PART 2: Reference BC tables from the data model
    // ══════════════════════════════════════════════════════════════
    Vendors = vendor,
    BC_Ledger = vendor_ledger_entry,

    // ══════════════════════════════════════════════════════════════
    // PART 3: Two-stage matching
    // PERF NOTES:
    //   Stage 1 (Vendor name):
    //     • Fast: Match vendor name (trimmed first 30 chars) against small vendor table (~3000 entries)
    //     • Returns vendor_no for ledger filtering
    //   Stage 2 (Invoice number on filtered ledger):
    //     • Step 1 (exact doc): Fast. Exact external_document_no match.
    //     • Step 2 (trimmed doc): Fast. Trimmed UPPERCASE match.
    //     • Step 3 (fuzzy): Only runs if vendor matched + exact/trimmed failed.
    //       Filters ledger to vendor first, then Levenshtein (small dataset, fast).
    //     • No amount/date match fallback.
    // ══════════════════════════════════════════════════════════════
    AddMatch = Table.AddColumn(Invoices, "MatchResult", each
        let
            inv_num_raw  = Text.From([invoice_number] ?? ""),
            inv_num_trim = Text.Upper(Text.Trim(inv_num_raw)),
            inv_vendor   = Text.Upper(Text.Trim(Text.From([vendor_name] ?? ""))),
            inv_vendor_prefix = Text.Start(inv_vendor, 30),

            // ── STAGE 1: Vendor name match (against vendor table) ──
            VendorMatch = try Table.SelectRows(Vendors, each
                let
                    bc_vendor = Text.Upper(Text.Trim(Text.From([vendor_name] ?? ""))),
                    bc_vendor_prefix = Text.Start(bc_vendor, 30)
                in
                    bc_vendor_prefix = inv_vendor_prefix and inv_vendor_prefix <> ""
            ) otherwise null,
            MatchedVendorNo = try
                if VendorMatch <> null and Table.RowCount(VendorMatch) > 0
                    then Table.FirstValue(Table.SelectColumns(VendorMatch, {"vendor_no"}))
                    else null
                otherwise null,

            // ── STAGE 2: Invoice number match (on filtered ledger or full ledger) ──
            // Step 1: Exact match
            ExactRows = try Table.SelectRows(BC_Ledger, each
                Text.From([external_document_no] ?? "") = inv_num_raw
            ) otherwise null,

            // Step 2: Trimmed UPPERCASE match
            TrimmedRows = try
                if ExactRows <> null and Table.RowCount(ExactRows) > 0 then null
                else Table.SelectRows(BC_Ledger, each
                    Text.Upper(Text.Trim(Text.From([external_document_no] ?? ""))) = inv_num_trim
                    and inv_num_trim <> ""
                )
                otherwise null,

            // Step 3: Fuzzy match (only if vendor matched; filter ledger to vendor_no)
            FuzzyRows = try
                if ExactRows <> null and Table.RowCount(ExactRows) > 0 then null
                else if TrimmedRows <> null and Table.RowCount(TrimmedRows) > 0 then null
                else if MatchedVendorNo = null then null
                else
                    let
                        VendorLedger = Table.SelectRows(BC_Ledger, each
                            [vendor_no] = MatchedVendorNo
                        )
                    in
                        Table.SelectRows(VendorLedger, each
                            let
                                bc_doc = Text.Upper(Text.Trim(Text.From([external_document_no] ?? ""))),
                                bothLongEnough = Text.Length(inv_num_trim) >= 3
                                    and Text.Length(bc_doc) >= 3,
                                containsMatch = bothLongEnough
                                    and (Text.Contains(bc_doc, inv_num_trim)
                                        or Text.Contains(inv_num_trim, bc_doc)),
                                similarMatch = if containsMatch or not bothLongEnough then false
                                    else LevenshteinSimilarity(inv_num_trim, bc_doc) >= FuzzyThreshold
                            in
                                containsMatch or similarMatch
                        )
                otherwise null,

            // ── Pick best result ──
            MatchKind = try
                if ExactRows <> null and Table.RowCount(ExactRows) > 0 then "exact"
                else if TrimmedRows <> null and Table.RowCount(TrimmedRows) > 0 then "trimmed"
                else if FuzzyRows <> null and Table.RowCount(FuzzyRows) > 0 then "fuzzy"
                else null
                otherwise null,

            MatchedTable = try
                if MatchKind = "exact" then ExactRows
                else if MatchKind = "trimmed" then TrimmedRows
                else if MatchKind = "fuzzy" then FuzzyRows
                else null
                otherwise null,

            MatchedDoc = try
                if MatchedTable <> null and Table.RowCount(MatchedTable) > 0
                then Table.FirstValue(Table.SelectColumns(MatchedTable, {"external_document_no"}))
                else null
                otherwise null
        in
            [matched_bc_external_document_no = MatchedDoc, matched_vendor_no = MatchedVendorNo, match_kind = MatchKind]
    ),

    // ══════════════════════════════════════════════════════════════
    // PART 4: Expand match result + set types
    // ══════════════════════════════════════════════════════════════
    ExpandMatch = Table.ExpandRecordColumn(AddMatch, "MatchResult", {
        "matched_bc_external_document_no", "matched_vendor_no", "match_kind"
    }),

    // ── Prefix priority for sort order (0 = low … 3 = highest) ──
    PrioMap = [PRIO_LOW = "0_PRIO_LOW", PRIO_MEDIUM = "1_PRIO_MEDIUM",
               PRIO_HIGH = "2_PRIO_HIGH", PRIO_HIGHEST = "3_PRIO_HIGHEST"],
    WithPrio = Table.TransformColumns(ExpandMatch, {
        {"priority", each try Record.Field(PrioMap, _) otherwise _, type text}
    }),

    Typed = Table.TransformColumnTypes(WithPrio, {
        {"received_datetime", type text},
        {"invoice_date", type date},
        {"due_date", type date},
        {"amount", type number},
        {"urgency_level", Int64.Type},
        {"bc_found", type logical}
    })
in
    Typed
