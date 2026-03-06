// Power Query for Process Runs Statistics
// Load and transform process_runs.json from emails folder
// Usage: In Power BI, use this as the source for a "ProcessRuns" table

let
    // Configuration
    EmailsFolder = "C:\Users\MatthiasDiBianco\planted foods AG\Circle Finance - Documents\Accounting\Mailbox\emails",
    ProcessRunsFile = EmailsFolder & "\process_runs.json",

    // Load JSON file
    Source = Json.Document(File.Contents(ProcessRunsFile)),

    // Convert to table
    ToTable = Table.FromList(Source, Splitter.SplitByNothing(), null, null, ExtraValues.Error),

    // Expand the column to get individual fields
    Expanded = Table.ExpandRecordColumn(ToTable, "Column1",
        {"timestamp", "type", "emails_processed", "categories_by_keywords", "categories_by_llm",
         "ven_rem_analyzed", "ven_followup_analyzed", "ven_inv_processed",
         "emails_archived", "human_completed_moved", "jsons_saved", "llm_calls_by_model"},
        {"timestamp", "type", "emails_processed", "categories_by_keywords", "categories_by_llm",
         "ven_rem_analyzed", "ven_followup_analyzed", "ven_inv_processed",
         "emails_archived", "human_completed_moved", "jsons_saved", "llm_calls_by_model"}),

    // Type conversions
    TypedColumns = Table.TransformColumnTypes(Expanded,
        {
            {"timestamp", type datetimezone},
            {"type", type text},
            {"emails_processed", Int64.Type},
            {"categories_by_keywords", Int64.Type},
            {"categories_by_llm", Int64.Type},
            {"ven_rem_analyzed", Int64.Type},
            {"ven_followup_analyzed", Int64.Type},
            {"ven_inv_processed", Int64.Type},
            {"emails_archived", Int64.Type},
            {"human_completed_moved", Int64.Type},
            {"jsons_saved", Int64.Type}
        }),

    // Add date column for daily grouping
    AddDate = Table.AddColumn(TypedColumns, "run_date", each Date.From([timestamp]), type date),

    // Sort by timestamp descending
    Sorted = Table.Sort(AddDate, {{"timestamp", Order.Descending}})
in
    Sorted
