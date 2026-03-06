// Power Query for Cleanup Runs Statistics
// Load and transform cleanup_runs.json from emails folder
// Usage: In Power BI, use this as the source for a "CleanupRuns" table

let
    // Configuration
    EmailsFolder = "C:\Users\MatthiasDiBianco\planted foods AG\Circle Finance - Documents\Accounting\Mailbox\emails",
    CleanupRunsFile = EmailsFolder & "\cleanup_runs.json",

    // Load JSON file
    Source = Json.Document(File.Contents(CleanupRunsFile)),

    // Convert to table
    ToTable = Table.FromList(Source, Splitter.SplitByNothing(), null, null, ExtraValues.Error),

    // Expand the column to get individual fields
    Expanded = Table.ExpandRecordColumn(ToTable, "Column1",
        {"timestamp", "type", "emails_in_folder", "emails_classified", "emails_with_pass2",
         "emails_moved_to_archive", "api_calls_used"},
        {"timestamp", "type", "emails_in_folder", "emails_classified", "emails_with_pass2",
         "emails_moved_to_archive", "api_calls_used"}),

    // Type conversions
    TypedColumns = Table.TransformColumnTypes(Expanded,
        {
            {"timestamp", type datetimezone},
            {"type", type text},
            {"emails_in_folder", Int64.Type},
            {"emails_classified", Int64.Type},
            {"emails_with_pass2", Int64.Type},
            {"emails_moved_to_archive", Int64.Type},
            {"api_calls_used", Int64.Type}
        }),

    // Add date column for daily grouping
    AddDate = Table.AddColumn(TypedColumns, "cleanup_date", each Date.From([timestamp]), type date),

    // Sort by timestamp descending
    Sorted = Table.Sort(AddDate, {{"timestamp", Order.Descending}})
in
    Sorted
