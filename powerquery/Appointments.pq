let
    FolderPath = "C:\Users\mgori\OneDrive\Desktop\BIOSTAT__TRAINEE\PROJECTS\CANCEL__REPORTS\raw\zanda_exports\",

    Source   = Folder.Files(FolderPath),
    CsvOnly  = Table.SelectRows(Source, each Text.Lower([Extension]) = ".csv"
                                        and not Text.StartsWith([Name], "~")),
    Parsed   = Table.AddColumn(CsvOnly, "Data", each
                 Table.PromoteHeaders(
                     Csv.Document([Content], [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
                     [PromoteAllScalars = true])),
    KeepData = Table.SelectColumns(Parsed, {"Data"}),
    Expanded = Table.ExpandTableColumn(KeepData, "Data",
                 {"Client Name","Therapist","Date","End Time","Day","Appointment Type",
                  "Type of Play Therapy","Duration (min)","Status","Reason for Cancellation"}),

    Filtered = Table.SelectRows(Expanded, each [Date] <> null and Text.Trim([Date]) <> ""
                                         and [Client Name] <> "Client Name"),
    Typed    = Table.TransformColumnTypes(Filtered, {
                 {"Client Name", type text}, {"Therapist", type text}, {"Date", type text},
                 {"End Time", type text}, {"Day", type text}, {"Appointment Type", type text},
                 {"Type of Play Therapy", type text}, {"Duration (min)", Int64.Type},
                 {"Status", type text}, {"Reason for Cancellation", type text}}),

    AddDateParsed = Table.AddColumn(Typed, "DateParsed", each
        let
            DateText = Text.Trim(Text.BeforeDelimiter([Date], " ")),
            Parts    = Text.Split(DateText, "/")
        in
            try #date(Number.From(Parts{2}), Number.From(Parts{1}), Number.From(Parts{0})) otherwise null,
        type date),

    AddStartOfWeek = Table.AddColumn(AddDateParsed, "Start of Week", each
        Date.StartOfWeek([DateParsed], Day.Monday), type date),

    AddMonthName = Table.AddColumn(AddStartOfWeek, "Month Name", each
        Date.ToText([Start of Week], "MM-MMMM"), type text),

    AddQuarter = Table.AddColumn(AddMonthName, "Quarter", each
        "Q" & Text.From(Date.QuarterOfYear([Start of Week])) & " " & Text.From(Date.Year([Start of Week])),
        type text),

    AddYearMonth = Table.AddColumn(AddQuarter, "Year-Month", each
        Date.ToText([Start of Week], "yyyy-MM-dd"), type text),

    AddCC = Table.AddColumn(AddYearMonth, "Cancelled_Completed", each
        if [Status] = "Completed" then "Completed" else "Cancelled", type text),

    AddYear = Table.AddColumn(AddCC, "Year", each
        Date.Year([Start of Week]), Int64.Type),

    AddMonthLabel = Table.AddColumn(AddYear, "Month Label", each
        Date.ToText([Start of Week], "yyyy-MM") & " | " & Date.ToText([Start of Week], "MMMM"),
        type text),

    AddQuarterLabel = Table.AddColumn(AddMonthLabel, "Quarter Label", each
        Text.From(Date.Year([Start of Week])) & "-Q" & Text.From(Date.QuarterOfYear([Start of Week])),
        type text),

    Sorted   = Table.Sort(AddQuarterLabel, {{"Client Name", Order.Ascending}, {"DateParsed", Order.Ascending}}),
    NoErrors = Table.RemoveRowsWithErrors(Sorted),
    #"Removed Columns" = Table.RemoveColumns(NoErrors,{"Quarter"})
in
    #"Removed Columns"