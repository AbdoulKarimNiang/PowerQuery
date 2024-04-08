// Get Team matching from the best 2 teams of both groups

// Group A query

let
    Source = Excel.CurrentWorkbook(){[Name="GroupA"]}[Content],
    ChangedType = Table.TransformColumnTypes(Source,{{"Group A", type text}, {"Points", Int64.Type}, {"NRR", type number}}),
    SortedRows = Table.Sort(ChangedType,{{"Points", Order.Descending}, {"NRR", Order.Descending}}),
    First2GroupA = Table.FirstN(SortedRows,2),
    DemotedHeaders = Table.DemoteHeaders(First2GroupA),
    Appended = Table.Combine({DemotedHeaders, GroupB}),
    RemoveRowsHeaders = Table.SelectRows(Appended, each ([Column2] <> "Points")),
    RemovedColumns = Table.RemoveColumns(RemoveRowsHeaders,{"Column2", "Column3"}),
    Team1 = List.Alternate(RemovedColumns[Column1],1,1,1),
    Team2 = List.Alternate(RemovedColumns[Column1],1,1,0),
    GetTable = Table.FromColumns({Team1,Team2}),
    RenamedColumns = Table.RenameColumns(GetTable,{{"Column1", "Team1"}, {"Column2", "Team2"}})
in
    RenamedColumns


// Group B query

let
    Source = Excel.CurrentWorkbook(){[Name="GroupB"]}[Content],
    ChangedTypeFirst = Table.TransformColumnTypes(Source,{{"Group B", type text}, {"Points", Int64.Type}, {"NRR", type number}}),
    SortingFirst = Table.Sort(ChangedTypeFirst,{{"Points", Order.Descending}, {"NRR", Order.Descending}}),
    First2GroupB = Table.FirstN(SortingFirst,2),
    DemotingHeaders = Table.DemoteHeaders(First2GroupB),
    FinalChangedType = Table.TransformColumnTypes(DemotingHeaders,{{"Column1", type text}, {"Column2", type any}, {"Column3", type any}})
in
    FinalChangedType
