// Calculate the total goal per match

let
    Source = Excel.CurrentWorkbook(){[Name="Matches"]}[Content],
    ChangedType = Table.TransformColumnTypes(Source,{{"Match", type text}, {"Team 1", type text}, {"Team 2", type text}, {"Result", type text}}),
    GetTotalGoals = Table.AddColumn(ChangedType, "Total Goals", each List.Sum (List.Transform(Text.Split([Result],"-"), Number.From)))[[Match],[Total Goals]],
    GroupForDuplicates = Table.Group(GetTotalGoals, {"Total Goals"}, {{"Count", each Table.RowCount(_), Int64.Type}, {"Table", each _, type table [Match=nullable text, Total Goals=number]}}),
    GetDuplicates = Table.SelectRows(GroupForDuplicates, each [Count] > 1),
    RemovedColumns = Table.RemoveColumns(GetDuplicates,{"Total Goals", "Count"}),
    ExpandedTable = Table.ExpandTableColumn(RemovedColumns, "Table", {"Match", "Total Goals"}, {"Match", "Total Goals"})
in
    ExpandedTable
