// Find all the possible matches without considering the home criteria

let
    Source = Excel.CurrentWorkbook(){[Name="TClubs"]}[Content],
    ChangedType = Table.TransformColumnTypes(Source,{{"Clubs", type text}}),
    SortClubs = Table.Sort(ChangedType,{{"Clubs", Order.Ascending}}),
    AddTeam2 = Table.AddColumn(SortClubs, "Team2", each SortClubs[Clubs]),
    Expantion = Table.ExpandListColumn(AddTeam2, "Team2"),
    Checking = Table.AddColumn(Expantion, "Custom", each [Clubs]=[Team2]),
    Filtering = Table.SelectRows(Checking, each ([Custom] = false))[[Clubs],[Team2]],
    Renaming = Table.RenameColumns(Filtering,{{"Clubs", "Team1"}}),
    NoSameMatches = Table.AddColumn(Renaming, "RemoveDuplicate", each Text.Combine(List.Sort(Text.ToList( Text.Remove(Text.Combine({[Team1],[Team2]}),{" "})),Order.Ascending))),
    RemovedDuplicates = Table.Distinct(NoSameMatches, {"RemoveDuplicate"})[[Team1],[Team2]]
in
    RemovedDuplicates
