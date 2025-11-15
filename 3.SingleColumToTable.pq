let
    Source = Excel.CurrentWorkbook(){[Name="Table2"]}[Content][Column1],
    ListSplit = List.Split(Source,4),
    TableFromRows = Table.FromRows(ListSplit, type table [Name= Text.Type, Surname = Text.Type, Age = Int16.Type, JobTitle = Text.Type])
in
    TableFromRows
