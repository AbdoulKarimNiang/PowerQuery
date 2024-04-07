// Get the US Presidents that have has B as first letter in name, middle, last name 

let
    Source = Excel.CurrentWorkbook(){[Name="USPresidents"]}[Content],
    SplitbySpace = Table.SplitColumn(Source, "US Presidents", Splitter.SplitTextByDelimiter(" ", QuoteStyle.Csv), {"US Presidents.1", "US Presidents.2", "US Presidents.3"}),
    ChangedType = Table.TransformColumnTypes(SplitbySpace,{{"US Presidents.1", type text}, {"US Presidents.2", type text}, {"US Presidents.3", type text}}),
    CombineFirstLetters = Table.AddColumn(ChangedType, "StartWithB", each Text.Combine({Text.Start([US Presidents.1],1),Text.Start([US Presidents.2],1),Text.Start([US Presidents.3],1)})),
    ContainsB = Table.SelectRows(CombineFirstLetters, each Text.Contains([StartWithB],"B")),
    GetNameBack = Table.CombineColumns(ContainsB,{"US Presidents.1", "US Presidents.2", "US Presidents.3"},Combiner.CombineTextByDelimiter(" ", QuoteStyle.None),"US Presidents"),
    RemoveColumn = Table.RemoveColumns(GetNameBack,{"StartWithB"})
in
    RemoveColumn
