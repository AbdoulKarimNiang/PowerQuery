// Get the TOP 3 character frequency of US Presidents Full Names 

let
    Source = Excel.CurrentWorkbook(){[Name="USPresidents"]}[Content],
    ChangedType = Table.TransformColumnTypes(Source,{{"US Presidents", type text}}),
    UpperCase = Table.TransformColumns(ChangedType,{{"US Presidents",Text.Upper, type text}}),
    RemoveCharacters = Table.TransformColumns(UpperCase,{{"US Presidents", each Text.Remove(_,{" ","."}), type text}}),
    GetList = Table.AddColumn(RemoveCharacters, "Alphabets", each Splitter.SplitTextByRepeatedLengths(1)([US Presidents])),
    AlphabetExpantion = Table.ExpandListColumn(GetList, "Alphabets"),
    RenamingAlphabet = Table.RemoveColumns(AlphabetExpantion,{"US Presidents"}),
    Grouped = Table.Group(RenamingAlphabet, {"Alphabets"}, {{"Frequency", each Table.RowCount(_), Int64.Type}}),
    GroupedFrequency = Table.Group(Grouped, {"Frequency"}, {{"Placeholder", each _, type table [Alphabets=text, Frequency=number]}}),
    Max3 = Table.MaxN(GroupedFrequency,"Frequency",3),
    Result = Table.ExpandTableColumn(Max3, "Placeholder", {"Alphabets", "Frequency"}, {"Alphabets", "Frequency.1"})[[Alphabets],[Frequency]]
in
    Result
