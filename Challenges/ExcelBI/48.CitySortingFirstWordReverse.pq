// Sort (case insensitive) the cities on the basis of the first word of the city names and sorting should be done on the basis of 
// first word read backword. San Diego and Houston both have "n" at the end of first word, but would be worted on the basis of "naS", "notsuoH".
// Hence, San Diego will come prior to Houston

let
    Source = Excel.CurrentWorkbook(){[Name="Cities"]}[Content],
    ChangedType = Table.TransformColumnTypes(Source,{{"Cities", type text}}),
    Reversing = Table.AddColumn(ChangedType, "FirsLetterBack", each try Text.Lower( Text.Reverse(Text.Start([Cities],Text.PositionOf([Cities]," ")))) otherwise Text.Lower(Text.Reverse([Cities])) ),
    SortedRows = Table.SelectColumns( Table.Sort(Reversing,{{"FirsLetterBack", Order.Ascending}}),"Cities")
in
    SortedRows
