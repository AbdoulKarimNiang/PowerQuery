// Get the sum of the numbers extracted from the string

let
    Source = Excel.CurrentWorkbook(){[Name="Table1"]}[Content],
    ChangedType = Table.TransformColumnTypes(Source,{{"String", type text}}),
    GetDigits = Table.AddColumn(ChangedType, "Digits", each Text.Select([String], {"0".."9"})),
    Summing = Table.AddColumn(GetDigits, "Sum of Digits", each List.Sum( 
                                                                        List.Transform( 
                                                                        Text.ToList( [Digits]), each Number.From(_))
                                                                        ))
in
    Summing
