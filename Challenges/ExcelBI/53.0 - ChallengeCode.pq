// Extract The first and last number in the string

let
    Source = Excel.CurrentWorkbook(){[Name=Table1]}[Content],
    FisrtNumber = Table.AddColumn(Source, First Number, each Text.Start( Text.Select([String],{0..9}),1)),
    LastNumber = Table.AddColumn(FisrtNumber, Last Number, each Text.End( Text.Select([String],{0..9}),1))
in
    LastNumber
