// The challenged required to get the only the first and last name without the middle name

let
    Source = Excel.CurrentWorkbook(){[Name="Names"]}[Content],
    ChangedType = Table.TransformColumnTypes(Source,{{"Names", type text}}),
    FinalResult= Table.AddColumn(ChangedType, "First&LastName", each
    [
        FirstName= Text.Start([Names], Text.PositionOf([Names]," ")),
        LastName= Text.End([Names],Text.PositionOf( Text.Reverse([Names])," ")),
        FirstLast= Text.Proper( FirstName)&" "&Text.Proper(LastName)
    ][FirstLast], type text
    )
in
    FinalResult
