// Calculate US Presidents' Tenure

let
    Source = Excel.CurrentWorkbook(){[Name="PresidentsTenure"]}[Content],
    ChangedType = Table.TransformColumnTypes(Source,{{"Presidents", type text}, {"Tenure", type text}}),
    Result=
            List.Sum( 
                      [
                          Start= Table.AddColumn(ChangedType, "Starting", each Text.Start([Tenure],4), type text),
                          End= Table.AddColumn(Start, "Ending", each if Text.Length([Tenure]) >= 7 then Text.End([Tenure], Text.Length([Tenure])-1 - Text.PositionOf([Tenure],"-")) else [Starting],type text),
                          Correction= Table.AddColumn(End, "Correct End", each if Text.Length([Ending]) <4 or [Ending] = "" then Text.Start([Tenure],2)&[Ending] else [Ending],type text),
                          Correction2022= Table.AddColumn(Correction, "Final End", each if (Text.PositionOf([Tenure],"-")=4 and [Starting]="2021" ) then Date.Year(DateTime.LocalNow()) else [Correct End], type number), 
                          Duration= Table.AddColumn(Correction2022, "Lasting", each Number.From([Final End])-Number.From([Starting]),type number),
                          OnlyLasting=Table.SelectColumns(Duration,"Lasting")
                      ][OnlyLasting][Lasting]
                      )

in
    Result
