let
    Source = Excel.CurrentWorkbook(){[Name="Table1"]}[Content],
    Splitted = Table.TransformColumns(Source, 
        {
            {"Column1", 
            each [
                SpitByComma = Text.Split( _,","),
                SplitBySpace = List.Transform(SpitByComma, each Text.Split(_, " ")),
                Combine = List.Combine(SplitBySpace)
            ][Combine]
            }
        }
    ),
    FinalTable = [
    NumberOfColumns = 3,
    RowsWithValues = Splitted{0}[Column1],
    NumberOfValues = List.Count(RowsWithValues),
    Columns = List.FirstN(RowsWithValues, NumberOfColumns),
    Values = List.LastN(RowsWithValues, NumberOfValues - NumberOfColumns),
    // NumberOfGroups = List.Count(Values) / NumberOfColumns,
    ListOfRows = List.Split(Values, NumberOfColumns) ,
    ActualTable = Table.FromRows(ListOfRows, Columns )
    ][ActualTable]
in
    FinalTable
