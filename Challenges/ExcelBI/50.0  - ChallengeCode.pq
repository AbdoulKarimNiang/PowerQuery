// Sort the table to get only the rows where word 1 and word 2 have the same letters.
// Don't consider spaces

let
    Source = Excel.CurrentWorkbook(){[Name="Words"]}[Content],
    ChangedType = Table.TransformColumnTypes(Source,{{"Word1", type text}, {"Word2", type text}}),
    Word1Tranformed=Table.AddColumn(Source,"New Word1",each List.Sort( Text.ToList(Text.Lower(Text.Remove([Word1]," "))), Order.Ascending)),
    Word2Tranformed=Table.AddColumn(Word1Tranformed,"New Word2",each List.Sort( Text.ToList(Text.Lower(Text.Remove([Word2]," "))), Order.Ascending)),
    Same= Table.AddColumn(Word2Tranformed, "Verifing", each [New Word1]=[New Word2]),
    FilteredRows = Table.SelectRows(Same, each ([Verifing] = true)),
    SelectColumns = Table.SelectColumns(FilteredRows,{"Word1", "Word2"})
in
    SelectColumns
