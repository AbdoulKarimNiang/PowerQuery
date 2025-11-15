let
    Source = Excel.CurrentWorkbook(){[Name="Table1"]}[Content],
    ChangedType = Table.TransformColumnTypes(Source,{{"Date", Int64.Type}, {"Description", type text}, {"Amount", Int64.Type}, {"Balance", Int64.Type}}),
    Categorization = Table.AddColumn(
        ChangedType,
        "Category", each [
            KeyWords = TransactionLookups[KeyWord],
            Mappings = TransactionLookups[Mapping],
            NumberIterations = List.Count(KeyWords),
            ListToIterate = {0..NumberIterations-1},
            StartCatorories = {},
            Categories = List.Accumulate (
                ListToIterate,
                StartCatorories,
                (previous, current) => if Text.Contains([Description], KeyWords{current}, Comparer.OrdinalIgnoreCase) then previous & {Mappings{current}} else previous
            ),
            AllCategories = Text.Combine(
                    List.Transform(
                        List.Distinct(Categories),
                        each Text.Proper(_)
                        ),
                        ", ")
            ][AllCategories]
    )

in
    Categorization
