let
    Source = Excel.CurrentWorkbook(){[Name="Data"]}[Content],
    ChangedType = Table.TransformColumnTypes(Source,{{"Start Range", type text}, {"End Range", type text}, {"Operator", type text}, {"Condition", Int64.Type}}),
    FinalResult = Table.AddColumn(ChangedType, "Answer", each 
                                                                        [
                                                                        a= Number.From(Text.Middle([Start Range],1,2))-2,
                                                                        b= Number.From(Text.Middle([End Range],1,2)),
                                                                        c= List.Range(Numbers[Data],a,b-a-1),
                                                                        selector= Text.From([Operator])& Text.From([Condition]),
                                                                        d= List.Select(c, each Expression.Evaluate(Text.From(_)& selector) ),
                                                                        e=List.Sum(d)
                                                                        ][e]
                                                                        )
in
    FinalResult
