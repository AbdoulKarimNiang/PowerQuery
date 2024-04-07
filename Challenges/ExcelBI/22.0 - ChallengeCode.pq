// Get the count of attendance for Employee per day

// Date

let
    Source = Excel.CurrentWorkbook(){[Name="Dates"]}[Content],
    DateType = Table.TransformColumnTypes(Source,{{"Date", type date}}),
    Group = Table.Group(DateType, {"Date", "Employees"}, {{"Count", each Table.RowCount(_), Int64.Type}}),
    Sorting = Table.Sort(Group,{{"Date", Order.Ascending}}),
    ChangeToText = Table.TransformColumnTypes(Sorting,{{"Date", type text}}),
    Index = Table.AddIndexColumn(ChangeToText, "Index", 0, 1, Int64.Type)
in
    Index



// Employee Query

let
    Source = Excel.CurrentWorkbook(){[Name="Employees"]}[Content],
    TextType = Table.TransformColumnTypes(Source,{{"Employees", type text}}), // Preparation for Text.Combine
    GetEmployee = Table.NestedJoin(TextType, {"Employees"}, Dates, {"Employees"}, "Dates", JoinKind.LeftOuter),
    ExpandeMerge = Table.ExpandTableColumn(GetEmployee, "Dates", {"Date", "Count"}, {"Date", "Count"}),
    PlaceHolder= Table.ReplaceValue(ExpandeMerge,null,"nothing",Replacer.ReplaceValue,{"Date"}),
    CombineText = Table.AddColumn(PlaceHolder, "Headers", each 
                                                          if [Date] = "nothing" then "nothing" 
                                                          else Text.Combine
                                                                          (
                                                                          { 
                                                                          Text.End( [Date],4),
                                                                           Text.Middle([Date],3,2),
                                                                           Text.Start([Date],2)
                                                                          },"-" 
                                                                          )
                                                                          ),
    RemoveColumnsBeforPivot = Table.RemoveColumns(CombineText,{"Date"}),
    PivotColumns = Table.Pivot(RemoveColumnsBeforPivot, List.Distinct(RemoveColumnsBeforPivot[Headers]), "Headers", "Count", List.Sum),
    RemovePlaceHolder = Table.RemoveColumns(PivotColumns,{"nothing"}),
    ReplaceNulls = Table.ReplaceValue(RemovePlaceHolder,null,0,Replacer.ReplaceValue,{"2022-10-01", "2022-10-02", "2022-10-05", "2022-10-03", "2022-10-04"})
in
    ReplaceNulls
