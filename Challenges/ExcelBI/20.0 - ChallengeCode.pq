// Get the list of dates that are Monday and after the 25 of the month 
let
    Source = Excel.CurrentWorkbook(){[Name="Data"]}[Content],
    ChangedType = Table.TransformColumnTypes(Source,{{"2022", type date}},"en-US"
),
    CreateDates = Table.AddColumn(ChangedType, "Custom", each 
                                                                List.Dates(  
                                                                Date.StartOfMonth([2022]), 
                                                                Number.From( Date.EndOfMonth([2022]))-Number.From(Date.StartOfMonth([2022]))+1
                                                                , 
                                                                #duration(1,0,0,0)
                                                                )),
    GetTableFromList = Table.AddColumn(CreateDates, "Table", each Table.FromList([Custom],Splitter.SplitByNothing() ,{"Dates"})),
    RemoveTableColumn = Table.SelectColumns(GetTableFromList,{"Table"}),
    ExpandeDatesColumn = Table.ExpandTableColumn(RemoveTableColumn, "Table", {"Dates"}, {"Table.Dates"}),
    GetDayOfWeek = Table.AddColumn(ExpandeDatesColumn, "DayofWeek", each Date.DayOfWeekName([Table.Dates])),
    GetDayNumber = Table.AddColumn(GetDayOfWeek, "Day", each Date.Day([Table.Dates]), Int64.Type),
    FiltringLunedì = Table.SelectRows(GetDayNumber, each ([Day] > 25 and [DayofWeek]="lunedì")),
    ChangedTypeFinal = Table.TransformColumnTypes(FiltringLunedì,{{"Table.Dates", type date}})
in
    ChangedTypeFinal
