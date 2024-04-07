// Get the list of 2022 dates where Monday is place after the 25th day

let
 StartDate=  #date (2022,01,01),
 EndDate=  #date (2022,12,31),
 Source =List.Dates(StartDate, 365, #duration (1,0,0,0) ),
 GetTableFromList = Table.FromList(Source,Splitter.SplitByNothing() ,{"Dates"}),
 DayofWeek = Table.AddColumn(GetTableFromList, "DayofWeek", each Date.DayOfWeekName([Dates])),
 DayNumber = Table.AddColumn(DayofWeek, "Day", each Date.Day([Dates]), Int64.Type),
 Filter = Table.SelectRows(DayNumber, each ([Day] > 25 and [DayofWeek]="lunedì")),
 ChangedType = Table.TransformColumnTypes(Filter,{{"Dates", type date}}),
    RemovedColumns = Table.RemoveColumns(ChangedType,{"DayofWeek", "Day"})
in
RemovedColumns
