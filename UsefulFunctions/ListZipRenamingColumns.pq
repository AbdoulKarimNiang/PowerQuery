let
  Source = ""
  OriginalColumNames = Table.ColumnNames(Source),
  RenamedColumns = List.Transform(OriginalColumNames, each Text.Lower(Text.Replace(_, " ","_"))),
  ColumnsZipped = List.Zip({OriginalColumNames, RenamedColumns}),
  Renaming = Table.RenameColumns(#"Split Column by Delimiter", ColumnsZipped)

in
Renaming
  
