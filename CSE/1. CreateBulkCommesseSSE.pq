// ######################################################################################## Starting Query ########################################################################################
let
    Source = Excel.CurrentWorkbook(){[Name="Table2"]}[Content],
    Intermediate = Source{0}[Column1],
    NumberStart = Text.Middle(Intermediate,3,Text.Length(Intermediate)-3),
    Final = Number.From( NumberStart)+1
in
    Final

// ######################################################################################## Padre Query ########################################################################################
let
    Source = Excel.CurrentWorkbook(){[Name="Table1"]}[Content],
    AddIndexFromStarting = Table.AddIndexColumn(Source, "IntermediateCommessa", Starting, 1, Int64.Type),
    SeletColumn = Table.SelectColumns(AddIndexFromStarting,{"CONDOMINIO", "Indirizzo","Comune", "IntermediateCommessa"}),
    PrefixTG = Table.TransformColumns(SeletColumn, {{"IntermediateCommessa", each "CTG" & Text.From(_, "it-IT"), type text}}),
    Combining = Tablem.Combine({UpperCaseColumnHeaders,Padre }),
    UpperCaseColumnHeaders = Table.TransformColumnNames(Combining, Text.Upper),
    UpperCaseCONDOMINIO = Table.TransformColumns(PrefixTG,{{"CONDOMINIO", Text.Upper, type text}, {"Indirizzo", Text.Upper, type text}, {"Comune", Text.Upper, type text}})
  in
    UpperCaseCONDOMINIO

// ######################################################################################## Commesse Query ########################################################################################

let
    Source = Excel.CurrentWorkbook(){[Name="Table1"]}[Content],
    AddIndexFromStarting = Table.AddIndexColumn(Source, "Index", Starting, 1, Int64.Type),
    SplitByRows = Table.ExpandListColumn(Table.TransformColumns(AddIndexFromStarting, {{"COMMESSA", Splitter.SplitTextByDelimiter("+", QuoteStyle.Csv), let itemType = (type nullable text) meta [Serialized.Text = true] in type {itemType}}}), "COMMESSA"),
    MergeIndexTypeCommessa = Table.AddColumn(SplitByRows, "IntermediateCommessa", each Text.Combine({Text.From([Index], "it-IT"), [COMMESSA]}, "-"), type text),
    PrefixTG = Table.TransformColumns(MergeIndexTypeCommessa, {{"IntermediateCommessa", each "CTG" & Text.From(_, "it-IT"), type text}}),
    Suffix01 = Table.TransformColumns(PrefixTG, {{"IntermediateCommessa", each _ & "01", type text}}),
    SeletColumn = Table.SelectColumns(Suffix01,{"CONDOMINIO", "Indirizzo","Comune", "IntermediateCommessa"}),
    Appending = Table.Combine({SeletColumn, Padre}),
    UpperCaseCondominio = Table.TransformColumns(Appending,{{"CONDOMINIO", Text.Upper, type text}, {"Indirizzo", Text.Upper, type text}, {"Comune", Text.Upper, type text}}),
    RenameColumn = Table.RenameColumns(UpperCaseCondominio,{"IntermediateCommessa", "CODICE COMMESSA"}),
    UpperCaseColumnHeaders = Table.TransformColumnNames(RenameColumn, Text.Upper),
    Sorting = Table.Sort(UpperCaseColumnHeaders,{{"CODICE COMMESSA", Order.Ascending}}),
    Provincia = Table.AddColumn(Sorting, "PROVINCIA", each "TORINO", Text.Type),
    COMM_OLD = Table.AddColumn(Provincia, "COMM-OLD", each null, Text.Type),
    TipoContratto = Table.AddColumn(COMM_OLD, "TIPO CONTRATTO", each if Text.Length([CODICE COMMESSA]) = 10 then "PADRE" else Text.Middle([CODICE COMMESSA], 11,2)),  
    Final = Table.SelectColumns(TipoContratto, {"CONDOMINIO", "INDIRIZZO", "COMUNE", "PROVINCIA","COMM-OLD","TIPO CONTRATTO","CODICE COMMESSA" })
in
    Final

// ######################################################################################## FindReplace Query ########################################################################################

let
  Source = Table.FromRecords(
    {[Find = "PADRE", Replace = "?"], [Find = "CM", Replace = "COM"], [Find = "MS", Replace = "MAN"], [Find = "LT", Replace = "LETTURE"]},
    type table [Find = Any.Type, Replace = Any.Type]
  )
in
  Source

SSE
let
  Source = Commesse, 
  SelectColumns = Table.SelectColumns(Source, {"CODICE COMMESSA", "CONDOMINIO", "INDIRIZZO", "COMUNE", "TIPO CONTRATTO"}), 
  BulkReplace = List.Accumulate(
    {0 .. List.Count(FindReplace[Find]) - 1}, 
    SelectColumns, 
    (state, current) => Table.ReplaceValue(state, FindReplace[Find]{current}, FindReplace[Replace]{current}, Replacer.ReplaceText, {"TIPO CONTRATTO"})
  ), 
  ReplacedQuotationMark = Table.ReplaceValue(BulkReplace, "?", null, Replacer.ReplaceValue, {"TIPO CONTRATTO"}), 
  CreazioneDescrizioneSSE = Table.AddColumn(
    ReplacedQuotationMark, 
    "Descrizione SSE", 
    each 
      if [TIPO CONTRATTO] is null then
        [CONDOMINIO] & "#" & [INDIRIZZO] & " " & [COMUNE]
      else
        [CONDOMINIO] & "|" & [TIPO CONTRATTO] & "#" & [INDIRIZZO] & " " & [COMUNE], 
    Text.Type
  ), 
  RemoveOtherCOlumns = Table.SelectColumns(CreazioneDescrizioneSSE, {"CODICE COMMESSA", "Descrizione SSE"}), 
  StatoAttivo = Table.AddColumn(RemoveOtherCOlumns, "Stato Commessa", each "Attivo")
in
  StatoAttivo


