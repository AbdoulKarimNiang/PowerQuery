// Get the distinct list of animals in one columns 
let
    Source = Excel.CurrentWorkbook(){[Name="Animals"]}[Content],
    
    Intermediate= 
    
    [ 
        
    ListA=Source[Animals1], ListB=Source[Animals2], ListC=Source[Animals3] ,

    CheckListA = List.RemoveItems(ListA,ListB&ListC),

    CheckListB= List.RemoveItems(ListB,ListA&ListC),

    CheckListC= List.RemoveItems(ListC,ListA&ListB),
    
    FinalList= CheckListA&CheckListB&CheckListC  
    
    ] [FinalList],

    ResultDistinct= List.Distinct(Intermediate),

    ToTable=Table.FromList(ResultDistinct,Splitter.SplitByNothing(),{"Answer Expected"})
in
    ToTable
