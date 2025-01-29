Attribute VB_Name = "Module1"
Sub InserirLinhas()
    Dim nItens As Integer
    Dim Campo As String
    Selection.GoTo What:=wdGoToBookmark, Name:="TabelaItens"
    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    
    nItens = Val(ActiveDocument.Variables.Item("TOTAL_DE_ITENS_TABELA1").Value)
        
    Selection.MoveRight Unit:=wdCell
    Selection.MoveRight Unit:=wdCell
    Selection.MoveRight Unit:=wdCell
    Selection.MoveRight Unit:=wdCell
        
    For K = 1 To nItens
        
        Selection.MoveRight Unit:=wdCell
        Campo = "DOCVARIABLE cCod" & Trim(Str(K))
        Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:=Campo, PreserveFormatting:=True
        
        Selection.MoveRight Unit:=wdCell
        Campo = "DOCVARIABLE cDesc" & Trim(Str(K))
        Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:=Campo, PreserveFormatting:=True
        
        Selection.MoveRight Unit:=wdCell
        Campo = "DOCVARIABLE cMarca" & Trim(Str(K))
        Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:=Campo, PreserveFormatting:=True
        
        Selection.MoveRight Unit:=wdCell
        Campo = "DOCVARIABLE cLocal" & Trim(Str(K))
        Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:=Campo, PreserveFormatting:=True
        
    Next K
End Sub
