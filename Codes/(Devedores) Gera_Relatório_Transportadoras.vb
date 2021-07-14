Sub PLAN_TRANSP()
'AUTOR: THIAGO BASTOS / git: thiagohbastos
'GERA UM RELATÓRIO PARA TV COM OS FILTROS ATUAIS

Dim ultima_linha As Integer

Application.ScreenUpdating = False

    Range("Tabela3[[#Headers],[Transportadora]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Workbooks.Add
    ActiveSheet.Paste
    Application.CutCopyMode = False
    'Selection.Delete Shift:=xlToLeft
    Range("C:C,E:E").Delete Shift:=xlToLeft
    Columns("P:Q").Cut
    Columns("B:B").Insert Shift:=xlToRight
    Range("Q1").FormulaR1C1 = "OBSERVAÇÕES"
    Range("C:C,E:E,A:A").EntireColumn.AutoFit
    Range("D1").FormulaR1C1 = "TES"
    Range("F1").FormulaR1C1 = "DOC"
    Range("I1").FormulaR1C1 = "Dependência"
    Range("R1").FormulaR1C1 = "ATUALIZAÇÃO TRASPORTADORA"
    Columns("J:L").EntireColumn.AutoFit
    Columns("O:P").EntireColumn.AutoFit
    Columns("Q:Q").ColumnWidth = 17
    Columns("I:I").ColumnWidth = 30
    
    Cells.Replace What:="NÃO NOTIFICADO", Replacement:="Notificado " & Date, LookAt:=xlPart _
        , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
    Columns("C:C").EntireColumn.AutoFit
    Range("A1").AutoFilter
    Range("Q1").Select
    
ultima_linha = Range("Q1").End(xlDown).Row
    
    Range(Selection, Selection.End(xlDown)).Copy
    Range("R1:R" & ultima_linha).Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Columns("R:R").Select
    Selection.ColumnWidth = 32
    Range("A1").Select
    
Application.ScreenUpdating = True
End Sub