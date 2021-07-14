Sub Controle_email()
'AUTOR: THIAGO BASTOS / git: thiagohbastos
'Macro responsável por iniciar o tratamento de lançamentos das transportadoras (tabela dinâmica).

Application.ScreenUpdating = False

'Limpando dados anteriores:
Sheets("Ressarcimento").Select
Range("A3:F1048576").ClearContents
    
'Transcrevendo os valores de "resumo" para "controle":
Range("I6:I7").Select
Range(Selection, Selection.End(xlToRight)).Select
Selection.Copy
Range("E3").Select
Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
    False, Transpose:=True
Application.CutCopyMode = False
Selection.Borders(xlDiagonalDown).LineStyle = xlNone
Selection.Borders(xlDiagonalUp).LineStyle = xlNone
With Selection.Borders(xlEdgeLeft)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Selection.Borders(xlEdgeTop)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Selection.Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Selection.Borders(xlEdgeRight)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Selection.Borders(xlInsideVertical)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Selection.Borders(xlInsideHorizontal)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
Range("f3").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    
'Adaptando fórmulas:
lastline = Range("E1048576").End(xlUp).Row - 1
    
If lastline >= 4 Then
    'CI
        Range("A3").Select
        ActiveCell.FormulaR1C1 = _
            "=IFERROR(INDEX(FAIXAS!C[6],MATCH(Ressarcimento!RC[4],FAIXAS!C[1])),"""")"
    CI = "A3:A" & lastline
        Selection.AutoFill Destination:=Range(CI)
        
    'Local
        Range("B3").Select
        ActiveCell.FormulaR1C1 = _
            "=IFERROR(INDEX(DEPENDENCIAS!C,MATCH(RC[3],DEPENDENCIAS!C[-1],0)),"""")"
    localidade = "B3:B" & lastline
        Selection.AutoFill Destination:=Range(localidade)
        
    'Transportadora
        Range("C3").Select
        ActiveCell.FormulaR1C1 = _
            "=IFERROR(INDEX(DEPENDENCIAS!C[3],MATCH(RC[2],DEPENDENCIAS!C[-2],0)),"""")"
    transportadora = "C3:C" & lastline
        Selection.AutoFill Destination:=Range(transportadora)
Else
    Range("A3").Select
    ActiveCell.FormulaR1C1 = _
    "=IFERROR(INDEX(FAIXAS!C[6],MATCH(Ressarcimento!RC[4],FAIXAS!C[1])),"""")"
    Range("B3").Select
    ActiveCell.FormulaR1C1 = _
    "=IFERROR(INDEX(DEPENDENCIAS!C,MATCH(RC[3],DEPENDENCIAS!C[-1],0)),"""")"
    Range("C3").Select
    ActiveCell.FormulaR1C1 = _
    "=IFERROR(INDEX(DEPENDENCIAS!C[3],MATCH(RC[2],DEPENDENCIAS!C[-2],0)),"""")"
    
End If

'Range("C" & lastline + 1, "D" & lastline + 1).Value = "TOTAL"

Application.ScreenUpdating = True

End Sub