Sub Gerar_PDFs()
'AUTOR: THIAGO BASTOS / git: thiagohbastos
'Macro responsável por copilar os dados da aba Ressarcimento em fichas contábeis.

Application.ScreenUpdating = False

Dim tempo As Double
Dim tempo_final As Double
Dim tempo_total As Double

tempo = Timer

Sheets("SID_FALTA_RESSARCIMENTO").Visible = True
    
'Transcrevendo os valores do "controle" para a tabela da ficha:

    Sheets("Ressarcimento").Select
    Range("E1048576").Select
    Selection.End(xlUp).Select
    
lastline = ActiveCell.Row - 1

Sheets("SID_FALTA_RESSARCIMENTO").Select

If lastline > 3 Then
    Range("BD6").Select
    ActiveCell.FormulaR1C1 = "=Ressarcimento!R[-3]C[-55]"
    Selection.AutoFill Destination:=Range("BD6:BD" & lastline + 3), Type:=xlFillDefault
    
    Range("BE6").Select
    ActiveCell.FormulaR1C1 = "=Ressarcimento!R[-3]C[-55]"
    Selection.AutoFill Destination:=Range("BE6:BE" & lastline + 3), Type:=xlFillDefault

    Range("BG6").Select
    Selection.FormulaR1C1 = "=Ressarcimento!R[-3]C[-53]"
    Selection.AutoFill Destination:=Range("BG6:BG" & lastline + 3), Type:=xlFillDefault

    'Gerando os valores que serão transferidos para a ficha a partir da aba de preenchimento.

    Range("BH6").Select
    tes = "BH6:BH" & lastline + 3
    Selection.AutoFill Destination:=Range(tes)
    
    Range("BI6").Select
    digito = "BI6:BI" & lastline + 3
    Selection.AutoFill Destination:=Range(digito)
    
    Range("BJ6").Select
    nome_arquivo = "BJ6:BJ" & lastline + 3
    Selection.AutoFill Destination:=Range(nome_arquivo)
    
Else
        Range("BD6").Select
    ActiveCell.FormulaR1C1 = "=Ressarcimento!R[-3]C[-55]"
    
    Range("BE6").Select
    ActiveCell.FormulaR1C1 = "=Ressarcimento!R[-3]C[-55]"

    Range("BG6").Select
    Selection.FormulaR1C1 = "=Ressarcimento!R[-3]C[-53]"

    'Gerando os valores que serão transferidos para a ficha a partir da aba de preenchimento.

    Range("BH6").Select
    tes = "BH6"
    
    Range("BI6").Select
    digito = "BI6"
    
    Range("BJ6").Select
    nome_arquivo = "BJ6"

End If
    
linha = 6
    
Do Until Cells(linha, 56) = ""

    
Dim pastanome As String
pastanome = "K:\GSAS\09 - Coordenacao Gestao Numerario\03 Gestao de Diferencas\DEVEDORES E CREDORES\PDFs Ressarcimento\" & Cells(linha, 62).Value & ".pdf"

    Range("AF26:AK28").Value = Cells(linha, 60).Value
    Range("AL26:AM28").Value = Cells(linha, 61).Value
    Range("AN41:BB42").Value = Cells(linha, 59).Value
    
    Cells(linha, 58).Select
    Selection.FormulaR1C1 = "=R15C2"
    
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        pastanome _
        , Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
        :=False, OpenAfterPublish:=False
        
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False

    
        linha = linha + 1
Loop

    Range("BH7:BJ7").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    
If lastline > 3 Then
    Sheets("Ressarcimento").Select
    Range("D3").Select
    ActiveCell.FormulaR1C1 = "=SID_FALTA_RESSARCIMENTO!R[3]C[54]"
    Selection.AutoFill Destination:=Range("D3:D" & lastline)
    Range("D3:D" & lastline).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Application.CutCopyMode = False
Else
    Sheets("Ressarcimento").Select
    Range("D3").Select
    ActiveCell.FormulaR1C1 = "=SID_FALTA_RESSARCIMENTO!R[3]C[54]"
    Range("D3").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Application.CutCopyMode = False
End If
'Apagando dados da ficha e ocultando-a:

Sheets("SID_FALTA_RESSARCIMENTO").Select
Range("BD6:BG1048576").Select
Selection.ClearContents
ActiveWindow.SelectedSheets.Visible = False
    
Sheets("Ressarcimento").Select
Range("A2").Select

'Calculando tempo para conclusão do VBA
tempo_final = Timer

tempo_total = Round(tempo_final - tempo, 1)

MsgBox ("Concluído em " & tempo_total & " s")

Application.ScreenUpdating = True

End Sub