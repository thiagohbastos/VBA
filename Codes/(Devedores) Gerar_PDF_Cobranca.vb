Sub Gerar_PDF_Cobranca()
'AUTOR: THIAGO BASTOS / git: thiagohbastos
'ESTA MACRO É RESPONSÁVEL POR GERAR ATAS DE COBRANÇA DAS PENDÊNCIAS AINDA NÃO NOTIFICADAS EM FORMATO PDF E GERAR AS PLANILHAS DE PENDÊNCAS ABERTAS GERAIS.

'---------
'Variáveis
'---------
Dim ultima_linha As Long
Dim pastanome As String
Dim linha As Integer
Dim linha_valores As Integer

Dim tempo As Double
Dim tempo_final As Double
Dim tempo_total As Double

tempo = Timer

Application.ScreenUpdating = False

'------------------------------------------
'Gerando planilhas com pendências em aberto
'------------------------------------------
Sheets("Máscara Notificação").Visible = True
Sheets("Devedores").Select
Range("Tabela3[[#Headers],[Diferença]]").Select
ActiveSheet.ListObjects("Tabela3").Range.AutoFilter Field:=12, Criteria1:= _
    "<>0", Operator:=xlAnd
ActiveSheet.ListObjects("Tabela3").Range.AutoFilter Field:=1, Criteria1:= _
    "=BRINKS", Operator:=xlAnd
Call PLAN_TRANSP
ActiveWorkbook.SaveAs Filename:= _
    "K:\GSAS\09 - Coordenacao Gestao Numerario\03 Gestao de Diferencas\DEVEDORES E CREDORES\PDFs de Cobrança\BRINKS - Pendências Gerais.xlsx" _
    , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
Workbooks("BRINKS - Pendências Gerais.xlsx").Close
ActiveSheet.ShowAllData

ActiveSheet.ListObjects("Tabela3").Range.AutoFilter Field:=12, Criteria1:= _
    "<>0", Operator:=xlAnd
ActiveSheet.ListObjects("Tabela3").Range.AutoFilter Field:=1, Criteria1:= _
    "=PROSEGUR", Operator:=xlAnd
Call PLAN_TRANSP
    ActiveWorkbook.SaveAs Filename:= _
    "K:\GSAS\09 - Coordenacao Gestao Numerario\03 Gestao de Diferencas\DEVEDORES E CREDORES\PDFs de Cobrança\PROSEGUR - Pendências Gerais.xlsx" _
    , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
Workbooks("PROSEGUR - Pendências Gerais.xlsx").Close
ActiveSheet.ShowAllData

ActiveSheet.ListObjects("Tabela3").Range.AutoFilter Field:=12, Criteria1:= _
    "<>0", Operator:=xlAnd
ActiveSheet.ListObjects("Tabela3").Range.AutoFilter Field:=1, Criteria1:= _
    "=PROTEGE", Operator:=xlAnd
Call PLAN_TRANSP
    ActiveWorkbook.SaveAs Filename:= _
    "K:\GSAS\09 - Coordenacao Gestao Numerario\03 Gestao de Diferencas\DEVEDORES E CREDORES\PDFs de Cobrança\PROTEGE - Pendências Gerais.xlsx" _
    , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
Workbooks("PROTEGE - Pendências Gerais.xlsx").Close
ActiveSheet.ShowAllData
    
Application.ScreenUpdating = False

'------------------------------------------------------------------------------------
'Organizando, filtrando e transcrevendo as pendências necessárias da tabela DEVEDORES
'------------------------------------------------------------------------------------
Range("Tabela3[[#Headers],[Diferença]]").Select
ActiveSheet.ListObjects("Tabela3").Range.AutoFilter Field:=12, Criteria1:= _
    "<>0", Operator:=xlAnd

ActiveSheet.ListObjects("Tabela3").Range.AutoFilter Field:=13, Criteria1:= _
    "="
Range("B1048576").Select
Selection.End(xlUp).Select
Selection.End(xlUp).Select
Range(Selection, Selection.End(xlUp)).Select
Selection.Copy

Sheets("Máscara Notificação").Select
Range("M7").Select
Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
    xlNone, SkipBlanks:=False, Transpose:=False
    
Application.CutCopyMode = False
ActiveSheet.Range("M8", Selection.End(xlDown)).RemoveDuplicates Columns:=1, Header:=xlYes
    
Application.ScreenUpdating = False

'------------------------------------------
'Filtrando tabela dinâmica e iniciando loop
'------------------------------------------
ActiveSheet.PivotTables("Tabela dinâmica1").PivotCache.Refresh
ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields("Data notificação"). _
    ClearAllFilters
ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields("Data notificação"). _
    CurrentPage = "(blank)"
ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields("Diferença"). _
    ClearAllFilters
With ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields("Diferença")
    .PivotItems(" $-   ").Visible = False
End With

linha = 8

'--------------
'Iniciando loop
'--------------
Do Until Cells(linha, 13) = ""
    ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields("Tesouraria"). _
        ClearAllFilters
    ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields("Tesouraria"). _
        CurrentPage = Cells(linha, 13).Value
            
    
    Application.ScreenUpdating = False
    '----------------------------------------------
    'Executando a formatação (ainda dentro do loop)
    '----------------------------------------------
    Range("A1048576").Select
    Selection.End(xlUp).Select
    
    linha_valores = ActiveCell.Row
    Range("G14:J" & linha_valores + 6).Value = Range("A8:D" & linha_valores).Value
    Range("G14:J" & linha_valores + 6).Borders.LineStyle = xlContinuous
    Range("J14:J" & linha_valores + 6).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
        
    Range("P6:S10").Copy
    Range("G" & linha_valores + 8).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    Application.ScreenUpdating = False
    '----------------------------------------------------
    'Transformando cada Ata em PDF(ainda dentro do loop)
    '----------------------------------------------------
    ultima_linha = Sheets("Máscara Notificação").Range("I1048576").End(xlUp).Row + 1
    pastanome = "K:\GSAS\09 - Coordenacao Gestao Numerario\03 Gestao de Diferencas\DEVEDORES E CREDORES\PDFs de Cobrança\" & Cells(6, 4).Value & " - " & Cells(6, 1).Value & ".pdf"
    
    Range("F1:K" & ultima_linha).Select
    
        ActiveSheet.PageSetup.PrintArea = "$F$1:$K$" & ultima_linha
        Application.PrintCommunication = True
    
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
            pastanome _
            , Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
            :=False, OpenAfterPublish:=False
           
    Range("G14:J" & linha_valores + 14).Select
    Selection.Delete Shift:=xlUp
           
    linha = linha + 1

Loop


Application.ScreenUpdating = False
'----------------------------------------------------------
'Limpando formatações e modificações feitas durante a MACRO
'----------------------------------------------------------

Sheets("Máscara Notificação").Select
Range("M7").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Delete Shift:=xlUp
Sheets("Máscara Notificação").Visible = False

Sheets("Devedores").Select
Range("Tabela3[[#Headers],[Diferença]]").Select
ActiveSheet.ShowAllData
    
Range("A1").Select

tempo_final = Timer

tempo_total = Round(tempo_final - tempo, 1)

Application.ScreenUpdating = True

MsgBox ("Concluído em " & tempo_total & " segundos")

End Sub