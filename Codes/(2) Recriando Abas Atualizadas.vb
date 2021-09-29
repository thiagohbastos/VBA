Sub Atualiza_Apoio()

Application.ScreenUpdating = False

Sheets("Apoio").Visible = True
Sheets("Exemplo").Visible = True

Sheets("Apoio").Select
Range("A:F").Select
Selection.ClearContents
Sheets("Apoio").Range("G2:H1048576").ClearContents

Range("A2").Select
ActiveCell.FormulaR1C1 = "=COUNTIFS(R1C5:RC[4],RC[4])"
Range("B2").Select
ActiveCell.FormulaR1C1 = "=RC[-1]&RC[3]"

Sheets("Apoio").Range("C:E").Value = Sheets("Base").Range("B:D").Value
Columns("C:E").Select
ActiveSheet.Range("$C:$E").RemoveDuplicates Columns:=Array(1, 2, 3), Header:= _
    xlYes

Ln = Range("E1048576").End(xlUp).Row

Range("C2:D" & Ln).SpecialCells(xlCellTypeBlanks).EntireRow.Delete

Range("A2:B2").Select
Selection.AutoFill Destination:=Range("A2:B" & Ln)

Columns("G:G").Value = Columns("E:E").Value
ActiveSheet.Columns("G:G").RemoveDuplicates Columns:=1, Header:=xlNo

Ln = Range("G1048576").End(xlUp).Row

Range("H2").Select
ActiveCell.FormulaR1C1 = "=COUNTIF(C[-3],RC[-1])"
Selection.AutoFill Destination:=Range("H2:H" & Ln)

Application.DisplayAlerts = False

For Each aba In ThisWorkbook.Sheets
   If aba.Name <> "Base" And aba.Name <> "Apoio" And aba.Name <> "Exemplo" And aba.Name <> "Crivo-Ticket" Then
        aba.Delete
   End If
Next

Application.DisplayAlerts = True

opera = Sheets("Apoio").Range("G1048576").End(xlUp).Row

For Each operacao In Sheets("Apoio").Range("G2:G" & opera)

    Sheets("Exemplo").Copy Before:=Sheets(1)
    Sheets("Exemplo (2)").Name = operacao.Value
    Range("B1").Value = operacao
    
    cities = Range("A1").Value + 2
    
    Range("A3").Value = 1
    Range("A3:A" & cities).Select
    Selection.DataSeries Rowcol:=xlColumns, Type:=xlLinear, Date:=xlDay, _
        Step:=1, Trend:=False
    Range("B3:J3").Select
    Selection.AutoFill Destination:=Range("B3:J" & cities)

Next

Sheets("Apoio").Visible = False
Sheets("Exemplo").Visible = False

Application.ScreenUpdating = True

End Sub