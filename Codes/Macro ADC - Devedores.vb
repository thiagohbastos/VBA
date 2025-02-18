'AUTOR: THIAGO BASTOS (B042786) 29/08/2022
'Macro responsável por inserir os novos registro da faixa Devedores
Sub Adiciona_Registros()

    Application.ScreenUpdating = False

    Sheets("Devedores").Select

    If Sheets("Devedores").Range("B2").Value = "SIM" Then
        Dim tempo As Double
        Dim tempo_final As Double
        Dim tempo_total As Double
        
        tempo = Timer

        last_line = Range("A1048576").End(xlUp).Row + 1

        Sheets("CAB D-1").Select

        On Error Resume Next
        ActiveSheet.ShowAllData
        On Error GoTo 0

        ActiveSheet.ListObjects("Tabela_SQLGDNP_GNU_TI_DIFERENCAS_CAB").Range. _
        AutoFilter Field:=17, Criteria1:="FALTA REC"

        Range("Tabela_SQLGDNP_GNU_TI_DIFERENCAS_CAB[#All]").Select
        Selection.Copy
        Sheets("Devedores").Select

        Data = Sheets("CAB D-1").Range("D2").Value
        Data = CDate(Data)
        dif = DateDiff("d", Date, Data)
        Data = DateAdd("d", dif, Date)

        Range("A" & last_line).Select
        ActiveSheet.Paste
        Application.CutCopyMode = False

        Rows(last_line & ":" & last_line).Select
        Selection.Delete Shift:=xlUp

        n_ll = Range("A1048576").End(xlUp).Row

        Range("D" & last_line - 1 & ":D" & n_ll).Value = Data
        Range("A" & last_line - 1 & ":S" & last_line - 1).Select
        Selection.Copy

        Range("A" & last_line & ":S" & n_ll).Select
        Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False

        Range("B2:C2").Value = "NÃO"

        Sheets("CAB D-1").Select

        On Error Resume Next
        ActiveSheet.ShowAllData
        On Error GoTo 0

        Range("A1").Select

        Sheets("Devedores").Select
        Range("A1").Select

        'Calculando tempo para conclusão do VBA
        tempo_final = Timer

        tempo_total = Round(tempo_final - tempo, 1)
        
        num_registros = n_ll - last_line + 1

        MsgBox "Execução finalizada em " & tempo_total & "s! Foram inseridos " & num_registros & " registros com êxito.", vbSystemModal, "INSERÇÃO FINALIZADA COM ÊXITO"
    Else
        MsgBox "Favor verificar o valor da célula 'B2'. A macro só poderá ser executada caso seu valor seja SIM.", vbInformation, "MACRO BLOQUEADA"
    End If
    
    Application.ScreenUpdating = True

End Sub
