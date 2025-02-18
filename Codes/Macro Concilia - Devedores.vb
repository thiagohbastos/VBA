'AUTOR: THIAGO BASTOS (B042786) 29/08/2022
'Macro responsável por regularizar os lançamentos da faixa Devedores
Sub Regulariza_Registros()

    Application.ScreenUpdating = False

    Sheets("Devedores").Select

    If Sheets("Devedores").Range("C2").Value = "SIM" Then
    
        Dim tempo As Double
        Dim tempo_final As Double
        Dim tempo_total As Double
        tempo = Timer


        '------------ LIMPANDO REGISTROS LANÇADOS ANTERIORMENTE -------------
        Sheets("CONFIG").Visible = True
        Sheets("CONFIG").Columns("C:U").Delete
        ll_regularizacoes = Sheets("REGULARIZAÇÕES").Range("A1048576").End(xlUp).Row + 1
        Sheets("REGULARIZAÇÕES").Range("A2:J" & ll_regularizacoes).Delete


        '--------- SELECIONANDO OS NOVOS REGISTROS PARA LANÇAMENTO ----------
        Sheets("CAB D-1").Select

        On Error Resume Next
        ActiveSheet.ShowAllData
        On Error GoTo 0

        ActiveSheet.ListObjects("Tabela_SQLGDNP_GNU_TI_DIFERENCAS_CAB").Range. _
        AutoFilter Field:=17, Criteria1:="=GNU*"

        Range("Tabela_SQLGDNP_GNU_TI_DIFERENCAS_CAB[#All]").Select
        Selection.Copy


        '------------------- TRATANDO OS NOVOS REGISTROS --------------------
        Sheets("CONFIG").Select
        Range("C1").Select
        ActiveSheet.Paste
        Application.CutCopyMode = False

        Range("C:C,E:E,G:G,I:I,K:K,M:V").Select
        Selection.Delete Shift:=xlToLeft

        ll_config = Range("C1048576").End(xlUp).Row


        '----- ORGANIZANDO A TABELA PRINCIPAL PARA RECEBER OS REGISTROS -----
        Sheets("Devedores").Select

        On Error Resume Next
        Range("Tabela3[[#Headers],[Transportadora]]").Select
        ActiveSheet.ShowAllData
        On Error GoTo 0

        ActiveWorkbook.Worksheets("Devedores").ListObjects("Tabela3").Sort.SortFields. _
        Clear
        ActiveWorkbook.Worksheets("Devedores").ListObjects("Tabela3").Sort.SortFields. _
        Add Key:=Range("Tabela3[[#All],[Data Diferença]]"), SortOn:=xlSortOnValues _
        , Order:=xlAscending, DataOption:=xlSortNormal
        With ActiveWorkbook.Worksheets("Devedores").ListObjects("Tabela3").Sort
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With

        ll_devedores = Range("J1048576").End(xlUp).Row
        erros = 0
        exitos = 0

        '------------------ INICIANDO LOOP DE REGULARIZAÇÕES ----------------
        For linha = 2 To ll_config

            'DEFINIÇÃO DE VARIÁVEIS DOS LANÇAMENTOS
            doc = Sheets("CONFIG").Range("E" & linha).Value
            tesouraria = Sheets("CONFIG").Range("C" & linha).Value
            valor = Sheets("CONFIG").Range("G" & linha).Value
            
            Data = Sheets("CONFIG").Range("D" & linha).Value
            Data = CDate(Data)
            dif = DateDiff("d", Date, Data)
            Data = DateAdd("d", dif, Date)

            atm = Sheets("CONFIG").Range("F" & linha).Value
            If atm < 1000 Then
                x = "0" & CStr(atm)
                atm = x
            End If
            
            'LIMPANDO FILTROS DA TABELA PRINCIPAL NOVAMENTE
            On Error Resume Next
            Range("Tabela3[[#Headers],[Transportadora]]").Select
            ActiveSheet.ShowAllData
            On Error GoTo 0

            'INSERINDO NOVOS FILTROS
            ActiveSheet.ListObjects("Tabela3").Range.AutoFilter Field:=2, Criteria1:=tesouraria
            ActiveSheet.ListObjects("Tabela3").Range.AutoFilter Field:=8, Criteria1:=atm
            ActiveSheet.ListObjects("Tabela3").Range.AutoFilter Field:=12, Criteria1:=">=" & valor

            'LINHA DO LANÇAMENTO FILTRADO MAIS RECENTE, QUANDO APLICÁVEL
            linha_registro = Range("Tabela3[[#Headers],[Dependência]]").End(xlDown).Row

            'ÚLTIMA LINHA EM BRANCO DA ABA 'REGULARIZAÇÕES'
            ll_regularizacoes = Sheets("REGULARIZAÇÕES").Range("A1048576").End(xlUp).Row + 1

            'INSERINDO DADOS DO LANÇAMENTO ATUAL
            Sheets("REGULARIZAÇÕES").Range("A" & ll_regularizacoes).Value = Data
            Sheets("REGULARIZAÇÕES").Range("B" & ll_regularizacoes).Value = tesouraria
            Sheets("REGULARIZAÇÕES").Range("C" & ll_regularizacoes).Value = atm
            Sheets("REGULARIZAÇÕES").Range("D" & ll_regularizacoes).Value = doc
            Sheets("REGULARIZAÇÕES").Range("E" & ll_regularizacoes).Value = valor
            Sheets("REGULARIZAÇÕES").Range("G" & ll_regularizacoes).Value = Date

            'CASO NÃO SEJA ENCONTRADA SOBRA QUE ATENDA AOS FILTROS
            If linha_registro = ll_devedores Then
                Sheets("REGULARIZAÇÕES").Range("F" & ll_regularizacoes).Value = "NÃO ENCONTRADO"
                erros = erros + 1
            
            'CASO SEJA ENCONTRADA SOBRA QUE ATENDA AOS FILTROS
            Else

                vlr_anterior = Sheets("Devedores").Range("K" & linha_registro).Value
                Sheets("Devedores").Range("K" & linha_registro).Value = vlr_anterior + valor
                Sheets("Devedores").Range("O" & linha_registro).Value = "Compensação"
                Sheets("Devedores").Range("P" & linha_registro).Value = Data

                Sheets("REGULARIZAÇÕES").Range("F" & ll_regularizacoes).Value = "EXITO"
                exitos = exitos + 1

            End If

        Next
        
        '-------------- ORGANIZANDO PLANILHA PARA FINALIZAR -------------
        Range("B2:C2").Value = "NÃO"

        Sheets("CAB D-1").Select

        On Error Resume Next
        ActiveSheet.ShowAllData
        On Error GoTo 0

        Range("A1").Select

        Sheets("Devedores").Select
        Range("Tabela3[[#Headers],[Transportadora]]").Select
        ActiveSheet.ShowAllData

        Range("A1").Select
        
        Sheets("CONFIG").Visible = False
        
        'Calculando tempo para conclusão do VBA
        tempo_final = Timer
        
        tempo_total = Round(tempo_final - tempo, 1)

        MsgBox "Execução finalizada em " & tempo_total & "s! Foram regularizados " & exitos & " registros com SUCESSO, e um total de " & erros & " FALHAS. Checar a aba 'REGULARIZAÇÕES' para obter detalhes dos lançamentos.", vbSystemModal, "EXECUÇÃO FINALIZADA COM ÊXITO"
    Else
        MsgBox "Favor verificar o valor da célula 'C2'. A macro só poderá ser executada caso seu valor seja SIM.", vbInformation, "MACRO BLOQUEADA"
    End If

        Application.ScreenUpdating = True

End Sub
