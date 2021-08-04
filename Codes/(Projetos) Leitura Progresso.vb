Sub Leitura_Projetos()

'AUTOR: THIAGO BASTOS / git: thiagohbastos

Application.ScreenUpdating = False

'---------
'Variáveis
'---------
Dim pasta As Object
Dim caminho_pasta As String
Dim matricula As String
Dim cod_projeto As Integer
Dim dtafim_projeto As String
Dim progresso_projeto As Double
Dim controlfile As String
Dim lastline As Integer

controlfile = ActiveWorkbook.Name

Sheets(2).Range("A2:ZZ1048576").Clear

'---------------------------------------
'ESTRUTURA DE REPETIÇÃO DAQUI PARA BAIXO
'---------------------------------------

caminho_pasta = "C:\Users\thiag\Desktop\TESTE\"

Set pasta = CreateObject("Scripting.FileSystemObject").getfolder(caminho_pasta)

For Each planilha In pasta.Files

    Workbooks.Open (planilha)
    
    For Each aba In Workbooks(planilha.Name).Sheets
    
        matricula = Workbooks(planilha.Name).Sheets(aba.Name).Range("C2").Value
        cod_projeto = Workbooks(planilha.Name).Sheets(aba.Name).Range("C3").Value
        dtafim_projeto = Workbooks(planilha.Name).Sheets(aba.Name).Range("C4").Value
        progresso_projeto = Workbooks(planilha.Name).Sheets(aba.Name).Range("C5").Value
        
        Windows(controlfile).Activate
        ThisWorkbook.Sheets(2).Select
        
        lastline = ThisWorkbook.Sheets(2).Range("A1048576").End(xlUp).Row + 1
        
        ThisWorkbook.Sheets(2).Cells(lastline, 1).Value = matricula
        ThisWorkbook.Sheets(2).Cells(lastline, 2).Value = cod_projeto
        ThisWorkbook.Sheets(2).Cells(lastline, 3).Value = dtafim_projeto
        ThisWorkbook.Sheets(2).Cells(lastline, 4).Value = progresso_projeto
    
    Next
    
    Workbooks(planilha.Name).Close

Next

'--------------------
'FORMATAÇÃO DOS DADOS
'--------------------

Columns("B:B").Select
Selection.NumberFormat = "0"
Columns("C:C").Select
Selection.NumberFormat = "d/m/yyyy"
Columns("D:D").Select
Selection.NumberFormat = "0.0%"

Application.ScreenUpdating = True

End Sub