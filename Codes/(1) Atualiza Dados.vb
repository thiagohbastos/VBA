Sub Atualiza_Dados()
'AUTOR: THIAGO BASTOS / git: thiagohbastos

Dim PathName As String
Dim ControlFile As String
Dim TabName As String
Dim Tabela As String
Dim NumPlan As Integer

Application.ScreenUpdating = False

'Desocultando abas
Sheets("Devedores").Visible = True
Sheets("Credores").Visible = True
Sheets("ERROS DE SAQUE").Visible = True
Sheets("ERROS DE DEPÓSITO").Visible = True
Sheets("Config VBA").Visible = True

NumPlan = 2
ControlFile = ActiveWorkbook.Name

'Abrindo planilhas e extraindo dados das tabelas (LOOP)

Do Until Sheets("Config VBA").Cells(NumPlan, 1).Value = ""

    PathName = Sheets("Config VBA").Range("C" & NumPlan).Value
    TabName = Sheets("Config VBA").Range("D" & NumPlan).Value
    Tabela = Sheets("Config VBA").Range("E" & NumPlan).Value
    Nomeplan = Sheets("Config VBA").Range("F" & NumPlan).Value
    
    Sheets(TabName).Range("A1:ZZ1048576").Clear
    
    Workbooks.Open Filename:=PathName
    Sheets(TabName).Select
    ActiveSheet.ListObjects(Tabela).Range.Copy
    Windows(ControlFile).Activate
    Sheets(TabName).Select
    Range("A1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Workbooks(Nomeplan).Close SaveChanges:=False

NumPlan = NumPlan + 1

Loop

'Ocultando abas
Sheets("Devedores").Visible = False
Sheets("Credores").Visible = False
Sheets("ERROS DE SAQUE").Visible = False
Sheets("ERROS DE DEPÓSITO").Visible = False
Sheets("Config VBA").Visible = False

Sheets("Links").Select

    "K:\GSAS\09 - Coordenacao Gestao Numerario\03 Gestao de Diferencas\(TESTE) LANÇAMENTOS ERRO SAQUE.pdf" _
    , Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
    :=False, OpenAfterPublish:=False

Application.ScreenUpdating = True

End Sub