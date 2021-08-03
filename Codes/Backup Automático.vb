Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)

'ThisWorkbook.SaveCopyAs "T:\gnu\DIFERENÇAS\GESTÃO DEVEDORES - 6864-3.xlsm"

ThisWorkbook.SaveCopyAs "K:\GSAS\09 - Coordenacao Gestao Numerario\03 Gestao de Diferencas\DEVEDORES E CREDORES\BACKUPS\BACKUP_DEVEDORES" _
& Day(Date) & "-" & Month(Date) & "-" & Year(Date) & "_" & Hour(Now) & "h" & Minute(Now) & "min" & ".xlsm"

End Sub