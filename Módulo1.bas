Attribute VB_Name = "Módulo1"
Sub QuitaryPonerKey()
Attribute QuitaryPonerKey.VB_ProcData.VB_Invoke_Func = " \n14"
'
' QuitaryPonerKey Macro
'

'
    ActiveSheet.Unprotect (XXXXX)
    Cells.Select
    Selection.Locked = True
    Selection.FormulaHidden = False
    ActiveSheet.Protect (XXXXX)
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
End Sub
