Attribute VB_Name = "Módulo1"
Sub QuitaryPonerKey()
Attribute QuitaryPonerKey.VB_ProcData.VB_Invoke_Func = " \n14"
'
' QuitaryPonerKey Macro
'

'
    ActiveSheet.Unprotect (DragonSony2010)
    Cells.Select
    Selection.Locked = True
    Selection.FormulaHidden = False
    ActiveSheet.Protect (DragonSony2010)
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
End Sub
