Attribute VB_Name = "Módulo9"
Sub ArregAudTar()
Attribute ArregAudTar.VB_Description = "Macro grabada el 08/11/2006 por Winston J. Guzmán Z"
Attribute ArregAudTar.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ArregAudTar Macro
' Macro grabada el 08/11/2006 por Winston J. Guzmán Z
'

'
    Range("A1:H1").Select
    Selection.Copy
    Range("A2").Select
    ActiveSheet.Paste
    Range("N1:Q1").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("N2").Select
    ActiveSheet.Paste
    Range("S1:T1").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("S2").Select
    ActiveSheet.Paste
    Range("AB1").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AB2").Select
    ActiveSheet.Paste
    Range("AC2").Select
    Selection.End(xlToRight).Select
    Range("EZ1").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("EZ2").Select
    ActiveSheet.Paste
    Range("EZ1").Select
    Application.CutCopyMode = False
    Selection.EntireRow.Delete
End Sub
