Attribute VB_Name = "Módulo11"
Sub AgregarData()
'
' AgregarData Macro
' Macro grabada el 29/05/2002 por wg006
'
    On Error GoTo errorTranf
    Sheets("Agregar Solicitud de Censo").Select
    Range("c5:c14").Select
    Selection.Copy
    Sheets("BD Ingreso Llave-Alicate").Select
    Range("A1").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Range("a1").Select
    Selection.PasteSpecial Paste:=xlValue, Operation:=xlNone, SkipBlanks:=False _
        , Transpose:=True
         Range("A2").Select
'    ActiveWindow.TabRatio = 0.9
    Selection.Sort Key1:=Range("A2"), Order1:=xlAscending, Header:=xlGuess, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
    Sheets("Asignación Llave - Alicate").Select
    Application.CutCopyMode = False
    Range("c5:c13").Select
    Selection.ClearContents
    Range("c15").Select
    Selection.ClearContents

    Range("c5").Select
    Exit Sub
errorTranf:
    Range("A1").Select
    ActiveCell.Offset(1, 0).Range("a1").Select
    Resume Next
End Sub



