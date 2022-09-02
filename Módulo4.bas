Attribute VB_Name = "Módulo4"
Sub AgregarData3()
'
' AgregarData Macro
' Macro grabada el 29/05/2002 por wg006
'
    On Error GoTo errorTranf
    Sheets("Cerrar Censo de Proyecto").Select
    Range("c4:c30").Select
    Selection.Copy
    Sheets("BD Conclusión de Censo de Proye").Select
    Range("A1").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Range("a1").Select
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:=False _
        , Transpose:=True
         Range("A2").Select
    Selection.Sort Key1:=Range("A2"), Order1:=xlAscending, Header:=xlGuess, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
    Sheets("Cerrar Censo de Proyecto").Select
    Application.CutCopyMode = False
    Range("c16:c29").Select
    Selection.ClearContents
    Range("c16").Select
    Exit Sub
errorTranf:
    Range("A1").Select
    ActiveCell.Offset(1, 0).Range("a1").Select
    Resume Next
End Sub





