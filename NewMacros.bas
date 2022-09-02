Attribute VB_Name = "NewMacros"
Sub txt()
Attribute txt.VB_Description = "Macro grabada el 27/04/01 por w2k"
Attribute txt.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.txt"
'
' txt Macro
' Macro grabada el 27/04/01 por w2k
'
    Selection.MoveRight Unit:=wdCharacter, Count:=4
    Selection.Delete Unit:=wdCharacter, Count:=1
    Selection.MoveRight Unit:=wdCharacter, Count:=2
    Selection.Delete Unit:=wdCharacter, Count:=1
    Selection.MoveRight Unit:=wdCharacter, Count:=8
    Selection.Delete Unit:=wdCharacter, Count:=1
    Selection.MoveRight Unit:=wdCharacter, Count:=9
    Selection.Delete Unit:=wdCharacter, Count:=1
    Selection.MoveRight Unit:=wdCharacter, Count:=9
    Selection.Delete Unit:=wdCharacter, Count:=1
    Selection.MoveRight Unit:=wdCharacter, Count:=8
    Selection.Delete Unit:=wdCharacter, Count:=1
    Selection.MoveRight Unit:=wdWord, Count:=3
    Selection.Delete Unit:=wdCharacter, Count:=1
    Selection.MoveRight Unit:=wdWord, Count:=1
    Selection.Delete Unit:=wdCharacter, Count:=1
    Selection.MoveRight Unit:=wdWord, Count:=1
    Selection.Delete Unit:=wdCharacter, Count:=1
    Selection.MoveRight Unit:=wdWord, Count:=1
    Selection.Delete Unit:=wdCharacter, Count:=1
    Selection.MoveRight Unit:=wdWord, Count:=1
    Selection.Delete Unit:=wdCharacter, Count:=1
    Selection.MoveRight Unit:=wdWord, Count:=1
    Selection.Delete Unit:=wdCharacter, Count:=1
End Sub
