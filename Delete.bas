Attribute VB_Name = "Module1"
Sub Очищение()
Attribute Очищение.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Очищение Макрос
'

'
    Range("A2:D5").Select
    Selection.ClearContents
    Range("A2").Select
End Sub
