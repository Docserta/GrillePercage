Attribute VB_Name = "Module1"
Sub catmain()


Dim RenSets As Collection
Dim Renset As Object
Dim test1 As Boolean
Dim test2 As Boolean


test1 = True
test2 = False

Set RenSets = New Collection

'Set Renset = test1
RenSets.Add test1, "PtA"
'Set Renset = test2
RenSets.Add test2, "PTB"

If RenSets.Item("PtA") = True Then
    MsgBox "PtA - " & RenSets.Item("PtA")
End If
If RenSets.Item("PTB") = True Then
    MsgBox "PTB -" & RenSets.Item("PTB")
End If


End Sub
