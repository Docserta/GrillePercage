VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_CartTxts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'c_CartTxts collection des textes du cartouche AIRBUS et des attributs correspondants
Private mcol As Collection

Public Function Add(nom As String, _
        Optional Attrib As String = "", _
        Optional Valeur As String = "") As c_carttxt
    
    Dim objNewMember As c_carttxt
    Set objNewMember = New c_carttxt
    
    objNewMember.nom = nom
    objNewMember.Attrib = Attrib
    objNewMember.Valeur = Valeur

    mcol.Add objNewMember, nom

    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As c_carttxt
  Set Item = mcol(vntIndexKey)
End Property

Public Property Get Items() As Collection
    Set Items = mcol
End Property

Public Property Get Count() As Long
    Count = mcol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    mcol.Remove vntIndexKey
End Sub

Private Sub Class_Initialize()
    Set mcol = New Collection
End Sub

Private Sub Class_Terminate()
    Set mcol = Nothing
End Sub




