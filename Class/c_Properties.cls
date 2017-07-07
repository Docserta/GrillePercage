VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_Properties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'c_Properties collection des propriétés AIRBUS
Private mcol As Collection

Public Function Add(nom As String, Optional Valeur As String = "") As c_Property
    
    Dim objNewMember As c_Property
    Set objNewMember = New c_Property
    
    objNewMember.nom = nom
    objNewMember.Valeur = Valeur

    mcol.Add objNewMember, nom

    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As c_Property
  Set Item = mcol(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mcol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    mcol.Remove vntIndexKey
End Sub

Private Sub Class_initialize()
    Set mcol = New Collection
End Sub

Private Sub Class_Terminate()
    Set mcol = Nothing
End Sub



