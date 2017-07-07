VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_Hbodies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'c_Hbodies
'Collection des HybridBodies Airbus de la part

Private mcol As Collection

Public Function Add(nom As String, _
                    Optional ShowStatus As CatVisPropertyShow = catVisPropertyNoShowAttr, _
                    Optional HBody As HybridBody = Nothing) As c_Hbody
    'Crée un nouvel objet
    Dim objNewMember As c_Hbody
    Set objNewMember = New c_Hbody

    'défini les propriétés transférées à la méthode
    objNewMember.nom = nom
    objNewMember.ShowStatus = ShowStatus
    objNewMember.HBody = HBody
        
    'Ajoute un objet à la collection
    mcol.Add objNewMember, nom
        
    'Renvoi l'objet crée
    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As c_Hbody
  Set Item = mcol(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mcol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    'est utilisé pour supprimer un élément de la collection
    'vntIndexKey contient l'index ou la clé de la collection,
    'donc il est déclaré comme Variant
    'Syntax: x.Remove(xyz)
 mcol.Remove vntIndexKey
End Sub

Private Sub Class_initialize()
    'crée la collection lorsque la classe est créée
    Set mcol = New Collection
End Sub

Private Sub Class_Terminate()
    'supprime la collection lorsque la classe est fermée
    Set mcol = Nothing
End Sub



