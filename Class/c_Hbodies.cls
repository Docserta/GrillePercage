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
    'Cr�e un nouvel objet
    Dim objNewMember As c_Hbody
    Set objNewMember = New c_Hbody

    'd�fini les propri�t�s transf�r�es � la m�thode
    objNewMember.nom = nom
    objNewMember.ShowStatus = ShowStatus
    objNewMember.HBody = HBody
        
    'Ajoute un objet � la collection
    mcol.Add objNewMember, nom
        
    'Renvoi l'objet cr�e
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
    'est utilis� pour supprimer un �l�ment de la collection
    'vntIndexKey contient l'index ou la cl� de la collection,
    'donc il est d�clar� comme Variant
    'Syntax: x.Remove(xyz)
 mcol.Remove vntIndexKey
End Sub

Private Sub Class_initialize()
    'cr�e la collection lorsque la classe est cr��e
    Set mcol = New Collection
End Sub

Private Sub Class_Terminate()
    'supprime la collection lorsque la classe est ferm�e
    Set mcol = Nothing
End Sub



