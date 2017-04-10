VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_Fasteners"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'c_Fasteners
'Collection des Fasteners d'une grille

Private mcol As Collection

Public Function Add(nom As String, _
                    Optional Comments As String = "", _
                    Optional FastDiam As String = "", _
                    Optional Xe As String = "", _
                    Optional Ye As String = "", _
                    Optional Ze As String = "", _
                    Optional Xdir As String = "", _
                    Optional Ydir As String = "", _
                    Optional Zdir As String = "", _
                    Optional X As Double = 0, _
                    Optional Y As Double = 0, _
                    Optional Z As Double = 0) As c_Fastener
                    
    Dim objNewMember As c_Fastener
    Set objNewMember = New c_Fastener

    'défini les propriétés transférées à la méthode
    objNewMember.nom = nom
    objNewMember.Comments = Comments
    objNewMember.FastDiam = FastDiam
    objNewMember.Xe = Xe
    objNewMember.Ye = Ye
    objNewMember.Ze = Ze
    objNewMember.Xdir = Xdir
    objNewMember.Ydir = Ydir
    objNewMember.Zdir = Zdir
    objNewMember.X = X
    objNewMember.Y = Y
    objNewMember.Z = Z
    
    'Ajoute un objet à la collection
    mcol.Add objNewMember, nom
        
    'Renvoi l'objet crée
    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As c_Fastener
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

Private Sub Class_Initialize()
    'crée la collection lorsque la classe est créée
    Set mcol = New Collection
End Sub

Private Sub Class_Terminate()
    'supprime la collection lorsque la classe est fermée
    Set mcol = Nothing
End Sub



