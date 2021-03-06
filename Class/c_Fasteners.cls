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

    'd�fini les propri�t�s transf�r�es � la m�thode
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
    
    'Ajoute un objet � la collection
    mcol.Add objNewMember, nom
        
    'Renvoi l'objet cr�e
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



