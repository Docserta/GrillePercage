VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_mesures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'c_mesures
'Collection des mesures des coordonnées réelles des fasteners

Option Explicit
Private mcol As Collection

Public Function Add(nom As String, _
                    Optional X As Double = 0, _
                    Optional Y As Double = 0, _
                    Optional Z As Double = 0, _
                    Optional Xe As Double = 0, _
                    Optional Ye As Double = 0, _
                    Optional Ze As Double = 0, _
                    Optional NomPtA As String = "", _
                    Optional PtAX As Double = 0, _
                    Optional PtAY As Double = 0, _
                    Optional PtAZ As Double) As c_mesure
                    
    'Crée un nouvel objet
    Dim objNewMember As c_mesure
    Set objNewMember = New c_mesure

    'défini les propriétés transférées à la méthode
    objNewMember.nom = nom
    objNewMember.X = X
    objNewMember.Y = Y
    objNewMember.Z = Z
    objNewMember.Xe = Xe
    objNewMember.Ye = Ye
    objNewMember.Ze = Ze
    objNewMember.NomPtA = NomPtA
    objNewMember.PtAX = PtAX
    objNewMember.PtAY = PtAY
    objNewMember.PtAZ = PtAZ
        
    'Ajoute un objet à la collection
    mcol.Add objNewMember, nom
        
    'Renvoi l'objet crée
    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As c_mesure
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
