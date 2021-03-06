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
'Collection des mesures des coordonn�es r�elles des fasteners

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
                    
    'Cr�e un nouvel objet
    Dim objNewMember As c_mesure
    Set objNewMember = New c_mesure

    'd�fini les propri�t�s transf�r�es � la m�thode
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
        
    'Ajoute un objet � la collection
    mcol.Add objNewMember, nom
        
    'Renvoi l'objet cr�e
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
