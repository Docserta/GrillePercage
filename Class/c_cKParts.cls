VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_cKParts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'c_cKParts
'Collection des parts d'une grille assemblée avec les attributs nécessaires au Check 3D
'variable local pour contenir la collection
Private mcol As Collection

Public Function Add(nom As String, _
                    ctPart As part, _
                    Optional chpSource As String = "", _
                    Optional chpDescr As String = "", _
                    Optional Masse As Double = 0, _
                    Optional Matiere As Material = Nothing, _
                    Optional Params As c_Properties = Nothing, _
                    Optional xBodies As c_Hbodies = Nothing, _
                    Optional Fasteners As c_Fasteners = Nothing) As c_ckPart
                    
    'Crée un nouvel objet
    Dim objNewMember As c_ckPart
    Set objNewMember = New c_ckPart

    'défini les propriétés transférées à la méthode
    objNewMember.nom = nom
    objNewMember.ctPart = ctPart
    objNewMember.chpSource = chpSource
    objNewMember.chpDescr = chpDescr
    objNewMember.Masse = Masse
    objNewMember.Matiere = Matiere
    objNewMember.xProperties = Params
    objNewMember.xBodies = xBodies
    objNewMember.Fasteners = Fasteners
    
    'Ajoute un objet à la collection
    mcol.Add objNewMember, nom
        
    'Renvoi l'objet crée
    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As c_ckPart
  Set Item = mcol(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mcol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
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



