VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_ckPart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'c_ckPart
'Objet part
'avec ses propriétés controlées par le check 3D

Private mvarNom As String               'Nom du part (partnumber)
Private mvarPart As part                'Part du part
Private mvarChpSource As String         'Valeur de l'attibut "source"
Private mvarChpDescr As String          'Valeur de l'attribut "DescriptionRef"
Private mMass As Double                 'Masse du CatPart
Private mMat As Material                'Matière du CattPart
Private mColProperties As c_Properties  'Collection des attibuts Grilles (x......)
Private mcolHBodies As c_Hbodies        'Collection des HybridBodies
Private mColFasteners As c_Fasteners    'Collection des fasteners

Public Property Let nom(ByVal vData As String)
    mvarNom = vData
End Property

Public Property Get nom() As String
    nom = mvarNom
End Property

Property Let ctPart(ByVal vPart As part)
    Set mvarPart = vPart
End Property

Public Property Get ctPart() As part
    Set ctPart = mvarPart
End Property

Public Property Let chpSource(ByVal vData As String)
    mvarChpSource = vData
End Property

Public Property Get chpSource() As String
    chpSource = mvarChpSource
End Property

Public Property Let chpDescr(ByVal vData As String)
    mvarChpDescr = vData
End Property

Public Property Get chpDescr() As String
    chpDescr = mvarChpDescr
End Property

Public Property Let Masse(ByVal vDbl As Double)
    mMass = vDbl
End Property

Public Property Get Masse() As Double
    Masse = mMass
End Property

Public Property Let Matiere(ByVal vMat As Material)
    Set mMat = vMat
End Property

Public Property Get Matiere() As Material
   Set Matiere = mMat
End Property

Public Property Get NomMatiere() As String
    'Renvoi le nom de la matière
    On Error Resume Next
    NomMatiere = Matiere.Name
    If Err.Number <> 0 Then
        Err.Clear
        NomMatiere = "Non renseignée"
    End If
End Property

Public Property Let xProperties(ByVal vCol As c_Properties)
    Set mColProperties = vCol
End Property

Public Property Get xProperties() As c_Properties
    Set xProperties = mColProperties
End Property

Public Property Let xBodies(ByVal vCol As c_Hbodies)
    Set mcolHBodies = vCol
End Property

Public Property Get xBodies() As c_Hbodies
    Set xBodies = mcolHBodies
End Property

Public Property Let Fasteners(ByVal vCol As c_Fasteners)
    Set mColFasteners = vCol
End Property

Public Property Get Fasteners() As c_Fasteners
    Set Fasteners = mColFasteners
End Property

