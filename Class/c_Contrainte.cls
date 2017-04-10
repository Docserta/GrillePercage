VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_Contrainte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'c_Contrainte
'Objet contrainte d'assemblage
'avec ses propriétés : Statut, ReferenceType

Option Explicit
Private mvarNom As String
Private mvarStatut As CatConstraintStatus
Private mvarRefType As CatConstraintRefType

Public Property Let nom(ByVal vData As String)
    mvarNom = vData
End Property

Public Property Get nom() As String
    nom = mvarNom
End Property

Public Property Let statut(ByVal vData As CatConstraintStatus)
    mvarStatut = vData
End Property

Public Property Get statut() As CatConstraintStatus
    statut = mvarStatut
End Property

Public Property Let RefType(ByVal vData As CatConstraintRefType)
    mvarRefType = vData
End Property

Public Property Get RefType() As CatConstraintRefType
    RefType = mvarRefType
End Property

