VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_Pt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'c_pt
Private mvarNom As String
Private mvarCrd As c_Coord
Private mvarParent As String

Public Property Let nom(ByVal vData As String)
    mvarNom = vData
End Property

Public Property Get nom() As String
    nom = mvarNom
End Property

Public Property Let Crd(ByVal vCrd As c_Coord)
    Set mvarCrd = vCrd
End Property

Public Property Get Crd() As c_Coord
    Set Crd = mvarCrd
End Property

Public Property Let Parent(ByVal vData As String)
    mvarParent = vData
End Property

Public Property Get Parent() As String
    Parent = mvarParent
End Property





