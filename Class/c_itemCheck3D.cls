VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_itemCheck3D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'c_Check3D
Private mCote As String
Private mCheck As Boolean
Private mStatut As String
Private mComment As String

Public Property Let Cote(ByVal vData As String)
    mCote = vData
End Property

Public Property Get Cote() As String
    Cote = mCote
End Property

Public Property Let Check(ByVal vBool As Boolean)
    mCheck = vBool
End Property

Property Get Check() As Boolean
    Check = mCheck
End Property

Public Property Let statut(ByVal vData As String)
    mStatut = vData
End Property

Property Get statut() As String
    statut = mStatut
End Property

Public Property Let Comment(ByVal vData As String)
    mComment = vData
End Property

Property Get Comment() As String
    Comment = mComment
End Property

