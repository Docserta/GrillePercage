VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_Fastener"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'c_Fastener
' infos sur les refférences externes isolées

Private mvarNom As String       ' le nom réel
Private mvarComments As String  ' le paramétre comment
Private mvarFastDiam As String      ' le Diamètre de perçage
Private mvarXe As String
Private mvarYe As String
Private mvarZe As String
Private mvarXdir As String
Private mvarYdir As String
Private mvarZdir As String
Private mvarX As Double         'Coordonnée X réelle du fastener
Private mvarY As Double
Private mvarZ As Double

Public Property Let nom(ByVal vData As String)
    mvarNom = vData
End Property

Public Property Get nom() As String
    nom = mvarNom
End Property

Public Property Let Comments(ByVal vData As String)
    mvarComments = vData
End Property

Public Property Get Comments() As String
    Comments = mvarComments
End Property

Public Property Let FastDiam(ByVal vData As String)
    mvarFastDiam = vData
End Property

Public Property Get FastDiam() As String
    FastDiam = mvarFastDiam
End Property

Public Property Let Xe(ByVal vData As String)
    mvarXe = vData
End Property

Public Property Get Xe() As String
    Xe = mvarXe
End Property

Public Property Let Ye(ByVal vData As String)
    mvarYe = vData
End Property

Public Property Get Ye() As String
    Ye = mvarYe
End Property

Public Property Let Ze(ByVal vData As String)
    mvarZe = vData
End Property

Public Property Get Ze() As String
    Ze = mvarZe
End Property

Public Property Let Xdir(ByVal vData As String)
    mvarXdir = vData
End Property

Public Property Get Xdir() As String
    Xdir = mvarXdir
End Property

Public Property Let Ydir(ByVal vData As String)
    mvarYdir = vData
End Property

Public Property Get Ydir() As String
    Ydir = mvarYdir
End Property

Public Property Let Zdir(ByVal vData As String)
    mvarZdir = vData
End Property

Public Property Get Zdir() As String
    Zdir = mvarZdir
End Property

Public Property Let X(ByVal vDbl As Double)
    mvarX = vDbl
End Property

Public Property Get X() As Double
    X = mvarX
End Property

Public Property Let Y(ByVal vDbl As Double)
    mvarY = vDbl
End Property

Public Property Get Y() As Double
    Y = mvarY
End Property

Public Property Let Z(ByVal vDbl As Double)
    mvarZ = vDbl
End Property

Public Property Get Z() As Double
    Z = mvarZ
End Property
