VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_Coord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'c_Coord

Private mvarX As Double
Private mvarY As Double
Private mvarZ As Double

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

