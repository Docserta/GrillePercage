VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_mesure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'c_mesure
'Collection des mesures des coordonnées réelles des fasteners

Private mNom As String
Private mFastX As Double
Private mFastY As Double
Private mFastZ As Double
Private mFastXe As Double
Private mFastYe As Double
Private mFastZe As Double
Private mNomPtA As String
Private mPtAX As Double
Private mPtAY As Double
Private mPtAZ As Double

Public Property Let nom(ByVal vData As String)
    mNom = vData
End Property

Public Property Get nom() As String
    nom = mNom
End Property

Public Property Let X(ByVal vDbl As Double)
    mFastX = vDbl
End Property

Property Get X() As Double
    X = mFastX
End Property

Public Property Let Y(ByVal vDbl As Double)
    mFastY = vDbl
End Property

Property Get Y() As Double
    Y = mFastY
End Property

Public Property Let Z(ByVal vDbl As Double)
    mFastZ = vDbl
End Property

Property Get Z() As Double
    Z = mFastZ
End Property

Public Property Let Xe(ByVal vDbl As Double)
    mFastXe = vDbl
End Property

Property Get Xe() As Double
    Xe = mFastXe
End Property

Public Property Let Ye(ByVal vDbl As Double)
    mFastYe = vDbl
End Property

Property Get Ye() As Double
    Ye = mFastYe
End Property

Public Property Let Ze(ByVal vDbl As Double)
    mFastZe = vDbl
End Property

Property Get Ze() As Double
    Ze = mFastZe
End Property

Public Property Let NomPtA(ByVal vData As String)
    mNomPtA = vData
End Property

Public Property Get NomPtA() As String
    NomPtA = mNomPtA
End Property

Public Property Let PtAX(ByVal vDbl As Double)
    mPtAX = vDbl
End Property

Property Get PtAX() As Double
    PtAX = mPtAX
End Property

Public Property Let PtAY(ByVal vDbl As Double)
    mPtAY = vDbl
End Property

Property Get PtAY() As Double
    PtAY = mPtAY
End Property

Public Property Let PtAZ(ByVal vDbl As Double)
    mPtAZ = vDbl
End Property

Property Get PtAZ() As Double
    PtAZ = mPtAZ
End Property
