VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_DefBague"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'c_DefBague

Private mvarRef As String
Private mvarDnom As Double
Private mvarMat As String
Private mvarD1 As Double
Private mvarD2 As Double
Private mvarD3 As Double
Private mvarL1 As Double
Private mvarL2 As Double
Private mvarNomFic As String

Public Property Let Ref(ByVal vData As String)
    mvarRef = vData
End Property

Public Property Get Ref() As String
    Ref = mvarRef
End Property

Public Property Let DNom(ByVal vData As Double)
    mvarDnom = vData
End Property

Public Property Get DNom() As Double
    DNom = mvarDnom
End Property

Public Property Let Mat(ByVal vData As String)
    mvarMat = vData
End Property

Public Property Get Mat() As String
    Mat = mvarMat
End Property

Public Property Let D1(ByVal vData As Double)
    mvarD1 = vData
End Property

Public Property Get D1() As Double
    D1 = mvarD1
End Property

Public Property Let D2(ByVal vData As Double)
    mvarD2 = vData
End Property

Public Property Get D2() As Double
    D2 = mvarD2
End Property

Public Property Let D3(ByVal vData As Double)
    mvarD3 = vData
End Property

Public Property Get D3() As Double
    D3 = mvarD3
End Property

Public Property Let L1(ByVal vData As Double)
    mvarL1 = vData
End Property

Public Property Get L1() As Double
    L1 = mvarL1
End Property

Public Property Let L2(ByVal vData As Double)
    mvarL2 = vData
End Property

Public Property Get L2() As Double
    L2 = mvarL2
End Property

Public Property Let NomFic(ByVal vData As String)
    mvarNomFic = vData
End Property

Public Property Get NomFic() As String
    NomFic = mvarNomFic
End Property

