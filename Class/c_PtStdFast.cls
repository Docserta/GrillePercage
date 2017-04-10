VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_PtStdFast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'c_PtStdFast
Private mvarnPtA As String
Private mvarnPtB As String
Private mvarnStd As String
Private mvarFxPtA As c_Pt
Private mvarFxPtB As c_Pt
Private mvarFast As c_Fastener

Public Property Let nPtA(ByVal vData As String)
     mvarnPtA = vData
End Property

Public Property Get nPtA() As String
    nPtA = mvarnPtA
End Property

Public Property Let nPtB(ByVal vData As String)
     mvarnPtB = vData
End Property

Public Property Get nPtB() As String
    nPtB = mvarnPtB
End Property

Public Property Let nstd(ByVal vData As String)
    mvarnStd = vData
End Property

Public Property Get nstd() As String
    nstd = mvarnStd
End Property

Public Property Let FxPtA(vCol As c_Pt)
    Set mvarFxPtA = vCol
End Property

Public Property Get FxPtA() As c_Pt
    Set FxPtA = mvarFxPtA
End Property

Public Property Let FxPtb(vCol As c_Pt)
    Set mvarFxPtB = vCol
End Property

Public Property Get FxPtb() As c_Pt
    Set FxPtb = mvarFxPtB
End Property

Public Property Let Fastener(vCol As c_Fastener)
    Set mvarFast = vCol
End Property

Public Property Get Fastener() As c_Fastener
    Set Fastener = mvarFast
End Property
