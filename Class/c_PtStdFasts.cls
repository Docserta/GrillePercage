VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_PtStdFasts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'c_PtStdFasts
Private mcol As Collection

Public Function Add(nPtA As String, _
                    nPtB As String, _
                    nstd As String, _
                    FxPtA As c_Pt, _
                    FxPtb As c_Pt, _
                    Fastener As c_Fastener) As c_PtStdFast

    Dim objNewMember As c_PtStdFast
    Set objNewMember = New c_PtStdFast

    objNewMember.nPtA = nPtA
    objNewMember.nPtB = nPtB
    objNewMember.nstd = nstd
    objNewMember.FxPtA = FxPtA
    objNewMember.FxPtb = FxPtb
    objNewMember.Fastener = Fastener

    mcol.Add objNewMember, nPtA

    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As c_PtStdFast
  Set Item = mcol(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mcol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    mcol.Remove vntIndexKey
End Sub

Private Sub Class_Initialize()
    Set mcol = New Collection
End Sub

Private Sub Class_Terminate()
    Set mcol = Nothing
End Sub




