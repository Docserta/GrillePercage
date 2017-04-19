VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_DefBagues"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'c_DefBagues collection des Bagues Spécifiques
Private mcol As Collection

Public Function Add(Ref As String, _
                    DNom As Double, _
                    Mat As String, _
                    D1 As Double, _
                    D2 As Double, _
                    D3 As Double, _
                    L1 As Double, _
                    L2 As Double, _
                    NomFic As String) As c_DefBague
    
    Dim objNewMember As c_DefBague
    Set objNewMember = New c_DefBague
    
    objNewMember.Ref = Ref
    objNewMember.DNom = DNom
    objNewMember.Mat = Mat
    objNewMember.D1 = D1
    objNewMember.D2 = D2
    objNewMember.D3 = D3
    objNewMember.L1 = L1
    objNewMember.L2 = L2
    objNewMember.NomFic = NomFic

    mcol.Add objNewMember, Ref

    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As c_DefBague
  Set Item = mcol(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mcol.Count
End Property

Public Property Get Items() As Collection
    Set Items = mcol
End Property

Public Sub Remove(vntIndexKey As Variant)
    mcol.Remove vntIndexKey
End Sub

Private Sub Class_initialize()
    Set mcol = New Collection
End Sub

Private Sub Class_Terminate()
    Set mcol = Nothing
End Sub




