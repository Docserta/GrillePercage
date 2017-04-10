VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FRM_Agrafage 
   Caption         =   "Type d'épinglage"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7650
   OleObjectBlob   =   "FRM_Agrafage.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FRM_Agrafage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit
Private Sub BtnAnnul_Click()
    Me.Hide
    Me.ChB_OkAnnule = False
End Sub

Private Sub BtnOK_Click()
Dim Erreur As Boolean
    Me.ChB_OkAnnule = True
    Erreur = False

If Not Erreur Then
    Me.Hide
End If
End Sub

Private Sub BtnSel_Click()
'Selection des Points A à percer.
Dim GrilleTemp As New c_PartGrille
GrilleTemp.GrilleSelection.Clear
Dim i As Long, j As Long
Dim Nb_Pt_Sel As Long
Dim NomUdfEC As String 'Nom de l'UDF en cours de traitement
Dim DiamUdfSel As String 'stocke le diamètre de perçage avion des UDF sélectionnées
ReDim Tab_Select_Points(2, 0)
Me.Hide
SelectPTA GrilleTemp

'Vérification que la sélection n'est pas vide
    Nb_Pt_Sel = GrilleTemp.GrilleSelection.Count

    If Nb_Pt_Sel = 0 Then
       MsgBox "Vous n'avez pas selectionné de points dans PointsA"
       Exit Sub
    End If
'Vérification que ce soient des Point A
'Ils doivent appartenir au Set "pointsA"
'#############
'
' a faire
'
'###############

'Ajout
    For i = 1 To Nb_Pt_Sel
        If GrilleTemp.GrilleSelection.Item(i).Type = "HybridShape" Then
            ReDim Preserve Tab_Select_Points(2, i - 1)
            Tab_Select_Points(0, i - 1) = GrilleTemp.GrilleSelection.Item(i).Value.Name
            'Recupération du nom de la ligne STD
            Tab_Select_Points(2, i - 1) = GrilleTemp.GrilleSelection.Item(i).Value.Element1.DisplayName
            NomUdfEC = Right(Tab_Select_Points(0, i - 1), Len(Tab_Select_Points(0, i - 1)) - InStr(1, CStr(Tab_Select_Points(0, i - 1)), "-", vbTextCompare))
        End If
    Next

'Ajout du nom des points sélectionnés dans le formulaire
    FRM_Agrafage.LB_SelTrous.ColumnCount = 3
    FRM_Agrafage.LB_SelTrous.List = TranspositionTabl(Tab_Select_Points)
    
    Set GrilleTemp = Nothing
    Me.Show

End Sub


Private Sub Logo_eXcent_Click()
'Chargement de la boite eXcent
    Load Frm_eXcent
    Frm_eXcent.Tbx_Version = VMacro
    Frm_eXcent.Show

    Unload Frm_eXcent
End Sub

Private Sub UserForm_Initialize()
Dim i As Long
 'Rempli la liste déroulantes agrafes
    Dim NumAgrafesTemp()
        NumAgrafesTemp = CollMachines.ListAgrafes
    For i = 0 To UBound(NumAgrafesTemp, 2)
        FRM_Agrafage.CBL_NumAgrafe.AddItem (NumAgrafesTemp(0, i))
    Next
End Sub
