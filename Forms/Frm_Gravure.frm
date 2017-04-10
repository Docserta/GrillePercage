VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm_Gravure 
   Caption         =   "Ajout d'une Grille de perçage"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9405
   OleObjectBlob   =   "Frm_Gravure.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Frm_Gravure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Btn_Fic_Excel_Click()
'Documente les champs de la boite de dialogue à partir
'du fichier excel DSCGP
On Error GoTo Erreur
Dim GrilleEnCours As New c_DSCGP
    Me.TB_FicDSCGP = CATIA.FileSelectionBox("Selectionner le fichier DSCGP", "*.xls;*.xlsx", CatFileSelection)

    GrilleEnCours.VersionDscgp = 2


    GrilleEnCours.OpenDSCGP = Me.TB_FicDSCGP
    
    'Vérifie que le non du DSCGP correspond à la grille active
    Me.Tbx_NoDSCGP = GrilleEnCours.NumGrilleNue
    If (Me.Tbx_NoGrille = GrilleEnCours.NumGrilleNue) Or (Me.Tbx_NoGrille = GrilleEnCours.NumGrilleSymNue) Then
        Me.Tbx_NoGrille.Visible = False
        Me.Tbx_NoDSCGP.Visible = False
        Me.Lbl_diff.Visible = False
        Me.Lbl_NoDSCGP.Visible = False
        Me.Lbl_Nogrille.Visible = False
    Else
        MsgBox "Attention le N0 du DSCGP ne correspond pas au N° de la grille", vbCritical, "Erreur de Numéro"
        Me.Tbx_NoGrille.BackColor = RGB(255, 212, 255)
        Me.Tbx_NoGrille.Visible = True
        Me.Tbx_NoDSCGP.BackColor = RGB(255, 212, 255)
        Me.Tbx_NoDSCGP.Visible = True
        Me.Lbl_diff.Visible = True
        Me.Lbl_NoDSCGP.Visible = True
        Me.Lbl_Nogrille.Visible = True
    End If

    If Me.TB_FicDSCGP <> "" Then
        ValDscgp.GravureSup = GrilleEnCours.GravSup
        ValDscgp.GravureInf = GrilleEnCours.GravInf
        ValDscgp.GravureLat1 = GrilleEnCours.GravLat1
        ValDscgp.GravureLat2 = GrilleEnCours.GravLat2
        ValDscgp.GravureLat3 = GrilleEnCours.GravLat3
        ValDscgp.GravureLat4 = GrilleEnCours.GravLat4
        Me.img_check.Visible = True
        Me.img_uncheck.Visible = False
    End If
    Me.CB_Face = "Face inf"
    Me.CB_Face = "Face Sup"
    GoTo Fin
    
Erreur:
    If Err.Number > vbObjectError + 512 Then
        MsgBox Err.Description, vbCritical, "Element manquant"
    Else
        MsgBox Err.Description, vbCritical, "Erreur system"
    End If
    End
Fin:
    'Libération des classes
    Set GrilleEnCours = Nothing
End Sub

Private Sub BtnAnnul_Click()
Me.Hide
Me.ChB_OkAnnule = False

End Sub

Private Sub BtnOK_Click()
Dim Erreur As Boolean
Dim TempStr As String
    If Me.TBX_Espace = "" Then
        Me.TBX_Espace.BackColor = RGB(255, 212, 255)
        Erreur = True
    ElseIf Me.TBX_Police = "" Then
        Me.TBX_Police.BackColor = RGB(255, 212, 255)
        Erreur = True
    ElseIf Me.TBX_Ratio = "" Then
        Me.TBX_Ratio.BackColor = RGB(255, 212, 255)
        Erreur = True
    ElseIf Me.TBX_Taille = "" Then
        Me.TBX_Taille.BackColor = RGB(255, 212, 255)
        Erreur = True
    Else
        Me.ChB_OkAnnule = True
        Erreur = False
    End If

    If Not Erreur Then
        For i = 0 To Me.LB_TextGravure.ListCount - 1
            TempStr = TempStr & Me.LB_TextGravure.List(i) & Chr(10)
        Next i
        Select Case Me.CB_Face
            Case "Face Sup"
                ValDscgp.GravureSup = TempStr
            Case "Face Inf"
                ValDscgp.GravureInf = TempStr
            Case "Face Lat1"
                ValDscgp.GravureLat1 = TempStr
            Case "Face Lat2"
                ValDscgp.GravureLat2 = TempStr
            Case "Face Lat3"
                ValDscgp.GravureLat3 = TempStr
            Case "Face Lat4"
                ValDscgp.GravureLat4 = TempStr
        End Select
        Me.Hide
    End If
End Sub

Private Sub CB_Face_change()
'met a jour le texte à graver dans la zone de texte
Dim TabTxtGrav() As String
ReDim TabTxtGrav(0)
TabTxtGrav(0) = ""
Select Case Me.CB_Face
    Case "Face Sup"
        TabTxtGrav() = StringtoTab(ValDscgp.GravureSup)
    Case "Face Inf"
        TabTxtGrav() = StringtoTab(ValDscgp.GravureInf)
    Case "Face Lat1"
        TabTxtGrav() = StringtoTab(ValDscgp.GravureLat1)
    Case "Face Lat2"
        TabTxtGrav() = StringtoTab(ValDscgp.GravureLat2)
    Case "Face Lat3"
        TabTxtGrav() = StringtoTab(ValDscgp.GravureLat3)
    Case "Face Lat4"
        TabTxtGrav() = StringtoTab(ValDscgp.GravureLat4)
End Select
    Me.LB_TextGravure.List = TabTxtGrav()
End Sub

Private Sub CBX_Taille_Change()
'Met a jour les champs
Me.TBX_Police = Me.CBX_Taille.Column(1)
Me.TBX_Taille = Me.CBX_Taille.Column(2)
Me.TBX_Ratio = Me.CBX_Taille.Column(3)
Me.TBX_Espace = Me.CBX_Taille.Column(4)
End Sub

Private Sub LB_TextGravure_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
'ouvre le formulaire de modif du texte de la ligne sélectionnée si double click
    Dim LineSel As Integer
    LineSel = Me.LB_TextGravure.ListIndex
    
    Load Frm_SaisieGrav
    If LineSel = -1 Then
        Frm_SaisieGrav.Tbx_texte = ""
    Else
        Frm_SaisieGrav.Tbx_texte = Me.LB_TextGravure.List(LineSel)
    End If
    Frm_SaisieGrav.Show
    
    If LineSel = -1 Then
        Me.LB_TextGravure.List(0) = Frm_SaisieGrav.Tbx_texte
    Else
        Me.LB_TextGravure.List(LineSel) = Frm_SaisieGrav.Tbx_texte
    End If
    Unload Frm_SaisieGrav
End Sub

Private Sub Logo_eXcent_Click()
'Chargement de la boite eXcent
    Load Frm_eXcent
    Frm_eXcent.Tbx_Version = VMacro
    Frm_eXcent.Show

    Unload Frm_eXcent
End Sub

Private Sub UserForm_Initialize()
    Me.CB_Face.AddItem "Face Sup"
    Me.CB_Face.AddItem "Face Inf"
    Me.CB_Face.AddItem "Face Lat1"
    Me.CB_Face.AddItem "Face Lat2"
    Me.CB_Face.AddItem "Face Lat3"
    Me.CB_Face.AddItem "Face Lat4"
    Me.img_check.Visible = False
    Me.img_uncheck.Visible = True
End Sub
