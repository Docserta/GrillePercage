VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm_AjoutGrille 
   Caption         =   "Ajout d'une Grille de perçage"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8475
   OleObjectBlob   =   "Frm_AjoutGrille.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Frm_AjoutGrille"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Btn_Fic_Excel_Click()
'Documente les champs de la boite de dialogue à partir
'du fichier excel DSCGP
On Erreur GoTo Erreur
Dim GrilleEnCours As New c_DSCGP
Me.TB_FicDSCGP = CATIA.FileSelectionBox("Selectionner le fichier DSCGP", "*.xls;*.xlsx", CatFileSelection)

    If Me.TB_FicDSCGP <> "" Then
        'Type de DSCGP
        If Me.Rbt_Dscgp1 Then
            GrilleEnCours.VersionDscgp = 1
        Else
            GrilleEnCours.VersionDscgp = 2
        End If
        GrilleEnCours.OpenDSCGP = Me.TB_FicDSCGP
        If Me.TBX_NomAss <> GrilleEnCours.NumduLot Then
            Dim msg As String
            msg = "Le Nom de l'assemblage du fichier DSCGP est différent de l'assemblage actuel!"
            msg = msg & vbCrLf
            msg = msg & "Nom d'assemblabe actuel : " & Me.TBX_NomAss
            msg = msg & vbCrLf
            msg = msg & "Nom d'assemblage dans fichier DSCGP : " & GrilleEnCours.NumduLot
            MsgBox msg, vbCritical, "Erreur de Nom d'assemblage"
        End If
        ValDscgp.NumGrilleAss = GrilleEnCours.NumGrille
        ValDscgp.NumGrilleAssSym = GrilleEnCours.NumGrilleSym
        ValDscgp.NumGrilleNue = GrilleEnCours.NumGrilleNue
        ValDscgp.NumGrilleNueSym = GrilleEnCours.NumGrilleSymNue
        
        If Me.ChB_Sym Then
            Me.TBX_NomGriAss = ValDscgp.NumGrilleAssSym
            Me.TBX_NomGriNue = ValDscgp.NumGrilleNueSym
            ValDscgp.Numout = GrilleEnCours.NumOutillageSym
            ValDscgp.design = GrilleEnCours.DesignGrilleSym
        Else
            Me.TBX_NomGriAss = ValDscgp.NumGrilleAss
            Me.TBX_NomGriNue = ValDscgp.NumGrilleNue
            ValDscgp.Numout = GrilleEnCours.NumOutillage
            ValDscgp.design = GrilleEnCours.DesignGrille
        End If
    
        ValDscgp.NumEnvAvion = GrilleEnCours.EnvAvionCAO
        ValDscgp.Mat = GrilleEnCours.MatGrille
        ValDscgp.NumPiecesPerc = GrilleEnCours.PiecesPercees
        ValDscgp.Site = GrilleEnCours.Site
        ValDscgp.NumProgAvion = GrilleEnCours.NoProgAvion
        ValDscgp.Observ = GrilleEnCours.Observations
        ValDscgp.Dtemplate = GrilleEnCours.Dtemplate
        ValDscgp.Exemplaire = GrilleEnCours.Exemplaire
       
        Me.img_check.Visible = True
        Me.img_uncheck.Visible = False
    End If
    GoTo Fin
    
Erreur:
    If Err.Number > vbObjectError + 512 Then
        MsgBox Err.Description, vbCritical, "Element manquant"
    Else
        MsgBox Err.Description, vbCritical, "Erreur system"
    End If
    End
Fin:
End Sub

Private Sub BtnAnnul_Click()
Me.Hide
Me.ChB_OkAnnule = False

End Sub

Private Sub BtnOK_Click()

Me.ChB_OkAnnule = True
Erreur = False
    ValDscgp.NumLot = Me.TBX_NomAss
    If Me.TBX_NomGriAss = "" Then
        Me.TBX_NomGriAss.BackColor = RGB(255, 212, 255)
        Erreur = True
    Else
        ValDscgp.NumGrilleAss = Me.TBX_NomGriAss 'grille ass
        Me.TBX_NomGriAss.BackColor = RGB(212, 255, 255)
    End If
    If Me.TBX_NomGriNue = "" Then
        Me.TBX_NomGriNue.BackColor = RGB(255, 212, 255)
        Erreur = True
    Else
       ValDscgp.NumGrilleNue = Me.TBX_NomGriNue 'grille nue
        Me.TBX_NomGriNue.BackColor = RGB(212, 255, 255)
    End If
    If Me.TBX_RepSave = "" Then
        Me.TBX_RepSave.BackColor = RGB(255, 212, 255)
        Erreur = True
    Else
        Me.TBX_RepSave.BackColor = RGB(212, 255, 255)
    End If

If Not Erreur Then
    Me.Hide
End If
End Sub


Private Sub ChB_Sym_Click()
'efface les champs sur changement sym / non sym
    If Me.ChB_Sym Then
        Me.TBX_NomGriAss = ValDscgp.NumGrilleAssSym
        Me.TBX_NomGriNue = ValDscgp.NumGrilleNueSym
    Else
        Me.TBX_NomGriAss = ValDscgp.NumGrilleAss
        Me.TBX_NomGriNue = ValDscgp.NumGrilleNue
    End If
End Sub



Private Sub ChB_U01_Change()
    If Me.ChB_U01 Then
        Me.Lb_U01.Enabled = True
        Me.TBX_NomU01.Enabled = True
    Else
        Me.Lb_U01.Enabled = False
        Me.TBX_NomU01.Enabled = False
        Me.TBX_NomU01 = ""
    End If
End Sub

Private Sub Logo_eXcent_Click()
'Chargement de la boite eXcent
    Load Frm_eXcent
    Frm_eXcent.Tbx_Version = VMacro
    Frm_eXcent.Show

    Unload Frm_eXcent
End Sub



Private Sub UserForm_Initialize()
'Active le bouton "DSCGP" de type 2
Me.Rbt_Dscgp2 = True
Me.ChB_U01 = False
Me.Lb_U01.Enabled = False
Me.TBX_NomU01.Enabled = False

    Me.img_check.Visible = False
    Me.img_uncheck.Visible = True
End Sub
