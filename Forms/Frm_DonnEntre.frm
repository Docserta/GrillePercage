VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm_DonnEntre 
   Caption         =   "Données d'entrée de la grilles de perçage"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10215
   OleObjectBlob   =   "Frm_DonnEntre.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Frm_DonnEntre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Btn_Fic_Excel_Click()
'Documente les champs de la boite de dialogue à partir
'du fichier excel DSCGP
On Error GoTo Erreur
Dim GrilleEnCours As New c_DSCGP
    Me.TB_FicDSCGP = CATIA.FileSelectionBox("Selectionner le fichier DSCGP", "*.xls;*.xlsx", CatFileSelectionModeOpen)
    
    If Me.TB_FicDSCGP <> "" Then
        If Me.Rbt_Dscgp1 Then
            GrilleEnCours.VersionDscgp = 1
        Else
            GrilleEnCours.VersionDscgp = 2
        End If
        GrilleEnCours.OpenDSCGP = Me.TB_FicDSCGP
        Me.TBX_CoteConcep = GrilleEnCours.CoteConception
        Me.TBX_NomAss = GrilleEnCours.NumduLot
        Me.TBX_NomGriAss = GrilleEnCours.NumGrille
        Me.TBX_NomGriAssSym = GrilleEnCours.NumGrilleSym
        Me.TBX_NomGriNue = GrilleEnCours.NumGrilleNue
        Me.TBX_NomGriNueSym = GrilleEnCours.NumGrilleSymNue
        Me.TBX_NomU01 = GrilleEnCours.NumPartU01
        Me.TBX_NomU01Sym = GrilleEnCours.NumPartU01Sym
        Me.TBX_NomDtromp = GrilleEnCours.NumPartDet
        ValDscgp.CoteAvion = GrilleEnCours.CoteAvion
        ValDscgp.design = GrilleEnCours.DesignGrille
        ValDscgp.DesignSym = GrilleEnCours.DesignGrilleSym
        ValDscgp.NumEnvAvion = GrilleEnCours.EnvAvionCAO
        ValDscgp.Mat = GrilleEnCours.MatGrille
        ValDscgp.NumPiecesPerc = GrilleEnCours.PiecesPercees
        ValDscgp.Site = GrilleEnCours.Site
        ValDscgp.NumProgAvion = GrilleEnCours.NoProgAvion
        ValDscgp.Observ = GrilleEnCours.Observations
        ValDscgp.Dtemplate = GrilleEnCours.Dtemplate
        ValDscgp.Numout = GrilleEnCours.NumOutillage
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

Private Sub Btn_Nav_Detromp_Click()
Dim NomComplet As String
NomComplet = CATIA.FileSelectionBox("Selectionner le part de détrompage", "*.CATPart", CatFileSelectionModeOpen)
Me.TBX_NomDtromp = NomComplet
'Unload FRM_Navigateur_Fic
End Sub

Private Sub Btn_Nav_Env_Click()
Dim NomComplet As String
NomComplet = CATIA.FileSelectionBox("Selectionner l'environnement avion", "*.CATProduct", CatFileSelectionModeOpen)
Me.TBX_FicEnv = NomComplet
'Unload FRM_Navigateur_Fic
End Sub

Private Sub Btn_Nav_RepSauv_Click()

Me.TBX_RepSave = GetPath("Dossier destination grilles")

End Sub

Private Sub BtnAnnul_Click()
Me.Hide
Me.ChB_OkAnnule = False

End Sub

Private Sub BtnOK_Click()

Me.ChB_OkAnnule = True
Erreur = False
    'test si le nom de l'assemblage est documenté
    If Me.TBX_NomAss = "" Then 'And Me.TBX_NomAssSym = "" Then 'aucun N° de lot renseigné
        ColorChamps 2, 3, 3, 3, 2, 3, 3, 3
        Erreur = True
    ElseIf Me.TBX_NomAss <> "" Then 'And Me.TBX_NomAssSym = "" Then 'N° de lot gauche renseigné
        ColorChamps 1, 4, 4, 4, 3, 3, 3, 3
        If Me.TBX_NomGriAss = "" Then                           'aucun N° de grille ass renseigné
            ColorChamps 0, 2, 0, 0, 0, 0, 0, 0
            Erreur = True
        Else
            ColorChamps 0, 1, 0, 0, 0, 0, 0, 0
        End If
        If Me.TBX_NomGriNue = "" Then                           'aucun N° de grille nue renseigné
            ColorChamps 0, 0, 2, 0, 0, 0, 0, 0
            Erreur = True
        Else
            ColorChamps 0, 0, 1, 0, 0, 0, 0, 0
        End If
        If Me.ChB_U01 And Me.TBX_NomU01 = "" Then               'Le nom de la part U1 n'est pas renseigné
            ColorChamps 0, 0, 0, 2, 0, 0, 0, 0
            Erreur = True
        ElseIf Me.ChB_U01 And Me.TBX_NomU01 <> "" Then
            ColorChamps 0, 0, 0, 1, 0, 0, 0, 0
        End If
    ElseIf Me.TBX_NomAss = "" Then 'And Me.TBX_NomAssSym <> "" Then 'N° de lot droit renseigné
        ColorChamps 3, 3, 3, 3, 1, 4, 4, 4
        If Me.TBX_NomGriAssSym = "" Then                           'aucun N° de grille ass renseigné
            ColorChamps 0, 0, 0, 0, 0, 2, 0, 0
            Erreur = True
        Else
            ColorChamps 0, 0, 0, 0, 0, 1, 0, 0
        End If
        If Me.TBX_NomGriNueSym = "" Then                           'aucun N° de grille nue renseigné
            ColorChamps 0, 0, 0, 0, 0, 0, 2, 0
            Erreur = True
        Else
            ColorChamps 0, 0, 0, 0, 0, 0, 1, 0
        End If
        If Me.ChB_U01 And Me.TBX_NomU01Sym = "" Then               'Le nom de la part U1 n'est pas renseigné
            ColorChamps 0, 0, 0, 0, 0, 0, 0, 2
            Erreur = True
        ElseIf Me.ChB_U01 And Me.TBX_NomU01Sym <> "" Then
            ColorChamps 0, 0, 0, 0, 0, 0, 0, 1
        End If
    
    
    End If
    
    

    ' teste si le fichier d'environnement existe
    If (Not FileExist(Me.TBX_FicEnv)) Then
        Me.TBX_FicEnv.BackColor = RGB(255, 212, 255)
        Erreur = True
    Else
        Me.TBX_FicEnv.BackColor = RGB(212, 255, 255)
    End If
    If Me.TBX_RepSave = "" Then
        Me.TBX_RepSave.BackColor = RGB(255, 212, 255)
        Erreur = True
    Else
        Me.TBX_RepSave.BackColor = RGB(212, 255, 255)
    End If
    'Test si la part de détrompage existe
    '### désactivé, on ne va pas cercher la part de détrompage, on la crée
    '## CR MLC - CFR 13/05/16
'    If Me.ChB_Detromp Then
'        If (Not FileExist(Me.TBX_NomDtromp)) Then
'            Me.TBX_NomDtromp.BackColor = RGB(255, 212, 255)
'            Erreur = True
'        Else
'            Me.TBX_NomDtromp.BackColor = RGB(212, 255, 255)
'        End If
'    End If

If Not Erreur Then
    Me.Hide
End If
End Sub

Private Sub ChB_Detromp_Change()
    If Me.ChB_Detromp Then
        Me.LB_Detromp.enabled = True
        Me.TBX_NomDtromp.enabled = True
        Me.Btn_Nav_Detromp.enabled = True
    Else
        Me.LB_Detromp.enabled = False
        Me.TBX_NomDtromp.enabled = False
        Me.Btn_Nav_Detromp.enabled = False
    End If
End Sub

Private Sub ChB_U01_Change()
    If Me.ChB_U01 Then
        Me.Lb_U01.enabled = True
        Me.TBX_NomU01.enabled = True
        Me.TBX_NomU01Sym.enabled = True
    Else
        Me.Lb_U01.enabled = False
        Me.TBX_NomU01.enabled = False
        Me.TBX_NomU01Sym.enabled = False
        'Me.TBX_NomU01 = ""
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
'active le bouton "DSCGP" de type 2
Me.Rbt_Dscgp2 = True
Me.ChB_Detromp = False
Me.LB_Detromp.enabled = False
Me.TBX_NomDtromp.enabled = False
Me.Btn_Nav_Detromp.enabled = False
Me.ChB_U01 = False
Me.Lb_U01.enabled = False
Me.TBX_NomU01Sym.enabled = False
Me.TBX_NomU01.enabled = False
Me.img_check.Visible = False
Me.img_uncheck.Visible = True
End Sub

Public Sub ColorChamps(Ag As Integer, GAg As Integer, GNg As Integer, U1g As Integer, Ad As Integer, GAd As Integer, GNd As Integer, U1d As Integer)
'Colore en rouge les champs du formulaire ne contenant pas les bonnes informations
    If Ag <> 0 Then Me.TBX_NomAss.BackColor = Choose(Ag, ColVert, ColRouge, ColGris, ColNeutre)
    If GAg <> 0 Then Me.TBX_NomGriAss.BackColor = Choose(GAg, ColVert, ColRouge, ColGris, ColNeutre)
    If GNg <> 0 Then Me.TBX_NomGriNue.BackColor = Choose(GNg, ColVert, ColRouge, ColGris, ColNeutre)
    If U1g <> 0 Then Me.TBX_NomU01.BackColor = Choose(U1g, ColVert, ColRouge, ColGris, ColNeutre)
    'If Ad <> 0 Then Me.TBX_NomAssSym.BackColor = Choose(Ad, ColVert, ColRouge, ColGris, ColNeutre)
    If GAd <> 0 Then Me.TBX_NomGriAssSym.BackColor = Choose(GAd, ColVert, ColRouge, ColGris, ColNeutre)
    If GNd <> 0 Then Me.TBX_NomGriNueSym.BackColor = Choose(GNd, ColVert, ColRouge, ColGris, ColNeutre)
    If U1d <> 0 Then Me.TBX_NomU01Sym.BackColor = Choose(U1d, ColVert, ColRouge, ColGris, ColNeutre)


End Sub
