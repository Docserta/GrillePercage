VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FRM_CheckDscgp 
   Caption         =   "Check DSCGP"
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12240
   OleObjectBlob   =   "FRM_CheckDscgp.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FRM_CheckDscgp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Btn_Annule_Click()
'Clic sur bouton Annuler
    Me.Hide
    Me.ChB_OkAnnule = False
End Sub

Private Sub Btn_Correct_Click()
'Met a jour les propriétés dans la part et le product
    'Designation
    If Me.Chb_Design Then
        GrAss.Prm_DescriptionRef = Me.TB_Design_DSCGP
        Me.TB_Design_Prod = Me.TB_Design_DSCGP
        GrNue.EcritParam nPrmDesign, Me.TB_Design_DSCGP
        Me.TB_Design_Part = Me.TB_Design_DSCGP
    End If
    'Matiere
    If Me.Chb_Mat Then
        GrNue.EcritParam nPrmMaterial, Me.TB_Mat_DSCGP
        Me.TB_Mat_Part = Me.TB_Mat_DSCGP
    End If
    'Recognition
    If Me.Chb_Recogn Then
        GrNue.EcritParam nPrmRecogn, Me.TB_Recogn_DSCGP
        Me.TB_Recogn_Part = Me.TB_Recogn_DSCGP
    End If
    'Observation
    If Me.Chb_Observ Then
        GrNue.EcritParam nPrmObserv, Me.TB_Observ_DSCGP
        Me.TB_Observ_Part = Me.TB_Observ_DSCGP
    End If
    'DTEMPLATE
    If Me.Chb_Dtemplate Then
        GrNue.EcritParam nPrmDtempl, Me.TB_Dtempl_DSCGP
        Me.TB_Dtempl_Part = Me.TB_Dtempl_DSCGP
    End If
    'NumOutillage
    If Me.Chb_NumOut Then
        GrNue.EcritParam nPrmNumout, Me.TB_NumOut_DSCGP
        Me.TB_NumOut_Part = Me.TB_NumOut_DSCGP
    End If
    'Exemplaire (Gestion indice)
    If Me.Chb_Exemplaire Then
        GrNue.EcritParam nPrmExempl, Me.TB_Exempl_DSCGP
        Me.TB_Exempl_part = Me.TB_Exempl_DSCGP
    End If
    'Site Client
    If Me.Chb_Site Then
        GrNue.EcritParam nPrmSite, Me.TB_Site_DSCGP
        Me.TB_Site_part = Me.TB_Site_DSCGP
    End If
    'Prog Avion
    If Me.Chb_ProgAv Then
        GrNue.EcritParam nPrmProgAv, Me.TB_ProgAv_DSCGP
        Me.TB_ProgAv_part = Me.TB_ProgAv_DSCGP
    End If
    
    CheckValProd
    CheckValPart

End Sub

Private Sub Btn_Fic_Excel_Click()
'Récupère les champs du fichier excel DSCGP
On Error GoTo Erreur
Dim DSCGP_EC As New c_DSCGP
    Me.TB_FicDSCGP = CATIA.FileSelectionBox("Selectionner le fichier DSCGP", "*.xls;*.xlsx", CatFileSelection)
    
    If Me.TB_FicDSCGP <> "" Then
        If Me.Rbt_Dscgp1 Then
            DSCGP_EC.VersionDscgp = 1
        Else
            DSCGP_EC.VersionDscgp = 2
        End If
        DSCGP_EC.OpenDSCGP = Me.TB_FicDSCGP
        ValDscgp.NumLot = DSCGP_EC.NumduLot
        'Me.TBX_NomAssSym = Dscgp_EC.NumduLotSym
        If DSCGP_EC.NumGrilleNue = "" Then
            ValDscgp.NumGrilleNue = DSCGP_EC.NumGrilleSymNue
            ValDscgp.NumGrilleAss = DSCGP_EC.NumGrilleSym
            ValDscgp.design = DSCGP_EC.DesignGrilleSym
        Else
            ValDscgp.NumGrilleNue = DSCGP_EC.NumGrilleNue
            ValDscgp.NumGrilleAss = DSCGP_EC.NumGrille
            ValDscgp.design = DSCGP_EC.DesignGrille
        End If
    '    ValDscgp.NumGrilleAssSym = Dscgp_EC.NumGrilleSym
    '    ValDscgp.NumGrilleNueSym = Dscgp_EC.NumGrilleSymNue
        ValDscgp.NumPartU01Sym = DSCGP_EC.NumPartU01
        ValDscgp.NumPartU01 = DSCGP_EC.NumPartU01Sym
        ValDscgp.NumDetromp = DSCGP_EC.NumPartDet
        ValDscgp.NumEnvAvion = DSCGP_EC.EnvAvionCAO
        ValDscgp.Mat = DSCGP_EC.MatGrille
        ValDscgp.NumPiecesPerc = DSCGP_EC.PiecesPercees
        ValDscgp.Site = DSCGP_EC.Site
        ValDscgp.NumProgAvion = DSCGP_EC.NoProgAvion
        ValDscgp.Observ = DSCGP_EC.Observations
        ValDscgp.Dtemplate = DSCGP_EC.Dtemplate
        ValDscgp.Numout = DSCGP_EC.NumOutillage
        ValDscgp.Exemplaire = DSCGP_EC.Exemplaire
           
        'affiche les infos dans le formulaire
        Me.TB_REFPrt_DSCGP = ValDscgp.NumGrilleNue
        Me.TB_REFPrd_DSCGP = ValDscgp.NumGrilleAss
        Me.TB_Design_DSCGP = ValDscgp.design
        Me.TB_Mat_DSCGP = ValDscgp.Mat
        Me.TB_Recogn_DSCGP = vRecogn 'toujours "PGRI"
        Me.TB_Observ_DSCGP = ValDscgp.Observ
        Me.TB_Dtempl_DSCGP = ValDscgp.Dtemplate
        Me.TB_NumOut_DSCGP = ValDscgp.Numout
        Me.TB_Exempl_DSCGP = ValDscgp.Exemplaire
        Me.TB_Site_DSCGP = ValDscgp.Site
        Me.TB_ProgAv_DSCGP = ValDscgp.NumProgAvion
        Me.TB_PiecesP_DSCGP = ValDscgp.NumPiecesPerc
        CheckValProd
        CheckValPart
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

'Private Sub Btn_Nav_Gril_Click()
'Dim NomComplet As String
'    NomComplet = CATIA.FileSelectionBox("Selectionner le product de la grille Assemblée.", "*.CATProduct", CatFileSelection)
'    Me.TB_REFP_Prod = NomComplet
'End Sub

Private Sub Btn_OK_Click()
'Cache la boite de dialogue et redonne la main au programme
    Me.Hide
    Me.ChB_OkAnnule = True

End Sub

Private Sub Btn_SelGrille_Click()
'Sélection du product de la grille ass
   
    Me.Hide
    'Selection de la grille
    Me.TB_REFP_Prod = Select_PartGrille(5)
    If InStr(1, Me.TB_REFP_Prod, ".", vbTextCompare) > 0 Then
        Me.TB_REFP_Prod = Left(Me.TB_REFP_Prod, InStr(1, Me.TB_REFP_Prod, ".", vbTextCompare) - 1)
    End If
    Set coll_docs = CATIA.Documents
    
    On Error Resume Next
    GrAss.ProductDocGrille = coll_docs.Item(Me.TB_REFP_Prod & ".CATProduct")
    If Err.Number <> 0 Then
        Err.Clear
        MsgBox "Vous devez sélectionner un CatProduct", vbExclamation, "Erreur de sélection"
    Else
        Me.TB_REFP_Prod = GrAss.Numero
        Me.TB_Design_Prod = GrAss.Prm_DescriptionRef
        CheckValProd
    End If
    Me.Show
    Set coll_docs = Nothing
    Set GrAss = Nothing
End Sub

Private Sub Btn_SelGrilleNue_Click()
'Sélection de la part de la grille nue
    
    Me.Hide
    'Selection de la grille
    Me.TB_REFP_Part = Select_PartGrille(3)
     
    Set coll_docs = CATIA.Documents
    GrNue.PG_partDocGrille = coll_docs.Item(Me.TB_REFP_Part & ".CATPart")
    Me.TB_Design_Part = GrNue.Prm_xDesignation
    Me.TB_Mat_Part = GrNue.Prm_Material
    Me.TB_Recogn_Part = GrNue.Prm_Recognition
    Me.TB_Observ_Part = GrNue.Prm_Observation
    Me.TB_Dtempl_Part = GrNue.Prm_Dtemplate
    Me.TB_NumOut_Part = GrNue.Prm_xNumoutillage
    Me.TB_Exempl_part = GrNue.Prm_xExemplaire
    Me.TB_Site_part = GrNue.Prm_xSite
    Me.TB_ProgAv_part = GrNue.Prm_xNoprogavion
    Me.TB_PiecesP_part = GrNue.Prm_xPiecepercees
    
    CheckValPart
    Me.Show
    
    Set coll_docs = Nothing
    Set GrNue = Nothing
End Sub

Private Sub CheckValPart()
'Met en rouge les champs ne correspondant pas au DSCGP
If Me.TB_Design_Part <> Me.TB_Design_DSCGP Then
    Me.TB_Design_Part.BackColor = RGB(255, 212, 255)  'rouge
Else
    Me.TB_Design_Part.BackColor = RGB(212, 255, 255) 'vert
End If
If Me.TB_Mat_Part <> Me.TB_Mat_DSCGP Then
    Me.TB_Mat_Part.BackColor = RGB(255, 212, 255)  'rouge
Else
    Me.TB_Mat_Part.BackColor = RGB(212, 255, 255) 'vert
End If
If Me.TB_Recogn_Part <> Me.TB_Recogn_DSCGP Then
    Me.TB_Recogn_Part.BackColor = RGB(255, 212, 255)  'rouge
Else
    Me.TB_Recogn_Part.BackColor = RGB(212, 255, 255) 'vert
End If

If Me.TB_Observ_Part <> Me.TB_Observ_DSCGP Then
    Me.TB_Observ_Part.BackColor = RGB(255, 212, 255)  'rouge
Else
    Me.TB_Observ_Part.BackColor = RGB(212, 255, 255) 'vert
End If
If Me.TB_Dtempl_Part <> Me.TB_Dtempl_DSCGP Then
    Me.TB_Dtempl_Part.BackColor = RGB(255, 212, 255)  'rouge
Else
    Me.TB_Dtempl_Part.BackColor = RGB(212, 255, 255) 'vert
End If
If Me.TB_NumOut_Part <> Me.TB_NumOut_DSCGP Then
    Me.TB_NumOut_Part.BackColor = RGB(255, 212, 255)  'rouge
Else
    Me.TB_NumOut_Part.BackColor = RGB(212, 255, 255) 'vert
End If
If Me.TB_Exempl_part <> Me.TB_Exempl_DSCGP Then
    Me.TB_Exempl_part.BackColor = RGB(255, 212, 255)  'rouge
Else
    Me.TB_Exempl_part.BackColor = RGB(212, 255, 255) 'vert
End If
If Me.TB_Site_part <> Me.TB_Site_DSCGP Then
    Me.TB_Site_part.BackColor = RGB(255, 212, 255)  'rouge
Else
    Me.TB_Site_part.BackColor = RGB(212, 255, 255) 'vert
End If
If Me.TB_ProgAv_part <> Me.TB_ProgAv_DSCGP Then
    Me.TB_ProgAv_part.BackColor = RGB(255, 212, 255)  'rouge
Else
    Me.TB_ProgAv_part.BackColor = RGB(212, 255, 255) 'vert
End If
If Me.TB_PiecesP_part <> Me.TB_PiecesP_DSCGP Then
    Me.TB_PiecesP_part.BackColor = RGB(255, 212, 255)  'rouge
Else
    Me.TB_PiecesP_part.BackColor = RGB(212, 255, 255) 'vert
End If

End Sub

Private Sub CheckValProd()
'Met en rouge les champs ne correspondant pas au DSCGP
If Me.TB_Design_Prod <> Me.TB_Design_DSCGP Then
    Me.TB_Design_Prod.BackColor = RGB(255, 212, 255)  'rouge
Else
    Me.TB_Design_Prod.BackColor = RGB(212, 255, 255) 'vert
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
Me.Rbt_Dscgp2.Value = True
Me.img_check.Visible = False
Me.img_uncheck.Visible = True
End Sub


