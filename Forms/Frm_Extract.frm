VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm_Extract 
   Caption         =   "Création Rapport de controle"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9330
   OleObjectBlob   =   "Frm_Extract.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Frm_Extract"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub Btn_SelAxis_Click()
'******************************************************************************************************
'* Création de 3 plans perpendiculaires aux 3 axes du triedre sélectionné
'*
'*
'* Création :  CFR Date : 29/1/15
'* Modification :
'*
'******************************************************************************************************

Me.Hide

Set coll_docs = CATIA.Documents
Dim MonPartdoc As PartDocument
Dim NomPartGrille As String
Dim mPartGrille As c_PartGrille
Dim MaSelection, InputObjectType, GeoSetSel
Dim SelAxis
Dim mHSFactory As HybridShapeFactory
Dim MonAxisSystem As AxisSystem
Dim MonAxisSystemName As String
Dim MonHBody As HybridBody
Dim MesHShape As HybridShapes
Dim StrRef As String
Dim RefTriedrePlan As Reference
Dim HSPlan, HSPlan10 As HybridShapePlaneOffset
Dim HSPlanName, HSPlan10Name As String
    
    If Rbt_GrilleG Or Rbt_GrilleSG Then
        NomPartGrille = Me.TBX_NomGriNueG & ".CATPart"
    ElseIf Rbt_GrilleD Or Rbt_GrilleSD Then
        NomPartGrille = Me.TBX_NomGriNueD & ".CATPart"
    End If

    'Initialisation des classes
    Set mPartGrille = New c_PartGrille

    mPartGrille.PG_partDocGrille = coll_docs.Item(CStr(NomPartGrille))
    Set MonPartdoc = mPartGrille.partDocGrille

    Set MaSelection = MonPartdoc.Selection
    MaSelection.Clear
    ReDim InputObjectType(0)
    InputObjectType(0) = "AxisSystem"
    GeoSetSel = MaSelection.SelectElement2(InputObjectType, "Selectionnez le triedre", False)
    If Not (GeoSetSel = "") Then
        Set SelAxis = MonPartdoc.Selection.Item(1).Value
    End If

    Set mHSFactory = mPartGrille.HShapeFactory
    Set MonAxisSystem = SelAxis
    MonAxisSystemName = MonAxisSystem.Name

    'Création du set géometrique "travail" s'il n'existe pas
    If Not (mPartGrille.Exist_HB(nHBTrav)) Then
        mPartGrille.Create_HyBridShape (nHBTrav)
    End If
    Set MonHBody = mPartGrille.Hb(nHBTrav)
    Set MesHShape = MonHBody.HybridShapes

    'Détection de la présence des plan de ref et suppression le cas échéant
    For i = 1 To 3
        Select Case i
        Case 1
            HSPlanName = "PlanX0"
            HSPlan10Name = "PlanX10"
        Case 2
            HSPlanName = "PlanY0"
            HSPlan10Name = "PlanY10"
        Case 3
            HSPlanName = "PlanZ0"
            HSPlan10Name = "PlanZ10"
        End Select
        On Error Resume Next
        Set HSPlan = MesHShape.Item(CStr(HSPlanName))
        If (Err.Number <> 0) Then
            ' "Plan_0" n'existe pas
            Err.Clear
        Else
            mHSFactory.DeleteObjectForDatum HSPlan
        End If
            Set HSPlan10 = MesHShape.Item(CStr(HSPlan10Name))
        If (Err.Number <> 0) Then
            ' "Plan_10" n'existe pas
            Err.Clear
        Else
            mHSFactory.DeleteObjectForDatum HSPlan10
        End If
    Next i

    For i = 1 To 3
        StrRef = "RSur:(Face:(Brp:(" & MonAxisSystemName & ";" & i & ");None:();Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)"
        Set RefTriedrePlan = mPartGrille.PartGrille.CreateReferenceFromBRepName(CStr(StrRef), MonAxisSystem)
        Set HSPlan = mHSFactory.AddNewPlaneOffset(RefTriedrePlan, 0#, False)
        Set HSPlan10 = mHSFactory.AddNewPlaneOffset(RefTriedrePlan, 0.01, False)
        MonHBody.AppendHybridShape HSPlan
        MonHBody.AppendHybridShape HSPlan10
        Select Case i
        Case 1
            HSPlan.Name = "PlanZ0"
            HSPlan10.Name = "PlanZ10"
        Case 2
            HSPlan.Name = "PlanX0"
            HSPlan10.Name = "PlanX10"
        Case 3
            HSPlan.Name = "PlanY0"
            HSPlan10.Name = "PlanY10"
        End Select
    Next i
    mPartGrille.PartGrille.Update
    Me.TB_AxisRef = MonAxisSystemName
    
    Me.Show
End Sub

Private Sub Btn_Fic_Excel_Click()
'Recupère les infos Grilles gauche et sym a partir du DSCGP
On Error GoTo Erreur
Dim DSCGP_EC As New c_DSCGP   'Grille
    Me.TB_FicDSCGP = CATIA.FileSelectionBox("Selectionner le fichier DSCGP", "*.xls;*.xlsx", CatFileSelection)
    'Type de DSCGP
    If Me.Rbt_Dscgp1 Then
        DSCGP_EC.VersionDscgp = 1
    Else
        DSCGP_EC.VersionDscgp = 2
    End If
    DSCGP_EC.OpenDSCGP = Me.TB_FicDSCGP
    Me.TBX_NomGriNueG = DSCGP_EC.NumGrilleNue
    Me.TBX_NomGriNueD = DSCGP_EC.NumGrilleSymNue
    With InfoDscgp
        .NumLot = DSCGP_EC.NumLot
        .NumGrilleAss = DSCGP_EC.NumGrille
        .NumGrilleAssSym = DSCGP_EC.NumGrilleSym
        .NumGrilleNue = DSCGP_EC.NumGrilleNue
        .NumGrilleNueSym = DSCGP_EC.NumGrilleSymNue
        .design = DSCGP_EC.DesignGrille
        .DesignSym = DSCGP_EC.DesignGrilleSym
        .Numout = DSCGP_EC.NumOutillage
        .NumEnvAvion = DSCGP_EC.EnvAvionCAO
        .Mat = DSCGP_EC.MatGrille
        .NumPiecesPerc = DSCGP_EC.PiecesPercees
        .Site = DSCGP_EC.Site
        .NumProgAvion = DSCGP_EC.NoProgAvion
        .Observ = DSCGP_EC.Observations
        .Dtemplate = DSCGP_EC.Dtemplate
        .SystNum = DSCGP_EC.SystemNum
        .Exemplaire = DSCGP_EC.Exemplaire
    End With
    Me.img_check.Visible = True
    Me.img_uncheck.Visible = False
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

Private Sub Btn_SelGrilleD_Click()
'Sélection du Product de la grille Droite
Frm_Extract.Hide
Dim BSGD_nomPart As String
'Selection de la grille
BSGD_nomPart = Select_PartGrille(1)

Frm_Extract.TBX_NomGriNueD = BSGD_nomPart
'Active la zone de sélection de l'Axis
Me.TB_AxisRef.Enabled = True
Me.Btn_SelAxis.Enabled = True
Frm_Extract.Show
End Sub

Private Sub Btn_SelGrilleG_Click()
'Sélection du Product de la grille gauche
Frm_Extract.Hide
Dim BSGG_nomPart As String
'Selection de la grille
BSGG_nomPart = Select_PartGrille(1)

Frm_Extract.TBX_NomGriNueG = BSGG_nomPart
'Active la zone de sélection de l'Axis
Me.TB_AxisRef.Enabled = True
Me.Btn_SelAxis.Enabled = True
Frm_Extract.Show
End Sub

Private Sub BtnAnnul_Click()
Me.Hide
Me.ChB_OkAnnule = False

End Sub

Private Sub BtnOK_Click()
Me.Hide
Me.ChB_OkAnnule = True
Erreur = False

End Sub

Private Sub Logo_eXcent_Click()
'Chargement de la boite eXcent
    Load Frm_eXcent
    Frm_eXcent.Tbx_Version = VMacro
    Frm_eXcent.Show

    Unload Frm_eXcent
End Sub

Private Sub Rbt_GrilleD_Click()
'Active/Desactive les boutons Axe de symétrie
Cache_RB_Axe 2 'Non sym
Cache_Tbx_Gri 1
'RAZ Axis de référence
RAZAxis
End Sub

Private Sub Rbt_GrilleG_Click()
'Active/Desactive les boutons Axe de symétrie
Cache_RB_Axe 2 'Non sym
Cache_Tbx_Gri 2
'RAZ Axis de référence
RAZAxis
End Sub

Private Sub Cache_Tbx_Gri(CTG_GDS As Integer)
'Active/Desactive les zônes de texte Info Grille
'CTG_GDS = type de grille
' 1 = Droite, 2 = Gauche, 3 = symétrique Droite, 4 = symétrique Gauche
Select Case CTG_GDS
    Case 1
        Me.Btn_SelGrilleD.Enabled = True
        Me.TBX_NomGriNueD.Enabled = True
        Me.Lbl_DGD.Enabled = True

        Me.Btn_SelGrilleG.Enabled = False
        Me.TBX_NomGriNueG.Enabled = False
        Me.Lbl_DGG.Enabled = False

    Case 2
        Me.Btn_SelGrilleD.Enabled = False
        Me.TBX_NomGriNueD.Enabled = False
        Me.Lbl_DGD.Enabled = False

        Me.Btn_SelGrilleG.Enabled = True
        Me.TBX_NomGriNueG.Enabled = True
        Me.Lbl_DGG.Enabled = True

    Case 3 To 4
        Me.Btn_SelGrilleD.Enabled = True
        Me.TBX_NomGriNueD.Enabled = True
        Me.Lbl_DGD.Enabled = True
        
        Me.Btn_SelGrilleG.Enabled = True
        Me.TBX_NomGriNueG.Enabled = True
        Me.Lbl_DGG.Enabled = True
    
End Select

End Sub

Private Sub Cache_RB_Axe(CRB_Choix)
'Active/Desactive les boutons Symétrie
'CRB_Choix = symétrique ou non
' 1 = Sym, 2 = non sym
Select Case CRB_Choix
    Case 1
        Me.Fr_Sym.Enabled = True
        Me.Rbt_X.Enabled = True
        Me.Rbt_Y.Enabled = True
        Me.Rbt_Z.Enabled = True
    Case 2
        Me.Fr_Sym.Enabled = False
        Me.Rbt_X.Enabled = False
        Me.Rbt_Y.Enabled = False
        Me.Rbt_Z.Enabled = False
End Select
End Sub

Private Sub Rbt_GrilleSD_Click()
'Active/Desactive les boutons Axe de symétrie
Cache_RB_Axe 1 'sym
Cache_Tbx_Gri 3
'RAZ Axis de référence
RAZAxis
End Sub

Private Sub Rbt_GrilleSG_Click()
'Active/Desactive les boutons Axe de symétrie
Cache_RB_Axe 1 'sym
Cache_Tbx_Gri 4
'RAZ Axis de référence
RAZAxis
End Sub

Private Sub UserForm_Initialize()
Me.Rbt_GrilleG = True
Me.Rbt_X = True
Me.Rbt_Num3D = True
Me.RBt_Fr = True
Me.TB_AxisRef.Enabled = False
Me.Btn_SelAxis.Enabled = False
Me.Lbl_Axis.Enabled = False
Me.Rbt_Dscgp2 = True
Me.img_check.Visible = False
Me.img_uncheck.Visible = True
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'Clic sur le bouton de fermeture dur formulaire
Exit Sub
End Sub

Private Sub RAZAxis()
'Désactive et remet a zéro le champs Axis de référence
Me.TB_AxisRef = "Référence de la part"
Me.TB_AxisRef.Enabled = False
Me.Lbl_Axis.Enabled = False
Me.Btn_SelAxis.Enabled = False

End Sub

