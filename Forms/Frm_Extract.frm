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
If Rbt_GrilleG Or Rbt_GrilleSG Then
    NomPartGrille = Me.TBX_NomGriNueG & ".CATPart"
ElseIf Rbt_GrilleD Or Rbt_GrilleSD Then
    NomPartGrille = Me.TBX_NomGriNueD & ".CATPart"
End If

Dim MonPartGrille As New c_PartGrille
MonPartGrille.PG_partDocGrille = coll_docs.Item(CStr(NomPartGrille))
Set MonPartdoc = MonPartGrille.partDocGrille

    Dim MaSelection, InputObjectType, GeoSetSel
    Set MaSelection = MonPartdoc.Selection
    Dim SelAxis
    MaSelection.Clear
    ReDim InputObjectType(0)
    InputObjectType(0) = "AxisSystem"
    GeoSetSel = MaSelection.SelectElement2(InputObjectType, "Selectionnez le triedre", False)
    If Not (GeoSetSel = "") Then
        Set SelAxis = MonPartdoc.Selection.Item(1).Value
    End If

Dim MonHSFactory As HybridShapeFactory
Set MonHSFactory = MonPartGrille.HShapeFactory

Dim MonAxisSystem As AxisSystem
Set MonAxisSystem = SelAxis

Dim MonAxisSystemName As String
MonAxisSystemName = MonAxisSystem.Name

Dim MonHBody As HybridBody
If Not (MonPartGrille.Exist_HB(nHBTrav)) Then
    MonPartGrille.Create_HyBridShape (nHBTrav)
End If
Set MonHBody = MonPartGrille.Hb(nHBTrav)
Dim MesHShape As HybridShapes
Set MesHShape = MonHBody.HybridShapes

Dim StrRef As String
Dim RefTriedrePlan As Reference
Dim HSPlan, HSPlan10 As HybridShapePlaneOffset
'Détection de la présence des plan de ref et suppression le cas échéant
For i = 1 To 3
    Dim HSPlanName, HSPlan10Name As String
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
        MonHSFactory.DeleteObjectForDatum HSPlan
    End If
        Set HSPlan10 = MesHShape.Item(CStr(HSPlan10Name))
    If (Err.Number <> 0) Then
        ' "Plan_0" n'existe pas
        Err.Clear
    Else
        MonHSFactory.DeleteObjectForDatum HSPlan10
    End If
Next i

For i = 1 To 3
    StrRef = "RSur:(Face:(Brp:(" & MonAxisSystemName & ";" & i & ");None:();Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)"
    Set RefTriedrePlan = MonPartGrille.PartGrille.CreateReferenceFromBRepName(CStr(StrRef), MonAxisSystem)
    Set HSPlan = MonHSFactory.AddNewPlaneOffset(RefTriedrePlan, 0#, False)
    Set HSPlan10 = MonHSFactory.AddNewPlaneOffset(RefTriedrePlan, 10#, False)
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
MonPartGrille.PartGrille.Update
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
Me.TB_AxisRef.enabled = True
Me.Btn_SelAxis.enabled = True
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
Me.TB_AxisRef.enabled = True
Me.Btn_SelAxis.enabled = True
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
        Me.Btn_SelGrilleD.enabled = True
        Me.TBX_NomGriNueD.enabled = True
        Me.Lbl_DGD.enabled = True

        Me.Btn_SelGrilleG.enabled = False
        Me.TBX_NomGriNueG.enabled = False
        Me.Lbl_DGG.enabled = False

    Case 2
        Me.Btn_SelGrilleD.enabled = False
        Me.TBX_NomGriNueD.enabled = False
        Me.Lbl_DGD.enabled = False

        Me.Btn_SelGrilleG.enabled = True
        Me.TBX_NomGriNueG.enabled = True
        Me.Lbl_DGG.enabled = True

    Case 3 To 4
        Me.Btn_SelGrilleD.enabled = True
        Me.TBX_NomGriNueD.enabled = True
        Me.Lbl_DGD.enabled = True
        
        Me.Btn_SelGrilleG.enabled = True
        Me.TBX_NomGriNueG.enabled = True
        Me.Lbl_DGG.enabled = True
    
End Select

End Sub

Private Sub Cache_RB_Axe(CRB_Choix)
'Active/Desactive les boutons Symétrie
'CRB_Choix = symétrique ou non
' 1 = Sym, 2 = non sym
Select Case CRB_Choix
    Case 1
        Me.Fr_Sym.enabled = True
        Me.Rbt_X.enabled = True
        Me.Rbt_Y.enabled = True
        Me.Rbt_Z.enabled = True
    Case 2
        Me.Fr_Sym.enabled = False
        Me.Rbt_X.enabled = False
        Me.Rbt_Y.enabled = False
        Me.Rbt_Z.enabled = False
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
Me.TB_AxisRef.enabled = False
Me.Btn_SelAxis.enabled = False
Me.Lbl_Axis.enabled = False
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
Me.TB_AxisRef.enabled = False
Me.Lbl_Axis.enabled = False
Me.Btn_SelAxis.enabled = False

End Sub

