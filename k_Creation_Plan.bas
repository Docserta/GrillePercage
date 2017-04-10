Attribute VB_Name = "k_Creation_Plan"
Option Explicit

Sub catmain()
' *****************************************************************
' * Création des vues de base d'une grille
' * Ajout des Notas
' * Remplissage du cartouche
' *
' *
' * Création CFR le 19/08/2016
' * Modification le 15/12/2016 Intégration à la macro générale des grilles
' *
' *
' *****************************************************************

'Log de l'utilisation de la macro
LogUtilMacro nPath, nFicLog, nMacro, "k_Creation_Plan", VMacro

'Documents
Dim mDocs As Documents
Dim mDoc As Document
Dim mProd As Product

Dim FileToRead As String
Dim mDrawDoc As DrawingDocument
Dim mSheets As DrawingSheets

Dim mgrille As c_PartGrille

Dim AxeViewFront() 'As Double = Coordonées du plan de projection
Dim TypePlan As String 'Type de plan (ensemble ou detail)
Dim OrientPlan As String
Dim FormatPlan As String
Dim gCC As Boolean, gVT As Boolean
Dim gSym As Boolean
Dim gVisTronq As Boolean
Dim DocTo2D As Document
Dim DoctoCart As Document

    Set mDocs = CATIA.Documents
    Set mDoc = CATIA.ActiveDocument
    
    'Chargement du formulaire et initialisation des champs
    Load FRM_Plan
    FRM_Plan.Show
    'Bouton "annuler" choisi, on decharge le formulaire et on quitte
    If Not FRM_Plan.ChB_OkAnnule Then
        Unload FRM_Plan
        Exit Sub
    End If
    
    'Stockage des infos du formulaire
    If FRM_Plan.RBt_Horiz Then OrientPlan = "H" Else OrientPlan = "V"
    If FRM_Plan.CBx_Format = "" Then FormatPlan = "A0" Else FormatPlan = FRM_Plan.CBx_Format
    If FRM_Plan.Rbt_Ens Then
        TypePlan = "E"
        Set DoctoCart = mDocs.Item(CStr(FRM_Plan.Cbx_FicaTraiter))
        Set DocTo2D = mDocs.Item(CStr(FRM_Plan.Cbx_FicGriNue))
    Else
        TypePlan = "D"
        Set DocTo2D = mDocs.Item(CStr(FRM_Plan.Cbx_FicaTraiter))
    End If
    If FRM_Plan.RBt_CC Then gCC = True Else gCC = False
    If FRM_Plan.RBt_VT Then gVT = True Else gVT = False
    If FRM_Plan.ChB_Sym Then gSym = True Else gSym = False
    If FRM_Plan.ChB_Tronq Then gVisTronq = True Else gVisTronq = False
    FicCartAirbus = FRM_Plan.CBx_Format
    FichtxtCart = "FichtxtCart" & FRM_Plan.CBx_Format & ".txt"
    Unload FRM_Plan
    
    'instanciation de la classe partgrille
    Set mgrille = New c_PartGrille
    mgrille.PG_partDocGrille = DocTo2D
    
    'Création du plan
    'Chargement du format de référence
    FileToRead = Get_Active_CATVBA_Path & "TEMPLATE\" & FicCartAirbus & "_A0.CATDrawing"
    Set mDrawDoc = mDocs.NewFrom(FileToRead)
    Set mSheets = mDrawDoc.Sheets
    
    'Calcul des axes de projection de la vue de face
    ' AxeViewFront = CalPlanRef(mgrille)
     AxeViewFront = CalProjectionView(mgrille)
     
    'Tracé des vues
    If FRM_Plan.Rbt_Ens Then
        DesignVues DocTo2D, mSheets, OrientPlan, FormatPlan, AxeViewFront
    Else
        DesignVues DoctoCart, mSheets, OrientPlan, FormatPlan, AxeViewFront
    End If
    'DesignVues DocTo2D, mSheets, OrientPlan, FormatPlan, AxeViewFront
    
    'Insetion des Ditos
    WriteRemarks mDrawDoc, gCC, gVT, gSym, gVisTronq, mgrille.nom, ValDscgp.NumGrilleNueSym
    
    'Remplissage du cartouche
    WriteCart mgrille, TypePlan, DoctoCart, mSheets
    
End Sub

Private Function CalProjectionView(mgrille)
'Renvoi les coordonnée du plan de projection de la vue
'Ce plan est crée a partir de la projection de la "ref_grille" sur le plan "ref"
'et d'une perpendiculaire a cette projection

Dim Orient As HybridShape
Dim Plan As HybridShape
Dim ref_planProj As Reference
Dim ProjOrient As HybridShapeProject
Dim extProjOrient As HybridShapePointOnCurve
Dim PerpOrient As HybridShapeLineAngle
'Dim PerpOrient As HybridShapeLineNormal
Dim PlanProj As HybridShapePlane2Lines
Dim MesurPlan(8) 'As Double
Dim oMesur 'as Measurable

'Suppression des références de vue déja présent dans le Part
    'SupProjectionView mgrille
    
'Création de la projection de la ligne d'orientation grille sur le plan ref_grille
    Set Orient = mgrille.OrientationGrille
    Set Plan = mgrille.PlanRef
    Set ProjOrient = mgrille.HShapeFactory.AddNewProject(Orient, Plan)
    ProjOrient.Name = nProjOrient
    With ProjOrient
        .SolutionType = 0
        .Normal = True
        .SmoothingType = 0
    End With
    mgrille.Hb(nHBTrav).AppendHybridShape ProjOrient

'Création d'un point a l'extrémité de la projection
    Set extProjOrient = mgrille.HShapeFactory.AddNewPointOnCurveFromPercent(ProjOrient, 0#, False)
    extProjOrient.Name = nExtProjOrient
    mgrille.Hb(nHBTrav).AppendHybridShape extProjOrient

'Création d'une ligne perpendiculaire
    Set PerpOrient = mgrille.HShapeFactory.AddNewLineAngle(ProjOrient, Plan, extProjOrient, False, 0#, 100#, 90#, False)
    'Set PerpOrient = mGrille.HShapeFactory.AddNewLineNormal(Plan, extProjOrient, 0#, 100#, False)
    PerpOrient.Name = nPerpProjOrient
    mgrille.Hb(nHBTrav).AppendHybridShape PerpOrient

'Création d'un plan par 2 droites
    Set PlanProj = mgrille.HShapeFactory.AddNewPlane2Lines(ProjOrient, PerpOrient)
    PlanProj.Name = nPlaProj2D
    mgrille.Hb(nHBTrav).AppendHybridShape PlanProj
    Set ref_planProj = mgrille.PartGrille.CreateReferenceFromObject(PlanProj)
    mgrille.PartGrille.Update
'Mesure du plan
    Set oMesur = mgrille.GrilleSPAWorkbench.GetMeasurable(ref_planProj)
    oMesur.GetPlane MesurPlan

CalProjectionView = MesurPlan

End Function

Private Sub SupProjectionView(mgrille)
'Supprime les élément de construction et le plan de projection de la vue 2D
'Créé par la fonction "CalProjectionView"
'Cette procédure ne fonctionne pas( message d'erreur lors du Delete)
Dim mSel As Selection
Dim mShapes As HybridShapes
Dim mShape As HybridShape

    Set mSel = mgrille.GrilleSelection
    mSel.Clear
    On Error Resume Next
    Set mgrille.partDocGrille = mgrille.Hb(nHBTrav).InWorkObject
    Set mShapes = mgrille.Hb(nHBTrav).HybridShapes
    
    mSel.Add mShapes.Item(nPlaProj2D)
    'mSel.Add mShapes.Item(nPerpProjOrient)
    'mSel.Add mShapes.Item(nExtProjOrient)
    'mSel.Add mShapes.Item(nProjOrient)
    mSel.Delete
    On Error GoTo 0

End Sub

Private Function CalPlanRef(mgrille)
'Renvoi les coordonnée du plan de référence de la grille
On Error GoTo Erreur
Dim refPlan As Reference
Dim MesurPlan(8) 'As Double
Dim oMesur 'as Measurable
    
    Set refPlan = mgrille.Ref_PlanRef

'Mesure du plan
    Set oMesur = mgrille.GrilleSPAWorkbench.GetMeasurable(refPlan)
    oMesur.GetPlane MesurPlan
    CalPlanRef = MesurPlan
    GoTo Fin

Erreur:
    If Err.Number > vbObjectError + 512 Then
        MsgBox Err.Description, vbCritical, "Element manquant"
    Else
        MsgBox Err.Description, vbCritical, "Element manquant"
    End If
    End
Fin:

End Function

Private Sub WriteRemarks(mdraw As DrawingDocument, _
                            mCC As Boolean, _
                            mVT As Boolean, _
                            Optional mSym As Boolean = False, _
                            Optional mTronq As Boolean = False, _
                            Optional NoPart As String = " ", _
                            Optional NoPartSym As String = " ")
'documente le nota "remarks"
Dim mDito As Dito
Dim mSheets As DrawingSheets
Dim mSheetDetail As DrawingSheet
Dim mSheetCible As DrawingSheet
Dim mDitoSources As DrawingViews
Dim mVueCibles As DrawingViews
Dim mVueCible As DrawingView
Dim nVueMAin As DrawingView
Dim mDitoCibles As DrawingComponents
Dim DimSheet As Pos2D, PosDitoEC As Pos2D
Dim TxtDito
Dim TxtInstancie As DrawingText
Dim ValTxtDito As String
Dim cPt As Integer 'Compteur de remarques
Dim Nremark As Integer 'ReNumérotation des remarques dans le calque des vues

    Set mSheets = mdraw.Sheets
'collection des Dito du calque de détails
    If mSheets.Item(2).Name = "Calque.2 (Détail)" Then 'pour les plans Français
        Set mSheetDetail = mSheets.Item("Calque.2 (Détail)")
    ElseIf mSheets.Item(2).Name = "Sheet.2 (Detail)" Then 'pour les plans anglais
        Set mSheetDetail = mSheets.Item("Sheet.2 (Detail)")
    End If

'Initialisation de la cible des ditos
    Set mSheetCible = mSheets.Item(1)
    Set nVueMAin = mSheetCible.Views.Item("Main View")
    Set mVueCible = mSheetCible.Views.Item("Background View")
    Set mDitoSources = mSheetDetail.Views
    Set mDitoCibles = mVueCible.Components

'Calcul du point haut/droite du plan
'pour servir de référence d'insertion des ditos
        PosDitoEC.Y = GetDimSheet(mSheetCible.PaperSize).Y
        PosDitoEC.X = GetDimSheet(mSheetCible.PaperSize).X
        PosDitoEC.Y = PosDitoEC.Y - 20
        PosDitoEC.X = PosDitoEC.X - 173
        
    For cPt = 1 To 18
        Select Case cPt
            Case 1  '"Remarks1"
                Nremark = Nremark + 1
                'Instanciation du Dito
                mDito = InstanceDito(mDitoSources, mDitoCibles, PosDitoEC, CStr(cPt), CStr(Nremark))
                'Explose le Dito
                ExplodeDito mdraw, mDito
                'Memorise le dernier texte instancié pour lier le suivant
                Set TxtInstancie = mVueCible.Texts.GetItem("TxtRemk" & cPt)
                'Calcule la position du Dito suivant
                PosDitoEC.Y = PosDitoEC.Y - (HLig * 4)  'Hauteur du Dito inséré
        
            Case 2  '"Remarks2"
                Nremark = Nremark + 1
                mDito = InstanceDito(mDitoSources, mDitoCibles, PosDitoEC, CStr(cPt), CStr(Nremark))
                ExplodeDito mdraw, mDito
                'Lie le texte instancié au précédent
                AssociatText TxtInstancie, mVueCible.Texts.GetItem("TxtRemk" & cPt)
                PosDitoEC.Y = PosDitoEC.Y - (HLig * 2) 'Hauteur du Dito inséré
                Set TxtInstancie = mVueCible.Texts.GetItem("TxtRemk" & cPt)
            Case 3  '"Remarks3"
                If mCC Then
                    Nremark = Nremark + 1
                    mDito = InstanceDito(mDitoSources, mDitoCibles, PosDitoEC, CStr(cPt), CStr(Nremark))
                    ExplodeDito mdraw, mDito
                    AssociatText TxtInstancie, mVueCible.Texts.GetItem("TxtRemk" & cPt)
                    Set TxtInstancie = mVueCible.Texts.GetItem("TxtRemk" & cPt)
                    PosDitoEC.Y = PosDitoEC.Y - (HLig * 2) 'Hauteur du Dito inséré
                End If
            
            Case 4  '"Remarks4"
                Nremark = Nremark + 1
                mDito = InstanceDito(mDitoSources, mDitoCibles, PosDitoEC, CStr(cPt), CStr(Nremark))
                ExplodeDito mdraw, mDito
                AssociatText TxtInstancie, mVueCible.Texts.GetItem("TxtRemk" & cPt)
                Set TxtInstancie = mVueCible.Texts.GetItem("TxtRemk" & cPt)
                PosDitoEC.Y = PosDitoEC.Y - (HLig * 2)  'Hauteur du Dito inséré
                
            Case 5  '"Remarks5"
                Nremark = Nremark + 1
                mDito = InstanceDito(mDitoSources, mDitoCibles, PosDitoEC, CStr(cPt), CStr(Nremark))
                'Remplacement du texte xRmAj par H7,H8,F8
                Set TxtDito = mDito.Cible.GetModifiableObject(1)
                If mCC Then ValTxtDito = "H8" Else ValTxtDito = "H7"
                TxtDito.Text = Replace(CStr(TxtDito.Text), "xRmAj", ValTxtDito, 1, , vbTextCompare)
                ExplodeDito mdraw, mDito
                AssociatText TxtInstancie, mVueCible.Texts.GetItem("TxtRemk" & cPt)
                Set TxtInstancie = mVueCible.Texts.GetItem("TxtRemk" & cPt)
                PosDitoEC.Y = PosDitoEC.Y - (HLig * 3)  'Hauteur du Dito inséré
                
            Case 6  '"Remarks6"
                Nremark = Nremark + 1
                mDito = InstanceDito(mDitoSources, mDitoCibles, PosDitoEC, CStr(cPt), CStr(Nremark))
                ExplodeDito mdraw, mDito
                AssociatText TxtInstancie, mVueCible.Texts.GetItem("TxtRemk" & cPt)
                Set TxtInstancie = mVueCible.Texts.GetItem("TxtRemk" & cPt)
                PosDitoEC.Y = PosDitoEC.Y - (HLig * 4) 'Hauteur du Dito inséré
                
            Case 7  '"Remarks7"
                Nremark = Nremark + 1
                mDito = InstanceDito(mDitoSources, mDitoCibles, PosDitoEC, CStr(cPt), CStr(Nremark))
                ExplodeDito mdraw, mDito
                AssociatText TxtInstancie, mVueCible.Texts.GetItem("TxtRemk" & cPt)
                Set TxtInstancie = mVueCible.Texts.GetItem("TxtRemk" & cPt)
                PosDitoEC.Y = PosDitoEC.Y - (HLig * 4) 'Hauteur du Dito inséré
                
            Case 8  '"Remarks8"
                Nremark = Nremark + 1
                mDito = InstanceDito(mDitoSources, mDitoCibles, PosDitoEC, CStr(cPt), CStr(Nremark))
                ExplodeDito mdraw, mDito
                AssociatText TxtInstancie, mVueCible.Texts.GetItem("TxtRemk" & cPt)
                Set TxtInstancie = mVueCible.Texts.GetItem("TxtRemk" & cPt)
                PosDitoEC.Y = PosDitoEC.Y - (HLig * 11) 'Hauteur du Dito inséré
            
            Case 9  '"Remarks9"
                Nremark = Nremark + 1
                mDito = InstanceDito(mDitoSources, mDitoCibles, PosDitoEC, CStr(cPt), CStr(Nremark))
                'Remplacement du texte "xRmPartNbr" par le part number
                Set TxtDito = mDito.Cible.GetModifiableObject(1)
                ValTxtDito = NoPart
                TxtDito.Text = Replace(CStr(TxtDito.Text), "xRmPartNbr", ValTxtDito, 1, , vbTextCompare)
                ExplodeDito mdraw, mDito
                AssociatText TxtInstancie, mVueCible.Texts.GetItem("TxtRemk" & cPt)
                Set TxtInstancie = mVueCible.Texts.GetItem("TxtRemk" & cPt)
                PosDitoEC.Y = PosDitoEC.Y - (HLig * 3) 'Hauteur du Dito inséré
            
            Case 10 '"Remarks10"
                If mSym Then
                    Nremark = Nremark + 1
                    mDito = InstanceDito(mDitoSources, mDitoCibles, PosDitoEC, CStr(cPt), CStr(Nremark))
                    'Remplacement du texte "xRmPartNbrSym" par le part number
                    Set TxtDito = mDito.Cible.GetModifiableObject(1)
                    ValTxtDito = NoPartSym
                    TxtDito.Text = Replace(CStr(TxtDito.Text), "xRmPartNbrSym", ValTxtDito, 1, , vbTextCompare)
                    ExplodeDito mdraw, mDito
                    AssociatText TxtInstancie, mVueCible.Texts.GetItem("TxtRemk" & cPt)
                    Set TxtInstancie = mVueCible.Texts.GetItem("TxtRemk" & cPt)
                    PosDitoEC.Y = PosDitoEC.Y - (HLig * 3) 'Hauteur du Dito inséré
                End If
                
            Case 11 '"Remarks11"
                Nremark = Nremark + 1
                mDito = InstanceDito(mDitoSources, mDitoCibles, PosDitoEC, CStr(cPt), CStr(Nremark))
                ExplodeDito mdraw, mDito
                AssociatText TxtInstancie, mVueCible.Texts.GetItem("TxtRemk" & cPt)
                Set TxtInstancie = mVueCible.Texts.GetItem("TxtRemk" & cPt)
                PosDitoEC.Y = PosDitoEC.Y - (HLig * 4) 'Hauteur du Dito inséré
                
            Case 12 '"Remarks12"
                Nremark = Nremark + 1
                mDito = InstanceDito(mDitoSources, mDitoCibles, PosDitoEC, CStr(cPt), CStr(Nremark))
                ExplodeDito mdraw, mDito
                AssociatText TxtInstancie, mVueCible.Texts.GetItem("TxtRemk" & cPt)
                Set TxtInstancie = mVueCible.Texts.GetItem("TxtRemk" & cPt)
                PosDitoEC.Y = PosDitoEC.Y - (HLig * 8) 'Hauteur du Dito inséré
            
            Case 13 '"Remarks13"
                Nremark = Nremark + 1
                mDito = InstanceDito(mDitoSources, mDitoCibles, PosDitoEC, CStr(cPt), CStr(Nremark))
                ExplodeDito mdraw, mDito
                AssociatText TxtInstancie, mVueCible.Texts.GetItem("TxtRemk" & cPt)
                Set TxtInstancie = mVueCible.Texts.GetItem("TxtRemk" & cPt)
                PosDitoEC.Y = PosDitoEC.Y - (HLig * 2) 'Hauteur du Dito inséré
                
            Case 14 '"Remarks14"
                If mVT Then
                    Nremark = Nremark + 1
                    mDito = InstanceDito(mDitoSources, mDitoCibles, PosDitoEC, CStr(cPt), CStr(Nremark))
                    ExplodeDito mdraw, mDito
                    AssociatText TxtInstancie, mVueCible.Texts.GetItem("TxtRemk" & cPt)
                    Set TxtInstancie = mVueCible.Texts.GetItem("TxtRemk" & cPt)
                    PosDitoEC.Y = PosDitoEC.Y - (HLig * 2) 'Hauteur du Dito inséré
                End If
                
            Case 15 '"Remarks15"
                If mVT Then
                    Nremark = Nremark + 1
                    mDito = InstanceDito(mDitoSources, mDitoCibles, PosDitoEC, CStr(cPt), CStr(Nremark))
                    ExplodeDito mdraw, mDito
                    AssociatText TxtInstancie, mVueCible.Texts.GetItem("TxtRemk" & cPt)
                    Set TxtInstancie = mVueCible.Texts.GetItem("TxtRemk" & cPt)
                    PosDitoEC.Y = PosDitoEC.Y - (HLig * 2) 'Hauteur du Dito inséré
                End If
                
            Case 16 '"Remarks16"
                Nremark = Nremark + 1
                mDito = InstanceDito(mDitoSources, mDitoCibles, PosDitoEC, CStr(cPt), CStr(Nremark))
                ExplodeDito mdraw, mDito
                AssociatText TxtInstancie, mVueCible.Texts.GetItem("TxtRemk" & cPt)
                Set TxtInstancie = mVueCible.Texts.GetItem("TxtRemk" & cPt)
                PosDitoEC.Y = PosDitoEC.Y - (HLig * 2) 'Hauteur du Dito inséré
                
            Case 17 '"Remarks17"
                Nremark = Nremark + 1
                mDito = InstanceDito(mDitoSources, mDitoCibles, PosDitoEC, CStr(cPt), CStr(Nremark))
                ExplodeDito mdraw, mDito
                AssociatText TxtInstancie, mVueCible.Texts.GetItem("TxtRemk" & cPt)
                Set TxtInstancie = mVueCible.Texts.GetItem("TxtRemk" & cPt)
                PosDitoEC.Y = PosDitoEC.Y - (HLig * 2) 'Hauteur du Dito inséré
                
            Case 18 '"Remarks18"
                If mVT And mTronq Then
                    Nremark = Nremark + 1
                    mDito = InstanceDito(mDitoSources, mDitoCibles, PosDitoEC, CStr(cPt), CStr(Nremark))
                    ExplodeDito mdraw, mDito
                    AssociatText TxtInstancie, mVueCible.Texts.GetItem("TxtRemk" & cPt)
                    Set TxtInstancie = mVueCible.Texts.GetItem("TxtRemk" & cPt)
                    PosDitoEC.Y = PosDitoEC.Y - (HLig * 3) 'Hauteur du Dito inséré
                End If
            
        End Select
    Next cPt
    
nVueMAin.Activate
End Sub

Public Sub DesignVues(mDoc As Document, _
                        mSheets As DrawingSheets, _
                        Orient As String, _
                        Frmt As String, _
                        AxeVueF())
'Trace 4 vues, 1 vue de face + 2 vues projetée + 1 vue ISO
'Orient = Orientation du plan ("H" ou "V")
'Frmt = Format du plan ("A0", "A1" etc..)
' AxeVueF = Coordonées des vecteur du plan de projection

Dim mViews As DrawingViews
Dim mViewFront As DrawingView, _
        mViewProj1 As DrawingView, _
        mViewProj2 As DrawingView, _
        mViewIso As DrawingView, _
        mViewMain As DrawingView, _
        mViewBack As DrawingView
Dim mViewFrontGB As DrawingViewGenerativeBehavior, _
        mViewProj1GB As DrawingViewGenerativeBehavior, _
        mViewProj2GB As DrawingViewGenerativeBehavior, _
        mViewIsoGB As DrawingViewGenerativeBehavior
Dim mSheet As DrawingSheet

Dim iX1 As Double, iX2 As Double, iY1 As Double, iY2 As Double, iZ1 As Double, iZ2 As Double
Dim xAxData As Double, yAxData As Double
Dim mProjView As CatProjViewType
Dim nView As String
Dim DecalViewX As Long, DecalViewY As Long

    If Orient = "H" Then
        iX1 = AxeVueF(3)
        iY1 = AxeVueF(4)
        iZ1 = AxeVueF(5)
        iX2 = AxeVueF(6)
        iY2 = AxeVueF(7)
        iZ2 = AxeVueF(8)
        mProjView = catBottomView
        nView = nVueDessous
        DecalViewX = 0
        DecalViewY = 100
    Else
        iX1 = AxeVueF(6)
        iY1 = AxeVueF(7)
        iZ1 = AxeVueF(8)
        iX2 = AxeVueF(3)
        iY2 = AxeVueF(4)
        iZ2 = AxeVueF(5)
        mProjView = catRightView
        nView = nVueCote
        DecalViewX = 100
        DecalViewY = 0
    End If
        xAxData = 15
        yAxData = 15
    
    Set mSheet = mSheets.ActiveSheet
    Set mViews = mSheet.Views
    Set mViewMain = mViews.Item("Main View") 'affecte la "vue de Face" à mViewFront
    Set mViewBack = mViews.Item("Background View") 'Affecte Calque de fond a mViewBack
    mViewMain.Activate
    
    'Création Front view
    Set mViewFront = mViews.Add(nVueFace)
    Set mViewFrontGB = mViewFront.GenerativeBehavior
    With mViewFrontGB
        ' Declare le document a tracer dans la vue
        .Document = mDoc.Product
        ' Définition du plan de projection
        .DefineFrontView iX1, iY1, iZ1, iX2, iY2, iZ2
        .ForceUpdate
    End With
     
    'Création vue 1
    Set mViewProj1 = mViews.Add(nView)
    Set mViewProj1GB = mViewProj1.GenerativeBehavior
    With mViewProj1GB
        .Document = mDoc.Product
        .DefineProjectionView mViewFrontGB, mProjView
        .ForceUpdate
    End With
    With mViewProj1
        If Orient = "H" Then
            .yAxisData = yAxData
        Else
            .xAxisData = xAxData
        End If
        .ReferenceView = mViewFrontGB
        .AlignedWithReferenceView

    End With

    'Création vue 2
    Set mViewProj2 = mViews.Add(nVueArriere)
    Set mViewProj2GB = mViewProj2.GenerativeBehavior
    With mViewProj2GB
        .Document = mDoc.Product
        .DefineProjectionView mViewProj1GB, mProjView
        .ForceUpdate
    End With
    With mViewProj2
        If Orient = "H" Then
            .yAxisData = yAxData
        Else
            .xAxisData = xAxData
        End If
        .ReferenceView = mViewFrontGB
        .AlignedWithReferenceView
    End With
    
    ' Position des vues dans le plan
    ' !! Attention a mettre apres la création des vues auxiliaires
    '    pour eviter de perturber l'axe de glissement
    mViewFront.X = 300
    mViewFront.Y = 300
    mViewProj1.X = mViewFront.X + DecalViewX
    mViewProj1.Y = mViewFront.X + DecalViewY
    mViewProj2.X = mViewFront.X + DecalViewX * 2
    mViewProj2.Y = mViewFront.X + DecalViewY * 2
    
    'Création vue Iso
    Set mViewIso = mViews.Add(nVueIso)
    Set mViewIsoGB = mViewIso.GenerativeBehavior
    With mViewIsoGB
        .Document = mDoc.Product
        '.DefineIsometricView 1192.578, 762.125, 2688.684, 250, -12.5, 56.25
        .DefineIsometricView 1000, 500, 2000, 250, -10, 50
        .ForceUpdate
    End With
        mViewIso.X = mViewFront.X + 400
        mViewIso.Y = mViewFront.X + 400
End Sub

Public Sub WriteCart(mgrille As c_PartGrille, TypePlan As String, mDoc As Document, mSheets As DrawingSheets)
'Rempli les champs du cartouche
'TypePlan = "E" ou "D"
'mSheets = planche des vues

Dim mSheet As DrawingSheet
Dim mCarTexts As c_CartTxts
Dim mCarTxt As c_carttxt
    
Dim mProd As Product
Dim mParams As Parameters

Dim mTxts As DrawingTexts
Dim mTxt As DrawingText

Dim mView As DrawingView

Dim frmControls 'As Controls
Dim FrmControl 'As Control
   
    'Chargement de la liste des Champs
    Set mCarTxt = New c_carttxt
    Set mCarTexts = LoadAttrib()
    Select Case TypePlan
        Case "E"
            Set mProd = mDoc.Product
            Set mParams = mProd.ReferenceProduct.UserRefProperties
           
           'Ajout des valeur des attributs du product
            For Each mCarTxt In mCarTexts.Items
                'Le product Number
                If mCarTxt.Attrib = "PartNumber" Then
                    mCarTxt.Valeur = ValDscgp.NumGrilleAss
                    'mCarTxt.Valeur = mProd.PartNumber
                ElseIf mCarTxt.Attrib = "PartNumberSym" Then
                    mCarTxt.Valeur = ValDscgp.NumGrilleAssSym
                ElseIf mCarTxt.Attrib = "xDESIGNATION" Then
                    mCarTxt.Valeur = mProd.DescriptionRef
                ElseIf mCarTxt.Attrib = "xDESIGNATIONsym" Then
                    mCarTxt.Valeur = ValDscgp.DesignSym
                End If
            Next
        Case "D"
            'Ajout des valeur des attributs du Part
            For Each mCarTxt In mCarTexts.Items
                'mCarTxt.Valeur = mgrille.CherchParam(mCarTxt.Attrib)
                mCarTxt.Valeur = mgrille.LectureParam(mCarTxt.Attrib)
            Next
            'Le partNumber
            For Each mCarTxt In mCarTexts.Items
                If mCarTxt.Attrib = "PartNumber" Then
                    mCarTxt.Valeur = mgrille.nom
                ElseIf mCarTxt.Attrib = "xDESIGNATION" Then
                    mCarTxt.Valeur = mgrille.Prm_DescriptionRef
                End If
            Next

        Case Else
    End Select
    'Ajout des attributs communs
    'Ajout des valeurs non contenues dans les attributs du part
    For Each mCarTxt In mCarTexts.Items
        Select Case mCarTxt.Attrib
            Case "DRN_Date"
                mCarTxt.Valeur = DateFormatAIF()
            Case "DRN"
                mCarTxt.Valeur = "EXCENT"
        End Select
    Next
    
    'Validation des informations à ecrire
    Load Frm_Cartouche
    Set frmControls = Frm_Cartouche.Controls
        For Each mCarTxt In mCarTexts.Items
            Set FrmControl = frmControls.Item(mCarTxt.nom)
            FrmControl.Value = mCarTxt.Valeur
        Next
    Frm_Cartouche.Show
    'Sauvegarde des info modifiées
    For Each mCarTxt In mCarTexts.Items
        Set FrmControl = frmControls.Item(mCarTxt.nom)
        mCarTxt.Valeur = FrmControl.Value
    Next
    Unload Frm_Cartouche
    
    'Ecriture des valeurs dans le drawing
    Set mSheet = mSheets.ActiveSheet
    Set mView = mSheet.Views.Item("Background View")
    Set mTxts = mView.Texts
    
    For Each mCarTxt In mCarTexts.Items
        'Nom de l'objet texte dans le cartouche
'        Dim nChammpsCart As String
'        nChammpsCart = mCarTxt.nom
         Set mTxt = mTxts.GetItem(mCarTxt.nom)
          mTxt.Text = mCarTxt.Valeur
    Next

End Sub

Private Sub ExplodeDito(mdraw As DrawingDocument, mDito As Dito)
'Explose le Dito
'La décomposition du Dito génère de nouveau éléments mais conserve le Dito.
'Il faut ensuite le supprimer
Dim DrawingSel As Selection
        
        mDito.Cible.Explode
        Set DrawingSel = mdraw.Selection
        DrawingSel.Clear
        DrawingSel.Add mDito.Cible
        DrawingSel.Delete
End Sub

Private Function InstanceDito(mDitoSources, mDitoCibles, pos As Pos2D, cPt As String, No As String) As Dito
'Instancie le Dito passé en argument
'mDitoSources = collection des Dito sources du plan
'mDitoCibles = collection des Dito instanciés
'cPt = Numéro d'identification du Dito dans le calque de détail
'No = Numerotation du Dito apres import dans le calque des vues
Dim mDito As Dito
Dim TxtDito
Dim nDito As String 'Nom d'identification du Dito dans le plan template

nDito = "Remarks" & cPt
    Set mDito.Source = mDitoSources.Item(nDito)
    Set mDito.Cible = mDitoCibles.Add(mDito.Source, pos.X, pos.Y)
    
    'Remplace le numero de la remarque
    'Remplacement du texte xRmNo par le numéro du compteur
    Set TxtDito = mDito.Cible.GetModifiableObject(1)
    TxtDito.Text = Replace(CStr(TxtDito.Text), "xRmNo", No, 1, , vbTextCompare)
    InstanceDito = mDito
      
End Function

Private Sub AssociatText(TxtParent As DrawingText, TxtEnfant As DrawingText)
'Lie un texte à un autre
    TxtEnfant.AssociativeElement = TxtParent
End Sub

Private Function LoadAttrib() As c_CartTxts
'Charge la liste des attributs dans le fichier texte
Dim cCarttxt As c_carttxt
Dim cCarttxts As c_CartTxts
Dim PathFicCart As String
Dim f, fs

    Set cCarttxt = New c_carttxt
    Set cCarttxts = New c_CartTxts
    PathFicCart = Get_Active_CATVBA_Path & "Template\" & FichtxtCart
    Set fs = CreateObject("scripting.filesystemobject")
    Set f = fs.opentextfile(PathFicCart, ForReading, 1)
            Do While Not f.AtEndOfStream
                Set cCarttxt = SplitCSV(f.ReadLine)
                cCarttxts.Add cCarttxt.nom, cCarttxt.Attrib
            Loop
    f.Close

    Set LoadAttrib = cCarttxts
    
'Libération des classes
Set cCarttxt = Nothing
Set cCarttxts = Nothing

End Function

Public Function SplitCSV(str As String) As c_carttxt
'Extrait les valeurs de la chaine str séparées par le séparateur du fichier CSV (SepCSV)
Dim oMember As c_carttxt
Set oMember = New c_carttxt
Dim oVal As Collection
Set oVal = New Collection
    Do While InStr(1, str, SepCSV, vbTextCompare) > 0
        oVal.Add Left(str, InStr(1, str, SepCSV, vbTextCompare) - 1)
        str = Right(str, Len(str) - InStr(1, str, SepCSV, vbTextCompare))
    Loop
    oVal.Add str
    On Error Resume Next 'si oVal ne contient pas 3 valeurs
    oMember.nom = oVal.Item(1)
    oMember.Attrib = oVal.Item(2)
    oMember.Valeur = oVal.Item(3)

    On Error GoTo 0
Set SplitCSV = oMember
Set oVal = Nothing
Set oMember = Nothing
End Function

Private Function DateFormatAIF() As String
'Renvoi la date du jour au format Airbus
    DateFormatAIF = Day(Date) & "-"
    DateFormatAIF = DateFormatAIF & Choose(Month(Date), "JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC")
    DateFormatAIF = DateFormatAIF & "-" & Right(Year(Date), 2)
End Function
