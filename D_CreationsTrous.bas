Attribute VB_Name = "D_CreationsTrous"

Option Explicit


'*********************************************************************
'* Macro : D_CreationsTrous
'*
'* Fonctions :  Perçage de la grille
'*
'* Version : 7
'* Création :  SVI
'* Modification : 05/08/14 CFR
'*                Ajout test PartBody/corps principal
'*                Prise en compte des points numéroté A1-asan..., A1- ou A1
'*                07/08/14 CFR
'*                Récupération des info de perçage dans un fichier excel
'* Modification : 02/05/15 CFR
'*                Affiche la liste des UDF sélectionnés et le Diamètre de perçage avion dans le Form
'*                Indique si dans la sélection un des diam de perçage est différent de la collection
'*                Memorise le N° de la machine dans un paramètre accroché au STD
'* Modification : 13/06/15 CFR
'*                Ajout de la création des trièdres sur les trous des grilles VT
'* Modification le : 24/06/16
'*                   remplacement du tableau Coll_RefExIsol par la classe c_Fasteners
'* Modification le : 12/04/17
'*                  Ajout d'une autre bibliotheque de bagues
'*
'**********************************************************************
Sub CATMain()
Dim StatusPartUpdated As Boolean
Dim instance_catpart_grille_nue As PartDocument
Dim GrilleActive As New c_PartGrille
Dim TestHBody As HybridBody
Dim TestHShape As HybridShape

'Log de l'utilisation de la macro
LogUtilMacro nPath, nFicLog, nMacro, "D_CreationsTrous", VMacro

CheminSourcesMacro = Get_Active_CATVBA_Path

'---------------------------
' Check de l'environnement
'---------------------------
    
    Err.Clear
    On Error Resume Next
    Set instance_catpart_grille_nue = CATIA.ActiveDocument
    If Err.Number <> 0 Then
        MsgBox "Le document de la fenêtre courante n'est pas un CATPart !", vbCritical, "Environnement incorrect"
        End
    End If
    On Error GoTo 0
    
    'test l'existence des Sets géométriques
    On Error GoTo Erreur
    Set TestHBody = GrilleActive.Hb(nHBPtConst)
    Set TestHBody = GrilleActive.Hb(nHBStd)
    Set TestHBody = GrilleActive.Hb(nHBTrav)
    Set TestHBody = GrilleActive.Hb(nHBGeoRef)
    'Test l'existance des elements géométriques
    Set TestHShape = GrilleActive.OrientationGrille
    Set TestHShape = GrilleActive.HS(nSurfSup, nHBTrav)
    Set TestHShape = GrilleActive.HS(nSurf0, nHBS0)
    Set TestHBody = Nothing
    Set TestHShape = Nothing
    On Error GoTo 0

'---------------------------
'Initialisation des variables
'---------------------------
    'Defini le chemin de la bibli des composants
    CheminBibliComposants = CorrigeDFS() & "\" & ComplementCheminBibliComposants
    'Construction des collections des paramètres de perçage
    Set CollBagues = New c_DefBagues
    'Set CollBagues = ReadXlsBagues("C:\CFR\Dropbox\Macros\Grilles-Percage" & "\" & NomFicInfoBagues)
    Set CollBagues = ReadXlsBagues(CheminBibliComposants & RepBaguesSprecif & "\" & NomFicInfoBagues)
    CollMachines.OpenBibliMachine = CheminBibliComposants & "\" & NomFicInfoMachines
    
    Set GrilleActive = New c_PartGrille

'Vérifie que le part est Update
    StatusPartUpdated = GrilleActive.PartGrille.IsUpToDate(GrilleActive.PartGrille)
    If StatusPartUpdated = False Then
        MsgBox "La part n'est pas a update. Faites la mise à jour avant de lancer cette macro.", vbInformation, "Part not Update"
        End
    End If
    
'Chargement de la boite de dialogue
    Load FRM_DiamPercage
    FRM_DiamPercage.Show
    If Not (FRM_DiamPercage.ChB_OkAnnule) Then
        Unload FRM_DiamPercage
        Exit Sub
    End If
    
    Creation_Trous GrilleActive
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

Private Sub Creation_Trous(ByRef GrilleActive)
   
Dim Std_Parameters As Parameters 'Collection des paramètre de la droite STD en cours
Dim NoMachine As String 'Paramètre qui va porter le Numéro Machine
Dim Diam_Trou As Double 'Diametre de perçage du nez machine
'Dim Diam_Lamage As String
Dim Diam_Lamage As Double 'Diamètre du lamage autour du nez machine
Dim Prof_Lamage As Double 'Profondeur du lamage
Dim Type_Trou As Integer
Dim RayonPerVisArretoir As Double 'Valeur du rayon de perçage de la vis arretoir
Dim DiamVisArretoir As String 'Valeur du taraudage de la vis arretoir
Dim ProfVisArretoir As Double 'Valeur de la profondeur du trou de la vis arretoir
Dim ProfTarauVisArretoir As Double 'Paramètre de la profondeur du taraudage de la vis arretoir
Dim AngleVisArretoir As Double 'Valeur de l'angle de la vis arretoir / a la ref grille
Dim RefBague As String, Ref_Vis As String 'Nom des bague et des vis arretoir

Dim TrouVis As Hole 'stocke le trou de la vis arrétoir pour le passer a la fonction de création de triedre

Dim cpt As Long
Dim j As Long
Dim Nom_Pt As String
Dim Nb_Pt_Sel As Long
   
Dim HBShape_STD As HybridShapes 'Collection des STD
Dim HBShape_Std_EC As HybridShape 'STD en cours

Dim Pt_Inter_StdSurf30 As HybridShapeIntersection 'Point d'intersection entre la surf à 30 et le STD
Dim Ref_Inter_StdSurf30 As Reference

Dim HBShape_PlanNormal As HybridShapePlaneNormal
Dim TempPoint As HybridShapeIntersection

Dim Mesure_PtInterSTDSurf30 'As Measurable
Dim Coord_PtInterSTDSurf30(2) As Variant

Dim PtAName As String  ' Nom du point A

Dim GrilleCC As Boolean, GrilleVT As Boolean, GrillePM As Boolean, TrouLame As Boolean
    GrilleCC = False: GrilleVT = False: GrillePM = False: TrouLame = False
    
Dim NBVisArretoir  As String
Dim NumVisArretoir As String, NumBague As String
         
Dim Axe_Trou As HybridShapeLinePtDir 'Axe du trou
Dim PlanFondLamage As HybridShapePlaneOffset 'Plan du Fon du Lamage
Dim PointImport As HybridShapeIntersection 'Point d'intersection entre le plan du fond de lamage et une droite
Dim mBar As New c_ProgressBar

'Stockage des infos du formulaire
    'Pour VT, CC et PM
    Diam_Trou = ChangeSingle(FRM_DiamPercage.TBX_DiamPercage)
    
    If FRM_DiamPercage.RB_GrilleCC Then 'Pour CC seulement
        NoMachine = FRM_DiamPercage.CBL_NumMachine
        GrilleCC = True
    ElseIf FRM_DiamPercage.RB_GrillePM Then 'Pour PM seulement
        NoMachine = FRM_DiamPercage.CBL_NumMachine
        NumBague = FRM_DiamPercage.TBX_NumBague
        GrillePM = True
    ElseIf FRM_DiamPercage.RB_GrilleVT Then 'Pour VT seulement
        NoMachine = FRM_DiamPercage.CBL_NumMachine
        GrilleVT = True
        RayonPerVisArretoir = ChangeSingle(FRM_DiamPercage.TBX_PosArret)
        DiamVisArretoir = FRM_DiamPercage.TBX_DiamArret
        ProfVisArretoir = ChangeSingle(FRM_DiamPercage.TBX_ProfArret)
        ProfTarauVisArretoir = ChangeSingle(FRM_DiamPercage.TBX_ProfTaraud)
        AngleVisArretoir = 35#
        NBVisArretoir = FRM_DiamPercage.CBL_NBVis
        NumVisArretoir = FRM_DiamPercage.TBX_NumVis
        NumBague = FRM_DiamPercage.TBX_NumBague
    End If
    'Pour Trou Lamé
    If FRM_DiamPercage.CB_Lamage Then
        TrouLame = True
        'Diam_Lamage = FRM_DiamPercage.TBX_DiamLamage
        Diam_Lamage = ChangeSingle(FRM_DiamPercage.TBX_DiamLamage)
        Prof_Lamage = 0.5
        Type_Trou = 2 '"catCounterboredHole" 'Trou lamé
    Else
        Type_Trou = 0 '"catSimpleHole" 'Trou simple
    End If
    
    Unload FRM_DiamPercage
    
 Nb_Pt_Sel = GrilleActive.GrilleSelection.Count

'Progress Barre
    Set mBar = New c_ProgressBar
    mBar.ProgressTitre 1, " Création des trous, veuillez patienter."
        
    Set HBShape_STD = GrilleActive.Hb(nHBStd).HybridShapes
      
    For cpt = 1 To Nb_Pt_Sel
        'Extraction du radical du nom du points A (avant le "-"
        If InStr(1, Left(Tab_Select_Points(0, cpt - 1), 4), "-", vbTextCompare) = 0 Then
            PtAName = Tab_Select_Points(0, cpt - 1)
        Else
            PtAName = Left(Tab_Select_Points(0, cpt - 1), InStr(1, Left(Tab_Select_Points(0, cpt - 1), 4), "-", vbTextCompare) - 1)
        End If
        mBar.ProgressEtape ((100 / Nb_Pt_Sel) * cpt), " Création du trou " & PtAName
        
        'Active le Set Travail
        GrilleActive.PartGrille.InWorkObject = GrilleActive.Hb(nHBTrav)
            
        Set HBShape_Std_EC = HBShape_STD.Item(Tab_Select_Points(2, cpt - 1))
            
    'Ajout des paramètres sur le STD en cours
        Set Std_Parameters = GrilleActive.PartGrille.Parameters.SubList(HBShape_Std_EC, True)
        
        If GrilleCC Then
            CreateParamExistString Std_Parameters, "NoMachine", NoMachine
            CreateParamExistString Std_Parameters, "NumBague", NumBague
            CreateParamExistDimension Std_Parameters, "DiamTrouNezMachine", Diam_Trou, "LENGTH"
        End If
        If GrillePM Then
            CreateParamExistString Std_Parameters, "NumBagueSF", NumBague
            CreateParamExistString Std_Parameters, "RefBagueSF", NoMachine
            CreateParamExistDimension Std_Parameters, "DiamTrouNezMachine", Diam_Trou, "LENGTH"
        End If
        If GrilleVT Then
            CreateParamExistString Std_Parameters, "NoMachine", NoMachine
            CreateParamExistDimension Std_Parameters, "DiamTrouNezMachine", Diam_Trou, "LENGTH"
            CreateParamExistDimension Std_Parameters, "RayonPerVisArretoir", RayonPerVisArretoir, "LENGTH"
            'CreateParamExistString Std_Parameters, "RayonPerVisArretoir", RayonPerVisArretoir
            'CreateParamExistDimension Std_Parameters, "DiamVisArretoir", DiamVisArretoir, "LENGTH"
            CreateParamExistString Std_Parameters, "DiamVisArretoir", DiamVisArretoir
            CreateParamExistDimension Std_Parameters, "ProfVisArretoir", ProfVisArretoir, "LENGTH"
            'CreateParamExistString Std_Parameters, "ProfVisArretoir", ProfVisArretoir
            CreateParamExistDimension Std_Parameters, "ProfTarauVisArretoir", ProfTarauVisArretoir, "LENGTH"
            CreateParamExistDimension Std_Parameters, "AngleVisArretoir", AngleVisArretoir, "ANGLE"
            'CreateParamExistString Std_Parameters, "AngleVisArretoir", AngleVisArretoir
            CreateParamExistString Std_Parameters, "NumBague", NumBague
            CreateParamExistString Std_Parameters, "NumVisArretoir", NumVisArretoir
            CreateParamExistString Std_Parameters, "NbVisArretoir", NBVisArretoir
        End If
        If TrouLame Then
            CreateParamExistDimension Std_Parameters, "DiamLamageTrouNezMachine", Diam_Lamage, "LENGTH"
            CreateParamExistDimension Std_Parameters, "ProfLamageTrouNezMachine", Prof_Lamage, "LENGTH"
            'CreateParamExistString Std_Parameters, "DiamLamageTrouNezMachine", Diam_Lamage
        End If
        
    'Création d'un point d'intersection entre la droite STD et la surface à 30
        GrilleActive.PartGrille.InWorkObject = GrilleActive.Hb(nHBPtConst)
        Set Pt_Inter_StdSurf30 = GrilleActive.HShapeFactory.AddNewIntersection(HBShape_Std_EC, GrilleActive.HS(nSurfSup, nHBTrav))
            Pt_Inter_StdSurf30.Name = "TempPt" & PtAName
            Pt_Inter_StdSurf30.PointType = 0
        GrilleActive.Hb(nHBPtConst).AppendHybridShape Pt_Inter_StdSurf30
        Set Ref_Inter_StdSurf30 = GrilleActive.PartGrille.CreateReferenceFromObject(Pt_Inter_StdSurf30)
        GrilleActive.PartGrille.UpdateObject Pt_Inter_StdSurf30
            
    'Récupération des coordonnées du point
        Set Mesure_PtInterSTDSurf30 = GrilleActive.GrilleSPAWorkbench.GetMeasurable(Ref_Inter_StdSurf30)
        Mesure_PtInterSTDSurf30.GetPoint Coord_PtInterSTDSurf30
        
    'Création du plan perpendiculaire à la droite STD passant par le point d'intersection entre la droite STD et la surface à 30
        Set HBShape_PlanNormal = GrilleActive.HShapeFactory.AddNewPlaneNormal(HBShape_Std_EC, Ref_Inter_StdSurf30)
        GrilleActive.Hb(nHBPtConst).AppendHybridShape HBShape_PlanNormal
        HBShape_PlanNormal.Name = "TempPlan" & PtAName
        
        GrilleActive.PartGrille.UpdateObject HBShape_PlanNormal
            
    'Active le Coprs Principal
        GrilleActive.PartGrille.InWorkObject = GrilleActive.PartGrille.MainBody
            
    'Création du trou de nez machine
        PercageNezMachine GrilleActive, Coord_PtInterSTDSurf30, HBShape_PlanNormal, Type_Trou, Diam_Trou, Diam_Lamage, PtAName, HBShape_Std_EC, mBar
    
    'Création du triedre d'import de la bague
        If GrilleVT Or GrillePM Then 'pour les grilles PM et VT
            If Type_Trou = 2 Then 'C'est un trou lamé
                'Création d'un plan parallèle au plan TempPtAxx de la valeur de la profondeur du lamage
                Set PlanFondLamage = Create_PlanLamage(GrilleActive, HBShape_PlanNormal, HBShape_Std_EC.Name)
                
                'Création du point d'intersection PlanLamage et STD
                Set PointImport = Creation_Pt_InterPlanLigne(GrilleActive, PlanFondLamage, HBShape_Std_EC, "PtInsertBague_" & HBShape_Std_EC.Name)
                        
                'Création du triedre sur le Pt d'intersection entre le fond du lamage et l'axe
                'Creation_Triedre_SurPt GrilleActive, PointImport, HBShape_Std_EC, "Bague" & PtAName
                Creation_Triedre_SurPt GrilleActive, PointImport, HBShape_Std_EC, "Bague" & Tab_Select_Points(0, cpt - 1)
            Else
                'Création du triedre sur le Pt TempPtxx
                'Creation_Triedre_SurPt GrilleActive, Pt_Inter_StdSurf30, HBShape_Std_EC, "Bague" & PtAName
                Creation_Triedre_SurPt GrilleActive, Pt_Inter_StdSurf30, HBShape_Std_EC, "Bague" & Tab_Select_Points(0, cpt - 1)
            End If
        End If
        
        'Creation du trou de la vis arretoir, si c'est une machine à double vis arretoir on relance la procédure pour la seconde vis
        If GrilleVT Then 'pour les grilles VT seules
            'Création du trou de vis
            Set TrouVis = PercageVisArretoir(GrilleActive, Coord_PtInterSTDSurf30, HBShape_PlanNormal, RayonPerVisArretoir, DiamVisArretoir, ProfVisArretoir, ProfTarauVisArretoir, PtAName, Pt_Inter_StdSurf30, HBShape_Std_EC.Name, Std_Parameters, 1, mBar)
            
            'Creation de l'axe du trou
            Set Axe_Trou = Create_Axe(GrilleActive, TrouVis, HBShape_Std_EC, "AxeVisArretoir1_" & HBShape_Std_EC.Name)
            If Type_Trou = 2 Then
                'Création du point d'intersection PlanLamage et Axe
                Set PointImport = Creation_Pt_InterPlanLigne(GrilleActive, PlanFondLamage, Axe_Trou, "PtInsertVis1_" & HBShape_Std_EC.Name)
                
                'Création du triedre d'import de la vis Arretoir sur le Pt d'intersection entre le fond du lamage et l'axe
                'Creation_Triedre_SurPt GrilleActive, PointImport, HBShape_Std_EC, "VisArretoir1" & PtAName
                Creation_Triedre_SurPt GrilleActive, PointImport, HBShape_Std_EC, "VisArretoir1" & Tab_Select_Points(0, cpt - 1)
            Else
                'Création du triedre d'import de la Vis Arretoir sur le Pt du sketch du trou
                'Creation_Triedre_SurTrou GrilleActive, TrouVis, HBShape_Std_EC, "VisArretoir1" & PtAName
                Creation_Triedre_SurTrou GrilleActive, TrouVis, HBShape_Std_EC, "VisArretoir1" & Tab_Select_Points(0, cpt - 1)
                
            End If
            
            If UCase(NBVisArretoir) = "DOUBLE" Then
                'Création du trou de vis
                Set TrouVis = PercageVisArretoir(GrilleActive, Coord_PtInterSTDSurf30, HBShape_PlanNormal, RayonPerVisArretoir, DiamVisArretoir, ProfVisArretoir, ProfTarauVisArretoir, PtAName, Pt_Inter_StdSurf30, HBShape_Std_EC.Name, Std_Parameters, 2, mBar)
                
                'Creation de l'axe du trou
                 Set Axe_Trou = Create_Axe(GrilleActive, TrouVis, HBShape_Std_EC, "AxeVisArretoir2_" & HBShape_Std_EC.Name)
                If Type_Trou = 2 Then
                    'Création du point d'intersection PlanLamage et Axe
                    Set PointImport = Creation_Pt_InterPlanLigne(GrilleActive, PlanFondLamage, Axe_Trou, "PtInsertVis2_" & HBShape_Std_EC.Name)
                    
                    'Création du triedre d'import de la vis
                    'Creation_Triedre_SurPt GrilleActive, PointImport, HBShape_Std_EC, "VisArretoir2" & PtAName
                    Creation_Triedre_SurPt GrilleActive, PointImport, HBShape_Std_EC, "VisArretoir2" & Tab_Select_Points(0, cpt - 1)
                Else
                    'Création du triedre d'import de la Vis Arretoir sur le Pt du sketch du trou
                    'Creation_Triedre_SurTrou GrilleActive, TrouVis, HBShape_Std_EC, "VisArretoir2" & PtAName
                    Creation_Triedre_SurTrou GrilleActive, TrouVis, HBShape_Std_EC, "VisArretoir2" & Tab_Select_Points(0, cpt - 1)
                End If
            End If
            
        End If
            
    Next
    GrilleActive.PartGrille.Update
    
'Libération des classes
    Set GrilleActive = Nothing
    Set mBar = Nothing
    
End Sub

Private Sub PercageNezMachine(GrilleActive, Coord_CentreTrou, PlanDeRef, TypeTrou, DTrou, DLamage, NomTrou, AxePercage, mBar)
'Création du trou du nez machine
'Coord_CentreTrou tableau des coordonnées du point TempPtAx
'PlanDeRef plan de perçage TempPlanAx
'TypeTrou 0 = Trou simple, 2 = Trou Lamé
'DTrou Diamètre du trou de nez machine
'DLamage = Diamètre du Lamage
'NomTrou nom du trou Ax
'AxePercage STD en cours

mBar.Etape = " Création du trou nez machine pour le point : " & NomTrou

Dim Percage As Hole
Dim LimitePercage As Limit
Dim DiamPercage As Length
Dim DiamLamage As Length
Dim ProfLamage As Length
Dim MonTrouSketch As Sketch
Dim factory2D1 As Factory2D
Dim ProjLigneStdinSketch As GeometricElements
Dim PointProjLigneStd As Geometry2D
Dim ProjLigneStd
Dim MTrouSketchContaintes As Constraints
Dim CentreTrouContainte As Constraint
Dim PTProjLineStdinSketch As Point2D
Dim Ref_PTProjLineStdinSketch As Reference
Dim Ref_PointProjLigneStd As Reference
Dim VisPropertySet_Sketch As VisPropertySet
Dim Selection_Sketch As Selection
'Paramètres pour trou lamé
Dim TrouParametres As Parameters
Dim tempParam As Parameter
Dim NomParam As String, DiamParam As String, Profparam As String
Dim LenParam As Integer
'Formules de la relation entre le paramètre et la fonction
Dim Formule_DiamTrouNez As Formula
Dim Formule_DiamLamageTrouNez As Formula
Dim Formule_ProfLamageTrouNez As Formula

    Set Selection_Sketch = GrilleActive.partDocGrille.Selection
    
    Set Percage = GrilleActive.PartShapeFactory.AddNewHoleFromPoint(Coord_CentreTrou(0), Coord_CentreTrou(1), Coord_CentreTrou(2), PlanDeRef, 50.032137)
        Percage.Type = TypeTrou
        Percage.AnchorMode = catExtremPointHoleAnchor
        Percage.BottomType = catTrimmedHoleBottom
        Percage.ThreadingMode = catSmoothHoleThreading
        Percage.ThreadSide = catRightThreadSide
        Percage.Reverse
        
    Set LimitePercage = Percage.BottomLimit
        LimitePercage.LimitMode = catUpThruNextLimit
    Set DiamPercage = Percage.Diameter
        DiamPercage.Value = DTrou
        Percage.Name = NomTrou
    Set Formule_DiamTrouNez = GrilleActive.GrilleRelations.CreateFormula("Formule_DiamTrouNez", "", DiamPercage, "`std\" & AxePercage.Name & "\DiamTrouNezMachine`")
    Set TrouParametres = GrilleActive.PartGrille.Parameters.SubList(Percage, True)

    NomParam = GrilleActive.PartGrille.Name
    NomParam = NomParam & "\" & GrilleActive.PartGrille.MainBody.Name
    'NomParam = NomParam & "\" & hole1.Name & "\HoleCounterBoredType.1\Diamètre"
    NomParam = NomParam & "\" & Percage.Name & "\HoleCounterBoredType"
    LenParam = Len(NomParam)
    
    If TypeTrou = 2 Then 'trou lamé
        For Each tempParam In TrouParametres
            If Left(tempParam.Name, LenParam) = NomParam And (Right(tempParam.Name, 8) = "Diamètre" Or Right(tempParam.Name, 8) = "Diameter") Then
                Set DiamLamage = TrouParametres.Item(tempParam.Name)
            ElseIf Left(tempParam.Name, LenParam) = NomParam And (Right(tempParam.Name, 10) = "Profondeur" Or Right(tempParam.Name, 5) = "Depth") Then
                Set ProfLamage = TrouParametres.Item(tempParam.Name)
            End If
        Next
        DiamLamage.Value = DLamage
        Set Formule_DiamLamageTrouNez = GrilleActive.GrilleRelations.CreateFormula("Formule_LamageTrouNez", "", DiamLamage, "`std\" & AxePercage.Name & "\DiamLamageTrouNezMachine`")
        ProfLamage.Value = 0.5
        Set Formule_ProfLamageTrouNez = GrilleActive.GrilleRelations.CreateFormula("Formule_ProfLamageTrouNez", "", ProfLamage, "`std\" & AxePercage.Name & "\ProfLamageTrouNezMachine`")
    End If
    
                
    Set MonTrouSketch = Percage.Sketch
    Set factory2D1 = MonTrouSketch.OpenEdition()
    Set ProjLigneStd = factory2D1.CreateProjections(AxePercage)
    Set PointProjLigneStd = ProjLigneStd.Item(1)
        PointProjLigneStd.Construction = True
    
    Set MTrouSketchContaintes = MonTrouSketch.Constraints
    Set ProjLigneStdinSketch = MonTrouSketch.GeometricElements
    Set PTProjLineStdinSketch = ProjLigneStdinSketch.Item("Point.1")
        PTProjLineStdinSketch.SetData Coord_CentreTrou(0), Coord_CentreTrou(1)
                  
    Set Ref_PTProjLineStdinSketch = GrilleActive.PartGrille.CreateReferenceFromObject(PTProjLineStdinSketch)
    Set Ref_PointProjLigneStd = GrilleActive.PartGrille.CreateReferenceFromObject(PointProjLigneStd)
    Set CentreTrouContainte = MTrouSketchContaintes.AddBiEltCst(catCstTypeOn, Ref_PTProjLineStdinSketch, Ref_PointProjLigneStd)
        CentreTrouContainte.Mode = catCstModeDrivingDimension

    'Masque le sketcher du trou
    Selection_Sketch.Clear
    Selection_Sketch.Add MonTrouSketch
    Set VisPropertySet_Sketch = Selection_Sketch.VisProperties
    'Set VisPropertySet_Sketch = VisPropertySet_Sketch.Parent
    VisPropertySet_Sketch.SetShow 1
    MonTrouSketch.CloseEdition
  
End Sub

Private Function PercageVisArretoir(GrilleActive, Coord_CentreTrou, PlanDeRef, Diam_Percage, Val_Taraudage, Prof_Trou, Prof_Taraud, NomTrou, PtCentre, Nom_Std, ParametresSTD, NB_Trou, mBar) As Hole
'Création du trou de la vis arretoir
'Coord_CentreTrou tableau des coordonnées du point TempPtAx
'PlanDeRef plan de perçage TempPlanAx
    'Diam_Percage Diametre sur lequel est percé la vis arretoir
    'Val_Taraudage Valeur du taraudage de la vis arretoir
    'Prof_Trou profondeur du perçage
    'Prof_Taraud profondeur du taraudage
'NomTrou nom du trou Ax
'PtCentre Point de centre du trou de nez machine (TempPtAx)
'Nom_Std Nom du STD en cours pour les relations entre les fonctions et les paramètres
'ParametresSTD Collection des paramètre ajoutés au STD
'NB_Trou = 1 ou 2 pour les doubles vis arrétoir. une a 35° l'autre a  35° + 180°

mBar.Etape = " Création de la vis arretoir pour le point : " & NomTrou

Dim Ref_TempPtEC As Reference
Dim Mesure_TempPtEC 'As Measurable

Dim Ref_TempPlanEC As Reference

Dim Ref_LigAngleVisArretoir As Reference 'Ligne pilotant l'angle de la vis arretoir
Dim Ref_ProjLigOrientGrille As Reference 'Projection de la ligne d'orientation de la grille dans le sketch

'Cercle dans l'esquisse du trou de la vis arretoir sur lequel est positioné le pt cible
Dim CerclePosVisArretoir As Circle2D 'Cercle sur lequel est positioné le pt cible du trous de la vis arretoir
Dim Ref_CerclePosVisArretoir As Reference
Dim PtCentreCerclePosVisArretoir As Point2D ' Pt centre du cercle de position du trou de la vis arretoir
Dim Ref_PtCentreCerclePosVisArretoir As Reference

Dim TrouVisArretoir As Hole 'Trou de la vis arretoir
Dim ProfVisArretoir As Limit 'Profondeur du trou
Dim ValProfVisArretoir As Length

Dim ProfTarauVisArretoir As Length

'Dim Valeur_LgTaraudVisArretoir As Double
'    Valeur_LgTaraudVisArretoir = 12#
 
Dim DiamTrouVisArretoir As StrParam 'Diamètre du taraudage
'Dim DiamTrouVisArretoir As Length 'Diamètre du trou
'Dim Valeur_DiamtrouVisArretoir As Double
'    Valeur_DiamtrouVisArretoir = 6#

Dim SketchTrouVisArretoir As Sketch 'Sketcher de la fonction perçage

Dim GeoElem_SketchTrouVisArretoir As GeometricElements
Dim PtCentreVisArretoir As Point2D
Dim Ref_PtCentreVisArretoir As Reference

Dim GeoElem_LigOrientGrille As GeometricElements

'Projection du TempPtEC dans l'esquisse du trou de vis Arrétoir
Dim GeoElem_ProjTempPtEC As GeometricElements
Dim Proj_TempPtEC As Geometry2D
Dim Ref_ProjTempPtEC As Reference

Dim ContraintesSketchTrouVisArr As Constraints
Dim AngleVisArretoir As Constraint
Dim Valeur_AngleVisArretoir As Angle, Valeur_AngleComplementaire As Angle

Dim CoincidCentreCercleTempPt As Constraint
Dim DiamPosVisArretoir As Constraint
Dim Valeur_DiamPosVisArretoir As Length
Dim CoincidcentreTrouBordCercle As Constraint
Dim LigAngleVisArretoir As Line2D
Dim Proj_LigOrientGrille As Geometry2D
Dim VisPropertySet_Sketch As VisPropertySet
Dim Selection_Sketch As Selection
    Set Selection_Sketch = GrilleActive.partDocGrille.Selection

'Formules de la relation entre le paramètre et la fonction
Dim Formule_RayonPerVisArretoir As Formula
Dim Formule_DiamVisArretoir As Formula
'Dim Formule_ProfVisArretoir As Formula
'Dim Formule_ProfTarauVisArretoir As Formula
Dim Formule_AngleVisArretoir As Formula
Dim Formule_AngleVisArretoir2 As Formula

GrilleActive.PartGrille.InWorkObject = GrilleActive.PartGrille.MainBody

    Set Ref_TempPlanEC = GrilleActive.PartGrille.CreateReferenceFromObject(PlanDeRef)
    Set Ref_TempPtEC = GrilleActive.PartGrille.CreateReferenceFromObject(PtCentre)
    
    'Création du Trou taraudé
    Set TrouVisArretoir = GrilleActive.PartGrille.ShapeFactory.AddNewHoleFromPoint(Coord_CentreTrou(0), Coord_CentreTrou(1), Coord_CentreTrou(2), Ref_TempPlanEC, 2)
        TrouVisArretoir.Type = catSimpleHole
        TrouVisArretoir.AnchorMode = catExtremPointHoleAnchor
        TrouVisArretoir.BottomType = catVHoleBottom                     'extrémitée pointue
        TrouVisArretoir.ThreadingMode = catThreadedHoleThreading
        TrouVisArretoir.ThreadSide = catRightThreadSide
        TrouVisArretoir.CreateStandardThreadDesignTable catHoleMetricThickPitch
        TrouVisArretoir.Reverse
        TrouVisArretoir.Name = "VisArretoir_" & NomTrou
    
    'Diamètre
        Set DiamTrouVisArretoir = TrouVisArretoir.HoleThreadDescription
        'DiamTrouVisArretoir.Value = Val_Taraudage
        Set Formule_DiamVisArretoir = GrilleActive.GrilleRelations.CreateFormula("Formule_DiamVisArretoir", "", DiamTrouVisArretoir, "`std\" & Nom_Std & "\DiamVisArretoir` ")
    'Profondeur perçage
        Set ProfVisArretoir = TrouVisArretoir.BottomLimit
        ProfVisArretoir.LimitMode = catOffsetLimit
        Set ValProfVisArretoir = ProfVisArretoir.Dimension
        ValProfVisArretoir.Value = Prof_Trou
        'Set Formule_ProfVisArretoir = GrilleActive.GrilleRelations.CreateFormula("Formule_ProfVisArretoir", "", ValProfVisArretoir, "std\" & Nom_Std & "\ProfVisArretoir")
    'Profondeur taraudage
        Set ProfTarauVisArretoir = TrouVisArretoir.ThreadDepth
        ProfTarauVisArretoir.Value = Prof_Taraud
        'Set Formule_ProfTarauVisArretoir = GrilleActive.GrilleRelations.CreateFormula("Formule_ProfTarauVisArretoir", "", ProfTarauVisArretoir, "std\" & Nom_Std & "\ProfTarauVisArretoir")
        
    'Set DiamTrouVisArretoir = TrouVisArretoir.Diameter
        'DiamTrouVisArretoir.Value = Valeur_DiamtrouVisArretoir
        'DiamTrouVisArretoir.Value = DiamTrou
        
    'Edition du sketch du point de centre du trou taraudé
        Set SketchTrouVisArretoir = TrouVisArretoir.Sketch
    
        GrilleActive.PartGrille.InWorkObject = SketchTrouVisArretoir
    
        Dim factory2D1 As Factory2D
        Set factory2D1 = SketchTrouVisArretoir.OpenEdition()
    
        Set GeoElem_SketchTrouVisArretoir = SketchTrouVisArretoir.GeometricElements
    
        Set PtCentreVisArretoir = GeoElem_SketchTrouVisArretoir.Item("Point.1")
            'PtCentreVisArretoir.SetData -21.764295, -6.442167
            PtCentreVisArretoir.SetData Coord_CentreTrou(0), Coord_CentreTrou(1)
    
        Set PtCentreCerclePosVisArretoir = factory2D1.CreatePoint(Coord_CentreTrou(0), Coord_CentreTrou(1))
            PtCentreCerclePosVisArretoir.ReportName = 4
    
        'Set CerclePosVisArretoir = factory2D1.CreateClosedCircle(-38.086216, 5.116158, 20#)
        Set CerclePosVisArretoir = factory2D1.CreateClosedCircle(Coord_CentreTrou(0), Coord_CentreTrou(1), 20#)
            CerclePosVisArretoir.CenterPoint = PtCentreCerclePosVisArretoir
            CerclePosVisArretoir.ReportName = 5
            CerclePosVisArretoir.Construction = True
    
        Set GeoElem_ProjTempPtEC = factory2D1.CreateProjections(Ref_TempPtEC)
        Set Proj_TempPtEC = GeoElem_ProjTempPtEC.Item("Mark.1") 'Mark.1 - Empreinte.1
            Proj_TempPtEC.Construction = True
        
        Set ContraintesSketchTrouVisArr = SketchTrouVisArretoir.Constraints
        
        Set Ref_PtCentreCerclePosVisArretoir = GrilleActive.PartGrille.CreateReferenceFromObject(PtCentreCerclePosVisArretoir)
        Set Ref_ProjTempPtEC = GrilleActive.PartGrille.CreateReferenceFromObject(Proj_TempPtEC)
        
        Set CoincidCentreCercleTempPt = ContraintesSketchTrouVisArr.AddBiEltCst(catCstTypeOn, Ref_PtCentreCerclePosVisArretoir, Ref_ProjTempPtEC)
            CoincidCentreCercleTempPt.Mode = catCstModeDrivingDimension
    
        Set Ref_CerclePosVisArretoir = GrilleActive.PartGrille.CreateReferenceFromObject(CerclePosVisArretoir)
        
        Set DiamPosVisArretoir = ContraintesSketchTrouVisArr.AddMonoEltCst(catCstTypeRadius, Ref_CerclePosVisArretoir)
            DiamPosVisArretoir.Mode = catCstModeDrivingDimension
        
        Set Valeur_DiamPosVisArretoir = DiamPosVisArretoir.Dimension
            'Valeur_DiamPosVisArretoir.Value = Diam_Percage
        Set Formule_RayonPerVisArretoir = GrilleActive.GrilleRelations.CreateFormula("Formule_RayonPerVisArretoir", "", Valeur_DiamPosVisArretoir, "`std\" & Nom_Std & "\RayonPerVisArretoir`")
        
        Set Ref_PtCentreVisArretoir = GrilleActive.PartGrille.CreateReferenceFromObject(PtCentreVisArretoir)
        
        Set CoincidcentreTrouBordCercle = ContraintesSketchTrouVisArr.AddBiEltCst(catCstTypeOn, Ref_PtCentreVisArretoir, Ref_CerclePosVisArretoir)
            CoincidcentreTrouBordCercle.Mode = catCstModeDrivingDimension
        
        Set LigAngleVisArretoir = factory2D1.CreateLine(-38.086216, 5.116158, -21.764295, -6.442167)
            LigAngleVisArretoir.ReportName = 6
            LigAngleVisArretoir.Construction = True
            LigAngleVisArretoir.StartPoint = PtCentreCerclePosVisArretoir
            LigAngleVisArretoir.EndPoint = PtCentreVisArretoir
    
        Set GeoElem_LigOrientGrille = factory2D1.CreateProjections(GrilleActive.Ref_OrientationGrille)
        
        Set Proj_LigOrientGrille = GeoElem_LigOrientGrille.Item("Mark.1") 'Mark.1 - Empreinte.1
            Proj_LigOrientGrille.Construction = True
        
        Set Ref_LigAngleVisArretoir = GrilleActive.PartGrille.CreateReferenceFromObject(LigAngleVisArretoir)
        Set Ref_ProjLigOrientGrille = GrilleActive.PartGrille.CreateReferenceFromObject(Proj_LigOrientGrille)
        
        Set AngleVisArretoir = ContraintesSketchTrouVisArr.AddBiEltCst(catCstTypeAngle, Ref_LigAngleVisArretoir, Ref_ProjLigOrientGrille)
            AngleVisArretoir.Mode = catCstModeDrivingDimension
            AngleVisArretoir.AngleSector = catCstAngleSector0
        
        Set Valeur_AngleVisArretoir = AngleVisArretoir.Dimension
        
        If NB_Trou = 1 Then
            Set Formule_AngleVisArretoir = GrilleActive.GrilleRelations.CreateFormula("Formule_AngleVisArretoir", "", Valeur_AngleVisArretoir, "`std\" & Nom_Std & "\AngleVisArretoir`")
        ElseIf NB_Trou = 2 Then
            'Valeur_AngleVisArretoir.Value = 215#
            Set Formule_AngleVisArretoir = GrilleActive.GrilleRelations.CreateFormula("Formule_AngleVisArretoir2", "", Valeur_AngleVisArretoir, "`std\" & Nom_Std & "\AngleVisArretoir` + 180deg")
        End If
            
        'Masque le sketcher du trou
            Selection_Sketch.Clear
            Selection_Sketch.Add SketchTrouVisArretoir
            Set VisPropertySet_Sketch = Selection_Sketch.VisProperties
            VisPropertySet_Sketch.SetShow 1
        
        SketchTrouVisArretoir.CloseEdition
    
    'GrilleActive.PartGrille.InWorkObject = TrouVisArretoir
    GrilleActive.PartGrille.InWorkObject = GrilleActive.PartGrille.MainBody
    On Error Resume Next
    GrilleActive.PartGrille.Update
   
'    If Err.Number <> 0 Then
'        MsgBox "Erreur durant l'update du part ! Verifiez les fonctions.", vbInformation, "Update impossible"
'    End If
Set PercageVisArretoir = TrouVisArretoir
End Function

Public Sub Creation_Triedre_SurPt(GrilleActive, CT_PTCentre, CT_DroiteDirection, CT_Name)
'Création d'un triedre sur le trou de nez machine
'CT_PTCentre point de centre du triedre (TempPTx pour les trous de nez machine) as HybridShapeIntersection
'CT_DroiteDirection axe Z du triedre (Ligne STD) as HS_STD_EC
' CT_Name nom du triedre

Dim New_Triedre As AxisSystem
Dim Ref_HS_TempPt_EC As Reference 'Reference sur le Pt Temp en cours
Dim Ref_HS_STD_EC As Reference 'Reference sur la droite STD en cours
Dim NomAxis As String
    NomAxis = "RepAss_" & CT_Name

'Test de la prééxistance d'un triedre de même nom
    If GrilleActive.Exist_AxisSystem(NomAxis) Then
        MsgBox "Un axis " & NomAxis & " existe déja dans la part ! Pensez à Supprimez le doublon.", vbCritical, "Elément en double"
    End If

    Set Ref_HS_TempPt_EC = GrilleActive.PartGrille.CreateReferenceFromObject(CT_PTCentre)

    Set Ref_HS_STD_EC = GrilleActive.PartGrille.CreateReferenceFromObject(CT_DroiteDirection)

    Set New_Triedre = GrilleActive.PartAxisSystems.Add()
        New_Triedre.OriginType = catAxisSystemOriginByPoint
        New_Triedre.OriginPoint = Ref_HS_TempPt_EC
        'New_Triedre.ZAxisType = catAxisSystemAxisSameDirection
        New_Triedre.ZAxisType = catAxisSystemAxisOppositeDirection
        New_Triedre.ZAxisDirection = Ref_HS_STD_EC
        New_Triedre.YAxisType = catAxisSystemAxisSameDirection
        New_Triedre.YAxisDirection = GrilleActive.Ref_OrientationGrille
        New_Triedre.XAxisType = catAxisSystemAxisSameDirection
        New_Triedre.Name = NomAxis
    
    GrilleActive.PartGrille.UpdateObject New_Triedre

    New_Triedre.IsCurrent = True

    GrilleActive.PartGrille.Update
    
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

Private Sub Creation_Triedre_SurTrou(GrilleActive, CT_Trou_Vis, CT_DroiteDirection, CT_Name)
'Création d'un triedre sur le trou de vis arrétoir
'CT_PTCentre point de centre du triedre (TempPTx pour les trous de nez machine) as HybridShapeIntersection
'CT_DroiteDirection axe Z du triedre (Ligne STD) as HS_STD_EC
' CT_Name nom du triedre

Dim New_Triedre As AxisSystem

Dim Sketch_Trou_Vis As Sketch
Dim Ref_Sketch_Trou_Vis As Reference
Dim Ref_HS_STD_EC As Reference
Dim NomAxis As String
    NomAxis = "RepAss_" & CT_Name

'Test de la prééxistance d'un triedre de même nom
If GrilleActive.Exist_AxisSystem(NomAxis) Then
    MsgBox "Un axis " & NomAxis & " existe déja dans la part ! Pensez à Supprimez le doublon.", vbCritical, "Elément en double"
    End
End If

Set Sketch_Trou_Vis = CT_Trou_Vis.Sketch
Set Ref_Sketch_Trou_Vis = GrilleActive.PartGrille.CreateReferenceFromObject(Sketch_Trou_Vis)
Set Ref_HS_STD_EC = GrilleActive.PartGrille.CreateReferenceFromObject(CT_DroiteDirection)

Set New_Triedre = GrilleActive.PartAxisSystems.Add()
    New_Triedre.OriginType = catAxisSystemOriginByPoint
    New_Triedre.OriginPoint = Ref_Sketch_Trou_Vis
    
    'New_Triedre.ZAxisType = catAxisSystemAxisSameDirection
    New_Triedre.ZAxisType = catAxisSystemAxisOppositeDirection
    New_Triedre.ZAxisDirection = Ref_HS_STD_EC
    
    New_Triedre.YAxisType = catAxisSystemAxisSameDirection
    New_Triedre.YAxisDirection = GrilleActive.Ref_OrientationGrille

    New_Triedre.XAxisType = catAxisSystemAxisSameDirection

    New_Triedre.Name = NomAxis

GrilleActive.PartGrille.UpdateObject New_Triedre

New_Triedre.IsCurrent = True

GrilleActive.PartGrille.Update
End Sub

Public Function Create_Axe(GrilleActive, CA_Trou, CA_Direction, CA_Name) As HybridShapeLinePtDir
'Création d'un axe passant par le centre d'un trou
' et d'une direction donnée par une droite
' CA_Trou as Hole
' CA_Direction as HybridShapeLinePtPt
' CA_Name Nom de l'axe

'Dim GrilleActive as new c_partgrille

Dim Sketch_Trou As Sketch
Dim Ref_Sketch_Trou As Reference
Dim Ref_Direction As Reference

Dim DirLigne As HybridShapeDirection
Dim LignePtDir As HybridShapeLinePtDir

'Set CA_Direction = GrilleActive.HB(nHBStd  ).HybridShapes.Item("Droite.12")
'Set CA_Trou = GrilleActive.PartGrille.MainBody.Shapes.Item.Item("VisArretoir_A17")
Set Sketch_Trou = CA_Trou.Sketch
Set Ref_Sketch_Trou = GrilleActive.PartGrille.CreateReferenceFromObject(Sketch_Trou)
Set Ref_Direction = GrilleActive.PartGrille.CreateReferenceFromObject(CA_Direction)

Set DirLigne = GrilleActive.HShapeFactory.AddNewDirection(Ref_Direction)
Set LignePtDir = GrilleActive.HShapeFactory.AddNewLinePtDir(Ref_Sketch_Trou, DirLigne, 0#, 20#, False)
    LignePtDir.Name = CA_Name

GrilleActive.Hb(nHBPtConst).AppendHybridShape LignePtDir
GrilleActive.PartGrille.InWorkObject = LignePtDir
GrilleActive.PartGrille.Update

Set Create_Axe = LignePtDir

End Function

Public Function Create_PlanLamage(GrilleActive, Plan_Ref, Nom_Std) As HybridShapePlaneOffset
'Crée un plan décale dont la valeur est égale au paramètre de profondeur du lamage du trou nez machine
' Plan_Ref = TempPlanAxx as HybridShapePlaneNormal
'

'Dim GrilleActive as new c_partgrille
'Dim Plan_Ref As HybridShapePlaneNormal

Dim Ref_Plan_Ref As Reference
Dim Plan_Decale As HybridShapePlaneOffset
'Formules de la relation entre le paramètre et la fonction
Dim Formule_ProfLamage As Formula
Dim ValProf_Lamage As Length

'Set Plan_Ref = GrilleActive.HB(nHBPtConst    ).HybridShapes.Item("TempPlanA39")
Set Ref_Plan_Ref = GrilleActive.PartGrille.CreateReferenceFromObject(Plan_Ref)
Set Plan_Decale = GrilleActive.HShapeFactory.AddNewPlaneOffset(Ref_Plan_Ref, 0.5, False)
    Plan_Decale.Name = "PlanLamage_" & Nom_Std
    
Set ValProf_Lamage = Plan_Decale.Offset
'Ajout de la formule qui lie la cote d'offset au parametre du STD
Set Formule_ProfLamage = GrilleActive.GrilleRelations.CreateFormula("Formule_ProfLamage", "", ValProf_Lamage, "`std\" & Nom_Std & "\ProfLamageTrouNezMachine`")


GrilleActive.Hb(nHBPtConst).AppendHybridShape Plan_Decale
GrilleActive.PartGrille.InWorkObject = Plan_Decale
GrilleActive.PartGrille.Update

Set Create_PlanLamage = Plan_Decale
End Function

Public Function Creation_Pt_InterPlanLigne(GrilleActive, CPI_Plan, CPI_Ligne, CPI_Name) As HybridShapeIntersection
'Création d'un point d'intersection entre un plan et une ligne
'CPI_Plan
'CPI_Ligne
'CPI_Name

'Dim grilleactive as new c_partgrille

'Dim HS_STD_EC As HybridShapeLinePtPt 'Droite STD en cours
Dim Ref_CPI_Ligne As Reference 'Reference sur la droite STD en cours

'Dim Plan_Lamage As HybridShapePlaneOffset
'Set Plan_Lamage = hybridShapes1.Item("Plan.98")

Dim Ref_CPI_Plan As Reference
Set Ref_CPI_Plan = GrilleActive.PartGrille.CreateReferenceFromObject(CPI_Plan)

'Set HS_STD_EC = grilleactive.HB(nHBStd  ).HybridShapes.Item("Droite.7")

Set Ref_CPI_Ligne = GrilleActive.PartGrille.CreateReferenceFromObject(CPI_Ligne)

Dim Pt_Intersection As HybridShapeIntersection
Set Pt_Intersection = GrilleActive.HShapeFactory.AddNewIntersection(Ref_CPI_Plan, Ref_CPI_Ligne)

Pt_Intersection.PointType = 0
Pt_Intersection.Name = CPI_Name
GrilleActive.Hb(nHBPtConst).AppendHybridShape Pt_Intersection
GrilleActive.PartGrille.InWorkObject = Pt_Intersection
GrilleActive.PartGrille.Update

Set Creation_Pt_InterPlanLigne = Pt_Intersection
End Function

