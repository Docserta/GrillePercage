VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_PartGrille"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*********************************************************************
'* Classe : PartGrille
'*
'* Fonctions :  Teste la présence des éléments normalisé des Parts grilles
'*
'*
'* Version : 2
'* Création :  CFR
'* Modification : 29/01/14
'* Modification le : 24/06/16
'*                   remplacement du tableau Coll_RefExIsol par la classe c_Fasteners
'* Modification le 27/09/16
'*                   Ajout gestion d'erreurs
'*
'**********************************************************************

Private PG_PartDoc As PartDocument
Private PG_Part As part
Private PG_Selection As Selection
Private PG_PartBodies As Bodies

Private PG_Hbodies As HybridBodies ', PG_SubHbodies As HybridBodies
Private oHbodies As c_Hbodies

Private PG_ShapeFact As ShapeFactory
Private PB_HShapes  As HybridShapes
Private PG_HShapeFact As HybridShapeFactory

Private PG_AxisSystems As AxisSystems
Private PG_AxisSystem As AxisSystem

Private PG_OrientGrille As HybridShape     'Ligne d'orientation grille
Private PG_Refgrille As HybridShape        'Plan de ref grille

Private PG_ProductParametres As Parameters
Private PG_PartParametres As Parameters
Private PG_Product As Product

Private PG_SPAWorkbench As Workbench

Private PG_Fasteners As c_Fasteners
Private PG_CollBodies As c_Hbodies
'Private PG_CollSubBodies As c_Hbodies

Private PG_Relations As Relations

'
'===========================
' Initialisation de l'objet
'===========================
'
Private Sub Class_Initialize()
'
    On Error Resume Next
'
    Set PG_PartDoc = CATIA.ActiveDocument
    Set PG_Part = PG_PartDoc.part
    Set PG_PartBodies = PG_Part.Bodies
    Set PG_Hbodies = PG_Part.HybridBodies
    Set PG_Product = PG_PartDoc.Product
    Set PG_ProductParametres = PG_Product.ReferenceProduct.UserRefProperties
    Set PG_PartParametres = PG_Part.Parameters
    Set PG_ShapeFact = PG_Part.ShapeFactory
    Set PG_HShapeFact = PG_Part.HybridShapeFactory
    Set PG_Selection = PG_PartDoc.Selection
    Set PG_Relations = PG_Part.Relations
    Set PG_AxisSystems = PG_Part.AxisSystems
    Set PG_SPAWorkbench = PG_PartDoc.GetWorkbench("SPAWorkbench")
    Set PG_CollBodies = LstHBodies
    Set PG_Fasteners = LstFasteners

End Sub

Property Let PG_partDocGrille(partDocGrille As PartDocument)
' Réinitialisation de l'objet
    Set PG_PartDoc = partDocGrille
    Set PG_Part = PG_PartDoc.part
    Set PG_PartBodies = PG_Part.Bodies
    Set PG_Hbodies = PG_Part.HybridBodies
    Set PG_Product = PG_PartDoc.Product
    Set PG_ProductParametres = PG_Product.ReferenceProduct.UserRefProperties
    Set PG_PartParametres = PG_Part.Parameters
    Set PG_ShapeFact = PG_Part.ShapeFactory
    Set PG_HShapeFact = PG_Part.HybridShapeFactory
    Set PG_Selection = PG_PartDoc.Selection
    Set PG_Relations = PG_Part.Relations
    Set PG_AxisSystems = PG_Part.AxisSystems
    Set PG_SPAWorkbench = PG_PartDoc.GetWorkbench("SPAWorkbench")
    Set PG_CollBodies = LstHBodies
    Set PG_Fasteners = LstFasteners

End Property

Property Get partDocGrille() As PartDocument
'Renvoi le partDocument
    Set partDocGrille = PG_PartDoc
End Property

Property Get GrilleSelection() As Selection
'Renvoi les élements sélectionés dans le document
    Set GrilleSelection = PG_Selection
End Property

Property Get PartGrille() As part
'Renvoi le part de la grille
    Set PartGrille = PG_Part
End Property

Property Get PartBodies() As Bodies
'Renvoi la collection des PartBodys
    Set PartBodies = PG_PartBodies
End Property

Property Get mBody() As Body
'Renvoi le mainBody
    Set mBody = PartGrille.MainBody
End Property

Property Get Hbodies() As HybridBodies
'renvoi la collection des HybridBodies de premier niveau
    Set Hbodies = PG_Part.HybridBodies
End Property

Property Get Product() As Product
'Renvoi le Product du Part
    Set Product = PG_Product
End Property

Property Get ProductParametres() As Parameters
'Renvoi la collection des paramètres du part
    Set ProductParametres = PG_ProductParametres
End Property

Property Get PartParametres() As Parameters
'Renvoi la collection des paramètres du part
    Set PartParametres = PG_PartParametres
End Property

Property Get nom() As String
'Renvoi le nom de la part
    nom = PG_Part.Name
End Property

Property Get GrilleSPAWorkbench() As Workbench
'Renvoi le SPAWorkbench
    Set GrilleSPAWorkbench = PG_SPAWorkbench
End Property

Property Get GrilleRelations() As Relations
'Renvoi la collection des relations
    Set GrilleRelations = PG_Relations
End Property

'#############################################################
'# Les Parts Bodies
'#############################################################

Property Get PartBody(vNom) As Body
'Renvoi le Body passé en argument
On Error GoTo Erreur
    Set PartBody = PG_PartBodies.Item(vNom)
    GoTo Fin
Erreur:
    Err.Raise vbObjectError + 518, , "Le Part Body '" & vNom & "' est manquant ou mal orthographié. veuillez corriger."
Fin:
End Property

Property Get PartShapeFactory() As ShapeFactory
'Renvoi le part shape Factory
    Set PartShapeFactory = PG_ShapeFact
End Property

'#############################################################
'# Les Sets Géométriques
'#############################################################

Private Function LstHBodies() As c_Hbodies
'Renvoi la collection des Hybrid Bodies

Dim tmpHbodies As c_Hbodies
Dim tBod As HybridBody, tSBod As HybridBody
Dim tSubBodies As HybridBodies
Dim tSel As Selection
Dim tShow As CatVisPropertyShow
Dim tVisProp As VisPropertySet

    Set tmpHbodies = New c_Hbodies
    Set tSel = GrilleSelection
    'Set tSel = CATIA.ActiveDocument.Selection
    
    For Each tBod In Hbodies
        tSel.Clear
        tSel.Add tBod
        Set tVisProp = tSel.VisProperties
        tVisProp.GetShow tShow
        If isNotDoublon(tmpHbodies, tBod.Name) Then
            tmpHbodies.Add tBod.Name, tShow, tBod
        End If
        'Recherche des sous set géométriques dans le set "Travail"
        If tBod.Name = nHBTrav Then
            Set tSubBodies = tBod.HybridBodies
            For Each tSBod In tSubBodies
                tSel.Clear
                tSel.Add tSBod
                Set tVisProp = tSel.VisProperties
                tVisProp.GetShow tShow
                If isNotDoublon(tmpHbodies, tSBod.Name) Then
                    tmpHbodies.Add tSBod.Name, tShow, tSBod
                End If
            Next
        End If
    Next
    Set LstHBodies = tmpHbodies
    Set tmpHbodies = Nothing
End Function

Private Function isNotDoublon(tClasse, tName) As Boolean
'detecte si un item est déja présent dans la classe
'si la classe est indexée, elle ne supporte pas les doublons
Dim tmpItem
On Error GoTo Erreur
    Set tmpItem = tClasse.Item(tName)
    If Err.Number = 0 Then
        isNotDoublon = False
    End If
    GoTo Fin
Erreur:
    isNotDoublon = True
Fin:
End Function

Property Get Exist_HB(ByVal vNom As String) As Boolean
'Détecte si le set géométrique passé en argument existe
    Exist_HB = Exist_Hbodi(vNom)
End Property

Property Get Hb(ByVal vNom As String) As HybridBody
'Renvoi le set géométrique passé en argument
    On Error GoTo Erreur
    Set Hb = PG_CollBodies.Item(vNom).HBody
    GoTo Fin
Erreur:
    Err.Raise vbObjectError + 513, , "Le Set Géométrique '" & vNom & "' est manquant ou mal orthographié. veuillez corriger."
Fin:
End Property

Private Function Exist_Hbodi(ByVal vNom As String) As Boolean
'Détecte si le set géométrique passé en argument existe
    Dim mHBody As HybridBody
    Dim mSubHbody As HybridBody
    Dim mSubBodies As HybridBodies
    Exist_Hbodi = False

    For Each mHBody In PG_Hbodies
        If mHBody.Name = vNom Then
            Exist_Hbodi = True
            Exit For
        ElseIf mHBody.Name = nHBTrav Then 'Recherche des sous set géométriques
            Set mSubBodies = mHBody.HybridBodies
            For Each mSubHbody In mSubBodies
                If mSubHbody.Name = vNom Then
                    Exist_Hbodi = True
                End If
            Next
        End If
    Next
End Function

Property Get HShapeFactory() As HybridShapeFactory
On Error GoTo Erreur
    Set HShapeFactory = PG_HShapeFact
    GoTo Fin
Erreur:
    Err.Raise vbObjectError + 519, , "Erreur lors de la récupération de l'objet HybridShapeFactory contactez le service Info."
Fin:
End Property

Property Get PartAxisSystems() As AxisSystems
'Renvoi la collection des systems d'axe
    Set PartAxisSystems = PG_AxisSystems
End Property

Property Get Exist_AxisSystem(NomAxis) As Boolean
'detecte si l'axis passé en argument existe
    On Error Resume Next
    Set PG_AxisSystem = PartAxisSystems.Item(NomAxis)
    If (Err.Number <> 0) Then
        ' axis inexistant
        Err.Clear
        Exist_AxisSystem = False
    Else
        Exist_AxisSystem = True
    End If
End Property

'#############################################################
'# Les surfaces
'#############################################################

Property Get Exist_Shape(ByVal nShape As String, ByVal nHB As String) As Boolean
'Detecte si la surface passée en argument existe
    Exist_Shape = Exist_shapeinHB(nShape, nHB)
End Property

Private Function Exist_shapeinHB(ByVal nShape As String, ByVal nHB As String) As Boolean
'Detecte si la surface passée en argument existe
    Dim tHS As HybridShape
    On Error Resume Next
    'vérification de l'existance du set géométrique
    If Exist_HB(nHB) Then
        Set tHB = Hb(nHB).HybridShapes.Item(nShape)
        If (Err.Number <> 0) Then
            Err.Clear
            Exist_shapeinHB = False
        Else
            Exist_shapeinHB = True
        End If
    Else
        Exist_shapeinHB = False
    End If
End Function

Property Get HS(ByVal nShape As String, ByVal nHB As String) As HybridShape
'Renvoi la surface demandée en argument dans le set passé en argument
'nShape = Nom de la surface
'nHB non de l'hybridBody dans lequel doit se trouver la surface
On Error GoTo Erreur
    Set HS = Hb(nHB).HybridShapes.Item(nShape)
    GoTo Fin
Erreur:
    Err.Raise vbObjectError + 514, , "La surface " & nShape & " est manquante ou mal orthographiée . Veuillez corriger."
Fin:
End Property

Property Get Ref_HS(ByVal nShape As String, ByVal nHB As String) As Reference
'Renvoi la référence à la surface passée en argument
    Set Ref_HS = PG_Part.CreateReferenceFromObject(HS(nShape, nHB))
End Property

'-Ligne d'orientation grille---------------------------------------------------------------------------

Property Get Exist_OrientationGrille() As Boolean
'Detecte si la ligne "Orientation Grille" existe
Dim SubHBGeoref As HybridBody
    'La ligne "orientation_grille" est dans le set "geometrie de reference"
    ' qui est dans le set géométrique "Travail" on vérifie d'abord si les set existe
    If Exist_HB(nHBTrav) And Exist_HB(nHBGeoRef) Then
        Set SubHBGeoref = Hb(nHBTrav).HybridBodies.Item(nHBGeoRef)
        On Error Resume Next
        Set PG_OrientGrille = SubHBGeoref.HybridShapes.Item(nOrientGri)
        If (Err.Number <> 0) Then
        'HybridShape "OrientationGrille" inexistante
            Err.Clear
            Exist_OrientationGrille = False
        Else
            Exist_OrientationGrille = True
        End If
    Else
        Exist_OrientationGrille = False
    End If
End Property

Property Get OrientationGrille() As HybridShape
'Renvoi la ligne d'orientation grilleHb(nHBGeoRef).HybridShapes
On Error GoTo Erreur
Dim tHB As HybridBody
    Set tHB = PG_CollBodies.Item(nHBGeoRef).HBody
    Set PG_OrientGrille = tHB.HybridShapes.Item(nOrientGri)
    Set OrientationGrille = PG_OrientGrille
    GoTo Fin
Erreur:
    Err.Raise vbObjectError + 515, , "La ligne 'orientation_grille' de la grille est manquante ou mal orthographiée . Veuillez corriger."
Fin:
End Property

Property Get Ref_OrientationGrille() As Reference
'Renvoi la référence à la ligne d'orintation grille
    Set Ref_OrientationGrille = PG_Part.CreateReferenceFromObject(OrientationGrille)
End Property

'-Plan ref grille---------------------------------------------------------------------------

Property Get PlanRef() As HybridShape
'Renvoi le plan "ref_grille"
On Error GoTo Erreur
Dim tHB As HybridBody
    Set tHB = PG_CollBodies.Item(nHBGeoRef).HBody
    Set PG_Refgrille = tHB.HybridShapes.Item(nRefGri)
    Set PlanRef = PG_Refgrille
    GoTo Fin
Erreur:
    Err.Raise vbObjectError + 516, , "Le plan " & nRefGri & " de la grille est manquant ou mal orthographiée . Veuillez corriger."
Fin:
End Property

Property Get Ref_PlanRef() As Reference
'Renvoi la référence à la ligne d'orintation grille
    Set Ref_PlanRef = PG_Part.CreateReferenceFromObject(PlanRef)
End Property

Property Get Exist_PT(EPT_Nom) As Boolean
'Detecte si le Point passé en argument existe dans le set "Point de construction"
Dim tHB As HybridBody
Dim EB_PT As HybridShapeIntersection
    Set tHB = PG_CollBodies.Item(nHBPtConst).HBody
    On Error Resume Next
    Set EB_PT = tHB.HybridShapes.Item(EPT_Nom)
    If (Err.Number <> 0) Then
        Err.Clear
        Exist_PT = False
    Else
        Exist_PT = True
    End If
End Property

'#############################################################
'# Les Paramètres
'#############################################################

Public Function TestParamExist(TPE_NomParam As String) As Boolean
'test si le paramètre passé en argument existe dans le part.
'si oui renvoi true sinon False,
On Error Resume Next
    Dim TPE_ParamCherche As StrParam
    Set TPE_ParamCherche = PG_ProductParametres.Item(TPE_NomParam)
    If (Err.Number <> 0) Then
        Err.Clear
        TestParamExist = False
    Else
        TestParamExist = True
    End If
End Function

Public Sub EcritParam(EP_NomParam As String, EP_ValeurParam As String)
'test si le paramètre passé en argument existe dans le part.
'si oui lui affecte la valeur passée en argument,
'sinon le crée et lui affecte la valeur passée en argument
Dim EP_ParamTemp As StrParam
    If TestParamExist(EP_NomParam) Then
        PG_ProductParametres.Item(EP_NomParam).Value = EP_ValeurParam
    Else
        Set EP_ParamTemp = PG_ProductParametres.CreateString(EP_NomParam, EP_ValeurParam)
    End If
End Sub

Public Function LectureParam(LP_NomParam As String) As String
'Renvoi le contenu du paramètre passé en argument dans le part de la grille.
'Si le paramètre n'existe pas renvois une chaine vide
    If TestParamExist(LP_NomParam) Then
        LectureParam = PG_ProductParametres.Item(LP_NomParam).Value
    Else
        LectureParam = ""
    End If
End Function

Public Function CherchParam(nParam As String) As String
'Cherche le paramètre passé en argument dans le part de la grille
'Cas ou le nom du paramètre comprte l'arborescence du part Part/propriètès/nomparam
'Si le paramètre existe renvoi sa valeur sino renvoi une chaine vide
Dim tmpParam As StrParam
    CherchParam = ""
    For Each tmpParam In PG_ProductParametres
        If InStr(1, tmpParam.Name, nParam, vbTextCompare) > 0 Then
            CherchParam = tmpParam.Value
        End If
    Next
End Function

Property Get Prm_Material()
'Renvoi le contenu du paramètre "MATERIAL"
    Prm_Material = LectureParam(nPrmMaterial)
End Property

Property Get Prm_Recognition()
'Renvoi le contenu du paramètre "RECOGNITION"
    Prm_Recognition = LectureParam(nPrmRecogn)
End Property

Property Get Prm_Observation()
'Renvoi le contenu du paramètre "OBSERVATIONS"
    Prm_Observation = LectureParam(nPrmObserv)
End Property

Property Get Prm_DescriptionRef()
'Renvoi le contenu du paramètre "DescriptionRef"
'Cas particulier paramètre de base sur les Parts
    Prm_DescriptionRef = PG_Product.DescriptionRef
End Property

Property Get Prm_Dtemplate()
'Renvoi le contenu du paramètre "DTEMPLATE"
    Prm_Dtemplate = LectureParam(nPrmDtempl)
End Property

Property Get Prm_xNumoutillage()
'Renvoi le contenu du paramètre "xNUMOUTILLAGE"
    Prm_xNumoutillage = LectureParam(nPrmNumout)
End Property

Property Get Prm_xExemplaire()
'Renvoi le contenu du paramètre "xExemplaire"
    Prm_xExemplaire = LectureParam(nPrmExempl)
End Property

Property Get Prm_xDesignation()
'Renvoi le contenu du paramètre "xDESIGNATION"
    Prm_xDesignation = LectureParam(nPrmDesign)
End Property

Property Get Prm_Param_Assembl()
'Renvoi le contenu du paramètre "Param_Assembl"
    Prm_Param_Assembl = LectureParam("Param_Assembl")
End Property

Property Get Prm_xPiecepercees()
'Renvoi le contenu du paramètre "xPIECESPERCEES"
    Prm_xPiecepercees = LectureParam(nPrmPiecPer)
End Property

Property Get Prm_xSite()
'Renvoi le contenu du paramètre "xSITE"
    Prm_xSite = LectureParam(nPrmSite)
End Property

Property Get Prm_xNoprogavion()
'Renvoi le contenu du paramètre "xNOPROGAVION"
    Prm_xNoprogavion = LectureParam(nPrmProgAv)
End Property

Public Sub Create_HyBridShape(PG_HSName As String)
'Crée le set géométrique passé en argument
Dim tHB As HybridBody
    Set tHB = Hbodies.Add()
    tHB.Name = HSName
End Sub

'#############################################################
'# Les Collections d'objets
'#############################################################
Private Function LstFasteners() As c_Fasteners
'Renvoi la collection des fasterners
Dim tHB As HybridBody
Dim tmpFasts As c_Fasteners
Set tmpFasts = New c_Fasteners
Dim tmpFast As c_Fastener
Set tmpFast = New c_Fastener
Dim tmp_Param As Parameter
Dim tFast As HybridShape
Dim i As Long

    On Error GoTo Erreur1
    Set tHB = Hb(nHBRefExtIsol)
    'For Each tFast In thb.HybridShapes
    On Error GoTo Erreur2
    For i = 1 To tHB.HybridShapes.Count
        Set tFast = tHB.HybridShapes.Item(i)
        'test si c'est un "Fastener"
        If (TypeName(tHB.HybridShapes.Item(i)) = "HybridShapeInstance") Then
        'If TypeName(tFast) = "HybridShapeInstance" Then
            'Collecte les attibuts
            tmpFast.nom = tFast.Name
            Set tmp_Param = tFast.GetParameter("Comments")
            tmpFast.Comments = tmp_Param.Value
            Set tmp_Param = tFast.GetParameter("Fastener_Diameter")
            tmpFast.FastDiam = tmp_Param.Value
            Set tmp_Param = tFast.GetParameter("Xe")
            tmpFast.Xe = tmp_Param.Value
            Set tmp_Param = tFast.GetParameter("Ye")
            tmpFast.Ye = tmp_Param.Value
            Set tmp_Param = tFast.GetParameter("Ze")
            tmpFast.Ze = tmp_Param.Value
            Set tmp_Param = tFast.GetParameter("Xdir")
            tmpFast.Xdir = tmp_Param.Value
            Set tmp_Param = tFast.GetParameter("Ydir")
            tmpFast.Ydir = tmp_Param.Value
            Set tmp_Param = tFast.GetParameter("Zdir")
            tmpFast.Zdir = tmp_Param.Value
            'Collecte des coordonnnées
'                tmpFast.X = CoordFast(tFast, "X")
'                tmpFast.Y = CoordFast(tFast, "Y")
'                tmpFast.Z = CoordFast(tFast, "Z")
            
           tmpFasts.Add tmpFast.nom, tmpFast.Comments, tmpFast.FastDiam, tmpFast.Xe, tmpFast.Ye, tmpFast.Ze, tmpFast.Xdir, tmpFast.Ydir, tmpFast.Zdir, tmpFast.X, tmpFast.Y, tmpFast.Z
        End If
    Next
    Set LstFasteners = tmpFasts
    GoTo Fin

Erreur1:
    Set LstFasteners = Nothing
    GoTo Fin
Erreur2:
    Err.Raise vbObjectError + 517, , "Erreur dans la collecte des Fasteners"
Fin:
    'liberation des objets
    Set tmpFast = Nothing
    Set tmpFasts = Nothing
End Function

Public Property Get Fasteners() As c_Fasteners
'Renvoi la collection des Fasteners
    Set Fasteners = PG_Fasteners
End Property


