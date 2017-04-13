Attribute VB_Name = "D2_Agrafes"

Option Explicit


'*********************************************************************
'* Macro : D2_Agrafes
'*
'* Fonctions :  Identification des agrafes
'*
'* Version : 8
'* Création :  CFR
'* Modification :
'*
'**********************************************************************
Sub CATMain()

'Log de l'utilisation de la macro
LogUtilMacro nPath, nFicLog, nMacro, "D2_Agrafes", VMacro

CheminSourcesMacro = Get_Active_CATVBA_Path

Dim HBShape_STD As HybridShapes 'Collection des STD
Dim HBShape_Std_EC As HybridShape 'STD en cours
Dim Std_Parameters As Parameters 'Collection des paramètre de la droite STD en cours
Dim StatusPartUpdated As Boolean
Dim NoAgrafe As String 'Paramètre qui va porter le Numéro de l'agrafe
Dim Nb_Pt_Sel As Long
Dim PtAName As String, PtImportName As String  ' Nom du point A et nom du point "TempPtax"
Dim cpt As Long
Dim PointImport As HybridShapeIntersection 'Point d'origine du triedre d'import du composant
Dim GrilleActive As New c_PartGrille
Dim instance_catpart_grille_nue As PartDocument
Dim mBar As c_ProgressBar
Dim TrouLame As Boolean
Dim ParamDiaLamage As Dimension
Dim TestHBody As HybridBody
Dim TestHShape As HybridShape

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
    
'Test le chemin de la bibli des composants
    CheminBibliComposants = CorrigeDFS()
    
    Set GrilleActive = New c_PartGrille

'test l'existence des Sets géométriques
    On Error GoTo Erreur
    Set TestHBody = GrilleActive.Hb(nHBPtConst)
    Set TestHBody = GrilleActive.Hb(nHBStd)
    'Test l'existance de la droite d'orientation grille
    Set TestHShape = GrilleActive.OrientationGrille
    Set TestHBody = Nothing
    Set TestHShape = Nothing
    On Error GoTo 0

'Vérifie que le part est Update
    StatusPartUpdated = GrilleActive.PartGrille.IsUpToDate(GrilleActive.PartGrille)
    If StatusPartUpdated = False Then
        MsgBox "La part n'est pas a update. Faites la mise à jour avant de lancer cette macro.", vbInformation, "Part not Update"
        End
    End If
    
'Chargement de la boite de dialogue
    CollMachines.OpenBibliMachine = CheminBibliComposants & "\" & NomFicInfoMachines
    Load FRM_Agrafage
    
    FRM_Agrafage.Show
    
    If Not (FRM_Agrafage.ChB_OkAnnule) Then
        Unload FRM_Agrafage
        Exit Sub
    End If
    
'Stockage des infos du formulaire
    NoAgrafe = FRM_Agrafage.CBL_NumAgrafe
    Unload FRM_Agrafage
    Nb_Pt_Sel = GrilleActive.GrilleSelection.Count
    
'Progress Barre
    Set mBar = New c_ProgressBar
    mBar.ProgressTitre 1, " Ajout des attributs d'agrafage sur les STD, veuillez patienter."

    Set HBShape_STD = GrilleActive.Hb(nHBStd).HybridShapes

    For cpt = 1 To Nb_Pt_Sel
        'Extraction du radical du nom du points A (avant le "-"
        If InStr(1, Left(Tab_Select_Points(0, cpt - 1), 4), "-", vbTextCompare) = 0 Then
            PtAName = Tab_Select_Points(0, cpt - 1)
        Else
            PtAName = Left(Tab_Select_Points(0, cpt - 1), InStr(1, Left(Tab_Select_Points(0, cpt - 1), 4), "-", vbTextCompare) - 1)
        End If
        
        mBar.ProgressTitre ((100 / Nb_Pt_Sel) * cpt), " Création du trou " & PtAName & ", veuillez patienter."
        
        'Active le Set Travail
        GrilleActive.PartGrille.InWorkObject = GrilleActive.Hb(nHBTrav)
        
        'Ajout des paramètres sur le STD en cours
        Set HBShape_Std_EC = HBShape_STD.Item(Tab_Select_Points(2, cpt - 1))
        Set Std_Parameters = GrilleActive.PartGrille.Parameters.SubList(HBShape_Std_EC, True)
        CreateParamExistString Std_Parameters, "NoAgrafe", NoAgrafe
        
        'Vérification si ce trou est lamé ou non
        On Error Resume Next
        Set ParamDiaLamage = Std_Parameters.Item("DiamLamageTrouNezMachine")
        If Err.Number <> 0 Then
            TrouLame = False
        Else
            TrouLame = True
        End If
        On Error GoTo 0
        
        If TrouLame Then
            'Récupération du point "PtInsertBague_xxxxx pour le point En cours
            PtImportName = "PtInsertBague_" & HBShape_Std_EC.Name
        Else
            'Récupération du point TempPtAx pour le point En cours
            PtImportName = "TempPt" & PtAName
        End If
        'test si ce point existe
        If GrilleActive.Exist_PT(PtImportName) Then
            Set PointImport = GrilleActive.Hb(nHBPtConst).HybridShapes.GetItem(PtImportName)
            'Création du triedre d'import de l'agrafe sur le Pt A
            Creation_Triedre_SurPt GrilleActive, PointImport, HBShape_Std_EC, "Agrafe" & PtAName
        Else
            MsgBox "le point " & PtImportName & "est absent du set Point de construction, l'agrafe ne sera pas importée.", vbCritical, "Elément manquant"
        End If
        
     Next
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
     Set mBar = Nothing

End Sub


