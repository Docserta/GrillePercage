Attribute VB_Name = "C_Creation_Lignes_PointsAB"
Option Explicit

'*********************************************************************
'* Macro : C_Creation_Lignes_PointsAB
'*
'* Fonctions :  Création des Pts A et des Pts B
'*              Nommage des points
'*
'* Version 7.0
'* Création :  SVI
'* Modification : 15/05/14 CFR
'*              Optimisation de la fonction de controle de la pré-éxistance des pts.
'*              Création d'une seule fonction (Check_PtExist)
'*              au lieu de 3 (check_Aexist, check_Bexist, check_faux_exist)
'*              Ajout d'une barre de progression
'* Modification : 31/07/14 CFR
'*              Regroupement des macros C1_Creation_Lignes_PointsAB et C2_Creation_Lignes_PointsAB_comments
'*              Ajout d'une boite de dialogue permetant de choisir le type de grille (avec ou sans Fastener)
'*              ajout d'une procédure de création des lignes par sélection graphique Point/direction pour les grille sans Fastener
'* Modification : 15/02/15 CFR
'*              Intégration de la classe PartGrille
'*              Ajout possibilité de creéer des pts et ligne sTD pour un seul UDF au lieu du set ref externe isolés entier
'* Modification : 22/04/15 CFR
'*              Modification de la sélection dans la procédure "SelectPoints"
'*              pour permettre la selection multiple
'* Modification : 24/06/16 CFR
'*                remplacement du tableau Coll_RefExIsol par la classe c_Fasteners
'* Modification : 19/01/17 CFR
'*               Ajout Inversion sens du STD
'* Modification : 10/04/17 CFR
'*               Ajout création ligne STD perpendiculaires surf 0
'**********************************************************************

Sub CATMain()

'Log de l'utilisation de la macro
LogUtilMacro nPath, nFicLog, nMacro, "C_Creation_Lignes_PointsAB", VMacro

Dim NumPt3D As Integer, SelPts As Integer
Dim TypeSTD As String
Dim GrilleSelection As Selection
Dim visProperties1 As VisPropertySet
Dim instance_catpart_grille_nue As PartDocument
Dim mBar As New c_ProgressBar
Dim GrilleActive As c_PartGrille
Dim TestHBody As HybridBody
Dim TestHShape As HybridShape
Dim InvertSTD As Boolean

'---------------------------
' Checker l'environnement
'---------------------------
  
    On Error Resume Next
    Set instance_catpart_grille_nue = CATIA.ActiveDocument
    If Err.Number <> 0 Then
        Err.Clear
        MsgBox "Le document de la fenêtre courante n'est pas un CATPart !", vbCritical, "Environnement incorrect"
        End
    End If
    On Error GoTo 0
    
Set GrilleActive = New c_PartGrille
    
    'Vérification de l'existence des sets géométriques
    On Error GoTo Erreur
    Set TestHBody = GrilleActive.Hb(nHBPtA)
    Set TestHBody = GrilleActive.Hb(nHBPtB)
    Set TestHBody = GrilleActive.Hb(nHBStd)
    Set TestHBody = GrilleActive.Hb(nHBRefExtIsol)
    'Vérification de l'existence des surf0 et surf100
    Set TestHShape = GrilleActive.HS(nSurf0, nHBS0)
    Set TestHShape = GrilleActive.HS(nSurf100, nHBS100)
    Set TestHBody = Nothing
    Set TestHShape = Nothing
    On Error GoTo 0

'Ouvre la boite de dlg "Frm_CreationPtA"
    Load Frm_CreationPtA
    Frm_CreationPtA.Show
 
'Sort du programme si click sur bouton Annuler dans FRM_DonnEntre
    If Not (Frm_CreationPtA.ChB_OkAnnule) Then
        Unload Frm_CreationPtA
        Exit Sub
    End If
    
'Stockage des choix de la boite de dialogue avant fermeture
    If Frm_CreationPtA.RbtNumNomStd Then
        NumPt3D = 1
    ElseIf Frm_CreationPtA.RbtNumCommentStd Then
        NumPt3D = 2
    ElseIf Frm_CreationPtA.RbtNumOrdre Then
        NumPt3D = 3
    End If
    
    If Frm_CreationPtA.Rbt_SelPts Then
        SelPts = 1
    ElseIf Frm_CreationPtA.Rbt_SelSetRef Then
        SelPts = 2
    End If
    
    If Frm_CreationPtA.Rbt_RefSTD Then
        TypeSTD = "RefSTD"
    ElseIf Frm_CreationPtA.Rbt_RefLEgacy Then
        TypeSTD = "RefLegacy"
    ElseIf Frm_CreationPtA.Rbt_RefPerpSurf0 Then
        TypeSTD = "RefPrpSurf0"
    End If
    
    'Invertion des STD
    If Frm_CreationPtA.CB_InvertSTD Then InvertSTD = True Else InvertSTD = False
    
    Frm_CreationPtA.Hide
    Unload Frm_CreationPtA

'cacher points de construction
    Set GrilleSelection = GrilleActive.GrilleSelection
    If GrilleActive.Exist_HB(nHBPtConst) Then
        GrilleSelection.Add GrilleActive.Hb(nHBPtConst)
        Set visProperties1 = GrilleSelection.VisProperties
        visProperties1.SetShow 1
    End If
    GrilleSelection.Clear

'affichage de la barre de progression
    mBar.ProgressTitre 1, " Création des Pts A et B, veuillez patienter."
    
'Création des Faux Pt A et B et des STD
    Select Case TypeSTD
        Case "RefSTD"   'Création des Faux Pt A et B et des STD
            'Selection des points
            mBar.Cache
            SelectPoints SelPts, GrilleActive
            mBar.Affiche
            'Creation des lignes Std
            If Not (CreateStdFastener(GrilleActive, NumPt3D, InvertSTD, mBar)) Then
                MsgBox "Erreur durant la création des Faux Points A, B et STD"
            End If
        Case "RefLegacy" 'Création des STD a partir des lignes Legacy
            If (CreateStdLegacy(GrilleActive, mBar) < 0) Then
                MsgBox "Erreur durant la création des STD Legacy"
            End If

        Case "RefPrpSurf0"    'Création des lignes perpendiculaires à la sur
            'Selection des points
            mBar.Cache
            SelectPoints SelPts, GrilleActive
            mBar.Affiche
            'Création des lignes perpendiculaires à la surface 0
            If Not (CreateStdPerpSurf0(GrilleActive, mBar)) Then
                MsgBox "Erreur durant la création des lignes perpendiculaires à la surface 0"
            End If
    End Select
    'End If
    
'Création des Pt A et B
    If Not (Create_Pt(GrilleActive, TypeSTD, NumPt3D, mBar)) Then
        MsgBox "Erreur durant la création des points A et B"
    End If
  
'Libération de la classe
   Set GrilleActive = Nothing
    Set mBar = Nothing
    
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

Sub SelectPoints(SP_Type As Integer, ByRef GrilleActive)
'Renvois une sélection des points a traiter
'Si type = 1 -> selection manuelle des points
'Si type =2 -> tous le set géométrique "Reférence externes isolées"
Dim i As Integer
Dim TypRef As String
Dim tab_selection(0)
    tab_selection(0) = "HybridShape"
Dim Retour_Selection As String
    Retour_Selection = ""
Dim MsgSel As String
    MsgSel = "Sélectionnez les UDF dans la fenètre graphique ou dans le set géométrique Ref externe isolées"
    
    If SP_Type = 1 Then ' selection manuelle des références externes isolées
        Retour_Selection = GrilleActive.GrilleSelection.SelectElement3(tab_selection, MsgSel, True, CATMultiSelTriggWhenUserValidatesSelection, False)
        If Retour_Selection = "Cancel" Then
            MsgBox "Selection graphique des UDF abandonée !", vbCritical, "Erreur de sélection"
            End
        End If
    ElseIf SP_Type = 2 Then 'tous le set géométrique "Reférence externes isolées"
        For i = 1 To GrilleActive.Hb(nHBRefExtIsol).HybridShapes.Count
            TypRef = TypeName(GrilleActive.Hb(nHBRefExtIsol).HybridShapes.Item(i))
            If TypRef = "HybridShapeInstance" Or TypRef = "HybridShapePointExplicit" Then
                GrilleActive.GrilleSelection.Add GrilleActive.Hb(nHBRefExtIsol).HybridShapes.Item(i)
            End If
        Next
    End If

End Sub
Function Create_Pt(ByRef GrilleActive, CP_TypeSTD As String, CP_Num As Integer, mBar) As Boolean
'Création des points A et Points B
'GrilleActive = Part Actif
'CP_typeSTD type de STD "RefSTD" , "RefLegacy" ou "RefPrpSurf0"
'CP_Num type de numérotation des points  1=(nom UDF), 2=(comments) ou 3=(A1)

Dim Comments As String, Name_STD As String

Dim Nom_FauxPtA_Parent As String, No_FauxPtA_Parent As String, Nom_NewPt As String, No_NewPt As String
Dim Nb_Std As Long
Dim cpt As Long, i As Integer
    cpt = 1
Dim mHSIntersectA As HybridShapeIntersection, mHSIntersectB As HybridShapeIntersection
Dim mHSIntersectSTD


    mBar.Titre = " Création des droites AB, veuillez patienter."

While (cpt <= GrilleActive.GrilleSelection.Count)
    Nb_Std = GrilleActive.Hb(nHBStd).HybridShapes.Count
    If CP_TypeSTD = "RefSTD" Then
        'Recupération du Numero du point d'origine de la ligne std pour nommer les points A et B
        For i = 1 To Nb_Std
            Nom_FauxPtA_Parent = GrilleActive.Hb(nHBStd).HybridShapes.Item(i).PtOrigine.DisplayName
            Nom_NewPt = Right(Nom_FauxPtA_Parent, Len(GrilleActive.GrilleSelection.Item(cpt).Value.Name))
            'Recherche parmis les lignes STD celle dont le point d'origine porte le même nom que l'UFD sélectionné
            If GrilleActive.GrilleSelection.Item(cpt).Value.Name = Nom_NewPt Then
                Set mHSIntersectSTD = GrilleActive.Hb(nHBStd).HybridShapes.Item(i)
                'Récupération du N° du point parent (si fauxA49-xxx on récupère 49)
                No_NewPt = Left(Nom_FauxPtA_Parent, InStr(Nom_FauxPtA_Parent, "-") - 1)
                No_NewPt = Right(No_NewPt, Len(No_NewPt) - 6) '"faux A " = 6 caractères
                Exit For
            End If
        Next
        Comments = GrilleActive.GrilleSelection.Item(cpt).Value.GetParameter("Comments").Value
        Name_STD = GrilleActive.GrilleSelection.Item(cpt).Value.Name
        If IsEmpty(mHSIntersectSTD) Then
            MsgBox "La ligne STD n'a pas été trouvée. Impossible de créer les points A et point B", vbInformation, "Elément manquant"
            End
        End If
    ElseIf CP_TypeSTD = "RefPrpSurf0" Then
        'Recupération du Numero du point de reference de la ligne std pour nommer les points A et B
        For i = 1 To Nb_Std
            Nom_FauxPtA_Parent = GrilleActive.Hb(nHBStd).HybridShapes.Item(i).Point.DisplayName
            If GrilleActive.GrilleSelection.Item(cpt).Value.Name = Nom_FauxPtA_Parent Then
                'Récupération du N° du point parent
                Set mHSIntersectSTD = GrilleActive.Hb(nHBStd).HybridShapes.Item(i)
                No_NewPt = i
                Exit For
            End If
        Next i
        If IsEmpty(mHSIntersectSTD) Then
            MsgBox "La ligne STD n'a pas été trouvée. Impossible de créer les points A et point B", vbInformation, "Elément manquant"
            End

        End If
    End If
    
'Maj Barre de progression
    mBar.Progression = 50 + (50 / GrilleActive.GrilleSelection.Count) * cpt

'Point A
    If (Check_PtExist(GrilleActive.Hb(nHBPtA), Name_STD) <> 1) Then
        'Création de l'intersection entre la ligne STD et la surface à 0
        Set mHSIntersectA = GrilleActive.HShapeFactory.AddNewIntersection(mHSIntersectSTD, GrilleActive.HS(nSurf0, nHBS0))
        mHSIntersectA.PointType = 0
        'Renommage du point
        Select Case CP_Num
        Case 1 ' A1-Nom du STD

            mHSIntersectA.Name = "A" & No_NewPt & "-" & Name_STD
            mBar.Etape = "A" & No_NewPt & "-" & Name_STD
        Case 2 ' A1-comments du STD
            mHSIntersectA.Name = "A" & No_NewPt & "-" & Comments
            mBar.Etape = "A" & No_NewPt & "-" & Comments
        Case 3 ' A1
            mHSIntersectA.Name = "A" & No_NewPt
            mBar.Etape = "A" & No_NewPt

        End Select
        GrilleActive.Hb(nHBPtA).AppendHybridShape mHSIntersectA
    End If
'Point B
    If (Check_PtExist(GrilleActive.Hb(nHBPtB), Name_STD) <> 1) Then
        'Création de l'intersection entre la ligne STD et la surface à 100
        Set mHSIntersectB = GrilleActive.HShapeFactory.AddNewIntersection(mHSIntersectSTD, GrilleActive.HS(nSurf100, nHBS100))
        mHSIntersectB.PointType = 0
        'Renommage du point
        Select Case CP_Num
        Case 1 ' B1-Nom du STD

            mHSIntersectB.Name = "B" & No_NewPt & "-" & Name_STD
            mBar.Etape = "B" & No_NewPt & "-" & Name_STD
        Case 2 ' B1-comments du STD
            mHSIntersectB.Name = "B" & No_NewPt & "-" & Comments
            mBar.Etape = "B" & No_NewPt & "-" & Comments
        Case 3 ' B1
            mHSIntersectB.Name = "B" & No_NewPt
            mBar.Etape = "B" & No_NewPt
        End Select
        GrilleActive.Hb(nHBPtB).AppendHybridShape mHSIntersectB
    End If
    
    cpt = cpt + 1
Wend

GrilleActive.PartGrille.Update
Create_Pt = True

End Function

Public Function CreateStdLegacy(ByRef GrilleActive, mBar) As Integer
'Création des droites STD a partir des lignes (legacy) collée dans le set références externes isolées
'la macro crée un point aux extrémités de la droite puis une droite entre ces point étendue de 100 mm de chaque cotés.

'Selection des points
Dim varfilter(0) As Variant
Dim objSel As Selection
Dim objSelLB As Object
Dim strReturn As String
    strReturn = "Normal"
Dim msg, Msg2, strMsgPt1, strMsgLine  As String
Dim CSL_HBShapeLinePtLineDir As HybridShapeLinePtDir
Dim CSL_HSDirection As HybridShapeDirection
Dim CSL_HShapeFactory As HybridShapeFactory
Set CSL_HShapeFactory = GrilleActive.HShapeFactory
Dim PtCoord1, LineDir As Reference
Dim LignDirName As String


    Set objSel = GrilleActive.partDocGrille.Selection
    Set objSelLB = objSel

    msg = "Pour chaque référence externe isolée, sélectionnez:" & Chr(13) & "1) le point de l'extrémité de la ligne dans le part." & Chr(13) & "2) puis la ligne." & Chr(13) & "Appuyez sur 'Echap' pour arréter la sélection."
    Msg2 = "Sélection des Legacy"
    strMsgPt1 = "Sélectionnez le point de l'extrémité de la ligne dans le part"
    strMsgLine = "Sélectionnez la ligne dans le part"

    MsgBox msg, vbInformation, Msg2
    
    Do While strReturn = "Normal"
        varfilter(0) = "Vertex"
        objSel.Clear
        strReturn = objSelLB.SelectElement2(varfilter, strMsgPt1, True)
        If (strReturn = "Cancel") Then Exit Do
        Set PtCoord1 = objSel.Item(1).Value
        
        objSel.Clear
        varfilter(0) = "Line"
        strReturn = objSelLB.SelectElement2(varfilter, strMsgLine, True)
        If (strReturn = "Cancel") Then Exit Do
        Set LineDir = objSel.Item(1).Value
        LignDirName = objSel.Item(1).Value.Name
        Set CSL_HSDirection = CSL_HShapeFactory.AddNewDirection(LineDir)
                
        Set CSL_HBShapeLinePtLineDir = CSL_HShapeFactory.AddNewLinePtDir(PtCoord1, CSL_HSDirection, -50#, 200#, True)
        GrilleActive.Hb(nHBStd).AppendHybridShape CSL_HBShapeLinePtLineDir
        CSL_HBShapeLinePtLineDir.Name = LignDirName
    
    Loop
    
    GrilleActive.PartGrille.Update
End Function
 
Public Function CreateStdFastener(ByRef GrilleActive, CSF_Numerotation As Integer, InvertSTD, mBar) As Boolean
'Création des droites STD a partir des fasteners collée dans le set références externes isolées
'la macro crée des points "FauxA" et FauxB" avec les coordonnées récupérées dans les paramètres du fasteners
'Puis crée une droite entre les pts FauxA et FauxB
On Error GoTo Err_CreateStdFastener
Dim tLisfast As c_Fasteners
'Set tLisfast = New c_Fasteners
Set tLisfast = GrilleActive.Fasteners
Dim tFast As c_Fastener
Set tFast = New c_Fastener

Dim cpt As Long
    cpt = 1
Dim Xe1 As Double, Ye1 As Double, Ze1 As Double
Dim Xe2 As Double, Ye2 As Double, Ze2 As Double
Dim Name_Input As String
Dim SelName As String
Dim i As Integer
Dim FauxA As HybridShapePointCoord, FauxB As HybridShapePointCoord

    While (cpt <= GrilleActive.GrilleSelection.Count)
        mBar.Progression = ((50 / GrilleActive.GrilleSelection.Count) * cpt)
        SelName = GrilleActive.GrilleSelection.Item(cpt).Value.Name
        'Recherche du fastener dans la collection
        Set tFast = tLisfast.Item(SelName)
        Select Case CSF_Numerotation
            Case 1
                Name_Input = "-" & tFast.nom
            Case 2
                Name_Input = "-" & tFast.Comments
            Case 3
                Name_Input = ""
        End Select
        
        If InvertSTD Then
            Xe1 = tFast.Xe
            Ye1 = tFast.Ye
            Ze1 = tFast.Ze
            Xe2 = tFast.Xe - 100 * tFast.Xdir
            Ye2 = tFast.Ye - 100 * tFast.Ydir
            Ze2 = tFast.Ze - 100 * tFast.Zdir
        Else
            Xe1 = tFast.Xe
            Ye1 = tFast.Ye
            Ze1 = tFast.Ze
            Xe2 = tFast.Xe + 100 * tFast.Xdir
            Ye2 = tFast.Ye + 100 * tFast.Ydir
            Ze2 = tFast.Ze + 100 * tFast.Zdir
        End If
        
        'Création faux pt A
        Set FauxA = Create_PtCoord(Xe1, Ye1, Ze1, "faux A" & cpt & Name_Input, GrilleActive)
        mBar.Etape = "faux A" & cpt & Name_Input

        'Création faux pt A
        Set FauxB = Create_PtCoord(Xe2, Ye2, Ze2, "faux B" & cpt & Name_Input, GrilleActive)
        mBar.Etape = "faux B" & cpt & Name_Input
        
        'Creation de la ligne STD
        If Create_Line_PtPt(FauxA, FauxB, GrilleActive, "Line." & cpt & Name_Input) Then
        End If
        cpt = cpt + 1
    Wend
    
    GrilleActive.PartGrille.Update
    CreateStdFastener = True
    GoTo Quit_CreateStdFastener

Err_CreateStdFastener:
    'MsgBox Err.Number & "  " & Err.Description, vbCritical, "Erreur"
    CreateStdFastener = False
    GoTo Quit_CreateStdFastener
   
Quit_CreateStdFastener:
    Set tFast = Nothing
    Set tLisfast = Nothing

End Function
 
Private Function CreateStdPerpSurf0(ByRef GrilleActive, mBar) As Boolean
'Création des droites STD a partir des points
'la macro crée une droite perpendiculaire à la surface 0 passant par le Point

Dim mHSfact As HybridShapeFactory
Dim HFSurf0 As HybridShape
Dim mPt As HybridShape
Dim NoSTD As Long
Dim mHSNormale As HybridShapeLineNormal

Dim cpt As Long
    cpt = 1

    Set mHSfact = GrilleActive.PartGrille.HybridShapeFactory
    Set HFSurf0 = GrilleActive.HS(nSurf0, nHBS0)
    GrilleActive.PartGrille.InWorkObject = GrilleActive.Hb(nHBStd)
    'Récupère le nombre de lignes STD pour poursuivre la numérotation
    NoSTD = GrilleActive.Hb(nHBStd).HybridShapes.Count
    'Boucle sur la sélection des points
    While (cpt <= GrilleActive.GrilleSelection.Count)
        NoSTD = NoSTD + 1
        Set mPt = GrilleActive.GrilleSelection.Item(cpt).Value
    
        Set mHSNormale = mHSfact.AddNewLineNormal(HFSurf0, mPt, 20#, 20#, False)
        mHSNormale.FirstUptoElem = HFSurf0
        mHSNormale.Name = "Line." & NoSTD
        GrilleActive.Hb(nHBStd).AppendHybridShape mHSNormale
        GrilleActive.PartGrille.InWorkObject = mHSNormale
        cpt = cpt + 1
        
    Wend
        GrilleActive.PartGrille.Update
        CreateStdPerpSurf0 = True
        GoTo Quit_CreateStdPerpSurf0
    

    
Err_CreateStdPerpSurf0:
    CreateStdPerpSurf0 = False
    GoTo Quit_CreateStdPerpSurf0
    
Quit_CreateStdPerpSurf0:

End Function

 
