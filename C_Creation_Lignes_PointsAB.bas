Attribute VB_Name = "C_Creation_Lignes_PointsAB"
Option Explicit

'*********************************************************************
'* Macro : C_Creation_Lignes_PointsAB
'*
'* Fonctions :  Cr�ation des Pts A et des Pts B
'*              Nommage des points
'*
'* Version 7.0
'* Cr�ation :  SVI
'* Modification : 15/05/14 CFR
'*              Optimisation de la fonction de controle de la pr�-�xistance des pts.
'*              Cr�ation d'une seule fonction (Check_PtExist)
'*              au lieu de 3 (check_Aexist, check_Bexist, check_faux_exist)
'*              Ajout d'une barre de progression
'* Modification : 31/07/14 CFR
'*              Regroupement des macros C1_Creation_Lignes_PointsAB et C2_Creation_Lignes_PointsAB_comments
'*              Ajout d'une boite de dialogue permetant de choisir le type de grille (avec ou sans Fastener)
'*              ajout d'une proc�dure de cr�ation des lignes par s�lection graphique Point/direction pour les grille sans Fastener
'* Modification : 15/02/15 CFR
'*              Int�gration de la classe PartGrille
'*              Ajout possibilit� de cre�er des pts et ligne sTD pour un seul UDF au lieu du set ref externe isol�s entier
'* Modification : 22/04/15 CFR
'*              Modification de la s�lection dans la proc�dure "SelectPoints"
'*              pour permettre la selection multiple
'* Modification : 24/06/16 CFR
'*                remplacement du tableau Coll_RefExIsol par la classe c_Fasteners
'* Modification : 19/01/17 CFR
'*               Ajout Inversion sens du STD
'**********************************************************************

Sub catmain()

'Log de l'utilisation de la macro
LogUtilMacro nPath, nFicLog, nMacro, "C_Creation_Lignes_PointsAB", VMacro

Dim NumPt3D As Integer, SelPts As Integer
Dim TypeSTD As String
Dim GrilleSelection As Selection
Dim visProperties1 As VisPropertySet
Dim instance_catpart_grille_nue As PartDocument
Dim Barre As New ProgressBarre
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
        MsgBox "Le document de la fen�tre courante n'est pas un CATPart !", vbCritical, "Environnement incorrect"
        End
    End If
    On Error GoTo 0
    
Set GrilleActive = New c_PartGrille
    
    'V�rification de l'existence des sets g�om�triques
    On Error GoTo Erreur
    Set TestHBody = GrilleActive.Hb(nHBPtA)
    Set TestHBody = GrilleActive.Hb(nHBPtB)
    Set TestHBody = GrilleActive.Hb(nHBStd)
    Set TestHBody = GrilleActive.Hb(nHBRefExtIsol)
    'V�rification de l'existence des surf0 et surf100
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
    Else
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
    Barre.ProgressTitre 1, " Cr�ation des Pts A et B, veuillez patienter."
    
'Cr�ation des Faux Pt A et B et des STD
    If TypeSTD = "RefSTD" Then 'Cr�ation des Faux Pt A et B et des STD
        Barre.Cache
        SelectPoints SelPts, GrilleActive
        Barre.Affiche
        'Creation des lignes Std
        If Not (CreateStdFastener(GrilleActive, NumPt3D, InvertSTD, Barre)) Then
            MsgBox "Erreur durant la cr�ation des Faux Points A, B et STD"
        End If
    ElseIf TypeSTD = "RefLegacy" Then 'Cr�ation des STD a partir des lignes Legacy
        If (CreateStdLegacy(GrilleActive, Barre) < 0) Then
            MsgBox "Erreur durant la cr�ation des STD Legacy"
        End If
    End If
    
'Cr�ation des Pt A et B
    If Not (Create_Pt(GrilleActive, TypeSTD, NumPt3D, Barre)) Then
        MsgBox "Erreur durant la cr�ation des points A et B"
    End If
  
'Lib�ration de la classe
   Set GrilleActive = Nothing
    Set Barre = Nothing
    
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
'Renvois une s�lection des points a traiter
'Si type = 1 -> selection manuelle des points
'Si type =2 -> tous le set g�om�trique "Ref�rence externes isol�es"
Dim i As Integer
Dim tab_selection(0)
    tab_selection(0) = "HybridShape"
Dim Retour_Selection As String
    Retour_Selection = ""
Dim MsgSel As String
    MsgSel = "S�lectionnez les UDF dans la fen�tre graphique ou dans le set g�om�trique Ref externe isol�es"
    
    If SP_Type = 1 Then ' selection manuelle des r�f�rences externes isol�es
        'v�rifie si une preselection existe et test si se sont des hybridshapes
        'If GrilleActive.GrilleSelection.Count > 0 Then ' une selection existe d�ja
        '    For i = 1 To GrilleActive.GrilleSelection.Count
        '        'Si la s�lection est un HybridShape on l'ajoute a la liste box du formulaire
        '        If GrilleActive.GrilleSelection.Item(i + 1).Value = "HybridShape" Then
        '
        '        End If
        '    Next
        'Else
            Retour_Selection = GrilleActive.GrilleSelection.SelectElement3(tab_selection, MsgSel, True, CATMultiSelTriggWhenUserValidatesSelection, False)
            If Retour_Selection = "Cancel" Then
                MsgBox "Selection graphique des UDF abandon�e !", vbCritical, "Erreur de s�lection"
                End
            End If
        'End If
    ElseIf SP_Type = 2 Then 'tous le set g�om�trique "Ref�rence externes isol�es"
        For i = 1 To GrilleActive.Hb(nHBRefExtIsol).HybridShapes.Count
            If TypeName(GrilleActive.Hb(nHBRefExtIsol).HybridShapes.Item(i)) = "HybridShapeInstance" Then
                GrilleActive.GrilleSelection.Add GrilleActive.Hb(nHBRefExtIsol).HybridShapes.Item(i)
            End If
        Next
    End If

End Sub
Function Create_Pt(ByRef GrilleActive, CP_TypeSTD As String, CP_Num As Integer, Barre) As Boolean
'Cr�ation des points A et Points B
'GrilleActive = Part Actif
'CP_typeSTD type de STD "RefSTD" ou "RefLegacy"
'CP_Num type de num�rotation des points  1=(nom UDF), 2=(comments) ou 3=(A1)

Dim Comments As String, Name_STD As String
Dim Nom_FauxPtA_Parent As String, No_FauxPtA_Parent As String, Nom_NewPt As String, No_NexPt As String
Dim hybridShapes_name As HybridShapes
Dim cPt As Long, i As Integer
    cPt = 1
Dim CP_HShapeIntersectionA As HybridShapeIntersection, CP_HShapeIntersectionB As HybridShapeIntersection
Dim CP_hybridShapeLineSTD

Barre.Titre = " Cr�ation des droites AB, veuillez patienter."

While (cPt <= GrilleActive.GrilleSelection.Count)
    'Recherche parmis les lignes STD celle dont le point d'origine porte le m�me nom que l'UFD s�lectionn�
    For i = 0 To GrilleActive.Hb(nHBStd).HybridShapes.Count - 1
        Nom_FauxPtA_Parent = GrilleActive.Hb(nHBStd).HybridShapes.Item(i + 1).PtOrigine.DisplayName
        Nom_NewPt = Right(Nom_FauxPtA_Parent, Len(GrilleActive.GrilleSelection.Item(cPt).Value.Name))
        If GrilleActive.GrilleSelection.Item(cPt).Value.Name = Nom_NewPt Then
            Set CP_hybridShapeLineSTD = GrilleActive.Hb(nHBStd).HybridShapes.Item(i + 1)
            'R�cup�ration du N� du point parent (si fauxA49-xxx on r�cup�re 49)
            No_NexPt = Left(Nom_FauxPtA_Parent, InStr(Nom_FauxPtA_Parent, "-") - 1)
            No_NexPt = Right(No_NexPt, Len(No_NexPt) - 6) '"faux A " = 6 caract�res
            Exit For
        End If
    Next
    If IsEmpty(CP_hybridShapeLineSTD) Then
        MsgBox "La ligne STD n'a pas �t� trouv�e. Impossible de cr�er les points A et point B", vbInformation, "El�ment manquant"
        End
    End If
    
'Maj barre de progression
    Barre.Progression = 50 + (50 / GrilleActive.GrilleSelection.Count) * cPt
    
    Comments = GrilleActive.GrilleSelection.Item(cPt).Value.GetParameter("Comments").Value
    Name_STD = GrilleActive.GrilleSelection.Item(cPt).Value.Name

'Point A
    If (Check_PtExist(GrilleActive.Hb(nHBPtA), Name_STD) <> 1) Then
        'Cr�ation de l'intersection entre la ligne STD et la surface � 0
        Set CP_HShapeIntersectionA = GrilleActive.HShapeFactory.AddNewIntersection(CP_hybridShapeLineSTD, GrilleActive.HS(nSurf0, nHBS0))
        CP_HShapeIntersectionA.PointType = 0
        'Renommage du point
        Select Case CP_Num
        Case 1 ' A1-Nom du STD
            CP_HShapeIntersectionA.Name = "A" & No_NexPt & "-" & Name_STD
            Barre.Etape = "A" & No_NexPt & "-" & Name_STD
        Case 2 ' A1-comments du STD
            CP_HShapeIntersectionA.Name = "A" & No_NexPt & "-" & Comments
            Barre.Etape = "A" & No_NexPt & "-" & Comments
        Case 3 ' A1
            CP_HShapeIntersectionA.Name = "A" & No_NexPt
            Barre.Etape = "A" & No_NexPt
        End Select
        GrilleActive.Hb(nHBPtA).AppendHybridShape CP_HShapeIntersectionA
    End If
'Point B
    If (Check_PtExist(GrilleActive.Hb(nHBPtB), Name_STD) <> 1) Then
        'Cr�ation de l'intersection entre la ligne STD et la surface � 100
        Set CP_HShapeIntersectionB = GrilleActive.HShapeFactory.AddNewIntersection(CP_hybridShapeLineSTD, GrilleActive.HS(nSurf100, nHBS100))
        CP_HShapeIntersectionB.PointType = 0
        'Renommage du point
        Select Case CP_Num
        Case 1 ' B1-Nom du STD
            CP_HShapeIntersectionB.Name = "B" & No_NexPt & "-" & Name_STD
            Barre.Etape = "B" & No_NexPt & "-" & Name_STD
        Case 2 ' B1-comments du STD
            CP_HShapeIntersectionB.Name = "B" & No_NexPt & "-" & Comments
            Barre.Etape = "B" & No_NexPt & "-" & Comments
        Case 3 ' B1
            CP_HShapeIntersectionB.Name = "B" & No_NexPt
            Barre.Etape = "B" & No_NexPt
        End Select
        GrilleActive.Hb(nHBPtB).AppendHybridShape CP_HShapeIntersectionB
    End If
    
    cPt = cPt + 1
Wend

GrilleActive.PartGrille.Update
Create_Pt = True

End Function

Public Function CreateStdLegacy(ByRef GrilleActive, Barre) As Integer
'Cr�ation des droites STD a partir des lignes (legacy) coll�e dans le set r�f�rences externes isol�es
'la macro cr�e un point aux extr�mit�s de la droite puis une droite entre ces point �tendue de 100 mm de chaque cot�s.

'Selection des points
Dim varfilter(0) As Variant
Dim objSel As Selection
Dim objSelLB As Object
Dim strReturn As String
    strReturn = "Normal"
Dim msg, Msg2, strMsgPt1, strMsgLine  As String

Set objSel = GrilleActive.partDocGrille.Selection
Set objSelLB = objSel

msg = "Pour chaque r�f�rence externe isol�e, s�lectionnez:" & Chr(13) & "1) le point de l'extr�mit� de la ligne dans le part." & Chr(13) & "2) puis la ligne." & Chr(13) & "Appuyez sur 'Echap' pour arr�ter la s�lection."
Msg2 = "S�lection des Legacy"
strMsgPt1 = "S�lectionnez le point de l'extr�mit� de la ligne dans le part"
strMsgLine = "S�lectionnez la ligne dans le part"

Dim CSL_HBShapeLinePtLineDir As HybridShapeLinePtDir
Dim CSL_HSDirection As HybridShapeDirection
Dim CSL_HShapeFactory As HybridShapeFactory
Set CSL_HShapeFactory = GrilleActive.HShapeFactory
Dim PtCoord1, LineDir As Reference
Dim LignDirName As String

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
 
Public Function CreateStdFastener(ByRef GrilleActive, CSF_Numerotation As Integer, InvertSTD, Barre) As Boolean
'Cr�ation des droites STD a partir des fasteners coll�e dans le set r�f�rences externes isol�es
'la macro cr�e des points "FauxA" et FauxB" avec les coordonn�es r�cup�r�es dans les param�tres du fasteners
'Puis cr�e une droite entre les pts FauxA et FauxB
On Error GoTo Err_CreateStdFastener
Dim tLisfast As c_Fasteners
'Set tLisfast = New c_Fasteners
Set tLisfast = GrilleActive.Fasteners
Dim tFast As c_Fastener
Set tFast = New c_Fastener

Dim cPt As Long
    cPt = 1
Dim Xe1 As Double, Ye1 As Double, Ze1 As Double
Dim Xe2 As Double, Ye2 As Double, Ze2 As Double
Dim Name_Input As String
Dim SelName As String
Dim i As Integer
Dim FauxA As HybridShapePointCoord, FauxB As HybridShapePointCoord

    While (cPt <= GrilleActive.GrilleSelection.Count)
        Barre.Progression = ((50 / GrilleActive.GrilleSelection.Count) * cPt)
        SelName = GrilleActive.GrilleSelection.Item(cPt).Value.Name
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
        
'        If InvertSTD Then
'            Xe1 = tFast.Xe + 100 * tFast.Xdir
'            Ye1 = tFast.Ye + 100 * tFast.Ydir
'            Ze1 = tFast.Ze + 100 * tFast.Zdir
'            Xe2 = tFast.Xe
'            Ye2 = tFast.Ye
'            Ze2 = tFast.Ze
'        Else
'            Xe1 = tFast.Xe
'            Ye1 = tFast.Ye
'            Ze1 = tFast.Ze
'            Xe2 = tFast.Xe + 100 * tFast.Xdir
'            Ye2 = tFast.Ye + 100 * tFast.Ydir
'            Ze2 = tFast.Ze + 100 * tFast.Zdir
'        End If
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
        
        'Cr�ation faux pt A
        Set FauxA = Create_PtCoord(Xe1, Ye1, Ze1, "faux A" & cPt & Name_Input, GrilleActive)
        Barre.Etape = "faux A" & cPt & Name_Input

        'Cr�ation faux pt A
        Set FauxB = Create_PtCoord(Xe2, Ye2, Ze2, "faux B" & cPt & Name_Input, GrilleActive)
        Barre.Etape = "faux B" & cPt & Name_Input
        
        'Creation de la ligne STD
        If Create_Line_PtPt(FauxA, FauxB, GrilleActive, "Line." & cPt & Name_Input) Then
        End If
        cPt = cPt + 1
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
 
 
 
