Attribute VB_Name = "H_Gravure"
Option Explicit
Sub catmain()

'*********************************************************************
'* Macro : H_Gravure
'*
'* Fonctions :  Création de gravures
'*              Crée un texte dans un drawing, le converti en IGES puis l'importe dans le part
'*
'* Version 2
'* Création : SVI
'* Modification : 02/06/2015
'*                Prise en compte de la classe "PartGrille"
'*                Ajout choix police
'* Modification : 21/04/2016
'*                Modification de la selection des éléments géométriques dans le fichier iges2D
'*                changement de set géometrique des sélection des esquisses gravure (travail\draftgravures ald \gravures
'*                Ajout des infos du DSCGP dans la gravure
'*
'**********************************************************************
'
Dim instance_catpart_grille_nue As PartDocument
Dim GrilleActive As c_PartGrille

Dim Ig2Doc As Document 'Fichier IGES temporaire
Dim Ig2Draw As DrawingDocument
Dim DrwDoc As DrawingDocument 'Fichier Catdrawing pour création du texte 2D
Dim IG2sheet As DrawingSheet
Dim Ig2View As DrawingView

Dim sketch_select As Selection 'sketcher support de la gravure
Dim name_sketch As String 'Nom du sketcher support de la gravure
Dim TargetSketch As Sketch

Dim DrwText As DrawingText 'Texte de la gravure
Dim ContenuText As String  'Contenu du texte a graver
Dim Txt_First1 As Long, Txt_indChar1 As Long, Txt_Val1 As Long 'paramètres de ratio du texte
Dim Txt_First2 As Long, Txt_indChar2 As Long, Txt_Val2 As Long 'paramètres d'espace du texte
Dim Txt_FontName As String, Txt_FontSize As String
Dim Ig2Selection As Selection 'Selection texte iges

Dim List_TaillesGravures As String 'Nom du fichier texte contenant les tailles de gravure
    List_TaillesGravures = "Grille_Taille_gravures.ini"
Dim Tab_TaillesGravures() As String 'Tableau des polices definies dans le fichier text Grille_Taille_gravures.ini
Dim NBParam_TailleGrav As Long ' Nombre de paramètre séparés par des ; pour chaque ligne de ce fichier
    NBParam_TailleGrav = 6
Dim Ligne_TailleGravure As String 'ligne du fichier List_TaillesGravures en cours de lecture
Dim fs, f
Dim Fichtxt As String
Dim i As Long, boucle As Long
Dim Liste_Sketch() As String
Dim Coll_GeomElems As GeometricElements
Dim GeomElem As GeometricElement
    
'On Error Resume Next
CheminSourcesMacro = Get_Active_CATVBA_Path

'Log de l'utilisation de la macro
LogUtilMacro nPath, nFicLog, nMacro, "H_Gravure", VMacro

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
    
    Set GrilleActive = New c_PartGrille

'test l'existence des Sets géométriques
    If Not (GrilleActive.Exist_HB(nHBGrav)) Then
        MsgBox "Le set Géométrique : " & (nHBGrav) & " est manquant ou mal orthograpié.", vbCritical, "Eléments manquants"
        End
    End If
    If Not (GrilleActive.Exist_HB(nHBDrGrav)) Then
        MsgBox "Le set Géométrique : " & (nHBDrGrav) & " est manquant ou mal orthograpié.", vbCritical, "Eléments manquants"
        End
    End If

'Fige l'affichage pour améliorer les performances
    'CATIA.HSOSynchronized = False
    CATIA.RefreshDisplay = False
    
'Récupère la liste des sketch dans le set "gravures"
    If GrilleActive.Hb(nHBDrGrav).HybridSketches.Count = 0 Then
        MsgBox "Le set géométrique Draft gravures ne contient aucun sketch support. Veuillez créez un sketch support avant de lancer la macro.", vbInformation, "Element manquant"
        End
    Else
        For i = 1 To GrilleActive.Hb(nHBDrGrav).HybridSketches.Count
            ReDim Preserve Liste_Sketch(i - 1)
            Liste_Sketch(i - 1) = GrilleActive.Hb(nHBDrGrav).HybridSketches.Item(i).Name
        Next
    End If

'Chargement de la boite de dialogue
    Load Frm_Gravure
    'Initialisation des listes déroulantes
    Frm_Gravure.CB_Support.List = Liste_Sketch
    'Liste des Tailles de texte
    Set fs = CreateObject("scripting.filesystemobject")
    Fichtxt = CheminSourcesMacro & List_TaillesGravures
    Set f = fs.opentextfile(Fichtxt, ForReading, 1)
    boucle = 0
    Do While Not f.AtEndOfStream
        Ligne_TailleGravure = f.ReadLine
        ReDim Preserve Tab_TaillesGravures(NBParam_TailleGrav - 1, boucle)
        For i = 1 To NBParam_TailleGrav
            Tab_TaillesGravures(i - 1, boucle) = Split_txt(Ligne_TailleGravure, i)
        Next i
        boucle = boucle + 1
    Loop
    Frm_Gravure.CBX_Taille.List = TranspositionTabl(Tab_TaillesGravures)
    Frm_Gravure.Tbx_NoGrille = GrilleActive.nom
                
    'Documente la référence et la désignation
    Frm_Gravure.Show
    If Not (Frm_Gravure.ChB_OkAnnule) Then
        End
    End If
          
    'Stockage des choix de la boite de dialogue
    name_sketch = Frm_Gravure.CB_Support
    'ContenuText = Frm_Gravure.LB_TextGravure
    Txt_First1 = CLng(Left(Frm_Gravure.TBX_Ratio, InStr(1, Frm_Gravure.TBX_Ratio, ",", vbTextCompare)))
    Txt_indChar1 = CLng(Mid(Frm_Gravure.TBX_Ratio, InStr(1, Frm_Gravure.TBX_Ratio, ",", vbTextCompare) + 1, InStr(2, Frm_Gravure.TBX_Ratio, ",", vbTextCompare) - 1))
    Txt_Val1 = CLng(Right(Frm_Gravure.TBX_Ratio, InStr(2, Frm_Gravure.TBX_Ratio, ",", vbTextCompare)))
    Txt_First2 = CLng(Left(Frm_Gravure.TBX_Espace, InStr(1, Frm_Gravure.TBX_Espace, ",", vbTextCompare)))
    Txt_indChar2 = CLng(Mid(Frm_Gravure.TBX_Espace, InStr(1, Frm_Gravure.TBX_Espace, ",", vbTextCompare) + 1, InStr(2, Frm_Gravure.TBX_Ratio, ",", vbTextCompare) - 1))
    Txt_Val2 = CLng(Right(Frm_Gravure.TBX_Espace, InStr(2, Frm_Gravure.TBX_Espace, ",", vbTextCompare)))
    Txt_FontName = Frm_Gravure.TBX_Police
    Txt_FontSize = Frm_Gravure.TBX_Taille
    
        
'Selection du sketcher support de gravure
    Set sketch_select = GrilleActive.GrilleSelection
    sketch_select.Clear
    
' create new drawing document
    Set DrwDoc = CATIA.Documents.Add("Drawing")
    
'Ajout des textes
    Select Case Frm_Gravure.CB_Face
    Case "Face Sup"
        ContenuText = InsLineSpace(ValDscgp.GravureSup)
    Case "Face Inf"
        ContenuText = InsLineSpace(ValDscgp.GravureInf)
    Case "Face Lat1"
        ContenuText = InsLineSpace(ValDscgp.GravureLat1)
    Case "Face Lat2"
        ContenuText = InsLineSpace(ValDscgp.GravureLat2)
    Case "Face Lat3"
        ContenuText = InsLineSpace(ValDscgp.GravureLat3)
    Case "Face Lat4"
        ContenuText = InsLineSpace(ValDscgp.GravureLat4)
End Select
    ContenuText = ContenuText & Chr(10)
    Set DrwText = DrwDoc.Sheets.Item(1).Views.Item(1).Texts.Add(ContenuText, 20, 20)

    Unload Frm_Gravure
      
        DrwText.TextProperties.FONTSIZE = CDbl(Txt_FontSize)
        DrwText.TextProperties.FONTNAME = Txt_FontName
        DrwText.TextProperties.Underline = 0
        
        'DrwText.SetParameterOnSubString catCharRatio, 0, 0, 60
        DrwText.SetParameterOnSubString catCharRatio, Txt_First1, Txt_indChar1, Txt_Val1
        
        'DrwText.SetParameterOnSubString catCharSpacing, 0, 0, 35
        DrwText.SetParameterOnSubString catCharSpacing, Txt_First2, Txt_indChar2, Txt_Val2
    '    DrwText.TextProperties.Update

' save document as .ig2 file
    CATIA.DisplayFileAlerts = False
    DrwDoc.ExportData "C:\temp\drawing1.ig2", "ig2"
    DrwDoc.Close

' open .ig2 document
    Set Ig2Doc = CATIA.Documents.Read("C:\temp\drawing1.ig2")
    
    Ig2Doc.SaveAs "c:\temp\ig2doc.catdrawing"
    CATIA.DisplayFileAlerts = True
    
    Ig2Doc.Close
    Set Ig2Draw = CATIA.Documents.Read("c:\temp\ig2doc.catdrawing")
    Set IG2sheet = Ig2Draw.Sheets.Item(1)
    Set Ig2View = IG2sheet.Views.Item(3)

    Set Coll_GeomElems = Ig2View.GeometricElements
    
    ' Set Ig2Selection = Ig2Draw.Selection Doesn't work
    Set Ig2Selection = CATIA.ActiveDocument.Selection
    
    For Each GeomElem In Coll_GeomElems
        Ig2Selection.Add GeomElem
    Next
    
' search for all polylines
'Ig2Selection.Search "type=Polyline + type=Line;all"
'Ig2Selection.Search "((CAT2DLSearch.2DPolyline + CATSketchSearch.2DPolyline) + CATDrwSearch.2DPolyline),all"
'Ig2Selection.Search "(((((((((CATStFreeStyleSearch.Line + CAT2DLSearch.2DLine) + CATSketchSearch.2DLine) + CATDrwSearch.2DLine) + CATPrtSearch.Line) + CATGmoSearch.Line) + CATSpdSearch.Line) + CAT2DLSearch.2DPolyline) + CATSketchSearch.2DPolyline) + CATDrwSearch.2DPolyline),all"

' copy polylines found to a buffer
    Ig2Selection.Copy

' get selection object of Part document
    GrilleActive.partDocGrille.Activate
    Set TargetSketch = GrilleActive.Hb(nHBDrGrav).HybridSketches.Item(name_sketch)

' open sketch for edition
    TargetSketch.OpenEdition

' paste contents of a buffer in the sketch
    GrilleActive.GrilleSelection.Clear
    GrilleActive.GrilleSelection.Add TargetSketch
    'ActiveDoc.Selection.Paste
    GrilleActive.GrilleSelection.Paste


' end sketch edition
    TargetSketch.CloseEdition
' update part
'GrilleActive.PartGrille.Update

' close Ig2 document
    Ig2Draw.Close
    Set Ig2Draw = Nothing
    
' restore default Selection perfomance
    CATIA.RefreshDisplay = True
    CATIA.DisplayFileAlerts = True
'CATIA.HSOSynchronized = True

'Libération des classes
Set GrilleActive = Nothing


End Sub


