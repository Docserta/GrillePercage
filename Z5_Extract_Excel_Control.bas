Attribute VB_Name = "Z5_Extract_Excel_Control"
Option Explicit

Sub CATMain()
' *****************************************************************
' * Génère le rapport de controle au format excel
' * utilise le template situé dans le répertoire de la macro.
' *
' * Création CFR le : 06/06/2014
' * modification le : 01/09/14 - recuperation des tolérance sur Diamètre de perçage dans l'onglet "Diametre"
' * modification le : 18/12/14 - Remplacement du N° de lot par le N0 de grille dans la case C12 du rapport de controle
' * modification le : 02/02/15 - Extraction de coordonnées de points à partir d'un triedre sélectionné.
' *                            - Prise en compte de la classe "PartGrille"
' * modification le : 18/08/15 - unification des ficher excel de rapport de controle en un seul multilingue
' *                            - changé formule TCPt_WorkSheet.range("F" & TCPt_Ligne) = "=2*(" & TCPt_OngletA & "!H" & TCPt_cpt & ")"
' *                              était TCPt_WorkSheet.range("F" & TCPt_Ligne) = "=" & TCPt_OngletA & "!H" & TCPt_cpt
' *                              demande Ludo (Localisation * 2)
' * modification le 22/06/16   - Ajout de la sélection du DSCGP pour récupérer les infos "Designation" ... dans le fichier de rapport
' *                              Supprimer l'obligation d'avoir physiquement une grille sym pour générer le rapport Sym
' *****************************************************************

'Log de l'utilisation de la macro
LogUtilMacro nPath, nFicLog, nMacro, "Z5_Extract_Excel_Control", VMacro

Dim NoGrilleD, NoGrilleG As String
Dim NumOutillageGrilleD, NumOutillageGrilleG  As String
Dim DesignGrilleD, DesignGrilleG As String
Dim Exemplaire As String

Dim TypeGrille As String
Dim CoteGrille As String
Dim OrigPts As String
Dim AxeSym As String
Dim PartGrilleD As New c_PartGrille
Dim PartGrilleG As New c_PartGrille
Dim mBar As c_ProgressBar
Dim NomRapportControl As String, NomRapportControlSym As String
Dim PartDloaded As Boolean, PartGLoaded As Boolean
Dim TableauPoints

    Langue = 1
    CheminSourcesMacro = Get_Active_CATVBA_Path
    
    'Appel la fenetre de selection des Parts Grilles
    Load Frm_Extract
    Frm_Extract.Show
    'Clic sur Bouton "Annule" dans formulaire
    If Not (Frm_Extract.ChB_OkAnnule) Then
        'Exit Sub
        End
    End If
       
    'Stockage des infos récupérées dans la boite de dialogue
    'Coté de grille
    If Frm_Extract.Rbt_GrilleD Then
        CoteGrille = "D"
    ElseIf Frm_Extract.Rbt_GrilleG Then
        CoteGrille = "G"
    ElseIf Frm_Extract.Rbt_GrilleSD Then
        CoteGrille = "DS"
    ElseIf Frm_Extract.Rbt_GrilleSG Then
        CoteGrille = "GS"
    End If
    
    'Origine des Coordonnées des points
    If Frm_Extract.TB_AxisRef = "Référence de la part" Then
        OrigPts = "oPart"
    Else
        OrigPts = "oRep"
    End If
    
    'Progress Barre
    Set mBar = New c_ProgressBar
    'Initialisation du fractionnement de la barre de progression
    ' Points A et B, Pinules et pieds, angles et diametres valent 1 = 6
    ' Remontage Airbus  valent 3 * 2 = 6
    ' donc NbEtapes = 12
    If CoteGrille = "DS" Or CoteGrille = "GS" Then
        nbEtapes = 26 'Deux fichiers excel a traiter
    Else
        nbEtapes = 13 'Un fichier excel a traiter
    End If
    
    noEtape = 1: noItem = 1: nbItems = 100: StrTitre = " Création de l'export excel, veuillez patienter."
    mBar.CalculProgression noEtape, nbEtapes, noItem, nbItems, StrTitre
     
    NoGrilleD = Frm_Extract.TBX_NomGriNueD
    NoGrilleG = Frm_Extract.TBX_NomGriNueG
    If Frm_Extract.Rbt_GrilleSG Or Frm_Extract.Rbt_GrilleSD Then
        If Frm_Extract.Rbt_X Then
            AxeSym = "X"
        ElseIf Frm_Extract.Rbt_Y Then
            AxeSym = "Y"
        ElseIf Frm_Extract.Rbt_Z Then
            AxeSym = "Z"
        Else
            AxeSym = "N" 'Grille non symétrique
        End If
    End If
    Unload Frm_Extract

Set coll_docs = CATIA.Documents

    'Test si les CatParts droite et gauche sont chargées
    PartDloaded = IsLoadPart(coll_docs, NoGrilleD & ".CATPart")
    PartGLoaded = IsLoadPart(coll_docs, NoGrilleG & ".CATPart")

    'Part de la grille droite
    If CoteGrille = "D" Then
        If (Not PartDloaded) Then
            MsgBox "La grille droite n'est pas chargée !", vbCritical, "Erreur"
            GoTo Exit_CATMain
        Else
            PartGrilleD.PG_partDocGrille = coll_docs.Item(NoGrilleD & ".CATPart")
            'Lecture des Attributs
            DesignGrilleD = PartGrilleD.Prm_xDesignation
            NumOutillageGrilleD = PartGrilleD.Prm_xNumoutillage
            Exemplaire = PartGrilleD.Prm_xExemplaire
            TypeGrille = Mid(PartGrilleD.Prm_Dtemplate, 3, 2)
        End If
    End If
    
    'Part de la grille Gauche
    If CoteGrille = "G" Then
        If (Not PartGLoaded) Then
            MsgBox "La grille gauche n'est pas chargée !", vbCritical, "Erreur"
            GoTo Exit_CATMain
        Else
            PartGrilleG.PG_partDocGrille = coll_docs.Item(NoGrilleG & ".CATPart")
            'Lecture des Attributs
            DesignGrilleG = PartGrilleG.Prm_xDesignation
            NumOutillageGrilleG = PartGrilleG.Prm_xNumoutillage
            Exemplaire = PartGrilleG.Prm_xExemplaire
            TypeGrille = Mid(PartGrilleG.Prm_Dtemplate, 3, 2)
        End If
    End If

    'Part des grilles G + Sym
    If CoteGrille = "GS" Then
        If NoGrilleD = "" Or NoGrilleG = "" Then
            MsgBox "Une des grilles (Droite ou Gauche) n'est pas renseignée !", vbCritical, "Erreur"
            GoTo Exit_CATMain
        Else
            'Lecture des Attributs Gauche
            PartGrilleG.PG_partDocGrille = coll_docs.Item(NoGrilleG & ".CATPart")
            TypeGrille = Mid(PartGrilleD.Prm_Dtemplate, 3, 2)
'            DesignGrilleG = PartGrilleG.Prm_xDesignation
'            NumOutillageGrilleG = PartGrilleG.Prm_xNumoutillage
'            Exemplaire = PartGrilleG.Prm_xExemplaire
            'Lecture des Attributs Droit
            If PartDloaded Then
                DesignGrilleD = PartGrilleD.Prm_xDesignation
                NumOutillageGrilleD = PartGrilleD.Prm_xNumoutillage
                Exemplaire = PartGrilleD.Prm_xExemplaire
            Else
                DesignGrilleD = InfoDscgp.DesignSym
                NumOutillageGrilleD = InfoDscgp.NumGrilleAssSym
                Exemplaire = InfoDscgp.Exemplaire
            End If
    
        End If
    End If
   
    'Part des grilles D + Sym
    If CoteGrille = "DS" Then
        If NoGrilleD = "" Or NoGrilleG = "" Then
            MsgBox "Une des grilles (Droite ou Gauche) n'est pas renseignée !", vbCritical, "Erreur"
            GoTo Exit_CATMain
        Else
            'Lecture des Attributs Droit
            PartGrilleD.PG_partDocGrille = coll_docs.Item(NoGrilleD & ".CATPart")
            TypeGrille = Mid(PartGrilleD.Prm_Dtemplate, 3, 2)
            'DesignGrilleD = PartGrilleD.Prm_xDesignation
            'NumOutillageGrilleD = PartGrilleD.Prm_xNumoutillage
            'Exemplaire = PartGrilleD.Prm_xExemplaire
            'Lecture des Attributs Gauche
            If PartGLoaded Then
                DesignGrilleG = PartGrilleG.Prm_xDesignation
                NumOutillageGrilleG = PartGrilleG.Prm_xNumoutillage
                Exemplaire = PartGrilleG.Prm_xExemplaire
            Else
                DesignGrilleG = InfoDscgp.DesignSym
                NumOutillageGrilleG = InfoDscgp.NumGrilleAssSym
                Exemplaire = InfoDscgp.Exemplaire
            End If
        End If
    End If
   
    Select Case CoteGrille
        Case "D"
            NomRapportControl = NoGrilleD & ".xlsm"
        Case "G"
            NomRapportControl = NoGrilleG & ".xlsm"
        Case "DS"
            NomRapportControl = NoGrilleD & ".xlsm"
            NomRapportControlSym = NoGrilleG & ".xlsm"
        Case "GS"
            NomRapportControl = NoGrilleG & ".xlsm"
            NomRapportControlSym = NoGrilleD & ".xlsm"
    End Select
       
    'verifie si un fichier de rapport est déja présent et l'efface
        If Not (EffaceFicNom(CheminDestRapport, NomRapportControl)) Then
            End
        End If
    'verifie si un fichier de rapport est déja présent et l'efface
        If Not (EffaceFicNom(CheminDestRapport, NomRapportControlSym)) Then
            End
        End If
        
    'Ouverture du Fichier de Controle Excel pour la grille de référence
    Dim objExcelControle
    Set objExcelControle = CreateObject("EXCEL.APPLICATION")
    objExcelControle.Visible = True
    
    'initialise les messages dans la langue choisie
    MG_Langue Langue
    Dim objWorkBookControle
    Dim objWorkSheet
    Dim WorkBookName As String
    
    'Test le chemin de la bibli des composants
    CheminBibliComposants = CorrigeDFS()    'Nom des fichiers de rapport de controle

        WorkBookName = CheminBibliComposants & "\" & ComplementCheminBibliComposants & "\" & NomTemplateExcelMulti
    
    On Error Resume Next
    Set objWorkBookControle = objExcelControle.Workbooks.Open(WorkBookName)
    If Err.Number <> 0 Then
        MsgBox "Le template du fichier de controle " & WorkBookName & " est introuvable !", vbCritical, "Fichier manquant"
        GoTo Exit_CATMain
    End If

    objWorkBookControle.ActiveSheet.Visible = True
    objWorkBookControle.screenupdating = False
    objExcelControle.Calculation = -4135
    objWorkBookControle.SaveAs (CheminDestRapport & NomRapportControl)
    
    'Ouverture du Fichier de Controle Excel pour la grille symétrique
    If CoteGrille = "DS" Or CoteGrille = "GS" Then
        Dim objWorkBookControleSym
        Dim objWorkSheetSym
        Set objWorkBookControleSym = objExcelControle.Workbooks.Open(CStr(CheminBibliComposants & "\" & ComplementCheminBibliComposants & "\" & NomTemplateExcelMulti))
        
        objWorkBookControleSym.ActiveSheet.Visible = True
        objWorkBookControleSym.screenupdating = False
        objWorkBookControleSym.SaveAs CheminDestRapport & NomRapportControlSym
    End If

    objExcelControle.WindowState = xLMinimized
    'Suppression des onglets inutiles
    objExcelControle.DisplayAlerts = False
    If CoteGrille = "G" Then  'Cote Gauche seul on efface les onglets droit
        SupOngletsExcel objWorkBookControle, "D"
    ElseIf CoteGrille = "D" Then 'Cote Droit seul on efface les onglet gauche
        SupOngletsExcel objWorkBookControle, "G"
    ElseIf CoteGrille = "DS" Then 'Cote droite  + sym on efface les onglet gauche dans le fichier droite et droit dans le sym
        SupOngletsExcel objWorkBookControle, "G"
        SupOngletsExcel objWorkBookControleSym, "D"
    ElseIf CoteGrille = "GS" Then 'Cote Gauche + sym on efface les onglet droit dans le fichier gauche et gauche dans le sym
        SupOngletsExcel objWorkBookControle, "D"
        SupOngletsExcel objWorkBookControleSym, "G"
    End If
    objExcelControle.DisplayAlerts = True
    
    'Collecte des coordonnées des points

'======================================================================================
'Coté Droit
    If CoteGrille = "D" Or CoteGrille = "DS" Then
    'export des points A
        noEtape = noEtape + 1
        If PartGrilleD.Exist_HB(nHBPtA) Then   'Test si le set géométrique existe
            Set objWorkSheet = objWorkBookControle.worksheets("cote droit A")
            If OrigPts = "oRep" Then
                TableauPoints = RecupListPtMesure(PartGrilleD.partDocGrille, nHBPtA)
            Else
                TableauPoints = RecupListPtCoord(PartGrilleD.partDocGrille, nHBPtA)
            End If
            ExportPt TableauPoints, objWorkSheet, nHBPtA, "N", mBar
        Else
            MsgBox "Le set Géométrique : " & (nHBPtA) & " est manquant ou mal orthograpié.", vbCritical, "Eléments manquant"
            End
        End If
    'export des points B
        noEtape = noEtape + 1
        If PartGrilleD.Exist_HB(nHBPtB) Then   'Test si le set géométrique existe
            Set objWorkSheet = objWorkBookControle.worksheets("cote droit B")
            If OrigPts = "oRep" Then
                TableauPoints = RecupListPtMesure(PartGrilleD.partDocGrille, nHBPtB)
            Else
                TableauPoints = RecupListPtCoord(PartGrilleD.partDocGrille, nHBPtB)
            End If
            ExportPt TableauPoints, objWorkSheet, nHBPtB, "N", mBar
        Else
            MsgBox "Le set Géométrique : " & (nHBPtA) & " est manquant ou mal orthograpié.", vbCritical, "Eléments manquant"
            End
        End If
    'export des Pieds
        noEtape = noEtape + 1
        If PartGrilleD.Exist_HB(nHBFeet) Then  'Test si le set géométrique existe
            Set objWorkSheet = objWorkBookControle.worksheets("pieds droit")
            If OrigPts = "oRep" Then
                TableauPoints = RecupListPtMesure(PartGrilleD.partDocGrille, nHBFeet)
            Else
                TableauPoints = RecupListPtCoord(PartGrilleD.partDocGrille, nHBFeet)
            End If
            ExportPt TableauPoints, objWorkSheet, nHBFeet, "N", mBar
        End If
    'export des pinnules
        noEtape = noEtape + 1
        If PartGrilleD.Exist_HB(nHBPin) Then 'Test si le set géométrique existe
            Set objWorkSheet = objWorkBookControle.worksheets("datum droit")
            If OrigPts = "oRep" Then
                TableauPoints = RecupListPtMesure(PartGrilleD.partDocGrille, nHBPin)
            Else
                TableauPoints = RecupListPtCoord(PartGrilleD.partDocGrille, nHBPin)
            End If
            ExportPt TableauPoints, objWorkSheet, nHBPin, "N", mBar
        End If
    'Export des angles
        noEtape = noEtape + 1
        Set objWorkSheet = objWorkBookControle.worksheets("cote droit angle")
        ExportAngle PartGrilleD.Hb(nHBPtA), objWorkSheet, "D", mBar
    'Export des Diamètres
        noEtape = noEtape + 1
        Set objWorkSheet = objWorkBookControle.worksheets("diamètre droit")
        ExportDiam PartGrilleD.Hb(nHBPtA), objWorkSheet, mBar
    End If
    If CoteGrille = "DS" Then
    'export des points A Sym
        noEtape = noEtape + 1
        If PartGrilleG.Exist_HB(nHBPtA) Then   'Test si le set géométrique existe
        Set objWorkSheetSym = objWorkBookControleSym.worksheets("cote gauche A")
            If OrigPts = "oRep" Then
                TableauPoints = RecupListPtMesure(PartGrilleD.partDocGrille, nHBPtA)
            Else
                TableauPoints = RecupListPtCoord(PartGrilleD.partDocGrille, nHBPtA)
            End If
            ExportPt TableauPoints, objWorkSheetSym, nHBPtA, AxeSym, mBar
        Else
            MsgBox "Le set Géométrique : " & (nHBPtA) & " est manquant ou mal orthograpié.", vbCritical, "Eléments manquant"
            End
        End If
    'export des points B Sym
        noEtape = noEtape + 1
        If PartGrilleG.Exist_HB(nHBPtB) Then   'Test si le set géométrique existe
            Set objWorkSheetSym = objWorkBookControleSym.worksheets("cote gauche B")
            If OrigPts = "oRep" Then
                TableauPoints = RecupListPtMesure(PartGrilleD.partDocGrille, nHBPtB)
            Else
                TableauPoints = RecupListPtCoord(PartGrilleD.partDocGrille, nHBPtB)
            End If
            ExportPt TableauPoints, objWorkSheetSym, nHBPtB, AxeSym, mBar
        Else
            MsgBox "Le set Géométrique : " & (nHBPtB) & " est manquant ou mal orthograpié.", vbCritical, "Eléments manquant"
            End
        End If
    'export des Pieds Sym
        noEtape = noEtape + 1
        If PartGrilleD.Exist_HB(nHBFeet) Then  'Test si le set géométrique existe
            Set objWorkSheetSym = objWorkBookControleSym.worksheets("pieds gauche")
            If OrigPts = "oRep" Then
                TableauPoints = RecupListPtMesure(PartGrilleD.partDocGrille, nHBFeet)
            Else
                TableauPoints = RecupListPtCoord(PartGrilleD.partDocGrille, nHBFeet)
            End If
            ExportPt TableauPoints, objWorkSheetSym, nHBFeet, AxeSym, mBar
        End If
    'export des pinnules Sym
        noEtape = noEtape + 1
        If PartGrilleD.Exist_HB(nHBPin) Then 'Test si le set géométrique existe
            
            Set objWorkSheetSym = objWorkBookControleSym.worksheets("datum gauche")
            If OrigPts = "oRep" Then
                TableauPoints = RecupListPtMesure(PartGrilleD.partDocGrille, nHBPin)
            Else
                TableauPoints = RecupListPtCoord(PartGrilleD.partDocGrille, nHBPin)
            End If
            ExportPt TableauPoints, objWorkSheetSym, nHBPin, AxeSym, mBar
        End If
    'Export des angles Sym
        noEtape = noEtape + 1
        Set objWorkSheetSym = objWorkBookControleSym.worksheets("cote gauche angle")
        ExportAngle PartGrilleD.Hb(nHBPtA), objWorkSheetSym, "G", mBar
    'Export des Diamètres
        noEtape = noEtape + 1
        Set objWorkSheetSym = objWorkBookControleSym.worksheets("diamètre gauche")
        ExportDiam PartGrilleD.Hb(nHBPtA), objWorkSheetSym, mBar
    End If

'Coté Droit Livraison Airbus
    If CoteGrille = "D" Or CoteGrille = "DS" Then
    'N° et Nom de grille Coté droit
        Set objWorkSheet = objWorkBookControle.worksheets("Livraison D Airbus")
        objWorkSheet.range("C11") = DesignGrilleD
        objWorkSheet.range("C12") = NumOutillageGrilleD
        objWorkSheet.range("C13") = "A" 'Indice
        objWorkSheet.range("F12") = Exemplaire
        RemontageAirbus objWorkBookControle, 223, "D", mBar
    End If
    If CoteGrille = "DS" Then
    'N° et Nom de grille Coté Sym
        Set objWorkSheetSym = objWorkBookControleSym.worksheets("Livraison G Airbus")
        objWorkSheetSym.range("C11") = DesignGrilleG
        objWorkSheetSym.range("C12") = NumOutillageGrilleG
        objWorkSheet.range("C13") = "A" 'Indice
        objWorkSheet.range("F12") = Exemplaire
        RemontageAirbus objWorkBookControleSym, 223, "G", mBar
    End If
'======================================================================================

'Coté Gauche
    If CoteGrille = "G" Or CoteGrille = "GS" Then
    'export des points A
        noEtape = noEtape + 1
        If PartGrilleG.Exist_HB(nHBPtA) Then   'Test si le set géométrique existe
            Set objWorkSheet = objWorkBookControle.worksheets("cote gauche A")
            If OrigPts = "oRep" Then
                TableauPoints = RecupListPtMesure(PartGrilleG.partDocGrille, nHBPtA)
            Else
                TableauPoints = RecupListPtCoord(PartGrilleG.partDocGrille, nHBPtA)
            End If
            ExportPt TableauPoints, objWorkSheet, nHBPtA, "N", mBar
        Else
            MsgBox "Le set Géométrique : " & (nHBPtA) & " est manquant ou mal orthograpié.", vbCritical, "Eléments manquant"
            End
        End If
    'export des points B
        noEtape = noEtape + 1
        If PartGrilleG.Exist_HB(nHBPtB) Then   'Test si le set géométrique existe
            Set objWorkSheet = objWorkBookControle.worksheets("cote gauche B")
            If OrigPts = "oRep" Then
                TableauPoints = RecupListPtMesure(PartGrilleG.partDocGrille, nHBPtB)
            Else
                TableauPoints = RecupListPtCoord(PartGrilleG.partDocGrille, nHBPtB)
            End If
            ExportPt TableauPoints, objWorkSheet, nHBPtB, "N", mBar
        Else
            MsgBox "Le set Géométrique : " & (nHBPtB) & " est manquant ou mal orthograpié.", vbCritical, "Eléments manquant"
            End
        End If
    'export des Pieds
        noEtape = noEtape + 1
        If PartGrilleG.Exist_HB(nHBFeet) Then  'Test si le set géométrique existe
            Set objWorkSheet = objWorkBookControle.worksheets("pieds gauche")
            If OrigPts = "oRep" Then
                TableauPoints = RecupListPtMesure(PartGrilleG.partDocGrille, nHBFeet)
            Else
                TableauPoints = RecupListPtCoord(PartGrilleG.partDocGrille, nHBFeet)
            End If
            
            ExportPt TableauPoints, objWorkSheet, nHBFeet, "N", mBar
        End If
    'export des pinnules
        noEtape = noEtape + 1
        If PartGrilleG.Exist_HB(nHBPin) Then 'Test si le set géométrique existe
            Set objWorkSheet = objWorkBookControle.worksheets("datum gauche")
            If OrigPts = "oRep" Then
                TableauPoints = RecupListPtMesure(PartGrilleG.partDocGrille, nHBPin)
            Else
                TableauPoints = RecupListPtCoord(PartGrilleG.partDocGrille, nHBPin)
            End If
            ExportPt TableauPoints, objWorkSheet, nHBPin, "N", mBar
        End If
    'Export des angles
        noEtape = noEtape + 1
        Set objWorkSheet = objWorkBookControle.worksheets("cote gauche angle")
        ExportAngle PartGrilleG.Hb(nHBPtA), objWorkSheet, "G", mBar
    'Export des Diamètres
        noEtape = noEtape + 1
        Set objWorkSheet = objWorkBookControle.worksheets("diamètre gauche")
        ExportDiam PartGrilleG.Hb(nHBPtA), objWorkSheet, mBar
    End If
     If CoteGrille = "GS" Then
    'export des points A Sym
        noEtape = noEtape + 1
        If PartGrilleG.Exist_HB(nHBPtA) Then   'Test si le set géométrique existe
            Set objWorkSheetSym = objWorkBookControleSym.worksheets("cote droit A")
            If OrigPts = "oRep" Then
                TableauPoints = RecupListPtMesure(PartGrilleG.partDocGrille, nHBPtA)
            Else
                TableauPoints = RecupListPtCoord(PartGrilleG.partDocGrille, nHBPtA)
            End If
            ExportPt TableauPoints, objWorkSheetSym, nHBPtA, AxeSym, mBar
        Else
            MsgBox "Le set Géométrique : " & (nHBPtA) & " est manquant ou mal orthograpié.", vbCritical, "Eléments manquant"
            End
        End If
    'export des points B Sym
        noEtape = noEtape + 1
        If PartGrilleG.Exist_HB(nHBPtB) Then   'Test si le set géométrique existe
            Set objWorkSheetSym = objWorkBookControleSym.worksheets("cote droit B")
            If OrigPts = "oRep" Then
                TableauPoints = RecupListPtMesure(PartGrilleG.partDocGrille, nHBPtB)
            Else
                TableauPoints = RecupListPtCoord(PartGrilleG.partDocGrille, nHBPtB)
            End If
            ExportPt TableauPoints, objWorkSheetSym, nHBPtB, AxeSym, mBar
        Else
            MsgBox "Le set Géométrique : " & (nHBPtB) & " est manquant ou mal orthograpié.", vbCritical, "Eléments manquant"
            End
        End If
    'export des Pieds
        noEtape = noEtape + 1
        If PartGrilleG.Exist_HB(nHBFeet) Then  'Test si le set géométrique existe
            Set objWorkSheetSym = objWorkBookControleSym.worksheets("pieds droit")
                        If OrigPts = "oRep" Then
                TableauPoints = RecupListPtMesure(PartGrilleG.partDocGrille, nHBFeet)
            Else
                TableauPoints = RecupListPtCoord(PartGrilleG.partDocGrille, nHBFeet)
            End If
            ExportPt TableauPoints, objWorkSheetSym, nHBFeet, AxeSym, mBar
        End If
    'export des pinnules
        noEtape = noEtape + 1
        If PartGrilleG.Exist_HB(nHBPin) Then 'Test si le set géométrique existe
            Set objWorkSheetSym = objWorkBookControleSym.worksheets("datum droit")
            If OrigPts = "oRep" Then
                TableauPoints = RecupListPtMesure(PartGrilleG.partDocGrille, nHBPin)
            Else
                TableauPoints = RecupListPtCoord(PartGrilleG.partDocGrille, nHBPin)
            End If
            ExportPt TableauPoints, objWorkSheetSym, nHBPin, AxeSym, mBar
        End If
    'Export des angles Sym
        noEtape = noEtape + 1
        Set objWorkSheetSym = objWorkBookControleSym.worksheets("cote droit angle")
        ExportAngle PartGrilleG.Hb(nHBPtA), objWorkSheetSym, "D", mBar
    'Export des Diamètres
        noEtape = noEtape + 1
        Set objWorkSheetSym = objWorkBookControleSym.worksheets("diamètre droit")
        ExportDiam PartGrilleG.Hb(nHBPtA), objWorkSheetSym, mBar
    End If
'Coté Gauche Livraison Airbus
    If CoteGrille = "G" Or CoteGrille = "GS" Then
    'N° et Nom de grille Coté gauche
        Set objWorkSheet = objWorkBookControle.worksheets("Livraison G Airbus")
        objWorkSheet.range("C11") = DesignGrilleG
        objWorkSheet.range("C12") = NoGrilleG
        objWorkSheet.range("C13") = "A" 'Indice
        objWorkSheet.range("F12") = Exemplaire
        RemontageAirbus objWorkBookControle, 223, "G", mBar
    End If
    If CoteGrille = "GS" Then
    'N° et Nom de grille Coté Sym
        Set objWorkSheetSym = objWorkBookControleSym.worksheets("Livraison D Airbus")
        objWorkSheetSym.range("C11") = DesignGrilleD
        objWorkSheetSym.range("C12") = NoGrilleD
        objWorkSheet.range("C13") = "A" 'Indice
        objWorkSheet.range("F12") = Exemplaire
        RemontageAirbus objWorkBookControleSym, 223, "D", mBar
    End If
    
    'Enregistrement et fermeture des fichiers Excel
    objExcelControle.Visible = True
    objExcelControle.Calculation = -4105
    objWorkBookControle.Save
    objWorkBookControle.screenupdating = True
    'objWorkBookControle.Close
    If CoteGrille = "DS" Or CoteGrille = "GS" Then
        objWorkBookControleSym.Save
        objWorkBookControleSym.screenupdating = True
        'objWorkBookControleSym.Close
    End If
    objExcelControle.WindowState = xLNormal
    
GoTo Exit_CATMain
    
Exit_CATMain:
'Libération des classes
    Set PartGrilleD = Nothing
    Set PartGrilleG = Nothing
    Set mBar = Nothing
    
End Sub

Public Sub ExportPt(EP_ListPts, EP_Worksheet, EP_NameConteneur As String, EP_AxeSym As String, mBar)
'Export des points (A ou B ou autre en fonction du conteneur) vers le fichier de controle excel
'Inversion du signe pour la valeur symétrique si elle est différente de X, Y ou Z
' EP_ListPts(3,x) = tableau de la liste des points
' EP_Worksheet =  feuille du fichier excel dans laquelle ecrire les coordonnées
' EP_NameConteneur = nom du set géométrique contenant les points
' EP_AxeSym = axe de symétrie
' mBar = objet "barre de progression"

Dim EP_Counter As Long

'Pas de point dans le set géométrique
On Error Resume Next
EP_Counter = UBound(EP_ListPts, 2)
    If (Err.Number <> 0) Then
    ' tableau vide
        Err.Clear
        Exit Sub
    End If

EP_Counter = 0
noEtape = 1: noItem = 1: nbItems = 100: StrTitre = " Création de l'export excel, veuillez patienter."
While (EP_Counter <= UBound(EP_ListPts, 2))

    mBar.CalculProgression noEtape, nbEtapes, EP_Counter, EP_Counter + 1, " Export des points : " & EP_NameConteneur

    'Nom du point
    If Frm_Extract.Rbt_Num3D Then
        EP_Worksheet.range("A" & EP_Counter + 2) = EP_ListPts(0, EP_Counter)
    Else
        EP_Worksheet.range("A" & EP_Counter + 2) = "A" & EP_Counter
    End If
    
    'Coordonnée des points
    If EP_AxeSym = "X" Then
        EP_Worksheet.range("B" & EP_Counter + 2) = -Round(EP_ListPts(1, EP_Counter), 3)
    Else
        EP_Worksheet.range("B" & EP_Counter + 2) = Round(EP_ListPts(1, EP_Counter), 3)
    End If
    EP_Worksheet.Columns("B").AutoFit
    If EP_AxeSym = "Y" Then
        EP_Worksheet.range("C" & EP_Counter + 2) = -Round(EP_ListPts(2, EP_Counter), 3)
    Else
        EP_Worksheet.range("C" & EP_Counter + 2) = Round(EP_ListPts(2, EP_Counter), 3)
    End If
    EP_Worksheet.Columns("C").AutoFit
    If EP_AxeSym = "Z" Then
        EP_Worksheet.range("D" & EP_Counter + 2) = -Round(EP_ListPts(3, EP_Counter), 3)
    Else
        EP_Worksheet.range("D" & EP_Counter + 2) = Round(EP_ListPts(3, EP_Counter), 3)
    End If
    EP_Worksheet.Columns("D").AutoFit
    
    Dim EP_Formule As String
    EP_Formule = "=RACINE((B" & Trim(str(EP_Counter + 2)) & "- E" & Trim(str(EP_Counter + 2)) & ")^ 2 + (C" & Trim(str(EP_Counter + 2)) & " - F" & Trim(str(EP_Counter + 2)) & ") ^ 2 + (D" & Trim(str(EP_Counter + 2)) & " - G" & Trim(str(EP_Counter + 2)) & ") ^ 2)"
    EP_Worksheet.range("H" & EP_Counter + 2).formulalocal = EP_Formule
           
    EP_Counter = EP_Counter + 1
Wend

'Mise en forme conditionnelle. Si l'écart est suppérieur à 0.2, le texte passe en rouge
With EP_Worksheet.range("H2:H" & EP_Counter + 2)
    .formatconditions.Delete
    .formatconditions.Add xLCellValue, xLGreater, "0,2"
    .formatconditions(1).Font.colorindex = 3
End With

'Coloriage des cellules en vert et jaune
EP_Counter = UBound(EP_ListPts, 2) + 1
CouleurCell EP_Worksheet, "B2", "D" & EP_Counter + 1, "vert"
CouleurCell EP_Worksheet, "E2", "G" & EP_Counter + 1, "jaune"
If EP_NameConteneur = nHBFeet Then
    CouleurCell EP_Worksheet, "I2", "I" & EP_Counter + 1, "jaune"
End If
End Sub

Public Function RecupListPtCoord(RLPC_PartDoc As Document, RLPC_NameConteneur) As String()
'Récupère dans un tableau les Coordonnées X, Y et Z des points du Part
'En récupérant les coordonnées Par rapport à l'origine de la part.
    Dim RLPC_TableauPt() As String

    Dim RLPC_PartGrille As New c_PartGrille
    RLPC_PartGrille.PG_partDocGrille = RLPC_PartDoc
    
    Dim RLPC_SPAworkbench As SPAWorkbench
    Set RLPC_SPAworkbench = RLPC_PartGrille.partDocGrille.GetWorkbench("SPAWorkbench")
    
    Dim RLPC_Coord(2)
    Dim RLPC_Measurable
    
    Dim RLPC_HybridBody As HybridBody
    If RLPC_PartGrille.Exist_HB(RLPC_NameConteneur) Then
        Set RLPC_HybridBody = RLPC_PartGrille.Hbodies.Item(RLPC_NameConteneur)
    Else
        MsgBox "Le set Géométrique : " & (RLPC_NameConteneur) & " est manquant ou mal orthograpié.", vbCritical, "Eléments manquant"
        End
    End If

Dim RLPC_Shapes As HybridShapes
Set RLPC_Shapes = RLPC_HybridBody.HybridShapes
If RLPC_Shapes.Count = 0 Then
    Exit Function
End If
Dim RLPC_counter As Integer
RLPC_counter = 1
Dim RLPC_CurPtShape As HybridShape

For RLPC_counter = 1 To RLPC_Shapes.Count
    ReDim Preserve RLPC_TableauPt(3, RLPC_counter - 1)
    Set RLPC_CurPtShape = RLPC_Shapes.Item(RLPC_counter)
    
    RLPC_TableauPt(0, RLPC_counter - 1) = RLPC_CurPtShape.Name
    Set RLPC_Measurable = RLPC_SPAworkbench.GetMeasurable(RLPC_CurPtShape)
    RLPC_Measurable.GetPoint RLPC_Coord
    
    RLPC_TableauPt(1, RLPC_counter - 1) = Round(RLPC_Coord(0), 3)
    RLPC_TableauPt(2, RLPC_counter - 1) = Round(RLPC_Coord(1), 3)
    RLPC_TableauPt(3, RLPC_counter - 1) = Round(RLPC_Coord(2), 3)
           
Next RLPC_counter

RecupListPtCoord = RLPC_TableauPt
End Function

Public Function RecupListPtMesure(RLPM_PartDoc As Document, RLPM_NameConteneur) As String()
'Récupère dans un tableu les Coordonnées X, Y et Z des points du Part
'En mesurant la distance mini entre le point et les plans PlanX0, PlanY0 et PlanZ0
    Dim RLPM_TableauPt() As String
    Dim RLPM_TableauPtTem(5) As Double
    Dim RLPM_HSPlanX As HybridShape, RLPM_HSPlanY As HybridShape, RLPM_HSPlanZ As HybridShape
    Dim RLPM_HSPlanX10 As HybridShape, RLPM_HSPlanY10 As HybridShape, RLPM_HSPlanZ10 As HybridShape
    Dim RLPM_SPAworkbench As SPAWorkbench
    Dim RLPM_XMeas, RLPM_YMeas As Measurable, RLPM_ZMeas As Measurable
    Dim RLPM_X10Meas As Measurable, RLPM_Y10Meas As Measurable, RLPM_Z10Meas As Measurable
    Dim RLPM_HBodyPts As HybridBody
    Dim RLPM_HbodyPlans As HybridBody
    Dim RLPM_counter As Long
    
    Dim RLPM_PartGrille As New c_PartGrille
    RLPM_PartGrille.PG_partDocGrille = RLPM_PartDoc
    
    Set RLPM_SPAworkbench = RLPM_PartGrille.GrilleSPAWorkbench
    
    'test l'existence des sets géométrique
    If RLPM_PartGrille.Exist_HB(RLPM_NameConteneur) Then
        Set RLPM_HBodyPts = RLPM_PartGrille.Hbodies.Item(RLPM_NameConteneur)
    Else
        MsgBox "Le set Géométrique : " & (RLPM_NameConteneur) & " est manquant ou mal orthograpié.", vbCritical, "Eléments manquant"
        End
    End If
    If RLPM_PartGrille.Exist_HB(nHBTrav) Then
        Set RLPM_HbodyPlans = RLPM_PartGrille.Hb(nHBTrav)
    Else
        MsgBox "Le set Géométrique : " & (RLPM_NameConteneur) & " est manquant ou mal orthograpié.", vbCritical, "Eléments manquant"
        End
    End If
 
 
Dim RLPM_Shapes As HybridShapes
Set RLPM_Shapes = RLPM_HBodyPts.HybridShapes
If RLPM_Shapes.Count = 0 Then
    RecupListPtMesure = RLPM_TableauPt()
    Exit Function
End If

On Error Resume Next
Set RLPM_HSPlanX = RLPM_HbodyPlans.HybridShapes.Item("PlanX0")
Set RLPM_HSPlanY = RLPM_HbodyPlans.HybridShapes.Item("PlanY0")
Set RLPM_HSPlanZ = RLPM_HbodyPlans.HybridShapes.Item("PlanZ0")
Set RLPM_HSPlanX10 = RLPM_HbodyPlans.HybridShapes.Item("PlanX10")
Set RLPM_HSPlanY10 = RLPM_HbodyPlans.HybridShapes.Item("PlanY10")
Set RLPM_HSPlanZ10 = RLPM_HbodyPlans.HybridShapes.Item("PlanZ10")
If (Err.Number <> 0) Then
    ' Un des plans est manquant
        Err.Clear
        MsgBox "Un des plans de référence est manquant ! re-sélectionnez le repère de référence.", vbCritical, "Plans de ref manquants"
        End
    End If
On Error GoTo 0

Set RLPM_XMeas = RLPM_SPAworkbench.GetMeasurable(RLPM_HSPlanX)
Set RLPM_X10Meas = RLPM_SPAworkbench.GetMeasurable(RLPM_HSPlanX10)
Set RLPM_YMeas = RLPM_SPAworkbench.GetMeasurable(RLPM_HSPlanY)
Set RLPM_Y10Meas = RLPM_SPAworkbench.GetMeasurable(RLPM_HSPlanY10)
Set RLPM_ZMeas = RLPM_SPAworkbench.GetMeasurable(RLPM_HSPlanZ)
Set RLPM_Z10Meas = RLPM_SPAworkbench.GetMeasurable(RLPM_HSPlanZ10)

RLPM_counter = 1
Dim RLPM_CurPtShape As HybridShape

For RLPM_counter = 1 To RLPM_Shapes.Count
    Set RLPM_CurPtShape = RLPM_Shapes.Item(RLPM_counter)
    
    RLPM_TableauPtTem(0) = RLPM_XMeas.GetMinimumDistance(RLPM_CurPtShape)
    RLPM_TableauPtTem(1) = RLPM_YMeas.GetMinimumDistance(RLPM_CurPtShape)
    RLPM_TableauPtTem(2) = RLPM_ZMeas.GetMinimumDistance(RLPM_CurPtShape)
    RLPM_TableauPtTem(3) = RLPM_X10Meas.GetMinimumDistance(RLPM_CurPtShape)
    RLPM_TableauPtTem(4) = RLPM_Y10Meas.GetMinimumDistance(RLPM_CurPtShape)
    RLPM_TableauPtTem(5) = RLPM_Z10Meas.GetMinimumDistance(RLPM_CurPtShape)
    
    ReDim Preserve RLPM_TableauPt(3, RLPM_counter - 1)

    If IsUpdatable(RLPM_PartGrille.PartGrille, RLPM_CurPtShape) Then
        RLPM_TableauPt(0, RLPM_counter - 1) = RLPM_CurPtShape.Name
        RLPM_TableauPt(1, RLPM_counter - 1) = SignePlusMoins(RLPM_TableauPtTem(0), RLPM_TableauPtTem(3)) * RLPM_TableauPtTem(0)
        RLPM_TableauPt(2, RLPM_counter - 1) = SignePlusMoins(RLPM_TableauPtTem(1), RLPM_TableauPtTem(4)) * RLPM_TableauPtTem(1)
        RLPM_TableauPt(3, RLPM_counter - 1) = SignePlusMoins(RLPM_TableauPtTem(2), RLPM_TableauPtTem(5)) * RLPM_TableauPtTem(2)
    End If
Next RLPM_counter

RecupListPtMesure = RLPM_TableauPt()
End Function

Public Sub ExportAngle(EA_HybridBodyA, EA_Worksheet, EA_Cote As String, mBar)
'Export des angles (Calcul de la distance entre pt A et pt B)
'pour le Théorique avec les coordonnées exportées du 3D reprises des onglets 'cote droit A' 'cote droit B'
'Pour le Pratique avec les mesures du fournisseur reprises des onglets 'cote droit A' 'cote droit B'
'EA_Cote = cote de grille a traiter = "D" ou "G"

Dim EA_ShapesA As HybridShapes
Set EA_ShapesA = EA_HybridBodyA.HybridShapes

Dim EA_ShapeA As HybridShape

'Initialisation du nom des onglets pour les formules excel
Dim EA_OngletA, EA_OngletB As String
If EA_Cote = "D" Then
    EA_OngletA = "'cote droit A'"
    EA_OngletB = "'cote droit B'"
ElseIf EA_Cote = "G" Then
    EA_OngletA = "'cote gauche A'"
    EA_OngletB = "'cote gauche B'"
End If
Dim EA_Formule As String

Dim cpt As Long
cpt = 1

While (cpt <= EA_ShapesA.Count)

mBar.CalculProgression noEtape, nbEtapes, cpt, EA_ShapesA.Count, " Export des angles."

    'Nom du point A
    Set EA_ShapeA = EA_ShapesA.Item(cpt)
    If Frm_Extract.Rbt_Num3D Then
        EA_Worksheet.range("A" & cpt + 1) = EA_ShapesA.Item(cpt).Name
    Else
        EA_Worksheet.range("A" & cpt + 1) = "A" & cpt
    End If
    
    EA_Worksheet.range("B" & cpt + 1) = "=((" & EA_OngletA & "!B" & cpt + 1 & " -" & EA_OngletB & "!B" & cpt + 1 & ")^2 + (" & EA_OngletA & "!C" & cpt + 1 & "-" & EA_OngletB & "!C" & cpt + 1 & ")^2 +(" & EA_OngletA & "!D" & cpt + 1 & "-" & EA_OngletB & "!D" & cpt + 1 & ")^2)^(1/2)"
    EA_Worksheet.range("C" & cpt + 1) = "=((" & EA_OngletA & "!E" & cpt + 1 & " -" & EA_OngletB & "!E" & cpt + 1 & ")^2 + (" & EA_OngletA & "!F" & cpt + 1 & "-" & EA_OngletB & "!F" & cpt + 1 & ")^2 +(" & EA_OngletA & "!G" & cpt + 1 & "-" & EA_OngletB & "!G" & cpt + 1 & ")^2)^(1/2)"
    
    EA_Formule = "=60*2*(ATAN(((((((" & EA_OngletA & "!B" & cpt + 1 & "-" & EA_OngletB & "!B" & cpt + 1 & ")/B" & cpt + 1 & ")-((" & EA_OngletA & "!E" & cpt + 1 & "-" & EA_OngletB & "!E" & cpt + 1 & ")/C" & cpt + 1 & "))^2+" _
    & "(((" & EA_OngletA & "!C" & cpt + 1 & "-" & EA_OngletB & "!C" & cpt + 1 & ")/B" & cpt + 1 & ")-((" & EA_OngletA & "!F" & cpt + 1 & "-" & EA_OngletB & "!F" & cpt + 1 & ")/C" & cpt + 1 & "))^2+(((" & EA_OngletA & "!D" & cpt + 1 & "-" & EA_OngletB & "!D" & cpt + 1 & ")/B" & cpt + 1 & ")-((" & EA_OngletA & "!G" & cpt + 1 & "-" & EA_OngletB & "!G" & cpt + 1 & ")/C" & cpt + 1 & "))^2)^(1/2))/2)/( -((((((" & EA_OngletA & "!B" & cpt + 1 & "-" & EA_OngletB & "!B" & cpt + 1 & ")/B" & cpt + 1 & ")-" _
    & "((" & EA_OngletA & "!E" & cpt + 1 & "-" & EA_OngletB & "!E" & cpt + 1 & ")/C" & cpt + 1 & "))^2+(((" & EA_OngletA & "!C" & cpt + 1 & "-" & EA_OngletB & "!C" & cpt + 1 & ")/B" & cpt + 1 & ")-((" & EA_OngletA & "!F" & cpt + 1 & "-" & EA_OngletB & "!F" & cpt + 1 & ")/C" & cpt + 1 & "))^2+(((" & EA_OngletA & "!D" & cpt + 1 & "-" & EA_OngletB & "!D" & cpt + 1 & ")/B" & cpt + 1 & ")-((" & EA_OngletA & "!G" & cpt + 1 & "-" & EA_OngletB & "!G" & cpt + 1 & ")/C" & cpt + 1 & "))^2)^(1/2))/2)*((((((" & EA_OngletA & "!B" & cpt + 1 & "-" & EA_OngletB & "!B" & cpt + 1 & ")/B" & cpt + 1 & ")-((" & EA_OngletA & "!E" & cpt + 1 & "-" & EA_OngletB & "!E" & cpt + 1 & ")/C" & cpt + 1 & "))^2+(((" & EA_OngletA & "!C" & cpt + 1 & "-" & EA_OngletB & "!C" & cpt + 1 & ")/B" & cpt + 1 & ")-((" & EA_OngletA & "!F" & cpt + 1 & "-" & EA_OngletB & "!F" & cpt + 1 & ")/C" & cpt + 1 & "))^2+(((" & EA_OngletA & "!D" & cpt + 1 _
    & "-" & EA_OngletB & "!D" & cpt + 1 & ")/B" & cpt + 1 & ")-((" & EA_OngletA & "!G" & cpt + 1 & "-" & EA_OngletB & "!G" & cpt + 1 & ")/C" & cpt + 1 & "))^2)^(1/2))/2)+1)^(1/2)))*180/3,1415926535898"
    
    EA_Worksheet.range("D" & cpt + 1).formulalocal = EA_Formule
    cpt = cpt + 1
Wend

'Mise en forme conditionnelle. Si l'écart est suppérieur à 10, le texte passe en rouge
With EA_Worksheet.range("D2:D" & cpt)
   .formatconditions.Delete
   .formatconditions.Add xLCellValue, xLGreater, "10"
   .formatconditions(1).Font.colorindex = 3
End With

End Sub

Public Sub ExportDiam(ED_HybridBodyA, ED_WorkSheet, mBar)
'Formatage de l'onglet des Diamètres de perçages (actuellement on ne sait pas récupérer la valeur des Diamètre)
'Création d'une ligne "Diam x" pour chaque point

Dim ED_Shapes As HybridShapes
Set ED_Shapes = ED_HybridBodyA.HybridShapes
Dim cpt As Long
cpt = 1

Dim ED_Shape As HybridShape

While (cpt <= ED_Shapes.Count)
mBar.CalculProgression noEtape, nbEtapes, cpt, ED_Shapes.Count, " Export des Diamètres "

    Set ED_Shape = ED_Shapes.Item(cpt)
    ED_WorkSheet.range("A" & cpt + 1) = "  Ø  " & ED_Shapes.Item(cpt).Name
    ED_WorkSheet.range("F" & cpt + 1) = "=RC[-3]-RC[-4]"
    cpt = cpt + 1
Wend
CouleurCell ED_WorkSheet, "B2", "E" & cpt, "jaune"
CouleurCell ED_WorkSheet, "G2", "G" & cpt, "jaune"

End Sub
Public Sub RemontageAirbus(RA_WorkBook, RA_Ligne As Integer, RA_Cote As String, mBar)
'RA_WorkBook Classeur excel
'RA_Ligne N° de ligne en court dans le fichier excel
'RA_Cote = Coté de la Grille ("D", "G")
Dim NomPt As String, NomPtA As String, NomPtB As String
Dim cpt As Long
Dim RA_cptA As String, RA_cptB As String
Dim FeuilleDatum As String, _
    FeuillePied As String, _
    FeuillePtA As String, _
    FeuillePtB As String, _
    FeuilleLivraison As String
Dim RA_WSheet_Datum, _
    RA_WSheet_Pieds, _
    RA_WSheet_PtA, _
    RA_WSheet_PtB, _
    RA_WSheet_Livraison

    'Initialisation du nom des calques exel en fonction du coté de grille
    If RA_Cote = "D" Then
        FeuilleDatum = "datum droit"
        FeuillePied = "pieds droit"
        FeuillePtA = "cote droit A"
        FeuillePtB = "cote droit B"
        FeuilleLivraison = "Livraison D Airbus"
    ElseIf RA_Cote = "G" Then
        FeuilleDatum = "datum gauche"
        FeuillePied = "pieds gauche"
        FeuillePtA = "cote gauche A"
        FeuillePtB = "cote gauche B"
        FeuilleLivraison = "Livraison G Airbus"
    End If
    Set RA_WSheet_Datum = RA_WorkBook.worksheets.Item(FeuilleDatum)
    Set RA_WSheet_Pieds = RA_WorkBook.worksheets.Item(FeuillePied)
    Set RA_WSheet_PtA = RA_WorkBook.worksheets.Item(FeuillePtA)
    Set RA_WSheet_PtB = RA_WorkBook.worksheets.Item(FeuillePtB)
    Set RA_WSheet_Livraison = RA_WorkBook.worksheets.Item(FeuilleLivraison)

    cpt = 2
    noEtape = noEtape + 2
'### Les Datum ###
    'Trace la ligne du Titre
    RA_cptA = "A" & RA_Ligne
    RA_cptB = "I" & RA_Ligne
    
    NomPt = RA_WSheet_Datum.range("A" & cpt)
    
    'Pas de la ligne de titre si pas de points
    If NomPt <> "" Then
        RA_WSheet_Livraison.range(RA_cptA, RA_cptB).MergeCells = True
        RA_WSheet_Livraison.range(RA_cptA) = MG_msg(10)
        BorduresCell RA_WSheet_Livraison, RA_cptA, RA_cptB
        RA_Ligne = RA_Ligne + 1
    End If
    
    While (NomPt <> "")

        mBar.CalculProgression noEtape, nbEtapes, cpt, NbDatums, " Remontage Airbus Pinnules : " & NomPt
        'Trace la ligne d'entète
        '############################
        'Ajouter une détection de haut de page
        '############################
    
        TraceCadreEnTete RA_WSheet_Livraison, RA_Ligne
        RA_Ligne = RA_Ligne + 1
        'Trace le cadre du point
        TraceCadreDatum RA_WSheet_Livraison, RA_Ligne, NomPt, cpt, CStr(FeuilleDatum)
        RA_Ligne = RA_Ligne + 3
        'Trace la ligne de localisation
        TraceCadreLoc RA_WSheet_Livraison, RA_Ligne
        RA_Ligne = RA_Ligne + 1
        'Trace une ligne vide
        TraceLigneVide RA_WSheet_Livraison, RA_Ligne
        RA_Ligne = RA_Ligne + 1
        
        cpt = cpt + 1
        NomPt = RA_WSheet_Datum.range("A" & cpt)
    Wend
    cpt = 2
    noEtape = noEtape + 2
    
'### Les Pieds ###
    'Trace la ligne du Titre
    RA_cptA = "A" & RA_Ligne
    RA_cptB = "I" & RA_Ligne
    
    NomPt = RA_WSheet_Pieds.range("A" & cpt)
    
    'Pas de ligne de titre si pas de points
    If NomPt <> "" Then
        RA_WSheet_Livraison.range(RA_cptA, RA_cptB).MergeCells = True
        RA_WSheet_Livraison.range(RA_cptA) = MG_msg(11)
        BorduresCell RA_WSheet_Livraison, RA_cptA, RA_cptB
        RA_WSheet_Livraison.range(RA_cptA, RA_cptB).cells.HorizontalAlignment = xLCenter
        RA_Ligne = RA_Ligne + 1
    End If
    While (NomPt <> "")
        mBar.CalculProgression noEtape, nbEtapes, cpt, NbFeets, " Remontage Airbus Pieds : " & NomPt
        
        'Trace la ligne d'entète
        '############################
        'Ajouter une détection de haut de page
        '############################
    
        TraceCadreEnTete RA_WSheet_Livraison, RA_Ligne
        RA_Ligne = RA_Ligne + 1
        'Trace le cadre du point
        TraceCadreFeet RA_WSheet_Livraison, RA_Ligne, NomPt, cpt, RA_Cote
        RA_Ligne = RA_Ligne + 1
        'Trace une ligne vide
        TraceLigneVide RA_WSheet_Livraison, RA_Ligne
        RA_Ligne = RA_Ligne + 1

        cpt = cpt + 1
        NomPt = RA_WSheet_Pieds.range("A" & cpt)
    Wend
    cpt = 2
    noEtape = noEtape + 2

'### Les Points ###
    'Trace la ligne du Titre
    RA_cptA = "A" & RA_Ligne
    RA_cptB = "I" & RA_Ligne

    RA_WSheet_Livraison.range(RA_cptA, RA_cptB).MergeCells = True
    RA_WSheet_Livraison.range(RA_cptA) = MG_msg(12)
    BorduresCell RA_WSheet_Livraison, RA_cptA, RA_cptB
    
    RA_cptA = "A" & RA_Ligne
    RA_cptB = "I" & RA_Ligne + 2
        
    NomPtA = RA_WSheet_PtA.range("A" & cpt)
    NomPtB = RA_WSheet_PtB.range("A" & cpt)
    RA_Ligne = RA_Ligne + 1
    
    While (NomPtA <> "")
        mBar.CalculProgression noEtape, nbEtapes, cpt, NbPts, " Remontage Airbus Points : " & NomPtA
        '############################
        'Ajouter une détection de haut de page
        '############################
    
        TraceCadreEnTete RA_WSheet_Livraison, RA_Ligne
        RA_Ligne = RA_Ligne + 1
        'Trace le cadre du point
        TraceCadrePT RA_WSheet_Livraison, RA_Ligne, NomPtA, NomPtB, cpt, RA_Cote
        RA_Ligne = RA_Ligne + 1
        'Trace une ligne vide
        TraceLigneVide RA_WSheet_Livraison, RA_Ligne
        RA_Ligne = RA_Ligne + 1
        
        cpt = cpt + 1
        NomPtA = RA_WSheet_PtA.range("A" & cpt)
        NomPtB = RA_WSheet_PtB.range("A" & cpt)
    Wend
    
    'Ajout page Synthèse
    TraceSynthese RA_WSheet_Livraison, RA_Ligne
    'Mise en page
    MisePage RA_WSheet_Livraison, RA_Ligne
    'RA_WSheet_Livraison.PageSetup.PrintArea = "$A$1:$I$" & RA_Ligne '"$A$1:$I$540"

End Sub

Public Sub TraceCadreDatum(TCD_WorkSheet, TCD_Ligne As Integer, TCD_PtNom, TCD_cpt As Long, TCD_Onglet As String)
'Trace le bloc des lignes de ccordonnées de points (X, Y, Z) pour les Datum
'TCD_WorkSheet = Feuille excel
'TCD_Ligne = N° de la premiere ligne du cadre
'TCD_PtNom = Nom du point
'TCD_cpt = N° de ligne des coordonnées du pt de l'onglet datum
'TCD_Onglet = Onglet excel 'datum gauche' ou 'datum droit'

TCD_Onglet = "'" & TCD_Onglet & "'"

Dim TCD_cPtA, TCD_cPtB As String

TCD_WorkSheet.range("A" & TCD_Ligne) = "Point  "

TCD_cPtA = "A" & TCD_Ligne
TCD_cPtB = "A" & TCD_Ligne + 2
TCD_WorkSheet.range(TCD_cPtA, TCD_cPtB).MergeCells = True
    
TCD_cPtA = "B" & TCD_Ligne
TCD_cPtB = "C" & TCD_Ligne + 2
TCD_WorkSheet.range(TCD_cPtA, TCD_cPtB).MergeCells = True
TCD_WorkSheet.range("B" & TCD_Ligne) = TCD_PtNom
    
TCD_WorkSheet.range("D" & TCD_Ligne) = "  X  "
TCD_WorkSheet.range("D" & TCD_Ligne + 1) = "  Y  "
TCD_WorkSheet.range("D" & TCD_Ligne + 2) = "  Z  "

'Nominale
TCD_cPtA = "E" & TCD_Ligne
TCD_cPtB = "E" & TCD_Ligne + 2
'Couleur intérieure
CouleurCell TCD_WorkSheet, TCD_cPtA, TCD_cPtB, "gris"
    
'Mesurée
TCD_WorkSheet.range("F" & TCD_Ligne) = "=" & TCD_Onglet & "!E" & TCD_cpt
TCD_WorkSheet.range("F" & TCD_Ligne + 1) = "=" & TCD_Onglet & "!F" & TCD_cpt
TCD_WorkSheet.range("F" & TCD_Ligne + 2) = "=" & TCD_Onglet & "!G" & TCD_cpt

'Ecart/axe et tolerances
TCD_cPtA = "G" & TCD_Ligne
TCD_cPtB = "I" & TCD_Ligne + 2
'Couleur intérieure
CouleurCell TCD_WorkSheet, TCD_cPtA, TCD_cPtB, "gris"
    
'Bordure
TCD_cPtA = "A" & TCD_Ligne
TCD_cPtB = "I" & TCD_Ligne + 2
BorduresCell TCD_WorkSheet, TCD_cPtA, TCD_cPtB

End Sub

Public Sub TraceCadreFeet(wSheet, iLigne As Integer, PtNom, cpt As Long, Cote As String)
'Trace le bloc des lignes de ccordonnées de points (X, Y, Z) pour les Pieds
'wSheet = Feuille excel
'iLigne = N° de la premiere ligne du cadre
'PtNom = Nom du point
'cPt = N° de ligne des coordonnées du pt de l'onglet Pieds
'Cote = Coté de la grille ("G" ou "D")
Dim wSeetFeet As String
Dim cPtA As String, cPtb As String
Dim TolPtAInf As Double, TolPtASup As Double

 'Initialisation du nom des onglets pour les formules excel
    If Cote = "D" Then
        wSeetFeet = "'pieds droit'"
    ElseIf Cote = "G" Then
        wSeetFeet = "'pieds gauche'"
    End If

    wSheet.range("A" & iLigne) = "Point  "
    
    cPtA = "A" & iLigne
    cPtb = "A" & iLigne + 2
    wSheet.range(cPtA, cPtb).MergeCells = True
        
    cPtA = "B" & iLigne
    cPtb = "C" & iLigne + 2
    wSheet.range(cPtA, cPtb).MergeCells = True
    wSheet.range("B" & iLigne) = PtNom
        
    wSheet.range("D" & iLigne) = "  X  "
    wSheet.range("D" & iLigne + 1) = "  Y  "
    wSheet.range("D" & iLigne + 2) = "  Z  "
    
    'Nominale
    wSheet.range("E" & iLigne) = "=" & wSeetFeet & "!B" & cpt
    wSheet.range("E" & iLigne + 1) = "=" & wSeetFeet & "!C" & cpt
    wSheet.range("E" & iLigne + 2) = "=" & wSeetFeet & "!D" & cpt

    'Mesurée
    wSheet.range("F" & iLigne) = "=" & wSeetFeet & "!E" & cpt
    wSheet.range("F" & iLigne + 1) = "=" & wSeetFeet & "!F" & cpt
    wSheet.range("F" & iLigne + 2) = "=" & wSeetFeet & "!G" & cpt
    
    'tolerances
    cPtA = "G" & iLigne
    cPtb = "H" & iLigne + 2
    'Couleur intérieure
    CouleurCell wSheet, cPtA, cPtb, "gris"
     
    'Ecart/axe
    wSheet.range("I" & iLigne) = "=F" & iLigne & "-E" & iLigne
    wSheet.range("I" & iLigne + 1) = "=F" & iLigne + 1 & "-E" & iLigne + 1
    wSheet.range("I" & iLigne + 2) = "=F" & iLigne + 2 & "-E" & iLigne + 2
    
    iLigne = iLigne + 3

    'Trace la ligne "VG"
    cPtA = "A" & iLigne
    cPtb = "C" & iLigne
    TolPtAInf = -0.1
    TolPtASup = 0.1
    wSheet.range(cPtA, cPtb).MergeCells = True
    wSheet.range("A" & iLigne) = MG_msg(20)
    wSheet.range("D" & iLigne) = "  D  "
    wSheet.range("E" & iLigne) = 0
    wSheet.range("F" & iLigne) = "=" & wSeetFeet & "!I" & cpt
    wSheet.range("G" & iLigne) = TolPtAInf
    wSheet.range("H" & iLigne) = TolPtASup
    wSheet.range("I" & iLigne) = "=F" & iLigne & "-E" & iLigne

    'Couleur rouge
    FormText wSheet, ("I" & iLigne), "rouge"

    'Mise en forme conditionnelle. Si l'écart est suppérieur à 10, le texte passe en rouge
    With wSheet.range("I" & iLigne)
       .formatconditions.Delete
       .formatconditions.Add xLCellValue, xLBetween, Formula1:=TolPtAInf, Formula2:=TolPtASup
       .formatconditions(1).Font.Color = -11489280
    End With

    'Bordure
    cPtA = "A" & iLigne - 4
    cPtb = "I" & iLigne
    BorduresCell wSheet, cPtA, cPtb

End Sub

Public Sub TraceCadrePT(wSheet, iLigne As Integer, nPtA, nPtB, cpt As Long, strCote As String)
'Trace le bloc des lignes de ccordonnées de points (X, Y, Z) pour les Points A et B
'wSheet = Feuille excel
'iLigne = N° de la premiere ligne du cadre
'nPtA, nPtB = Nom des points A et B
'TCPt_Pt_cpt = N° de ligne des coordonnées du pt des onglet points A, B et angles
'strCote = Coté de la grille ("G" ou "D")
Dim strPtA As String, strPtB As String
Dim nwsPtA  As String, nwsPtB  As String, nwsAngle  As String, nwsDiam As String
Dim TolPtASup As Double, TolPtAInf As Double
Dim TolDiaSup As String, TolDiaInf As String

    'Initialisation du nom des onglets pour les formules excel
    If strCote = "D" Then
        nwsPtA = "'cote droit A'"
        nwsPtB = "'cote droit B'"
        nwsAngle = "'cote droit angle'"
        nwsDiam = "'diamètre droit'"
    ElseIf strCote = "G" Then
        nwsPtA = "'cote gauche A'"
        nwsPtB = "'cote gauche B'"
        nwsAngle = "'cote gauche angle'"
        nwsDiam = "'diamètre gauche'"
    End If
    
    'Nom Pt A
    wSheet.range("A" & iLigne) = "Point  "
    
    strPtA = "A" & iLigne
    strPtB = "A" & iLigne + 2
    wSheet.range(strPtA, strPtB).MergeCells = True
        
    strPtA = "B" & iLigne
    strPtB = "C" & iLigne + 2
    wSheet.range(strPtA, strPtB).MergeCells = True
    wSheet.range("B" & iLigne) = nPtA
        
    wSheet.range("D" & iLigne) = "  X  "
    wSheet.range("D" & iLigne + 1) = "  Y  "
    wSheet.range("D" & iLigne + 2) = "  Z  "
    
    'Nom Pt B
    wSheet.range("A" & iLigne + 3) = "Point  "
    
    strPtA = "A" & iLigne + 3
    strPtB = "A" & iLigne + 5
    wSheet.range(strPtA, strPtB).MergeCells = True
        
    strPtA = "B" & iLigne + 3
    strPtB = "C" & iLigne + 5
    wSheet.range(strPtA, strPtB).MergeCells = True
    wSheet.range("B" & iLigne + 3) = nPtB
        
    wSheet.range("D" & iLigne + 3) = "  X  "
    wSheet.range("D" & iLigne + 4) = "  Y  "
    wSheet.range("D" & iLigne + 5) = "  Z  "
    
    'Nominale pt A
    wSheet.range("E" & iLigne) = "=" & nwsPtA & "!B" & cpt
    wSheet.range("E" & iLigne + 1) = "=" & nwsPtA & "!C" & cpt
    wSheet.range("E" & iLigne + 2) = "=" & nwsPtA & "!D" & cpt
    
    'Nominale pt B
    wSheet.range("E" & iLigne + 3) = "=" & nwsPtB & "!B" & cpt
    wSheet.range("E" & iLigne + 4) = "=" & nwsPtB & "!C" & cpt
    wSheet.range("E" & iLigne + 5) = "=" & nwsPtB & "!D" & cpt
    
    'mesures pt A
    wSheet.range("F" & iLigne) = "=" & nwsPtA & "!E" & cpt
    wSheet.range("F" & iLigne + 1) = "=" & nwsPtA & "!F" & cpt
    wSheet.range("F" & iLigne + 2) = "=" & nwsPtA & "!G" & cpt
    
    'mesures pt B
    wSheet.range("F" & iLigne + 3) = "=" & nwsPtB & "!E" & cpt
    wSheet.range("F" & iLigne + 4) = "=" & nwsPtB & "!F" & cpt
    wSheet.range("F" & iLigne + 5) = "=" & nwsPtB & "!G" & cpt

    'les tolerances pt A
    TolPtASup = 0.2
    TolPtAInf = -0.2
    wSheet.range("G" & iLigne) = TolPtAInf
    wSheet.range("G" & iLigne + 1) = TolPtAInf
    wSheet.range("G" & iLigne + 2) = TolPtAInf
    
    wSheet.range("H" & iLigne) = TolPtASup
    wSheet.range("H" & iLigne + 1) = TolPtASup
    wSheet.range("H" & iLigne + 2) = TolPtASup
        
    'les tolerances pt B
    strPtA = "G" & iLigne + 3
    strPtB = "H" & iLigne + 5
    'Couleur intérieure
    CouleurCell wSheet, strPtA, strPtB, "gris"
    
    'Ecart/Axe pt A
    wSheet.range("I" & iLigne) = "=F" & iLigne & "-E" & iLigne
    wSheet.range("I" & iLigne + 1) = "=F" & iLigne + 1 & "-E" & iLigne + 1
    wSheet.range("I" & iLigne + 2) = "=F" & iLigne + 2 & "-E" & iLigne + 2
    
    'Couleur Rouge par defaut
    FormText wSheet, ("I" & iLigne & ":I" & iLigne + 2), "rouge"
    'Mise en forme conditionnelle. Si l'écart est compris entre les tolérances, le texte passe en vert
    With wSheet.range("I" & iLigne & ":I" & iLigne + 2)
       .formatconditions.Delete
       .formatconditions.Add xLCellValue, xLBetween, Formula1:=TolPtAInf, Formula2:=TolPtASup
       .formatconditions(1).Font.Color = -11489280
    End With
    
    'Ecart/Axe pt B
    wSheet.range("I" & iLigne + 3) = "=F" & iLigne + 3 & "-E" & iLigne + 3
    wSheet.range("I" & iLigne + 4) = "=F" & iLigne + 4 & "-E" & iLigne + 4
    wSheet.range("I" & iLigne + 5) = "=F" & iLigne + 5 & "-E" & iLigne + 5
          
    iLigne = iLigne + 6
        
    'Localisation
    strPtA = "A" & iLigne
    strPtB = "C" & iLigne
    TolPtASup = 0.4
    TolPtAInf = 0
    wSheet.range(strPtA, strPtB).MergeCells = True
    wSheet.range("A" & iLigne) = MG_msg(40)
    wSheet.range("A" & iLigne).cells.HorizontalAlignment = xLDroite
    wSheet.range("D" & iLigne) = "mm"
    wSheet.range("E" & iLigne) = 0
    wSheet.range("F" & iLigne) = "=2*(" & nwsPtA & "!H" & cpt & ")"
    wSheet.range("G" & iLigne) = TolPtAInf
    wSheet.range("H" & iLigne) = TolPtASup
    wSheet.range("I" & iLigne) = "=F" & iLigne & "-E" & iLigne
    
    'Couleur Rouge par defaut
    FormText wSheet, ("I" & iLigne & ":I" & iLigne), "rouge"
    'Mise en forme conditionnelle. Si l'écart est compris entre les tolérances, le texte passe en vert
    With wSheet.range("I" & iLigne & ":I" & iLigne)
       .formatconditions.Delete
       .formatconditions.Add xLCellValue, xLBetween, Formula1:=TolPtAInf, Formula2:=TolPtASup
       .formatconditions(1).Font.Color = -11489280
    End With
    
    iLigne = iLigne + 1
    
    'Angle CYL
    strPtA = "A" & iLigne
    strPtB = "C" & iLigne
    TolPtASup = 10
    TolPtAInf = -10
    wSheet.range(strPtA, strPtB).MergeCells = True
    wSheet.range("A" & iLigne) = MG_msg(41)
    wSheet.range("A" & iLigne).cells.HorizontalAlignment = xLDroite
    wSheet.range("D" & iLigne) = "min"
    wSheet.range("E" & iLigne) = 0
    wSheet.range("F" & iLigne) = "=" & nwsAngle & "!D" & cpt
    wSheet.range("G" & iLigne) = TolPtAInf
    wSheet.range("H" & iLigne) = TolPtASup
    wSheet.range("I" & iLigne) = "=F" & iLigne & "-E" & iLigne
    
    'Couleur Rouge par defaut
    FormText wSheet, ("I" & iLigne & ":I" & iLigne), "rouge"
    'Mise en forme conditionnelle. Si l'écart est compris entre les tolérances, le texte passe en vert
    With wSheet.range("I" & iLigne & ":I" & iLigne)
       .formatconditions.Delete
       .formatconditions.Add xLCellValue, xLBetween, Formula1:=TolPtAInf, Formula2:=TolPtASup
       .formatconditions(1).Font.Color = -11489280
    End With
    
    iLigne = iLigne + 1
    
    'Diametres
    strPtA = "A" & iLigne
    strPtB = "C" & iLigne
    
    wSheet.range(strPtA, strPtB).MergeCells = True
    wSheet.range("A" & iLigne) = MG_msg(42)
    wSheet.range("A" & iLigne).cells.HorizontalAlignment = xLDroite
    wSheet.range("D" & iLigne) = "  Ø  "
    wSheet.range("E" & iLigne) = "=" & nwsDiam & "!B" & cpt
    wSheet.range("F" & iLigne) = "=" & nwsDiam & "!C" & cpt
    wSheet.range("G" & iLigne) = "=" & nwsDiam & "!D" & cpt
    wSheet.range("H" & iLigne) = "=" & nwsDiam & "!E" & cpt
    wSheet.range("I" & iLigne) = "=F" & iLigne & "-E" & iLigne
    
    TolDiaSup = "=$H$" & iLigne
    TolDiaInf = "=$G$" & iLigne
    
    'Couleur Rouge par defaut
    FormText wSheet, ("I" & iLigne & ":I" & iLigne), "rouge"
    'Mise en forme conditionnelle. Si l'écart est compris entre les tolérances, le texte passe en vert
    With wSheet.range("I" & iLigne & ":I" & iLigne)
       .formatconditions.Delete
       .formatconditions.Add xLCellValue, xLBetween, Formula1:=TolDiaInf, Formula2:=TolDiaSup
       .formatconditions(1).Font.Color = -11489280
    End With
    
    'Bordure
    strPtA = "A" & iLigne - 8
    strPtB = "I" & iLigne
    BorduresCell wSheet, strPtA, strPtB

End Sub

Public Sub TraceCadreEnTete(wSheet, iLigne As Integer)
'Trace le bloc d'Entete des lignes de ccordonnées de points
'wSheet = Feuille excel
'iLigne = N° de la ligne d'entete
Dim strPtA As String, strPtB As String

    strPtA = "A" & iLigne
    strPtB = "C" & iLigne
    wSheet.range(strPtA, strPtB).MergeCells = True
    strPtB = "I" & iLigne
    
    wSheet.range("A" & iLigne) = MG_msg(0)
    wSheet.range("D" & iLigne) = "U"
    wSheet.range("E" & iLigne) = MG_msg(1)
    wSheet.range("F" & iLigne) = MG_msg(2)
    wSheet.range("G" & iLigne) = MG_msg(3)
    wSheet.range("H" & iLigne) = MG_msg(4)
    wSheet.range("I" & iLigne) = MG_msg(5)
    wSheet.range(strPtA & ":" & strPtB).cells.HorizontalAlignment = xLCenter
    wSheet.Columns("A:I").AutoFit
    
    'Bordure
    BorduresCell wSheet, strPtA, strPtB
    
End Sub

Public Sub TraceCadreLoc(TCL_WorkSheet, TCL_Ligne As Integer)
'Trace le bloc de la ligne "Localisation" des pinules
'TCL_WorkSheet = Feuille excel
'TCL_Ligne = N° de la ligne Localisation
Dim TCl_cptA As String, TCl_cptB As String

TCl_cptA = "A" & TCL_Ligne
TCl_cptB = "B" & TCL_Ligne
TCL_WorkSheet.range(TCl_cptA, TCl_cptB).MergeCells = True
TCL_WorkSheet.range("A" & TCL_Ligne) = "Localisation Ø 0.2"
TCL_WorkSheet.range("C" & TCL_Ligne).formulalocal = "=L(-3)C(-1)"
TCL_WorkSheet.Columns("C").AutoFit
TCL_WorkSheet.range("D" & TCL_Ligne) = "mm"

'Bordure
    TCl_cptA = "A" & TCL_Ligne
    TCl_cptB = "I" & TCL_Ligne
    BorduresCell TCL_WorkSheet, TCl_cptA, TCl_cptB

'Couleur intérieure
TCl_cptA = "E" & TCL_Ligne
TCl_cptB = "I" & TCL_Ligne
CouleurCell TCL_WorkSheet, TCl_cptA, TCl_cptB, "gris"

End Sub
Public Sub TraceLigneVide(TLV_WorkSheet, TLV_Ligne As Integer)
'Trace une ligne vide Fusionnée
'TCV_WorkSheet = Feuille excel
'TCV_Ligne = N° de la ligne Vide
Dim TLV_cptA As String, TLV_cptB As String

TLV_cptA = "A" & TLV_Ligne
TLV_cptB = "I" & TLV_Ligne
TLV_WorkSheet.range(TLV_cptA, TLV_cptB).MergeCells = True

End Sub

Public Sub TraceSynthese(TS_WorkSheet, TS_Ligne As Integer)
'Ajoute la page "Analyse et synthèse des résultats" en fin de rapport
'TS_WorkSheet = Feuille excel
'TS_Ligne = N° de la ligne Localisation
'TS_Langue = Langue du rapport 1=FR, 2=EN
Dim TS_cptA As String, TS_cptB As String

TS_cptA = "A" & TS_Ligne
TS_cptB = "I" & TS_Ligne
TS_WorkSheet.range(TS_cptA, TS_cptB).MergeCells = True

TS_WorkSheet.range("A" & TS_Ligne) = MG_msg(30)
BorduresCell TS_WorkSheet, TS_cptA, TS_cptB
FormText TS_WorkSheet, "A" & TS_Ligne, "Titre"
TS_Ligne = TS_Ligne + 1
TS_cptA = "A" & TS_Ligne
TS_cptB = "I" & TS_Ligne + 32
TS_WorkSheet.range(TS_cptA, TS_cptB).MergeCells = True
TS_WorkSheet.range("A" & TS_Ligne) = MG_msg(31)
BorduresCell TS_WorkSheet, TS_cptA, TS_cptB
FormText TS_WorkSheet, "A" & TS_Ligne, "rouge"
TS_Ligne = TS_Ligne + 33
TS_cptA = "A" & TS_Ligne
TS_cptB = "I" & TS_Ligne
TS_WorkSheet.range(TS_cptA, TS_cptB).MergeCells = True
TS_WorkSheet.range("A" & TS_Ligne) = MG_msg(32)
BorduresCell TS_WorkSheet, TS_cptA, TS_cptB
FormText TS_WorkSheet, "A" & TS_Ligne, "Titre"
TS_Ligne = TS_Ligne + 1
TS_cptA = "A" & TS_Ligne
TS_cptB = "I" & TS_Ligne + 13
TS_WorkSheet.range(TS_cptA, TS_cptB).MergeCells = True
TS_WorkSheet.range("A" & TS_Ligne) = MG_msg(33)
BorduresCell TS_WorkSheet, TS_cptA, TS_cptB
FormText TS_WorkSheet, "A" & TS_Ligne, "vert"
TS_Ligne = TS_Ligne + 13

End Sub

Public Sub SupOngletsExcel(SOE_workbook, SOE_CoteGrille As String)
'supprimes les onglets inutiles dans le fichier excel passé en argument
Dim SOE_WorkSheet
Dim ListeOnglets(6) As String
Dim i As Long

If SOE_CoteGrille = "D" Then
    ListeOnglets(0) = "cote droit A"
    ListeOnglets(1) = "cote droit B"
    ListeOnglets(2) = "datum droit"
    ListeOnglets(3) = "pieds droit"
    ListeOnglets(4) = "diamètre droit"
    ListeOnglets(5) = "cote droit angle"
    ListeOnglets(6) = "Livraison D Airbus"
ElseIf SOE_CoteGrille = "G" Then
    ListeOnglets(0) = "cote gauche A"
    ListeOnglets(1) = "cote gauche B"
    ListeOnglets(2) = "datum gauche"
    ListeOnglets(3) = "pieds gauche"
    ListeOnglets(4) = "diamètre gauche"
    ListeOnglets(5) = "cote gauche angle"
    ListeOnglets(6) = "Livraison G Airbus"
End If
    For i = 0 To 6
        Set SOE_WorkSheet = SOE_workbook.worksheets(ListeOnglets(i))
        SOE_WorkSheet.Delete
    Next i
End Sub



Public Sub MisePage(MP_WorkSheet, MP_NBLigne As Integer)
'Fixe la zone d'impression et met à jour les numéro de pages dans le sommaire
Dim MP_P1 As String, MP_P2 As String, MP_P3 As String, MP_P4 As String, MP_P5 As String, MP_pts As String, MP_Px As String, MP_Page As String
MP_P1 = "'" & MP_WorkSheet.Name & "'" & "!$A$1:$I$57"
MP_P2 = "'" & MP_WorkSheet.Name & "'" & "!$A$58:$I$79"
MP_P3 = "'" & MP_WorkSheet.Name & "'" & "!$A$80:$I$124"
MP_P4 = "'" & MP_WorkSheet.Name & "'" & "!$A$125:$I$173"
MP_P5 = "'" & MP_WorkSheet.Name & "'" & "!$A$174:$I$222"
MP_pts = "'" & MP_WorkSheet.Name & "'" & "!$A$223:$I$" & MP_NBLigne - 50
MP_Px = "'" & MP_WorkSheet.Name & "'" & "!$A$" & MP_NBLigne - 49 & ":$I$" & MP_NBLigne
MP_Page = MP_P1 & ";" & MP_P2 & ";" & MP_P3 & ";" & MP_P4 & ";" & MP_P5 & ";" & MP_pts & ";" & MP_Px
'MP_WorkSheet.PageSetup.PrintArea = MP_P1 & ";" & MP_P2 & ";" & MP_P3 & ";" & MP_P4 & ";" & MP_P5 & ";" & MP_pts & ";" & MP_Px
'MP_WorkSheet.PageSetup.PrintArea = MP_P1 & ";" & MP_P3
MP_WorkSheet.PageSetup.PrintArea = "'" & MP_WorkSheet.Name & "'!$A$1:$I" & MP_NBLigne
'='Livraison G Airbus'!$A$1:$I$25;'Livraison G Airbus'!$A$27:$I$43
End Sub

Public Sub MG_Langue(Langue)
'gestion temporaire de la langues des fichier excel
'En attendant une gestion par fichiers texte de messages
ReDim MG_msg(42)
Select Case Langue
Case 1
    MG_msg(0) = "COTES"
    MG_msg(1) = "Nominale"
    MG_msg(2) = "Mesurée"
    MG_msg(3) = "Tol inf."
    MG_msg(4) = "Tol sup."
    MG_msg(5) = "Ecart/Axe"
    '-------------------------
    MG_msg(10) = "LES POINTS DE PINULES:"
    MG_msg(11) = "LES PIEDS:"
    MG_msg(12) = "LES POINTS A, B ET ANGLES POUR CHAQUE ALESAGE:"
    '-------------------------
    MG_msg(20) = "Ecart VG selon la normale au point de contact"
    '-------------------------
    MG_msg(30) = "ANALYSE & SYNTHÈSE DES RÉSULTATS"
    MG_msg(31) = "100% des diamètres des douilles ont été contrôlés avec des tampons passe/passe pas. Les certificats d'étalonnage sont fournis en annexe de ce rapport."
    MG_msg(32) = "CONCLUSION SUR LA CONFORMITÉ DE L'OUTILLAGE"
    MG_msg(33) = "L'outillage est conforme aux exigences AIAH-G-005_V13"
    '-------------------------
    MG_msg(40) = "Localisation Ø 0.4"
    MG_msg(41) = "Angle Cyl."
    MG_msg(42) = "Ø Alésage"

Case 2
     MG_msg(0) = "DIMENSIONS"
     MG_msg(1) = "Nominale"
     MG_msg(2) = "Measured"
     MG_msg(3) = "Tol inf."
     MG_msg(4) = "Tol sup."
     MG_msg(5) = "Mismatch"
     '-------------------------
     MG_msg(10) = "INSPECTION POINTS :"
     MG_msg(11) = "FEET:"
     MG_msg(12) = "POINTS A, B AND ANGLES FOR EACH HOLE :"
     '-------------------------
     MG_msg(20) = "True position according to normal to surface"
     '-------------------------
     MG_msg(30) = "ANALYS OF THE RESULTS"
     MG_msg(31) = "100% of dasa bushes diameters have been inspected with GO / NO GO plug gauges. Calibration certificates are supplied with this report."
     MG_msg(32) = "CONCLUSION OF COMPLIANCE"
     MG_msg(33) = "The tooling complies with the AIAH-G-005 V13 specification."
     '-------------------------
     MG_msg(40) = "Localisation Ø 0.4"
     MG_msg(41) = "Cyl. deviation"
     MG_msg(42) = "Cylinder Ø"

End Select
End Sub



