Attribute VB_Name = "DECLARATIONS_PUBLIQUES"
Option Explicit

'Fonction de récupération du username
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

'Version de la macro
Public Const VMacro As String = "Version 10.10.5 du 15/05/17"
Public Const nMacro As String = "Grille de perçage"
Public Const nPath As String = "\\srvxsiordo\xLogs\01_CatiaMacros"
Public Const nFicLog As String = "logUtilMacro.txt"

'Les documents Catia
Public coll_docs As Documents
Public ActiveDoc As Document

' Variables Location Macro ini
    Public CheminSourcesMacro As String
    Public Const CheminDestRapport As String = "c:\temp\"
    Public Const NomTemplateExcelMulti As String = "fiche contrôle terminée.xlsm"
    
    Public Const NomFicInfoBagues As String = "Catalogue_Manual_Drilling_Bush.xlsx"
    Public Const NomFicInfoMachines As String = "InfoMachinesPercage.xlsx"
    Public CheminBibliComposants As String
    Public Const ComplementCheminBibliComposants = "19-BD STANDARDS GRILLES"
    
    Public Const RepBaguesSprecif As String = "\SPECIFIC FUNCTION\Bush\Manual drilling bush"
    Public Const RepBagues As String = "\BUSH\Standard reference\TBU-xxxxx\ST Parts"
    Public Const RepVis As String = "\OTHER STANDARDS\Locking screw\LOS-xxxxx\ST Parts"
    Public Const RepAgrafes As String = "\OTHER STANDARDS\Fasteners - Agrafes\serie 80\agrafes équipées serie 80"

'## Nom des objets Grilles ##
' nom des set géométriques standards
    Public Const nHBRefExtIsol As String = "références externes isolées" '"ref_ext_isol"   'ex "références externes isolées"
    Public Const nHBS0 As String = "surf0"
    Public Const nHBPin As String = "pinules"
    Public Const nHBFeet As String = "feet"
    Public Const nHBPtA As String = "pointsA"
    Public Const nHBStd As String = "std"
    Public Const nHBS100 As String = "surf100"
    Public Const nHBPtB As String = "pointsB"
    Public Const nHBGrav As String = "gravures"
    Public Const nHBPtConst As String = "points_de_construction"
    Public Const nHBTrav  As String = "travail"
    Public Const nHBGeoRef As String = "geometrie de reference" '"geom_ref"           'ex "geometrie de reference"
    Public Const nHBSDetr As String = "detrompage"
    Public Const nHBDrFeet As String = "draft feet" '"draft_feet"         'ex "draft feet"
    Public Const nHBDrPin As String = "draft_pinules"       'ex "draft pinules"
    Public Const nHBDrGrav As String = "draft_gravures"     'ex "draft gravures"

'Nom des surfaces
    Public Const nSurf0 As String = "surf0"
    Public Const nSurf100 As String = "Surfacea100"
    Public Const nSurfSup As String = "Surf_Sup_Grille"
    Public Const nSurfInf As String = "Surf_Inf_Grille"

'Nom des autres géométries
    Public Const nRefGri As String = "ref_grille"
    Public Const nOrientGri As String = "orientation_grille"
    
'Nom des accessoires
    Public Const nAnod As String = "Anodisation"
    Public Const nColcod As String = "Colour coding"

'Valeur des paramètres
    Public Const vDescGNue As String = "UNEQUIPPED TEMPLATE"
    Public Const vDescFloat As String = "FLOATING PART"
    Public Const vRecogn As String = "PGRI"

'Nom des paramètres
    Public Const nPrmMaterial As String = "MATERIAL"
    Public Const nPrmRecogn As String = "RECOGNITION"
    Public Const nPrmObserv As String = "OBSERVATIONS"
    Public Const nPrmDtempl As String = "DTEMPLATE"
    Public Const nPrmNumout As String = "xNUMOUTILLAGE"
    Public Const nPrmDesign As String = "xDESIGNATION"
    Public Const nPrmExempl As String = "xEXEMPLAIRE"
    Public Const nPrmPiecPer As String = "xPIECESPERCEES"
    Public Const nPrmSite As String = "xSITE"
    Public Const nPrmProgAv As String = "xNOPROGAVION"

'Format du Fichier excel du Check 3D
    Public Const NomTemplateCheck3D As String = "RapportCheck3D.xls"
    
    Public Const nSheetCk As String = "Checks"
    Public Const PremLigSheetCk As Integer = 7
    Public Const Cel_NumGrille As String = "C2"
    Public Const Cel_Date As String = "C3"
    Public Const Cel_Controleur As String = "C4"
    Public Const c_Cote As Integer = 2
    Public Const c_Check As Integer = 4
    Public Const c_Statut As Integer = 5
    Public Const c_Comment As Integer = 7
    
    Public Const nSheetMes  As String = "Measures"
    Public Const PremLigSheetMes    As Integer = 8
    Public Const c_nomFas As Integer = 1
    Public Const c_Xr As Integer = 2
    Public Const c_Yr As Integer = 3
    Public Const c_Zr As Integer = 4
    Public Const c_Xe As Integer = 5
    Public Const c_Ye As Integer = 6
    Public Const c_Ze As Integer = 7
    Public Const c_Xec As Integer = 8
    Public Const c_Yec As Integer = 9
    Public Const c_Zec As Integer = 10
    Public Const c_nomPtA As Integer = 14
    Public Const c_PtAX As Integer = 15
    Public Const c_PtAY As Integer = 16
    Public Const c_PtAZ As Integer = 17
    
' Fichier excel des perçages et bagues
    Public CollMachines As New NumMachines
    Public CollBagues As c_DefBagues
    
    Public ValDscgp As tDSCGP
    'Public GrilleAttributs() As String
       
'Tableau du nom, du diam perçage avion et du nom du STD des points sélectionnés
    Public Tab_Select_Points() As String
    'Tab_Select_Points(0,x) Nom du Point A
    'Tab_Select_Points(1,x) Nom de la ligne STD
    'Tab_Select_Points(2,x) Diamètres de perçage Avion
    
' Variables Chemin pour traitement par lot
    Public CheminFicLot As String

' Variables pour Barre de progression
    Public nbEtapes As Long
    Public noEtape As Long
    Public noItem As Long
    Public nbItems As Long
    Public StrTitre As String
    Public NbPts As Long
    Public NbFeets As Long
    Public NbDatums As Long

' Tableau des textes traduit
    Public MG_msg() As String
    Public Langue As Integer

' Constantes de Excel
    Public Const xLCenter As Long = -4108
    Public Const xLHaut As Long = -4160
    Public Const xLDroite As Long = -4152
    Public Const xLMoyen As Long = -4138
    Public Const xLNormal As Long = -4143
    Public Const xLMinimized As Long = -4140
    Public Const xLBetween As Long = 1
    Public Const xLCellValue As Long = 1
    Public Const xLGreater As Long = 5
    Public Const xlSolid As Long = 1
    Public Const xlAutomatic As Long = -4105
    Public Const xLTextString As Long = 9
    Public Const xlContains As Long = 0

' Constantes pour la lecture des fichiers textes
' Object("scripting.filesystemobject")
    Public Const ForReading As Long = 1
    
'Couleurs
    'Public Const ColVert = 13959167
    Public Const ColVert = 16777172
    'Public Const ColRouge = 16776960
    Public Const ColRouge = 16766207
    Public Const ColGris = 12632256
    'Public Const ColNeutre = -2147483643
    Public Const ColNeutre = 16777215

'Type DSCGP
Public Type tDSCGP
    NumLot As String
'    NumLotSym As String
    NumGrilleAss As String
    NumGrilleAssSym   As String
    NumGrilleNue As String
    NumGrilleNueSym   As String
    NumPartU01 As String
    NumPartU01Sym As String
    NumDetromp As String
    design As String
    DesignSym As String
    Numout As String
    NumEnvAvion As String
    Mat As String
    NumPiecesPerc As String
    Site As String
    NumProgAvion As String
    Observ As String
    Anod As String
    Dtemplate As String
    Color As String
    SystNum As String
    CoteAvion As String
    Exemplaire As String
    Accessoires As String
    Pres_Pinules As String
    Nb_Pinules As String
    GravureInf As String
    GravureSup As String
    GravureLat1 As String
    GravureLat2 As String
    GravureLat3 As String
    GravureLat4 As String
End Type

Public Type Coord
    X As Double
    Y As Double
    Z As Double
End Type

Public InfoDscgp As tDSCGP
Public ReportLog() As String

'Les collections
Public Col_PropAirbus As Collection
Public Col_SetAirbus As Collection
Public col_Mes As c_mesures

'Variable pour mise en plan2D

'Nom des vues 2D
    Public Const nVueFace As String = "Vue de face"
    Public Const nVueCote As String = "Vue de coté"
    Public Const nVueDessus As String = "Vue de dessus"
    Public Const nVueDessous As String = "Vue de dessous"
    Public Const nVueArriere As String = "Vue de derrière"
    Public Const nVueIso As String = "Vue isomètrique"

'Nom du fichier des textes des cartouches Airbus
    Public FichtxtCart As String
'Nom des fond de plan Airbus
    Public FicCartAirbus As String

'## Nom des elements de construction créés par les macros ##
'Plan de projection 2D
    Public Const nProjOrient As String = "LnProjOrient"
    Public Const nExtProjOrient As String = "PtExProj"
    Public Const nPerpProjOrient As String = "LnPerpProjOrient"
    Public Const nPlaProj2D As String = "PlProj2D"

'Objet Dito
    Public Type Dito
        Name As String
        X As Integer
        Y As Integer
        Size As Double
        Source As DrawingView
        Cible As DrawingComponent
    End Type

Public Type Pos2D
    X As Integer
    Y As Integer
End Type

'Séparateur des fichiers CSV
    Public Const SepCSV As String = ";"
    
Public Const HLig As Double = 5.75 'Hauteur d'une ligne de nota 2D

Public Sub init_ColPropAirbus()
'Documente la collection des propriétés Airbus
Set Col_PropAirbus = New Collection
    Col_PropAirbus.Add nPrmMaterial
    Col_PropAirbus.Add "THICKNESS/DIAMETER"
    Col_PropAirbus.Add "LENGTH"
    Col_PropAirbus.Add "WIDTH"
    Col_PropAirbus.Add "MASS"
    Col_PropAirbus.Add nPrmObserv
    Col_PropAirbus.Add nPrmRecogn
    Col_PropAirbus.Add nPrmDtempl
End Sub

Public Sub Init_ColSetAirbus()
'Documente la collection des Set géométriques Airbus
Set Col_SetAirbus = New Collection
    Col_SetAirbus.Add nHBRefExtIsol
    Col_SetAirbus.Add nHBS0
    Col_SetAirbus.Add nHBPin
    Col_SetAirbus.Add nHBFeet
    Col_SetAirbus.Add nHBPtA
    Col_SetAirbus.Add nHBStd
    Col_SetAirbus.Add nHBS100
    Col_SetAirbus.Add nHBPtB
    Col_SetAirbus.Add nHBGrav
    Col_SetAirbus.Add nHBPtConst
    Col_SetAirbus.Add nHBTrav
End Sub

'Liste des erreurs
'    No                  Module          Description
'vbObjectError + 513, c_PartGrille , Set Géométrique manquant ou mal orthographié
'vbObjectError + 514, c_PartGrille , La surface manquante ou mal orthographiée
'vbObjectError + 515, c_PartGrille , Ligne 'orientation_grille' de la grille manquante ou mal orthographiée
'vbObjectError + 516, c_PartGrille , Plan Ref_Grille de la grille manquant ou mal orthographiée
'vbObjectError + 517, c_PartGrille , Erreur dans la collecte des Fasteners
'vbObjectError + 518, c_PartGrille , Part Body manquant ou mal orthographié
 'vbObjectError + 519, c_PartGrille , "Erreur lors de la récupération de l'objet HybridShapeFactory contactez le service Info."

'vbObjectError + 525, c_DSCGP , Fichier DSCGP non renseigné
'vbObjectError + 526, c_DSCGP , Fichier DSCGP Introuvable
'vbObjectError + 527, c_DSCGP , Erreur dans la collecte des Info du DSCGP
'vbObjectError + 528, c_DSCGP , Le Type d'arborescence n'est pas documenté dans le DSCGP !
'vbObjectError + 529, c_DSCGP , Le N° de grille n'est pas documenté ou ne correspond pas au coté de conception !
'vbObjectError + 530, c_DSCGP , Le fichier DSCGP n'est pas conforme (onglets "DSCGP" ou "DSCGP" manquants) !
'vbObjectError + 531, c_DSCGP , Le coté de conception est mal documenté dans le DSCGP !

'vbObjectError + 540, grilleLot, Erreur dans la collecte des contraintes de positionnement de l'environnement
