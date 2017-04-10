VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_DSCGP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*********************************************************************
'* Classe : DSCGP
'*
'* Fonctions :  Récupération des attributs d'une grille dans un fichier excel (DSCGP)
'*              Calcule le numéro des grilles nues et assemblée
'*
'* Version : 8
'* Création :  CFR
'* Modification : 15/04/14
'* Modification : 20/01/15 - commenté l'action "fermeture du fichier excel DSCGP"
'* Modification : 25/06/15 - Modification de la numérotation pour Grille DROITE seule
'* Modification : 26/02/16 - remplacement tableau ListChamps_DSCGP par type tDSCGP
'* Modification le 03/10/16 - Ajout gestion d'erreurs
'* Modification : 05/12/16 - Le numéro de grille complet est renseigné dans le DSCGP
'*                           Il n'est plus nécéssaire de calculer le N° de grille ASS et nue
'* Modification : 16/12/16 - Reprise de la numérotation des Lot, Grilles Ass et grilles nues
'*
'**********************************************************************

Private NomCelDSCGP As tDSCGP 'nom des cellule contenant les infos dans le DSCGP
Private ValDscgp As tDSCGP
Private VerDscgp As Integer
'
'===========================
' Initialisation de l'objet
'===========================
'
Private Sub Class_Initialize()

'Initialise la liste des noms de cellule du DSCGP
'Les cellules sont nommées dans le DSCGP pour éviter la sensibilité à la mise en forme du DSCGP
    NomCelDSCGP.NumLot = "n_lot"
    NomCelDSCGP.NumGrilleAss = "n_grille_g" 'correspond aussi au n° d'outilage G
    NomCelDSCGP.NumGrilleAssSym = "n_grille_d" 'correspond aussi au n° d'outilage D
    NomCelDSCGP.NumDetromp = "n_detromp"
    NomCelDSCGP.design = "nom_grille_g"
    NomCelDSCGP.DesignSym = "nom_grille_d"
    NomCelDSCGP.NumEnvAvion = "n_env_avion"
    NomCelDSCGP.Mat = "matiere_grille"
    NomCelDSCGP.NumPiecesPerc = ""
    NomCelDSCGP.Site = "site"
    NomCelDSCGP.NumProgAvion = "prog_avion"
    NomCelDSCGP.Anod = "anod"
    NomCelDSCGP.Dtemplate = nPrmDtempl
    NomCelDSCGP.Color = "color"
    NomCelDSCGP.SystNum = "type_arbo"
    NomCelDSCGP.CoteAvion = "cote_avion"
    NomCelDSCGP.Exemplaire = "exemplaire"
    NomCelDSCGP.Pres_Pinules = "Pres_Pinules"
    NomCelDSCGP.Nb_Pinules = "Nb_Pinules"
    NomCelDSCGP.Accessoires = "lst_access"
    NomCelDSCGP.GravureInf = "lst_Gravures_Inf"
    NomCelDSCGP.GravureSup = "lst_Gravures_Sup"
    NomCelDSCGP.GravureLat1 = "lst_Gravures_Lat1"
    NomCelDSCGP.GravureLat2 = "lst_Gravures_Lat2"
    NomCelDSCGP.GravureLat3 = "lst_Gravures_Lat3"
    NomCelDSCGP.GravureLat4 = "lst_Gravures_Lat4"

End Sub

Property Let VersionDscgp(ByVal VDscgp As Integer)
'Défini la version du DSCGP
    VerDscgp = VDscgp
End Property

Property Let OpenDSCGP(NomFicDSCGP As String)

Dim objexcel
Dim objWorkBook
Dim objWSDscgp, objWSMacro
Dim ErrChampsDSCGP As String

    If NomFicDSCGP = "" Then
        GoTo Erreur1 'Fichier DSCGP non renseigné
    End If
    
    On Error GoTo Erreur2 'Fichier DSCGP introuvable
    Set objexcel = CreateObject("EXCEL.APPLICATION")
    Set objWorkBook = objexcel.Workbooks.Open(CStr(NomFicDSCGP))
    
    On Error GoTo Erreur3
    If VerDscgp = 1 Then
       'DSCGP Version 1
       
       Set objWSDscgp = objWorkBook.Sheets("Page1")
        ValDscgp.NumLot = objWSDscgp.cells(8, 4) ' N° du lot
        ValDscgp.NumGrilleAss = objWSDscgp.cells(9, 4) ' N° de la grille 'N° d'outillage
        ValDscgp.NumGrilleAssSym = objWSDscgp.cells(10, 4) ' N° de la symétrique 'N° d'outillage sym
        ValDscgp.NumDetromp = objWSDscgp.cells(52, 8) 'N° de détrompage
        ValDscgp.design = objWSDscgp.cells(9, 8) ' Désignation de la grille
        ValDscgp.DesignSym = objWSDscgp.cells(10, 8) ' Désignation de la grille sym
        ValDscgp.NumEnvAvion = objWSDscgp.cells(12, 9) ' N° env avion CAO
        ValDscgp.Mat = objWSDscgp.cells(30, 4) ' Matière grillle
        ValDscgp.NumPiecesPerc = objWSDscgp.cells(18, 8) ' N° pièces percées
        ValDscgp.Site = objWSDscgp.cells(16, 8) ' Site
        ValDscgp.NumProgAvion = objWSDscgp.cells(8, 9) ' N° programme avion
        ValDscgp.Anod = objWSDscgp.cells(57, 4) ' Anodisation
        ValDscgp.Dtemplate = objWSDscgp.cells(63, 4) ' DTEMPLATE
        ValDscgp.Color = objWSDscgp.cells(57, 8) ' Color Coding
        ValDscgp.SystNum = objWSDscgp.cells(76, 7) ' Sytem Numérotation
        ValDscgp.CoteAvion = objWSDscgp.cells(17, 8) ' Coté Avion
        ValDscgp.Exemplaire = objWSDscgp.cells(11, 5) 'Exemplaire
        ValDscgp.Pres_Pinules = ""
        ValDscgp.Nb_Pinules = ""
        ValDscgp.GravureInf = ""
        ValDscgp.GravureSup = ""
        ValDscgp.GravureLat1 = ""
        ValDscgp.GravureLat2 = ""
        ValDscgp.GravureLat3 = ""
        ValDscgp.GravureLat4 = ""
        
    ElseIf VerDscgp = 2 Then
        'DSCGP Version 2
        On Error GoTo Erreur4
            Set objWSDscgp = objWorkBook.Sheets("DSCGP")
            Set objWSMacro = objWorkBook.Sheets("MACRO")
        On Error GoTo Erreur3
            ErrChampsDSCGP = NomCelDSCGP.CoteAvion
        ValDscgp.CoteAvion = objWSDscgp.range(ExtractCell(objWorkBook.names.Item(NomCelDSCGP.CoteAvion))) ' Coté Avion
            ErrChampsDSCGP = NomCelDSCGP.NumLot
        ValDscgp.NumLot = objWSDscgp.range(ExtractCell(objWorkBook.names.Item(NomCelDSCGP.NumLot))) ' N° du lot
            ErrChampsDSCGP = NomCelDSCGP.NumGrilleAss
        ValDscgp.NumGrilleAss = objWSDscgp.range(ExtractCell(objWorkBook.names.Item(NomCelDSCGP.NumGrilleAss))) ' N° de la grille dessinée 'N° d'outillage
            ErrChampsDSCGP = NomCelDSCGP.NumGrilleAssSym
        ValDscgp.NumGrilleAssSym = objWSDscgp.range(ExtractCell(objWorkBook.names.Item(NomCelDSCGP.NumGrilleAssSym))) ' N° de la symétrique 'N° d'outillage
            ErrChampsDSCGP = NomCelDSCGP.design
        ValDscgp.design = objWSDscgp.range(ExtractCell(objWorkBook.names.Item(NomCelDSCGP.design))) ' ' Désignation de la grille dessinée
            ErrChampsDSCGP = NomCelDSCGP.DesignSym
        ValDscgp.DesignSym = objWSDscgp.range(ExtractCell(objWorkBook.names.Item(NomCelDSCGP.DesignSym))) ' Désignation de la grille sym
            ErrChampsDSCGP = NomCelDSCGP.NumDetromp
        ValDscgp.NumDetromp = objWSDscgp.range(ExtractCell(objWorkBook.names.Item(NomCelDSCGP.NumDetromp))) 'N° de détrompage
            ErrChampsDSCGP = NomCelDSCGP.NumEnvAvion
        ValDscgp.NumEnvAvion = objWSDscgp.range(ExtractCell(objWorkBook.names.Item(NomCelDSCGP.NumEnvAvion))) ' N° env avion CAO
            ErrChampsDSCGP = NomCelDSCGP.Mat
        ValDscgp.Mat = objWSDscgp.range(ExtractCell(objWorkBook.names.Item(NomCelDSCGP.Mat))) ' Matière grillle
            'ErrChampsDSCGP
        ValDscgp.NumPiecesPerc = "" ' N° pièces percées non impléménté sur DSCGP version 2
            ErrChampsDSCGP = NomCelDSCGP.Site
        ValDscgp.Site = objWSDscgp.range(ExtractCell(objWorkBook.names.Item(NomCelDSCGP.Site))) ' Site
            ErrChampsDSCGP = NomCelDSCGP.NumProgAvion
        ValDscgp.NumProgAvion = objWSDscgp.range(ExtractCell(objWorkBook.names.Item(NomCelDSCGP.NumProgAvion))) ' N° programme avion
            ErrChampsDSCGP = NomCelDSCGP.Accessoires
        ValDscgp.Anod = ChercheVal(objWSDscgp.range(ExtractCell(objWorkBook.names.Item(NomCelDSCGP.Accessoires))), nAnod)
            ErrChampsDSCGP = NomCelDSCGP.Dtemplate
        ValDscgp.Dtemplate = objWSDscgp.range(ExtractCell(objWorkBook.names.Item(NomCelDSCGP.Dtemplate))) ' DTEMPLATE
            ErrChampsDSCGP = NomCelDSCGP.Accessoires
        ValDscgp.Color = ChercheVal(objWSDscgp.range(ExtractCell(objWorkBook.names.Item(NomCelDSCGP.Accessoires))), nColcod)
            ErrChampsDSCGP = NomCelDSCGP.SystNum
        ValDscgp.SystNum = objWSDscgp.range(ExtractCell(objWorkBook.names.Item(NomCelDSCGP.SystNum))) ' Sytem Numérotation
            ErrChampsDSCGP = NomCelDSCGP.Exemplaire
        ValDscgp.Exemplaire = objWSDscgp.range(ExtractCell(objWorkBook.names.Item(NomCelDSCGP.Exemplaire))) ' exemplaire
            ErrChampsDSCGP = NomCelDSCGP.Pres_Pinules
        ValDscgp.Pres_Pinules = objWSDscgp.range(ExtractCell(objWorkBook.names.Item(NomCelDSCGP.Pres_Pinules))) ' Présence Pinules
            ErrChampsDSCGP = NomCelDSCGP.Nb_Pinules
        ValDscgp.Nb_Pinules = objWSDscgp.range(ExtractCell(objWorkBook.names.Item(NomCelDSCGP.Nb_Pinules))) ' Nombre de pinules
        'Récupération de la liste des gravures
        ErrChampsDSCGP = NomCelDSCGP.GravureSup
            ValDscgp.GravureSup = ConvString(objWSMacro.range(ExtractCell(objWorkBook.names.Item(NomCelDSCGP.GravureSup))))
        ErrChampsDSCGP = NomCelDSCGP.GravureInf
            ValDscgp.GravureInf = ConvString(objWSMacro.range(ExtractCell(objWorkBook.names.Item(NomCelDSCGP.GravureInf))))
        ErrChampsDSCGP = NomCelDSCGP.GravureLat1
            ValDscgp.GravureLat1 = ConvString(objWSMacro.range(ExtractCell(objWorkBook.names.Item(NomCelDSCGP.GravureLat1))))
        ErrChampsDSCGP = NomCelDSCGP.GravureLat2
            ValDscgp.GravureLat2 = ConvString(objWSMacro.range(ExtractCell(objWorkBook.names.Item(NomCelDSCGP.GravureLat2))))
        ErrChampsDSCGP = NomCelDSCGP.GravureLat3
            ValDscgp.GravureLat3 = ConvString(objWSMacro.range(ExtractCell(objWorkBook.names.Item(NomCelDSCGP.GravureLat3))))
        ErrChampsDSCGP = NomCelDSCGP.GravureLat4
            ValDscgp.GravureLat4 = ConvString(objWSMacro.range(ExtractCell(objWorkBook.names.Item(NomCelDSCGP.GravureLat4))))

    End If
    
    objWorkBook.Close False

GoTo Fin

Erreur1:
    Err.Raise vbObjectError + 525, , "Fichier DSCGP non renseigné"
    GoTo Fin
Erreur2:
    Err.Raise vbObjectError + 526, , "Fichier DSCGP Introuvable ! "
    GoTo Fin
Erreur3:
    Err.Raise vbObjectError + 527, , "Erreur dans la collecte des Info du DSCGP champs : " & ErrChampsDSCGP
    GoTo Fin
Erreur4:
    Err.Raise vbObjectError + 530, , "Erreur dans le fichier DSCGP !"
    GoTo Fin
Fin:
End Property

Property Get NumduLot()
'N° de fichier du lot de grille
    'Si le N° du Lot est vide on utilise le n° de la grille
    If ValDscgp.NumLot = "" Then
        If CoteConception = "GAUCHE" Or CoteConception = "CENTRE" Then
            NumduLot = ValDscgp.NumGrilleAss
        Else
            NumduLot = ValDscgp.NumGrilleAssSym
        End If
    Else
        NumduLot = ValDscgp.NumLot
    End If
    If NumduLot <> "" Then
        Select Case SystemNum
            Case 1, 11
                NumduLot = NumduLot & "-00000X01"
            Case 2, 12
                If NumLot = "" Then
                    NumduLot = Left(NumLot, Len(NumduLot) - 3)
                End If
            Case 3, 13
                NumduLot = Left(NumduLot, 14) & "000"
        End Select
    End If
End Property

Property Get chkFrmNumLot() As Boolean
'Check si le Nbe de carratère du Numéro de lot est correct
    chkFrmNumLot = True
    Select Case SystemNum
        Case 1, 11
            If Len(Me.NumduLot) <> 19 Then
                chkFrmNumLot = False
            ElseIf Right(Me.NumLot, 9) <> "-00000X01" Then
                chkFrmNumLot = False
            End If
        Case 2, 12
            If Me.NumLot = "" And (Len(Me.NumduLot) <> 14) Then
                chkFrmNumLot = False
            ElseIf Me.NumLot <> "" And (Me.NumduLot <> Me.NumLot) Then
                    chkFrmNumLot = False
            End If
        Case 3, 13
            If Len(Me.NumduLot) <> 17 Then
                chkFrmNumLot = False
            ElseIf Right(Me.NumLot, 3) <> "000" Then
                chkFrmNumLot = False
            End If
    End Select
End Property

Property Get NumLot()
'Renvoi le numéro du lot dnas le DSCGP (sans l'extension)
    NumLot = ValDscgp.NumLot
End Property

Property Get NumRadGrille()
'N° de la grille (Radical)
    Select Case CoteConception
        Case "GAUCHE"
            NumRadGrille = ValDscgp.NumGrilleAss
        Case "DROIT"
            NumRadGrille = ValDscgp.NumGrilleAssSym
        Case "CENTRE"
            NumRadGrille = ValDscgp.NumGrilleAss
    End Select
End Property

Property Get CoteConception()
'Cote de conception
Dim lst As String
    lst = "DROITE-GAUCHE-CENTRE"
    CoteConception = ValDscgp.CoteAvion
    If InStr(1, lst, CoteConception, vbTextCompare) = 0 Or CoteConception = "" Then
        Err.Raise vbObjectError + 531, , "Le coté de conception est mal documenté dans le DSCGP"
    End If
    
End Property

Property Get NumGrille()
'N° du fichier de la grille assemblée
    NumGrille = ""
    Select Case CoteConception
        Case "GAUCHE"
            NumGrille = ValDscgp.NumGrilleAss
        Case "DROIT"
            NumGrille = ValDscgp.NumGrilleAssSym
        Case "CENTRE"
            NumGrille = ValDscgp.NumGrilleAss
    End Select
    If NumGrille <> "" Then
        Select Case SystemNum
            Case 1, 11
                NumGrille = NumGrille & "-00000Y01"
            Case 2, 12
                NumGrille = NumGrille
            Case 3, 13
                NumGrille = NumGrille 'Le numéro de grille complet est renseigné dans le DSCGP modif 05/12/16
            Case Else
                Err.Raise vbObjectError + 528, , "Le Type d'arborescence n'est pas documenté dans le DSCGP !"
                GoTo Fin
        End Select
    Else
        Err.Raise vbObjectError + 529, , "Le N° de grille n'est pas documenté ou ne correspond pas au coté de conception !"
        GoTo Fin:
    End If
Fin:
End Property

Property Get NumGrilleNue()
'N° du fichier de la grille nue
    NumGrilleNue = ""
    Select Case CoteConception
        Case "GAUCHE"
            NumGrilleNue = ValDscgp.NumGrilleAss
        Case "DROIT"
            NumGrilleNue = ValDscgp.NumGrilleAssSym
        Case "CENTRE"
            NumGrilleNue = ValDscgp.NumGrilleAss
    End Select
    If NumGrilleNue <> "" Then
        Select Case SystemNum
            Case 1
                NumGrilleNue = NumGrilleNue & "-00002P01"
            Case 2
                NumGrilleNue = NumGrilleNue & "2"
            Case 3
                NumGrilleNue = Left(ValDscgp.NumGrilleAss, 14) & ExtNumGrilleNue(Right(ValDscgp.NumGrilleAss, 4))
            Case 11
                NumGrilleNue = NumGrilleNue & "-00003P01"
            Case 12
                NumGrilleNue = NumGrilleNue & "3"
            Case 13
                NumGrilleNue = Left(ValDscgp.NumGrilleAssSym, 14) & ExtNumGrilleNue(Right(ValDscgp.NumGrilleAssSym, 4))
            Case Else
                Err.Raise vbObjectError + 528, , "Le Type d'arborescence n'est pas documenté dans le DSCGP !"
                GoTo Fin:
        End Select
    Else
        Err.Raise vbObjectError + 529, , "Le N° de grille n'est pas documenté ou ne correspond pas au coté de conception !"
        GoTo Fin
    End If
Fin:
End Property

Property Get NumGrilleSym()
'N° du fichier de la grille symétrique assemblée
    NumGrilleSym = ""
    Select Case CoteConception
        Case "GAUCHE"
            NumGrilleSym = ValDscgp.NumGrilleAssSym
        Case "DROIT"
            NumGrilleSym = ValDscgp.NumGrilleAss
        Case "CENTRE"
            NumGrilleSym = ""
    End Select
    If NumGrilleSym <> "" Then
        Select Case SystemNum
            Case 1, 11
                NumGrilleSym = NumGrilleSym & "-00000Y01"
            Case 2, 12
                NumGrilleSym = NumGrilleSym
            Case 3, 13
                NumGrilleSym = NumGrilleSym 'Le numéro de grille complet est renseigné dans le DSCGP modif 05/12/16
            Case Else
                 Err.Raise vbObjectError + 528, , "Le Type d'arborescence n'est pas documenté dans le DSCGP !"
                 GoTo Fin
        End Select
    'Else
        'Err.Raise vbObjectError + 529, , "Le N° de grille n'est pas documenté ou ne correspond pas au coté de conception !"
        'GoTo fin
    End If
Fin:
End Property

Property Get NumGrilleSymNue()
'N° du fichier de la grille la symétrique nue
    NumGrilleSymNue = ""
    Select Case CoteConception
        Case "GAUCHE"
            NumGrilleSymNue = ValDscgp.NumGrilleAssSym
        Case "DROIT"
            NumGrilleSymNue = ValDscgp.NumGrilleAss
        Case "CENTRE"
        NumGrilleSymNue = ""
    End Select
    If NumGrilleSymNue <> "" Then
        Select Case SystemNum
            Case 1
                NumGrilleSymNue = NumGrilleSymNue & "-00003P01"
            Case 2
                NumGrilleSymNue = NumGrilleSymNue & "3"
            Case 3
                NumGrilleSymNue = Left(ValDscgp.NumGrilleAssSym, 14) & ExtNumGrilleNue(Right(ValDscgp.NumGrilleAssSym, 4))
            Case 11
                NumGrilleSymNue = NumGrilleSymNue & "-00002P01"
            Case 12
                NumGrilleSymNue = NumGrilleSymNue & "2"
            Case 13
                NumGrilleSymNue = Left(ValDscgp.NumGrilleAss, 14) & ExtNumGrilleNue(Right(ValDscgp.NumGrilleAss, 4))
            Case Else
                 Err.Raise vbObjectError + 528, , "Le Type d'arborescence n'est pas documenté dans le DSCGP !"
                 GoTo Fin
        End Select
    'Else
        'Err.Raise vbObjectError + 529, , "Le N° de grille n'est pas documenté ou ne correspond pas au coté de conception !"
        'GoTo fin
    End If
Fin:
End Property

Property Get NumPartU01()
'N° de la part U01
    NumPartU01 = ""
    Select Case CoteConception
        Case "GAUCHE"
            NumPartU01 = ValDscgp.NumGrilleAss
        Case "DROIT"
            NumPartU01 = ValDscgp.NumGrilleAssSym
        Case "CENTRE"
            NumPartU01 = ValDscgp.NumGrilleAss
    End Select
    If NumPartU01 <> "" Then
        Select Case SystemNum
            Case 1, 11
                NumPartU01 = NumPartU01 & "-00001U01"
            'Case 2, 12 'Pas defini a ce jour
            '    NumPartU01 = NumPartU01 & ""
            'Case 3, 13 'Pas defini a ce jour
            '    NumPartU01 = NumPartU01 & ""
            Case Else
                NumPartU01 = ""
        End Select
    End If
End Property

Property Get NumPartU01Sym()
'N° de la part U01
    NumPartU01Sym = ""
    Select Case CoteConception
        Case "GAUCHE"
            NumPartU01Sym = ValDscgp.NumGrilleAssSym
        Case "DROIT"
            NumPartU01Sym = ValDscgp.NumGrilleAss
        Case "CENTRE"
            NumPartU01Sym = ""
    End Select
    If NumPartU01Sym <> "" Then
        Select Case SystemNum
            Case 1, 11
                NumPartU01Sym = NumPartU01Sym & "-00001U01"
            'Case 2, 12 'Pas defini a ce jour
            '    NumPartU01Sym = NumPartU01Sym & ""
            'Case 3, 13 'Pas defini a ce jour
            '    NumPartU01Sym = NumPartU01Sym & ""
            Case Else
                NumPartU01Sym = ""
        End Select
    End If
End Property

Property Get NumPartDet()
'N° de la part de détrompage
    NumPartDet = ValDscgp.NumDetromp
End Property

Property Get DesignGrille()
'Désignation de la grille
    Select Case CoteConception
        Case "GAUCHE"
            DesignGrille = ValDscgp.design
        Case "DROIT"
            DesignGrille = ValDscgp.DesignSym
        Case "CENTRE"
            DesignGrille = ValDscgp.design
    End Select
End Property

Property Get DesignGrilleSym()
'Désignation de la grille sym
    Select Case CoteConception
        Case "GAUCHE"
            DesignGrilleSym = ValDscgp.DesignSym
        Case "DROIT"
            DesignGrilleSym = ValDscgp.design
        Case "CENTRE"
            DesignGrilleSym = ""
    End Select
End Property

Property Get EnvAvionCAO()
'N° env avion CAO
    EnvAvionCAO = ValDscgp.NumEnvAvion
End Property

Property Get MatGrille()
'Matière grillle
    MatGrille = ValDscgp.Mat
End Property

Property Get PiecesPercees()
'N° pièces percées
    PiecesPercees = ValDscgp.NumPiecesPerc
End Property

Property Get Site()
'Site
    Site = ValDscgp.Site
End Property

Property Get NoProgAvion()
'N° programme avion
    NoProgAvion = ValDscgp.NumProgAvion
End Property

Property Get CoteAvion()
'Coté de conception de la grille
    CoteAvion = ValDscgp.CoteAvion
End Property

Property Get Observations()
'Anodisation et Color Coding
    Dim str As String
    If (UCase(ValDscgp.Anod)) = "OUI" Then
        str = nAnod
    End If
    If (UCase(ValDscgp.Color)) = "OUI" Then
        If str = "" Then
            str = nColcod
        Else
            str = str & " / " & nColcod
        End If
    End If
    Observations = str
End Property

Property Get Dtemplate()
'DTEMPLATE
    Dtemplate = ValDscgp.Dtemplate
End Property

Property Get Exemplaire()
'Gestion des indices et exemplaire
    Exemplaire = ValDscgp.Exemplaire
End Property

Property Get Pres_Pinules()
'Presence de pinules
    Pres_Pinules = ValDscgp.Pres_Pinules
End Property

Property Get Nb_Pinules()
'Nombre de pinules
    Nb_Pinules = ValDscgp.Nb_Pinules
End Property

Property Get NumOutillage()
'Numero d'outillage
    NumOutillage = ValDscgp.NumGrilleAss
End Property

Property Get NumOutillageSym()
'Numero d'outillage Symétrique
    NumOutillageSym = ValDscgp.NumGrilleAssSym
End Property

'## modif du 18/03/16. info MLC -> le N° d'outillage = n_grille_g du DSCGP
'Property Get NumOutillage()
''Numero d'outillage
'    'Si le N° du Lot est vide on utilise le n° de la grille
'    If DSCGP.NumLot = "" Then
'        NumOutillage = DSCGP.NumGrilleAss
'    Else
'        NumOutillage = DSCGP.NumLot
'    End If
'End Property
'
'Property Get NumOutillageSym()
''Numero d'outillage Symétrique
'    'Si le N° du Lot est vide on utilise le n° de la grille
'    If DSCGP.NumLot = "" Then
'        NumOutillage = DSCGP.NumGrilleAssSym
'    Else
'        NumOutillage = DSCGP.NumLot
'    End If
'End Property
'## fin modif du 18/03/16

Property Get SystemNum()
'System de Numérotation
'renvois un code pour les 3 system de numérotation
'1 pour A350 Standard, 2 pour Legacy etc
'Ajoute 10 pour les grille droite seule
'11 pour A350 Standard droite, 12 pour Legacy Droite etc..

If VerDscgp = 1 Then
    If ValDscgp.SystNum = "Numérotation A350 standard (Broughton, Nantes, Saint-Nazaire)" Then
        SystemNum = 1
    ElseIf ValDscgp.SystNum = "Numérotation LEGACY (tout programme autre qu'A350)" Then
        SystemNum = 2
    ElseIf ValDscgp.SystNum = "Numérotation A350 type LEGACY (Toulouse)" Then
        SystemNum = 3
    Else
        GoTo Erreur
    End If
Else
    If ValDscgp.SystNum = "A350 standard (Broughton, Nantes, Saint-Nazaire)" Then
        SystemNum = 1
    ElseIf ValDscgp.SystNum = "LEGACY (tout programme autre qu'A350)" Then
        SystemNum = 2
    ElseIf ValDscgp.SystNum = "A350 type LEGACY (Toulouse)" Then
        SystemNum = 3
    Else
        GoTo Erreur
    End If
End If
    If InStr(1, ValDscgp.CoteAvion, "DROIT", vbTextCompare) > 0 Then
        SystemNum = SystemNum + 10
    End If
GoTo Fin

Erreur:
    Err.Raise vbObjectError + 528, , "Le Type d'arborescence n'est pas documenté dans e DSCGP !"
Fin:
End Property

Property Get GravSup()
'Renvois le text de la gravure sup
    GravSup = ValDscgp.GravureSup
End Property

Property Get GravInf()
'Renvois le text de la gravure Inf
    GravInf = ValDscgp.GravureInf
End Property

Property Get GravLat1()
'Renvois le text de la face Latéral 1
    GravLat1 = ValDscgp.GravureLat1
End Property

Property Get GravLat2()
'Renvois le text de la face Latéral 2
    GravLat2 = ValDscgp.GravureLat2
End Property

Property Get GravLat3()
'Renvois le text de la face Latéral 3
    GravLat3 = ValDscgp.GravureLat3
End Property

Property Get GravLat4()
'Renvois le text de la face Latéral 4
    GravLat4 = ValDscgp.GravureLat4
End Property

Private Function ExtractCell(strg As String) As String
'extrait de la string passée en argument la partie correspondant a une cellule
'strg = "DSCGP!$C$5" => ExtractCell = "C5"
'strg = "Macro!$C$29:$C$63" =>
Dim tmpstrg As String
    If InStr(1, strg, "!", vbTextCompare) > 0 Then
        tmpstrg = Right(strg, Len(strg) - InStr(1, strg, "!", vbTextCompare))
    Else
        tmpstrg = ""
    End If
    If InStr(1, tmpstrg, "$", vbTextCompare) > 0 Then
        If InStr(2, tmpstrg, "$", vbTextCompare) > 0 Then
            ExtractCell = Mid(tmpstrg, InStr(1, tmpstrg, "$", vbTextCompare) + 1, Len(tmpstrg) - InStr(1, tmpstrg, "$", vbTextCompare) - InStr(2, tmpstrg, "$", vbTextCompare))
            ExtractCell = ExtractCell & Right(tmpstrg, Len(tmpstrg) - InStr(2, tmpstrg, "$", vbTextCompare))
        End If
    End If
End Function

Private Function ExtractPlage(strg As String) As String
'Extrait de la string la partie correspondant à une plage de cellules.
'srtg =DSCGP!$C$184:$C$188 => ExtractPlage(str) = C184:C188
Dim tmpstrg(1 To 2) As String
Dim tmpplage(1 To 2) As String
Dim i As Integer
    If InStr(1, strg, "!", vbTextCompare) > 0 Then
        strg = Right(strg, Len(strg) - InStr(1, strg, "!", vbTextCompare))
    Else
        strg = ""
    End If
    If InStr(1, strg, ":", vbTextCompare) > 0 Then
        tmpstrg(2) = Right(strg, Len(strg) - InStr(1, strg, ":", vbTextCompare))
        tmpstrg(1) = Left(strg, InStr(1, strg, ":", vbTextCompare) - 1)
    End If
    For i = 1 To 2
        If InStr(1, tmpstrg(i), "$", vbTextCompare) > 0 Then
            If InStr(2, tmpstrg(i), "$", vbTextCompare) > 0 Then
                tmpplage(i) = Mid(tmpstrg(i), InStr(1, tmpstrg(i), "$", vbTextCompare) + 1, Len(tmpstrg(i)) - InStr(1, tmpstrg(i), "$", vbTextCompare) - InStr(2, tmpstrg(i), "$", vbTextCompare))
                tmpplage(i) = tmpplage(i) & Right(tmpstrg(i), Len(tmpstrg(i)) - InStr(2, tmpstrg(i), "$", vbTextCompare))
            End If
        End If
    Next i
    
    ExtractPlage = tmpplage(1) & ":" & tmpplage(2)
    
End Function

Private Function ExtractLig(strg As String) As Integer
'Extrait de la string passée en argument la partie correspondant à la ligne de la cellule
'strg = "G5" => ExtractLig = 5
Dim tmpstrg As String
Dim i As Integer
For i = 1 To Len(strg)
    If IsNumeric(Mid(strg, i, 1)) Then
        tmpstrg = tmpstrg & Mid(strg, i, 1)
    End If
Next
    ExtractLig = CInt(tmpstrg)
End Function

Private Function ExtractCol(strg As String) As String
'Extrait de la string passée en argument la partie correspondant à la colonne de la cellule
'strg = "G5" => ExtractCol = "G"
Dim tmpstrg As String
Dim i As Integer
For i = 1 To Len(strg)
    If IsNumeric(Mid(strg, i, 1)) Then
        Exit For
    Else
        tmpstrg = tmpstrg & Mid(strg, i, 1)
    End If
Next
ExtractCol = tmpstrg
End Function

Private Function ChercheVal(ByRef Plage, val As String) As String
'Recherche dans la plage "lst_acces" (=DSCGP!$C$184:$C$188)
'la présence des termes "color coding" et "anodisation"
'renvois "OUI" ou  "NON" si ça match
Dim i As Integer
ChercheVal = "NON"
    For i = 1 To Plage.Count
        If Plage(i) = val Then
            ChercheVal = "OUI"
        End If
    Next
End Function

Private Function ConvString(ByRef Plage) As String
'Converti le tableau passe en argument en string avec retours a la ligne
'supprime les lignes vides
Dim i As Integer
ConvString = ""
    For i = 1 To Plage.Count
        If Plage(i) <> "" Then
            ConvString = ConvString & Plage(i) & Chr(10)
        End If
    Next
    'ajoute une ligne vide a la fin
    ConvString = ConvString & "" & Chr(10)
End Function

Private Function ExtNumGrilleNue(str As String) As String
'Calcule l'extention du numéro de grille nue en fonction de l'extension du numéro de grille Ass
'A000 -> 0002, A001 -> 0003, A002 -> 0004 etc
Dim istr   As Integer
    On Error Resume Next
    istr = CInt(Right(str, 3))
    If Err.Number <> 0 Then
        Err.Clear
        ExtNumGrilleNue = ""
    Else
        ExtNumGrilleNue = form4C(istr + 2)
    End If
On Error GoTo 0
End Function

Private Function form4C(i As Integer) As String
'formate l'entier passé en argument en string de 4 carractère (avec des 0 devant)
Dim str As String
    str = CStr(i)
    While Len(str) < 4
        str = "0" & str
    Wend
    form4C = str
End Function

