Attribute VB_Name = "C1_Rename_AB"
Option Explicit
'*********************************************************************
'* Macro : C1_Rename_AB
'*
'* Fonctions :  Macro de renommage des points A, B et des STD
'*              Renomme les points dans l'odre de classement dans le set géométrique
'*
'* Version : 4
'* Création :  SVI
'* Modification : 05/08/14 CFR
'*                Ajout test PartBody/corps principal
'*                Ajout test existance Sets géométriques
'* Modification : 14/02/15 CFR
'*                Prise en compte de la class PartGrille
'* Modification : 15/04/15 CFR
'*                modification de renumérotation pour conserver la partie complémentaire du nom des points (asnaxxxxxxx)
'* Modification : 24/04/15 CFR
'*                Refonte de la numérotation
'*                Recupération du nom des fauxPtA pour renommer les pts A, B et STD
'* Modification : 18/04/16 CFR
'*                Prise en compte des droites explicite (std isolé) sans point d'origine ni d'etrémité
'*                Prise en compte des points A et B isolés
'* Modification : 17/07/16 CFR
'*                Ajout boite de dialogue options de renommage
'*                Ajout du renommage dans l'ordre des Fasteners
'* Modification : 19/01/17 CFR
'*                Ajout renommage éléments isolés
'**********************************************************************
Sub catmain()

'Log de l'utilisation de la macro
LogUtilMacro nPath, nFicLog, nMacro, "C1_Rename_AB", VMacro

Dim instance_catpart_grille_nue As PartDocument
Dim nParent As String, nFxPtA As String, nfxPtb As String
Dim PtA As HybridShape, PtB As HybridShape
Dim FxPtA As HybridShape, FxPtb As HybridShape
Dim LigParent As HybridShape
Dim i As Long, j As Long
Dim Distance As Double
Dim fastOK As Boolean 'Le Fastener à été trouvé dans la collection
Dim PtBOK As Boolean 'Le pts B à été trouvé dans la collection
Dim TestHBody As HybridBody
'Memorise les choix de la boite de dialogue
Dim OrdreFast As Boolean
Dim typeNum As String, ComplementNom As String

Dim coordPtA As c_Coord, CoordPtB As c_Coord

Dim tcPtA As c_Pt 'Collection des points A
Dim tcPtAs As c_Pts
Dim tcPtB As c_Pt 'Collection des points B
Dim tcPtBs As c_Pts
Dim tcFxPtA As c_Pt 'Collection des points faux A
Dim tcFxPtAs As c_Pts
Dim tcFxPtB As c_Pt 'Collection des points faux B
Dim tcFxPtBs As c_Pts
Dim tFast As c_Fastener 'Colllection des Fasteners
Dim tcrdFast As c_Coord
Dim tcPtStdFast As c_PtStdFast 'Colection de points, faux points , std et fasteners
Dim tcPtStdFasts As c_PtStdFasts
Dim tcPtStdFastTri As c_PtStdFasts 'Colection de points, faux points , std et fasteners triés
Dim PtIsoles As Boolean ' Choix du renommage d'éléments sans historique
Dim IsolPtA As Boolean, IsolPtB As Boolean, IsolFeet As Boolean, IsolPin As Boolean
Dim RenSets As Collection
'Dim Renset As Object

'Ouvre la boite de dlg "Frm_Renomage"
    Load Frm_Renomage
    Frm_Renomage.Show
    
'Mémorisation des choix
    If Frm_Renomage.CB_Isol = True Then
        PtIsoles = True
        IsolPtA = Frm_Renomage.CB_PtA.Value
        IsolPtB = Frm_Renomage.CB_PtB.Value
        IsolFeet = Frm_Renomage.CB_Feet.Value
        IsolPin = Frm_Renomage.CB_Pinules.Value
        Set RenSets = New Collection
        RenSets.Add IsolPtA, nHBPtA
        RenSets.Add IsolPtB, nHBPtB
        RenSets.Add IsolFeet, nHBFeet
        RenSets.Add IsolPin, nHBPin
    Else
        PtIsoles = False
        If Frm_Renomage.Rbt_RefSTD Then OrdreFast = True
        If Frm_Renomage.RbtNumNomFast Then typeNum = "nFast"
        If Frm_Renomage.RbtNumCommentFast Then typeNum = "Comment"
        If Frm_Renomage.RbtNumPtA Then typeNum = "PtA"
    End If
    
'Sort du programme si click sur bouton Annuler dans Frm_Renomage
    If Not (Frm_Renomage.ChB_OkAnnule) Then
        Unload Frm_Renomage
        Exit Sub
    End If
    Unload Frm_Renomage
    
    '---------------------------
    ' Checker l'environnement
    '---------------------------
        
    Err.Clear
    On Error Resume Next
    Set instance_catpart_grille_nue = CATIA.ActiveDocument
    If Err.Number <> 0 Then
        MsgBox "Le document de la fenêtre courante n'est pas un CATPart !", vbCritical, "Environnement incorrect"
        End
    End If
    On Error GoTo 0
    
    Dim GrilleActive As New c_PartGrille
      
    If PtIsoles Then 'Renomage d'éléments sans historique
        For i = 1 To RenSets.Count
            If RenSets.Item(i) = True Then
                RenomPtIsoles GrilleActive, i
            End If
        Next
        'liberation des classe
        Set GrilleActive = Nothing
        Set RenSets = Nothing
        MsgBox "Renommage des points isolés terminé !", vbInformation, "Fin de traitement"
        End
    Else
    
        'Initialisation des classes
        Set coordPtA = New c_Coord
        Set CoordPtB = New c_Coord
        Set tcPtA = New c_Pt
        Set tcPtAs = New c_Pts
        Set tcPtB = New c_Pt
        Set tcPtBs = New c_Pts
        Set tcFxPtA = New c_Pt
        Set tcFxPtAs = New c_Pts
        Set tcFxPtB = New c_Pt
        Set tcFxPtBs = New c_Pts
        Set tFast = New c_Fastener
        Set tcrdFast = New c_Coord
        Set tcPtStdFast = New c_PtStdFast
        Set tcPtStdFasts = New c_PtStdFasts
        Set tcPtStdFastTri = New c_PtStdFasts
    
        'Vérification de l'existence des sets géométriques
        On Error GoTo Erreur
        Set TestHBody = GrilleActive.Hb(nHBPtA)
        Set TestHBody = GrilleActive.Hb(nHBPtB)
        Set TestHBody = GrilleActive.Hb(nHBStd)
        Set TestHBody = GrilleActive.Hb(nHBRefExtIsol)
        Set TestHBody = Nothing
        On Error GoTo 0
    'GrilleActive.PartGrille.InWorkObject = GrilleActive.HB_RefExtIsole
    
    'Vérification que le set est updaté
        Dim StatusPartUpdated As Boolean
        StatusPartUpdated = GrilleActive.PartGrille.IsUpToDate(GrilleActive.Hb(nHBRefExtIsol))
     
        If StatusPartUpdated = False Then
            MsgBox "Le Part n'est pas a jour. Updatez le !", vbCritical, "Erreur"
            End
        End If
    
        'Collecte des Pts B
        For i = 1 To GrilleActive.Hb(nHBPtB).HybridShapes.Count
            Set PtB = GrilleActive.Hb(nHBPtB).HybridShapes.Item(i)
            Set tcPtB = InfoPt(PtB, GrilleActive.Hb(nHBStd))
            tcPtBs.Add tcPtB.nom, tcPtB.Crd, tcPtB.Parent
        Next i
        
        'Collecte des PtsA
        For i = 1 To GrilleActive.Hb(nHBPtA).HybridShapes.Count
            Set PtA = GrilleActive.Hb(nHBPtA).HybridShapes.Item(i)
            Set tcPtA = InfoPt(PtA, GrilleActive.Hb(nHBStd))
            tcPtAs.Add tcPtA.nom, tcPtA.Crd, tcPtA.Parent
        Next i
        
        'collecte des faux Points A et des std a partir de la liste des points A
        For i = 1 To tcPtAs.Count
            Set tcPtA = tcPtAs.Item(i)
            nParent = tcPtA.Parent
            If GrilleActive.Hb(nHBStd).HybridShapes.Item(nParent) Is Nothing Then 'La ligne STD n'a pas été trouvée dans le set STD
                nFxPtA = "faux A" & i
                nfxPtb = "faux B" & i
                With tcFxPtA
                    .nom = nFxPtA
                    .Crd = Nothing
                    .Parent = "Ligne_STD_Inconnue"
                End With
                With tcFxPtB
                    .nom = nfxPtb
                    .Crd = Nothing
                    .Parent = "Ligne_STD_Inconnue"
                End With
            Else
                Set LigParent = GrilleActive.Hb(nHBStd).HybridShapes.Item(nParent)
                nFxPtA = LigParent.PtOrigine.DisplayName
                On Error Resume Next
                Set FxPtA = GrilleActive.Hb(nHBPtConst).HybridShapes.Item(nFxPtA)
                If Err.Number <> 0 Then 'La ligne STD n'a pas d'origine
                    Err.Clear
                    With tcFxPtA
                        .nom = "faux A" & i
                        .Crd = Nothing
                        .Parent = nParent
                    End With
                    On Error GoTo 0
                Else
                    Set coordPtA = CoordPt(GrilleActive.PartGrille, FxPtA)
                    With tcFxPtA
                        .nom = nFxPtA
                        .Crd = coordPtA
                        .Parent = nParent
                    End With
                End If
                
                nfxPtb = LigParent.PtExtremity.DisplayName
                On Error Resume Next
                Set FxPtb = GrilleActive.Hb(nHBPtConst).HybridShapes.Item(nfxPtb)
                If Err.Number <> 0 Then 'La ligne STD n'a pas d'origine
                    Err.Clear
                    With tcFxPtB
                        .nom = "faux B" & i
                        .Crd = Nothing
                        .Parent = nParent
                    End With
                    On Error GoTo 0
                Else
                    CoordPtB = CoordPt(GrilleActive.PartGrille, FxPtb)
                    With tcFxPtB
                        .nom = nfxPtb
                        .Crd = CoordPtB
                        .Parent = nParent
                    End With
                End If
                
                '# recercher le PtB dans la collection des points B a partir du nom du parent
                Set tcPtB = Nothing
                PtBOK = False
                For j = 1 To tcPtBs.Count
                    If tcPtBs.Item(j).Parent = nParent Then
                        Set tcPtB = tcPtBs.Item(j)
                        PtBOK = True
                        Exit For
                    End If
                Next j
                If Not PtBOK Then
                    With tcPtB
                        .nom = "Point_B_non_trouve"
                        .Crd = Nothing
                        .Parent = nParent
                    End With
                End If
            End If
            
            'Recherche du fastener associé
            fastOK = False
            For j = 1 To GrilleActive.Fasteners.Count
                'verifie si la distance du fastener / au faux PtA est égale à 0 à +/-0.02 mm près
                If tcFxPtA.Crd Is Nothing Then
                    Set tFast = Nothing
                Else
                    Set tFast = GrilleActive.Fasteners.Item(j)
                    tcrdFast.X = tFast.Xe
                    tcrdFast.Y = tFast.Ye
                    tcrdFast.Z = tFast.Ze
                    Distance = DistMat(tcFxPtA.Crd, tcrdFast)
                    If Abs(Distance) < 0.02 Then
                        fastOK = True
                        Exit For
                    End If
                End If
            Next j
            If Not fastOK Then 'le fastener n'est pas dans la collection
                Set tFast = Nothing
            End If
            tcPtStdFasts.Add tcPtA.nom, tcPtB.nom, nParent, tcFxPtA, tcFxPtB, tFast
            Set tcPtA = New c_Pt
            Set tcPtB = New c_Pt
            nParent = ""
            Set tcFxPtA = New c_Pt
            Set tcFxPtB = New c_Pt
            Set tFast = New c_Fastener
        Next i
     
        If OrdreFast Then 'tri par ordre des fasteners
            For i = 1 To GrilleActive.Fasteners.Count
                For j = 1 To tcPtStdFasts.Count
                    'Set tcPtStdFast = tcPtStdFasts.Item(j)
                    If tcPtStdFasts.Item(j).Fastener.nom = GrilleActive.Fasteners.Item(i).nom Then
                        tcPtStdFastTri.Add tcPtStdFasts.Item(j).nPtA, _
                            tcPtStdFasts.Item(j).nPtB, _
                            tcPtStdFasts.Item(j).nstd, _
                            tcPtStdFasts.Item(j).FxPtA, _
                            tcPtStdFasts.Item(j).FxPtb, _
                            tcPtStdFasts.Item(j).Fastener
                    End If
                Next j
            Next i
            Set tcPtStdFasts = tcPtStdFastTri 'Remplacement de la collection par la collection triée
        End If
       
       'Renommage
       For i = 1 To tcPtStdFasts.Count
            Select Case typeNum
                Case "nFast"
                    ComplementNom = i & "-" & tcPtStdFasts.Item(i).Fastener.nom
                Case "Comment"
                    ComplementNom = i & "-" & tcPtStdFasts.Item(i).Fastener.Comments
                Case "PtA"
                    ComplementNom = i
            End Select
               
            'renommage du PtA
            GrilleActive.Hb(nHBPtA).HybridShapes.Item(tcPtStdFasts.Item(i).nPtA).Name = "A" & ComplementNom
            'renommage du PtB
            GrilleActive.Hb(nHBPtB).HybridShapes.Item(tcPtStdFasts.Item(i).nPtB).Name = "B" & ComplementNom
            'renommage du Std
            GrilleActive.Hb(nHBStd).HybridShapes.Item(tcPtStdFasts.Item(i).nstd).Name = "Line." & ComplementNom
            'renommage du Faux PtA
            GrilleActive.Hb(nHBPtConst).HybridShapes.Item(tcPtStdFasts.Item(i).FxPtA.nom).Name = "Faux A" & ComplementNom
            'renommage du Faux PtB
            GrilleActive.Hb(nHBPtConst).HybridShapes.Item(tcPtStdFasts.Item(i).FxPtb.nom).Name = "Faux B" & ComplementNom
       Next i
        GoTo Fin
   End If
Erreur:
    If Err.Number > vbObjectError + 512 Then
        MsgBox Err.Description, vbCritical, "Element manquant"
    Else
        MsgBox Err.Description, vbCritical, "Erreur system"
    End If
Fin:
'Libération des classes
    Set GrilleActive = Nothing
    Set coordPtA = Nothing
    Set CoordPtB = Nothing
    Set tcPtA = Nothing
    Set tcPtAs = Nothing
    Set tcPtB = Nothing
    Set tcPtBs = Nothing
    Set tcFxPtA = Nothing
    Set tcFxPtAs = Nothing
    Set tcFxPtB = Nothing
    Set tcFxPtBs = Nothing
    Set tFast = Nothing
    Set tcrdFast = Nothing
    Set tcPtStdFast = Nothing
    Set tcPtStdFasts = Nothing
    Set tcPtStdFastTri = Nothing
    
End Sub

Private Function InfoPt(HsPt As HybridShape, HBody As HybridBody) As c_Pt
'collecte les infos d'un point
' nom, coordonnées, Parent
Dim nParent As String
Dim tPt As c_Pt
Set tPt = New c_Pt
        With tPt
            .nom = HsPt.Name
            '.Crd = tcoorPt 'CoordPt(GrilleActive.PartGrille, HsPt) la coordonée du HsPt n'est pas utile
        End With
        nParent = HsPt.Element1.DisplayName
        If HBody.HybridShapes.Item(nParent) Is Nothing Then 'La ligne STD n'a pas été trouvée dans le set STD
            tPt.Parent = "Ligne_STD_inconnue"
        Else
            tPt.Parent = HBody.HybridShapes.Item(nParent).Name
        End If
        Set InfoPt = tPt
        
'Libération des classes
Set tPt = Nothing

End Function

Private Sub RenStd(HbStd As HybridBody)
' Renomme les Droites STD
' HbStd = set géométrique des STD

Dim i As Long

    For i = 1 To HbStd.HybridShapes.Count
    '        On Error Resume Next
    '        NomfauxPtParent = GrilleActive.Hb(nHBStd).HybridShapes.Item(i).PtExtremity.DisplayName
    '        If Err.Number <> 0 Then
    '            Nom_Point = i
    '        Else
    '            Nom_Point = Right(NomfauxPtParent, Len(NomfauxPtParent) - 6 - Len(CStr(i)))
    '        End If
    '        Err.Clear
    '        On Error GoTo 0
    '        GrilleActive.Hb(nHBStd).HybridShapes.Item(i).Name = "Line." & i & Nom_Point    'Renomme la droite STD
    Next

End Sub

Private Sub RenomPtIsoles(GrilleActive, num)
'Renomme les points sans historique dans le set pointsA, pointsB, feet et pinules
Dim HBody As HybridBody
Dim HShapes As HybridShapes
Dim HShape As HybridShape
Dim i As Long
Dim nSet As String, nPt As String

nSet = Choose(num, nHBPtA, nHBPtB, nHBFeet, nHBPin)
nPt = Choose(num, "A", "B", "F", "P")
    'Vérification de l'existence des sets géométriques spécifiques
    On Error GoTo Erreur
    Set HBody = GrilleActive.Hb(nSet)
    On Error GoTo 0
    
    Set HShapes = HBody.HybridShapes
    For i = 1 To HShapes.Count
        Set HShape = HShapes.Item(i)
        HShape.Name = nPt & i
    Next i
    GoTo Fin
    
Erreur:
    MsgBox "le set géométrique " & nSet & " est manquant ou mal orthographiè!", vbCritical, "Element manquant"
Fin:
    Set HBody = Nothing
End Sub
