Attribute VB_Name = "G_Recalcule_PtAB"
Option Explicit
'*********************************************************************
'* Macro : G_Recalcule_PtAB
'*
'* Fonctions :  Modification des coordonnées des  des faux-Pts A et faux-Pts B
'*              remplacement avec les coordonnée des ref externes isolées
'*              permet de corriger une grille dont les ref on évoluées
'*              ou d'adapter une grille existante sur une autre zone avion.
'*
'* Version 6
'* Création : 16/02/15 CFR
'* Modification :
'*
'**********************************************************************
'
Sub CATMain()
 
'Log de l'utilisation de la macro
LogUtilMacro nPath, nFicLog, nMacro, "G_Recalcule_PtAB", VMacro
 
Dim Report() As String

Dim LengthX As Length, LengthY As Length, LengtZ As Length
Dim CheminParam As String
Dim HS_RefExtIsoles As HybridShapes
Dim collection_param_udf As Parameters
Dim index_param_udf As Long
Dim nom_param As String
Dim Tmp_Shape As HybridShape
'Dim Tabl_RefExt() As String
Dim HS_FauxPts As HybridShapes
Dim NoPtTraite As Long
Dim NomPtTraite As String
Dim cpt_hs As Long, cpt_RefExt As Long, cpt_HSMax As Long, i As Long
    cpt_hs = 1
    cpt_RefExt = 0
Dim Mon_HShapePointCoord As HybridShapePointCoord
Dim Ligne_Repport As String
Dim mBar As c_ProgressBar
Dim instance_catpart_grille_nue As PartDocument
Dim tLisfast As c_Fasteners
'Set tLisfast = New c_Fasteners
Dim GrilleActive As c_PartGrille
Dim tFast As c_Fastener
Dim TestHBody As HybridBody

Set tFast = New c_Fastener

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
    
    Set GrilleActive = New c_PartGrille
    
    'Vérification de l'existence des sets géométriques
    On Error GoTo Erreur
    Set TestHBody = GrilleActive.Hb(nHBRefExtIsol)
    Set TestHBody = GrilleActive.Hb(nHBPtConst)
    Set TestHBody = Nothing
    On Error GoTo 0

'------------------------
' Collecte des fasteners
'------------------------
    Set HS_RefExtIsoles = GrilleActive.Hb(nHBRefExtIsol).HybridShapes
    Set tLisfast = GrilleActive.Fasteners
    'Tabl_RefExt() = GrilleActive.Coll_RefExIsol()
    
'----------------------------------
' Comptage des Faux PtA et Faux PtB
'----------------------------------
    Set HS_FauxPts = GrilleActive.Hb(nHBPtConst).HybridShapes
    For i = 1 To HS_FauxPts.Count
        If Left(HS_FauxPts.Item(i).Name, 4) = "faux" Then
            cpt_HSMax = cpt_HSMax + 1
        End If
    Next
    GrilleActive.GrilleSelection.Clear

'Progress Barre
    Set mBar = New c_ProgressBar
    mBar.ProgressTitre 1, " Recalcul des coordonnées X,Y et Z des faux Pts A et faux Pts B, veuillez patienter."
    
'-----------------------------------------------------
' Modification de coordonnées des Faux PtA et Faux PtB
'-----------------------------------------------------

    While (cpt_hs <= cpt_HSMax)
        'Maj barre de progression
        mBar.Progression = (100 / cpt_HSMax) * cpt_hs
        
        
 ' 1ere, le nom réel
' 2eme, le paramétre comment
' 3eme, le Diamètre de perçage
' 4eme, Xe
' 5eme, Ye
' 6eme, Ze
' 7eme, Xdir
' 8eme, Ydir
' 9eme, Zdir
        Set tFast = tLisfast.Item(cpt_hs)
        
        Set Mon_HShapePointCoord = HS_FauxPts.Item(cpt_hs)
        
        If (cpt_hs <= tLisfast.Count * 2) Then
            If Len(HS_FauxPts.Item(cpt_hs).Name) > 7 + Len(CStr(cpt_hs)) Then
                NomPtTraite = Mid(HS_FauxPts.Item(cpt_hs).Name, 6, 1)
                NoPtTraite = CInt(Mid(HS_FauxPts.Item(cpt_hs).Name, 7, Len(CStr(cpt_hs))))
            Else
                NomPtTraite = ""
                NoPtTraite = 0
            End If
            'Set Mon_HShapePointCoord = HS_FauxPts.Item(cpt_hs)
            If NomPtTraite = "A" Then
                
                Mon_HShapePointCoord.X.Value = CDbl(tFast.Xe)
                Mon_HShapePointCoord.Y.Value = CDbl(tFast.Ye)
                Mon_HShapePointCoord.Z.Value = CDbl(tFast.Ze)
                Mon_HShapePointCoord.Name = "faux A" & cpt_RefExt + 1 & "-" & tFast.nom
                   
            ElseIf NomPtTraite = "B" Then
                Mon_HShapePointCoord.X.Value = CDbl(tFast.Xe) + 100 * CDbl(tFast.Xdir)
                Mon_HShapePointCoord.Y.Value = CDbl(tFast.Ye) + 100 * CDbl(tFast.Ydir)
                Mon_HShapePointCoord.Z.Value = CDbl(tFast.Ze) + 100 * CDbl(tFast.Zdir)
                Mon_HShapePointCoord.Name = "faux B" & cpt_RefExt + 1 & "-" & tFast.nom

                cpt_RefExt = cpt_RefExt + 1 'On passe à la ligne suivante dans le tableau
            End If
            GrilleActive.PartGrille.UpdateObject Mon_HShapePointCoord
        Else
            GrilleActive.GrilleSelection.Add Mon_HShapePointCoord
        End If
        cpt_hs = cpt_hs + 1
    Wend
     
' On supprime les points en trop
    If GrilleActive.GrilleSelection.Count > 0 Then GrilleActive.GrilleSelection.Delete
    
    On Error Resume Next
    GrilleActive.PartGrille.Update
    
' Renommage des points et lignes
    C1_Rename_AB.CATMain
    
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
    Set GrilleActive = Nothing
    Set mBar = Nothing
    
End Sub



