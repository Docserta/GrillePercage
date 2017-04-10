Attribute VB_Name = "B_CreationOffsets"
Option Explicit
'*********************************************************************
'* Macro : B_CreationOffsets
'*
'* Fonctions : Création des surfaces à 30 et 100
'*
'* Version :
'* Création :  SVI
'* Modification : 14/02/2015
'*                Prise en compte de la classe "PartGrille
'* Modification : 15/04/15 CFR
'*                Ajout check Environnement
'*
'**********************************************************************

Sub catmain()

'Log de l'utilisation de la macro
LogUtilMacro nPath, nFicLog, nMacro, "B_CreationOffsets", VMacro

Dim HS_Factory As HybridShapeFactory
Dim Ref_Surf0 As Reference
Dim HS_OffsetSurf100 As HybridShapeOffset, HS_OffsetSurf30 As HybridShapeOffset, HS_OffsetSurf10 As HybridShapeOffset
Dim ValOffsetSurf100 As Integer, ValEPGrille As Integer, ValEPPieds As Integer, ValEPGrillePieds As Integer
Dim instance_catpart_grille_nue As PartDocument
Dim GrilleActive As New c_PartGrille
Dim specsAndGeomWindow1 As SpecsAndGeomWindow
Dim viewer3D1 As Viewer3D
Dim TestHShape As HybridShape

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
    
'Detecte si la 'surf0' existe
    On Error GoTo Fin
    Set TestHShape = GrilleActive.HS(nSurf0, nHBS0)
    On Error GoTo 0

'Detecte si les set géométriques "surf100" et "surf100" existe sinon les crées
    If Not (GrilleActive.Exist_HB(nHBS100)) Then
        GrilleActive.Create_HyBridShape (nHBS100)
    End If
    If Not (GrilleActive.Exist_HB(nHBTrav)) Then
        GrilleActive.Create_HyBridShape (nHBTrav)
    End If
    
    Set Ref_Surf0 = GrilleActive.Ref_HS(nSurf0, nHBS0)
    Set HS_Factory = GrilleActive.HShapeFactory

'Initialisation des epaisseurs de grille et de pieds
    ValOffsetSurf100 = 100
    
    ValEPGrille = InputBox("Select epaisseur : ")
    ValEPPieds = InputBox("Select hauteur pieds : ")

'Création de la surface à 100
    Set HS_OffsetSurf100 = HS_Factory.AddNewOffset(Ref_Surf0, ValOffsetSurf100, False, 0.01)
    HS_OffsetSurf100.OffsetDirection = True
    HS_OffsetSurf100.Name = nSurf100

    GrilleActive.Hb(nHBS100).AppendHybridShape HS_OffsetSurf100
    'GrilleActive.InWorkObject = HS_OffsetSurf100
    GrilleActive.PartGrille.Update

'Recadre la vue
    Set specsAndGeomWindow1 = CATIA.ActiveWindow
    Set viewer3D1 = specsAndGeomWindow1.ActiveViewer
    viewer3D1.Reframe
    
'Intérogation de l'utilisateur pour valider le sens de l'offset
'Si clic sur "Non" alors inversion de sens
    Do While MsgBox("Le sens de l'offset de la surface à 100 est il bon ? ", vbYesNo, "Offset Direction") = vbNo
        If HS_OffsetSurf100.OffsetDirection Then
            HS_OffsetSurf100.OffsetDirection = False
        Else
            HS_OffsetSurf100.OffsetDirection = True
        End If
        GrilleActive.PartGrille.Update
    Loop

'Création de la surface "Epaisseur pieds" dite surface à 30
    ValEPGrillePieds = ValEPGrille + ValEPPieds
    Set HS_OffsetSurf30 = HS_Factory.AddNewOffset(Ref_Surf0, ValEPGrillePieds, False, 0.01)
    'Inversion du sens de l'offset
    HS_OffsetSurf30.OffsetDirection = Not (HS_OffsetSurf100.OffsetDirection)

    GrilleActive.Hb(nHBTrav).AppendHybridShape HS_OffsetSurf30
    HS_OffsetSurf30.Name = nSurfSup

'Création de la surface a 10
    Set HS_OffsetSurf10 = HS_Factory.AddNewOffset(Ref_Surf0, ValEPPieds, False, 0.01)
    HS_OffsetSurf10.OffsetDirection = Not (HS_OffsetSurf100.OffsetDirection)
    GrilleActive.Hb(nHBTrav).AppendHybridShape HS_OffsetSurf10
    HS_OffsetSurf10.Name = nSurfInf
    GoTo Fin
Erreur:
    If Err.Number > vbObjectError + 512 Then
        MsgBox Err.Description, vbCritical, "Element manquant"
    Else
        MsgBox Err.Description, vbCritical, "Erreur system"
    End If
    End
Fin:
GrilleActive.PartGrille.Update
Set GrilleActive = Nothing
End Sub


