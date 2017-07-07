Attribute VB_Name = "E_CréationU01"
Option Explicit
'*********************************************************************
'* Macro : A_CreationPartU01
'*
'* Fonctions :  Création de la part de controle U01
'*              copie / colle sans liens:
'*                les surfaces 0 et à 100
'*                 Les pts A et B
'*                 les pinnules et les pieds
'*
'* Version : 9
'* Création :  CFR
'*
'* Modification :
'*
'*
'**********************************************************************

Sub CATMain()

'Log de l'utilisation de la macro
LogUtilMacro nPath, nFicLog, nMacro, "E_CréationU01", VMacro

Dim i As Integer
Dim PartGriDoc As PartDocument
Dim PartU1Doc As PartDocument
Dim PartU1 As c_PartGrille
Dim GrilleActive As c_PartGrille
Dim Sel_Source As Selection, Sel_Cible As Selection
Dim TestHBody As HybridBody
Dim TestHShape As HybridShape

 'Sélection des parts grille et U01
    Load FRM_PartU1
    FRM_PartU1.Show

'Clic sur Bouton "Annule" dans formulaire
    If Not (FRM_PartU1.ChB_OkAnnule) Then
        End
    End If

    Set coll_docs = CATIA.Documents

'Ouvre le fichier de la part grille nue
    If FileExist(FRM_PartU1.TBX_NomGrille) Then
        Set PartGriDoc = coll_docs.Open(FRM_PartU1.TBX_NomGrille)
    Else
        MsgBox "Le fichier de la part grille nue est introuvable!", vbCritical, "Fichier introuvable"
        End
    End If
'Ouvre le fichier de la part U01
    If FileExist(FRM_PartU1.TBX_NomU01) Then
        Set PartU1Doc = coll_docs.Open(FRM_PartU1.TBX_NomU01)
    Else
        MsgBox "Le fichier de la part U01 est introuvable!", vbCritical, "Fichier introuvable"
        End
    End If
    
    Set PartU1 = New c_PartGrille
    PartU1.PG_partDocGrille = PartU1Doc

    Set GrilleActive = New c_PartGrille
    GrilleActive.PG_partDocGrille = PartGriDoc
    
    'Vérification de l'existence des sets géométriques dans la parts grille nue
    On Error GoTo Erreur
    Set TestHBody = GrilleActive.Hb(nHBStd)
    Set TestHBody = GrilleActive.Hb(nHBPtA)
    Set TestHBody = GrilleActive.Hb(nHBPtB)
    Set TestHBody = GrilleActive.Hb(nHBPin)
    Set TestHBody = GrilleActive.Hb(nHBFeet)
    'Vérification de l'existence des surf0 et surf100 dans la parts grille nue
        Set TestHShape = GrilleActive.HS(nSurf0, nHBS0)
    Set TestHShape = GrilleActive.HS(nSurf100, nHBS100)
    Set TestHBody = Nothing
    Set TestHShape = Nothing
    On Error GoTo 0

    'Vérification de l'existence des sets géométriques dans la parts de controle
    'S'ls n'existent pas on les crés
    If Not (PartU1.Exist_HB(nSurf0)) Then
        Ajout1Set PartU1Doc, nHBS0
    End If
    If Not (PartU1.Exist_HB(nHBS100)) Then
        Ajout1Set PartU1Doc, nHBS100
    End If
    If Not (PartU1.Exist_HB(nHBPtA)) Then
        Ajout1Set PartU1Doc, nHBPtA
    End If
    If Not (PartU1.Exist_HB(nHBPtB)) Then
        Ajout1Set PartU1Doc, nHBPtB
    End If
    If Not (PartU1.Exist_HB(nHBStd)) Then
        Ajout1Set PartU1Doc, nHBStd
    End If
    If Not (PartU1.Exist_HB(nHBPin)) Then
        Ajout1Set PartU1Doc, nHBPin
    End If
    If Not (PartU1.Exist_HB(nHBFeet)) Then
        Ajout1Set PartU1Doc, nHBFeet
    End If
    
    'Copie des éléments géométriques dans la part de controle
    Set Sel_Source = GrilleActive.GrilleSelection
    Set Sel_Cible = PartU1.GrilleSelection

'Copie de la surf0
    Sel_Source.Clear
    Sel_Cible.Clear
    Sel_Source.Add GrilleActive.HS(nSurf0, nHBS0)
    Sel_Source.Copy
    Sel_Cible.Add PartU1.Hb(nHBS0)
    Sel_Cible.PasteSpecial "CATPrtResult"
    
'Copie de la surf100
    Sel_Source.Clear
    Sel_Cible.Clear
    Sel_Source.Add GrilleActive.HS(nSurf100, nHBS100)
    Sel_Source.Copy
    Sel_Cible.Add PartU1.Hb(nHBS100)
    Sel_Cible.PasteSpecial "CATPrtResult"
    
'Copie des PtA
    Sel_Source.Clear
    Sel_Cible.Clear
    If GrilleActive.Hb(nHBPtA).HybridShapes.Count > 0 Then
        For i = 1 To GrilleActive.Hb(nHBPtA).HybridShapes.Count
            Sel_Source.Add GrilleActive.Hb(nHBPtA).HybridShapes.Item(i)
        Next i
        Sel_Source.Copy
        
        Sel_Cible.Add PartU1.Hb(nHBPtA)
        Sel_Cible.PasteSpecial "CATPrtResult"
    End If
    
'Copie des PtB
    Sel_Source.Clear
    Sel_Cible.Clear
    
    If GrilleActive.Hb(nHBPtB).HybridShapes.Count > 0 Then
        For i = 1 To GrilleActive.Hb(nHBPtB).HybridShapes.Count
            Sel_Source.Add GrilleActive.Hb(nHBPtB).HybridShapes.Item(i)
        Next i
        Sel_Source.Copy
        
        Sel_Cible.Add PartU1.Hb(nHBPtB)
        Sel_Cible.PasteSpecial "CATPrtResult"
    End If
'Copie des Std
    Sel_Source.Clear
    Sel_Cible.Clear
    If GrilleActive.Hb(nHBStd).HybridShapes.Count > 0 Then
        For i = 1 To GrilleActive.Hb(nHBStd).HybridShapes.Count
            Sel_Source.Add GrilleActive.Hb(nHBStd).HybridShapes.Item(i)
        Next i
        Sel_Source.Copy
        
        Sel_Cible.Add PartU1.Hb(nHBStd)
        Sel_Cible.PasteSpecial "CATPrtResult"
    End If
    
'Copie des Pinnules
    Sel_Source.Clear
    Sel_Cible.Clear
    If GrilleActive.Hb(nHBPin).HybridShapes.Count > 0 Then
        For i = 1 To GrilleActive.Hb(nHBPin).HybridShapes.Count
            Sel_Source.Add GrilleActive.Hb(nHBPin).HybridShapes.Item(i)
        Next i
        Sel_Source.Copy
    
        Sel_Cible.Add PartU1.Hb(nHBPin)
        Sel_Cible.PasteSpecial "CATPrtResult"
    End If
    
'Copie des Pieds
    Sel_Source.Clear
    Sel_Cible.Clear
    If GrilleActive.Hb(nHBFeet).HybridShapes.Count > 0 Then
        For i = 1 To GrilleActive.Hb(nHBFeet).HybridShapes.Count
            Sel_Source.Add GrilleActive.Hb(nHBFeet).HybridShapes.Item(i)
        Next i
        Sel_Source.Copy
        
        Sel_Cible.Add PartU1.Hb(nHBFeet)
        Sel_Cible.PasteSpecial "CATPrtResult"
    End If

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
    Set PartU1 = Nothing
    Unload FRM_PartU1
    
End Sub


