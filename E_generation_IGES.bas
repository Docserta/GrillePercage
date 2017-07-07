Attribute VB_Name = "E_generation_IGES"
Option Explicit

'*********************************************************************
'* Macro : E_generation_IGES (ex E_sets_visible)
'*
'* Fonctions : Masque tous les sets g�om�triques
'*             puis Passe une partie des sets g�om�triques en statut "Visible"
'*             puis exporte la part en IGES
'*
'* Version : 9.0.0
'* Cr�ation :  SVI
'* Modification : 1/08/14 CFR
'*                Mise � jour de la liste des sets g�om�triques � afficher
'*                Masquage des autres sets
'*                int�rogation user sur masquage ou affichage du set "gravures"
'* Modification : 26/02/16 CFR
'*                Correction d�tection doc actif = catpart
'*                Ajout sauvegarde en C:\temp si part non sauvegarde pr�alablement
'*
'**********************************************************************
Sub CATMain()

'Log de l'utilisation de la macro
LogUtilMacro nPath, nFicLog, nMacro, "E_generation_IGES", VMacro

Dim MP_temp As Variant
Dim MaSel_visProperties As VisPropertySet
Dim my_path As String
Dim PartGrille As New c_PartGrille

'V�rification qu'un part est actif
If Check_partActif() Then
    'Active le set "travail"
    PartGrille.PartGrille.InWorkObject = PartGrille.Hb(nHBTrav)
    
    'Creation de la collection des s�lections
    Dim MaSelection As Selection
    Set MaSelection = PartGrille.GrilleSelection
    
    'Masquage de tous les set g�om�triques
    For Each MP_temp In PartGrille.Hbodies
    'For Each MP_temp In MaPart.HybridBodies
        MaSelection.Add MP_temp
    Next
    Set MaSel_visProperties = PartGrille.GrilleSelection.VisProperties
    MaSel_visProperties.SetShow 1
    MaSelection.Clear
 
    'S�lection des sets g�om�triques pour export IGES
    If PartGrille.Exist_HB(nSurf0) Then MaSelection.Add PartGrille.Hb(nHBS0)
    If PartGrille.Exist_HB(nHBS100) Then MaSelection.Add PartGrille.Hb(nHBS100)
    If PartGrille.Exist_HB(nHBPtA) Then MaSelection.Add PartGrille.Hb(nHBPtA)
    If PartGrille.Exist_HB(nHBPtB) Then MaSelection.Add PartGrille.Hb(nHBPtB)
    If PartGrille.Exist_HB(nHBPin) Then MaSelection.Add PartGrille.Hb(nHBPin)
    If PartGrille.Exist_HB(nHBFeet) Then MaSelection.Add PartGrille.Hb(nHBFeet)
    If MsgBox("Souhaitez vous afficher le set 'gravures' ?", vbYesNo, "Export IGES") = vbYes Then
        If PartGrille.Exist_HB(nHBGrav) Then MaSelection.Add PartGrille.Hb(nHBGrav)
    End If

    'Affichage des sets
    Set MaSel_visProperties = PartGrille.GrilleSelection.VisProperties
    MaSel_visProperties.SetShow 0

    'Export IGES
    my_path = PartGrille.partDocGrille.Path
    'enregistre dans "C:\temp" si le part n'est pas encore sauvegard� et que la propri�t� PartGrille.partDocGrille.Path = ""
    If my_path = "" Then
        my_path = "C:\temp"
        MsgBox "l'IGES a �t� sauvegard� dans C:\temp car le r�pertoire de sauvegarde du part est inconnu"
    End If
    PartGrille.partDocGrille.ExportData my_path & "\" & "CMM_" & PartGrille.nom & "_indA", "igs"
Else
    MsgBox "Le document actif n'est pas un part!", vbCritical, "Erreur"
    Exit Sub
End If

Set PartGrille = Nothing
End Sub
