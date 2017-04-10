Attribute VB_Name = "F_sets_hide"
Option Explicit

'*********************************************************************
'* Macro : F_sets_hide
'*
'* Fonctions : Masque tous les sets géométriques
'*             puis Passe une partie des sets géométriques en statut "Visible"
'*             et active le partBody
'*
'* Version :
'* Création :  SVI
'* Modification : 1/08/14 CFR
'*                Traitement par lot
'*                Mise à jour de la liste des sets géométriques à afficher
'*                Masquage des autres sets
'*                Activation du part Body
'* Modification : 26/02/16
'*                dépersonalisation(cation, ..) du formulaire Frm_ListeFichiers
'*                pour l'utiliser dans un autre module ( A2_CreationLot)
'*
'**********************************************************************

Sub catmain()

'Log de l'utilisation de la macro
LogUtilMacro nPath, nFicLog, nMacro, "F_sets_hide", VMacro

Dim i As Long, j As Long
Dim NomFicListEC As String
Dim LogFile As String
Dim ReportLog() As String
Dim cfdate As String
Dim PartGrilleEC As c_PartGrille
Dim IndSelection As Long 'indice qui permet de parcourir la listbox de Frm_Donnees

'Ouvre la boite de dlg "Frm_ListeFichiers"
    Load Frm_ListeFichiers
    Frm_ListeFichiers.Caption = "Masquage des Sets Géométriques"
    Frm_ListeFichiers.Tbx_Extent = "*.CATPart"
    Frm_ListeFichiers.TBX_FicDest.Visible = False
    Frm_ListeFichiers.Lbl_FicDest.Visible = False
    Frm_ListeFichiers.Btn_Nav_Dest.Visible = False
    
    Frm_ListeFichiers.TBX_EnvAvion.Visible = False
    Frm_ListeFichiers.Lbl_Env.Visible = False
    Frm_ListeFichiers.Btn_Nav_Env.Visible = False
    
    
    'Test si un Catpart est Actif
    If Check_partActif Then
        Frm_ListeFichiers.CB_Catpartactif.enabled = True
    Else
        Frm_ListeFichiers.CB_Catpartactif.enabled = False
    End If
    
    Frm_ListeFichiers.Show
    
 'Sort du programme si click sur bouton Annuler le formulaire
    If Not (Frm_ListeFichiers.ChB_OkAnnule) Then
        Unload Frm_ListeFichiers
        Exit Sub
    End If
   
    Set PartGrilleEC = New c_PartGrille
    
    If Frm_ListeFichiers.CB_Catpartactif Then 'Traitement du catpart actif seul
        ReDim Preserve ReportLog(0)
        ReportLog(0) = PartGrilleEC.nom & Chr(13)
        ReportLog(0) = ReportLog(0) & MasqueSetPart(PartGrilleEC)
        MsgBox ReportLog(0), vbInformation, "Traitement effectué"
        PartGrilleEC.partDocGrille.Save
    Else 'Traitement de la liste des Catpart
        IndSelection = 0
        'Lancement du traitement des fichiers
        For i = 0 To Frm_ListeFichiers.ListBox1.ListCount - 1
            'Boucle sur la liste des fichiers et test si le fichier est sélectionné
            If Frm_ListeFichiers.ListBox1.Selected(i) Then
                'Création de la ligne du report
                ReDim Preserve ReportLog(1, IndSelection)
                NomFicListEC = Frm_ListeFichiers.ListBox1.List(i)
                ReportLog(0, IndSelection) = NomFicListEC & Chr(13)
                'Traitement du catpart
                PartGrilleEC.PG_partDocGrille = CATIA.Documents.Open(CheminFicLot & NomFicListEC)
                ReportLog(1, IndSelection) = MasqueSetPart(PartGrilleEC)
                IndSelection = IndSelection + 1
                'Sauvegarde
                PartGrilleEC.partDocGrille.Save
                PartGrilleEC.partDocGrille.Close
            End If
        Next i
    
        If IndSelection = 0 Then
            MsgBox "Pas de fichier sélectionné!", vbInformation, "Pas de fichier sélectionné"
            Exit Sub
        Else
            MsgBox "Fin de traitement des fichiers.", vbInformation, "Fin de traitement"
            'affichage du log
            WriteLog ReportLog, CheminFicLot, "F_set_hide"
        End If
    End If

Unload Frm_ListeFichiers
Set PartGrilleEC = Nothing

End Sub

Public Function MasqueSetPart(MSP_PartActif) As String
'Traitement d'un part et renvoi du log

'Initialisation des variables
Dim MSP_Log As String
Dim MP_temp As Variant
Dim MaSel_visProperties As VisPropertySet

        'Active le PartBody
        MSP_PartActif.PartGrille.InWorkObject = MSP_PartActif.mBody
        MSP_Log = MSP_Log & " - " & "Corps principal Activé" & Chr(13)

        'Masquage de tous les set géométriques
        For Each MP_temp In MSP_PartActif.Hbodies
            MSP_PartActif.GrilleSelection.Add MP_temp
        Next
        Set MaSel_visProperties = MSP_PartActif.GrilleSelection.VisProperties
        MaSel_visProperties.SetShow 1
        MSP_Log = MSP_Log & " - " & "tous les sets géométriques masqués" & Chr(13)
        MSP_PartActif.GrilleSelection.Clear

        'Sélection des sets géométriques a afficher
        If MSP_PartActif.Exist_HB(nHBFeet) Then
            MSP_PartActif.GrilleSelection.Add MSP_PartActif.Hb(nHBFeet)
            MSP_Log = MSP_Log & " - " & "Set géométrique 'feet' affiché" & Chr(13)
        Else
            MSP_Log = MSP_Log & " - " & "Pas de Set géométrique 'feet' trouvé" & Chr(13)
        End If
        If MSP_PartActif.Exist_HB(nHBGrav) Then
            MSP_PartActif.GrilleSelection.Add MSP_PartActif.Hb(nHBGrav)
            MSP_Log = MSP_Log & " - " & "Set géométrique 'gravures' affiché" & Chr(13)
        Else
            MSP_Log = MSP_Log & " - " & "Pas de Set géométrique 'gravure' trouvé" & Chr(13)
        End If
        
        'Affichage des sets
        Set MaSel_visProperties = MSP_PartActif.GrilleSelection.VisProperties
        MaSel_visProperties.SetShow 0
        
MasqueSetPart = MSP_Log
End Function

