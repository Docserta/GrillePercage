VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FRM_DiamPercage 
   Caption         =   "Diametre des Perçages"
   ClientHeight    =   7845
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9900
   OleObjectBlob   =   "FRM_DiamPercage.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FRM_DiamPercage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Sub BtnAnnul_Click()
    Me.Hide
    Me.ChB_OkAnnule = False
End Sub

Private Sub BtnOK_Click()
Dim Erreur As Boolean
    Me.ChB_OkAnnule = True
    Erreur = False

    If Me.TBX_DiamPercage = "" Then
        Me.TBX_DiamPercage.BackColor = RGB(255, 212, 255)
        Erreur = True
    Else
        Me.TBX_DiamPercage.BackColor = RGB(212, 255, 255)
    End If
If Not Erreur Then
    Me.Hide
End If
End Sub

Private Sub BtnSel_Click()
'Selection des Points A à percer.
Dim GrilleTemp As New c_PartGrille
GrilleTemp.GrilleSelection.Clear

Dim tLisfast As c_Fasteners
'Set tLisfast = New c_Fasteners
Set tLisfast = GrilleTemp.Fasteners
Dim tFast As c_Fastener
Set tFast = New c_Fastener

Dim i As Long, j As Long
Dim Nb_Pt_Sel As Long
Dim NomUdfEC As String 'Nom de l'UDF en cours de traitement
Dim DiamUdfSel As String 'stocke le diamètre de perçage avion des UDF sélectionnées
ReDim Tab_Select_Points(2, 0)
Me.Hide
SelectPTA GrilleTemp

'Vérification que la sélection n'est pas vide
    Nb_Pt_Sel = GrilleTemp.GrilleSelection.Count

    If Nb_Pt_Sel = 0 Then
       MsgBox "Vous n'avez pas selectionné de points dans PointsA"
       Exit Sub
    End If
'Vérification que ce soient des Point A
'Ils doivent appartenir au Set "pointsA"
'#############
'
' a faire
'
'###############

'Ajout des Diamètre de perçage avion et du nom du STD dans le tableau pour chaque point selectionné
    For i = 1 To Nb_Pt_Sel
        If GrilleTemp.GrilleSelection.Item(i).Type = "HybridShape" Then
            ReDim Preserve Tab_Select_Points(2, i - 1)
            Tab_Select_Points(0, i - 1) = GrilleTemp.GrilleSelection.Item(i).Value.Name
            'Recupération du nom de la ligne STD
            Tab_Select_Points(2, i - 1) = GrilleTemp.GrilleSelection.Item(i).Value.Element1.DisplayName
            NomUdfEC = Right(Tab_Select_Points(0, i - 1), Len(Tab_Select_Points(0, i - 1)) - InStr(1, CStr(Tab_Select_Points(0, i - 1)), "-", vbTextCompare))
            'Récupération des diamètres de perçage Avion
            'Recherche du fastener dans la collection
            
            On Error Resume Next
            Set tFast = tLisfast.Item(NomUdfEC)
            If Err.Number <> 0 Then
                Err.Clear
                Tab_Select_Points(1, i - 1) = "Inconnu"
            Else
                Tab_Select_Points(1, i - 1) = tFast.FastDiam
            End If
            On Error GoTo 0
'            For j = 0 To UBound(GrilleTemp.Coll_RefExIsol(), 2)
'                If GrilleTemp.Coll_RefExIsol(0, j) = NomUdfEC Then
'                    Tab_Select_Points(1, i - 1) = GrilleTemp.Coll_RefExIsol(2, j)
'                    Exit For
'                End If
'            Next
            
        End If
    Next
        
'Ajout des nom de points sélectionnés et des diamètres de perçage avion dans le formulaire
    FRM_DiamPercage.LB_SelTrous.ColumnCount = 2
    FRM_DiamPercage.LB_SelTrous.List = TranspositionTabl(Tab_Select_Points)
        
'vérification que tous les UDF sont de même diamètres
    DiamUdfSel = Tab_Select_Points(1, 0)
    FRM_DiamPercage.TBX_DiamTrouAvUDF = DiamUdfSel
    For i = 0 To UBound(Tab_Select_Points, 2)
        If Tab_Select_Points(1, i) <> DiamUdfSel Then
            MsgBox "Tous les UDF sélectionnées ne sont pas de même diamètres !", vbCritical, "Sélection invalide"
            FRM_DiamPercage.LB_SelTrous.BackColor = RGB(255, 212, 255) 'Rouge
            Exit For
        End If
    Next
         
    Set GrilleTemp = Nothing
    Me.Show

End Sub

Private Sub CB_Lamage_Click()
    If Me.CB_Lamage Then
        Me.LB_DiamLamage.Enabled = True
        Me.TBX_DiamLamage.Enabled = True
    Else
        Me.LB_DiamLamage.Enabled = False
        Me.TBX_DiamLamage.Enabled = False
    End If
End Sub

Private Sub CBL_NumMachine_Change()
Dim NumMachine As String
    If Me.CBL_NumMachine <> "" Then
        NumMachine = Me.CBL_NumMachine
        'Récupère les info pour la machine choisie
        If Me.RB_GrilleCC Then
            Me.TBX_DiamTrouAvion = CollMachines.DiamPercageAvionCC(NumMachine)
            Me.TBX_DiamPercage = CollMachines.DiamPercageGrilleCC(NumMachine)
        ElseIf Me.RB_GrilleVT Then
            Me.TBX_DiamTrouAvion = CollMachines.DiamPercageAvionVT(NumMachine)
            Me.TBX_DiamPercage = CollMachines.DiamPercageGrilleVT(NumMachine)
            Me.TBX_DiamArret = CollMachines.DiamArretVT(NumMachine)
            Me.TBX_PosArret = CollMachines.PosArretVT(NumMachine)
            Me.TBX_ProfArret = CollMachines.ProfArretVT(NumMachine)
            Me.TBX_ProfTaraud = CollMachines.ProfTarauArretVT(NumMachine)
            Me.CBL_NBVis = CollMachines.NBVisArretoirVT(NumMachine)
            Me.TBX_DiamLamage = CollMachines.DiamLamageVT(NumMachine)
            Me.TBX_NumBague = CollMachines.RefBagueVT(NumMachine)
            Me.TBX_NumVis = CollMachines.RefVisArretoirVT(NumMachine)
        ElseIf Me.RB_GrillePM Then
            Me.TBX_DiamPercage = CollBagues.Item(NumMachine).D2
            Me.TBX_DiamLamage = CollBagues.Item(NumMachine).D3
            Me.TBX_NumBague = CollBagues.Item(NumMachine).NomFic
            
        End If
    End If
End Sub

Private Sub Logo_eXcent_Click()
'Chargement de la boite eXcent
    Load Frm_eXcent
    Frm_eXcent.Tbx_Version = VMacro
    Frm_eXcent.Show
    
    Unload Frm_eXcent
End Sub

Private Sub RB_GrilleCC_Click()
Dim NumMachineTemp()
Dim i As Long
    'Vidage de la liste déroulante
    Me.CBL_NumMachine.Clear
    ClearTBX
    If Me.RB_GrilleCC Then
        'Rempli la liste déroulantes des numéro machine
        NumMachineTemp = CollMachines.ListeMachinesCC
        For i = 0 To UBound(NumMachineTemp, 2)
            FRM_DiamPercage.CBL_NumMachine.AddItem (NumMachineTemp(0, i))
        Next
        AffChampsTrouBague "CC"

    End If
End Sub

Private Sub RB_GrillePM_Click()
Dim i As Long
Dim oBague As c_DefBague

    'Vidage de la liste déroulante
    Me.CBL_NumMachine.Clear
    ClearTBX
    If Me.RB_GrillePM Then
        'Rempli la liste déroulante des Numéros de bagues
        For Each oBague In CollBagues.Items
            FRM_DiamPercage.CBL_NumMachine.AddItem oBague.Ref
        Next
        AffChampsTrouBague "PM"
    
    End If
    
'Lbération des classes
Set oBague = Nothing
End Sub

Private Sub RB_GrilleVT_Click()
Dim i As Long
Dim NumMachineTemp()

    'Vidage de la liste déroulante
    Me.CBL_NumMachine.Clear
    ClearTBX
    If Me.RB_GrilleVT Then
        'Rempli la liste déroulantes des premiers terme des numéro machine
        NumMachineTemp = CollMachines.ListeMachinesVT
        For i = 0 To UBound(NumMachineTemp, 2)
            FRM_DiamPercage.CBL_NumMachine.AddItem (NumMachineTemp(0, i))
        Next
        AffChampsTrouBague "VT"

    End If
End Sub

Private Sub TBX_DiamTrouAvion_Change()
'Vérifie que le diamètre de perçage avion correspond aux diamètre enregistré dans les UDF
'sinon passe la case en rouge
    If Me.TBX_DiamTrouAvion <> Me.TBX_DiamTrouAvUDF Then
        Me.TBX_DiamTrouAvion.BackColor = RGB(255, 212, 255) 'Rouge
    Else
        Me.TBX_DiamTrouAvion.BackColor = RGB(212, 255, 255) 'Vert
    End If
End Sub

Private Sub UserForm_Initialize()
Me.RB_GrilleCC = True
Me.CBL_NBVis.AddItem "SIMPLE"
Me.CBL_NBVis.AddItem "DOUBLE"
Me.CBL_NBVis = "SIMPLE"
End Sub

Public Sub ClearTBX()
'vide les textebox
Me.TBX_DiamPercage = ""
Me.TBX_DiamTrouAvion = ""
Me.TBX_DiamArret = ""
Me.TBX_PosArret = ""
Me.TBX_ProfArret = ""
Me.TBX_ProfTaraud = ""
Me.TBX_DiamLamage = ""
Me.TBX_DiamLamage = ""
Me.CBL_NBVis = "SIMPLE"
Me.TBX_NumBague = ""
Me.TBX_NumVis = ""
End Sub
Private Sub AffChampsTrouBague(Typgrille As String)
'Masque/affiche ou renomme les controles en fonction dtype de grille
'Typgrille = "CC" ou "VT" ou "PM"
    Select Case Typgrille
        Case "CC"
            'redimentione le cadre
            Me.Fr_Machine.Height = 102
            Me.Fr_Machine.Width = 276
            'Renomme les labels
            Me.Lbl_NoMachine = "Numéro de Machine"
            Me.LB_NumBague = "Numéro de Vis Arretoir"
            'Affiche les controles
            Me.LB_NumBague.Visible = False
            Me.TBX_NumBague.Visible = False
            Me.LB_DiamTrouAvion.Visible = True
            Me.TBX_DiamTrouAvion.Visible = True
            'Masquage des champs Vis Arretoire
            AffChampsVisArretoir False
            'Choix Lamage
            AffChampsLamage False
            
        Case "VT"
            'redimentione le cadre
            Me.Fr_Machine.Height = 204
            Me.Fr_Machine.Width = 390
            'Renomme les labels
            Me.Lbl_NoMachine = "Numéro de Machine"
            Me.LB_NumBague = "Numéro de Vis Arretoir"
            'Affiche les controles
            Me.LB_NumBague.Visible = True
            Me.TBX_NumBague.Visible = True
            Me.LB_DiamTrouAvion.Visible = True
            Me.TBX_DiamTrouAvion.Visible = True
            'Affichage des champs Vis Arretoire
            AffChampsVisArretoir True
            'Choix Lamage
            AffChampsLamage True
            
        Case "PM"
            Me.Fr_Machine.Height = 204
            Me.Fr_Machine.Width = 390
            'Renomme les labels
            Me.Lbl_NoMachine = "Numéro de bague"
            Me.LB_NumBague = "N° JT de la bague"
            'Affiche les controles
            Me.LB_NumBague.Visible = True
            Me.TBX_NumBague.Visible = True
            Me.LB_DiamTrouAvion.Visible = False
            Me.TBX_DiamTrouAvion.Visible = False
            'Affichage des champs
            AffChampsVisArretoir False
            'Choix Lamage
            AffChampsLamage True
        
    End Select
End Sub

Private Sub AffChampsVisArretoir(Visible As Boolean)
'Masque ou affiche les controles Vis Arretoire
    Me.TBX_DiamArret.Visible = Visible
    Me.LB_DiamArret.Visible = Visible
    Me.TBX_PosArret.Visible = Visible
    Me.LB_PosArret.Visible = Visible
    Me.TBX_ProfArret.Visible = Visible
    Me.LB_ProfArret.Visible = Visible
    Me.LB_ProfTaraud.Visible = Visible
    Me.TBX_ProfTaraud.Visible = Visible
    Me.LB_NBVis.Visible = Visible
    Me.CBL_NBVis.Visible = Visible
    Me.LB_NumVis.Visible = Visible
    Me.TBX_NumVis.Visible = Visible
End Sub

Private Sub AffChampsLamage(Visible As Boolean)
'Affiche ou masque les controles Lamage
    Me.CB_Lamage.Visible = Visible
    Me.LB_DiamLamage.Visible = Visible
    Me.TBX_DiamLamage.Visible = Visible
    'Les controles sont désactivés par défaut
    Me.CB_Lamage.Value = False
    Me.LB_DiamLamage.Enabled = False
    Me.TBX_DiamLamage.Enabled = False
End Sub




