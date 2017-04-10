VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FRM_DiamPercage 
   Caption         =   "Diametre des Per�ages"
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
'Selection des Points A � percer.
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
Dim DiamUdfSel As String 'stocke le diam�tre de per�age avion des UDF s�lectionn�es
ReDim Tab_Select_Points(2, 0)
Me.Hide
SelectPTA GrilleTemp

'V�rification que la s�lection n'est pas vide
    Nb_Pt_Sel = GrilleTemp.GrilleSelection.Count

    If Nb_Pt_Sel = 0 Then
       MsgBox "Vous n'avez pas selectionn� de points dans PointsA"
       Exit Sub
    End If
'V�rification que ce soient des Point A
'Ils doivent appartenir au Set "pointsA"
'#############
'
' a faire
'
'###############

'Ajout des Diam�tre de per�age avion et du nom du STD dans le tableau pour chaque point selectionn�
    For i = 1 To Nb_Pt_Sel
        If GrilleTemp.GrilleSelection.Item(i).Type = "HybridShape" Then
            ReDim Preserve Tab_Select_Points(2, i - 1)
            Tab_Select_Points(0, i - 1) = GrilleTemp.GrilleSelection.Item(i).Value.Name
            'Recup�ration du nom de la ligne STD
            Tab_Select_Points(2, i - 1) = GrilleTemp.GrilleSelection.Item(i).Value.Element1.DisplayName
            NomUdfEC = Right(Tab_Select_Points(0, i - 1), Len(Tab_Select_Points(0, i - 1)) - InStr(1, CStr(Tab_Select_Points(0, i - 1)), "-", vbTextCompare))
            'R�cup�ration des diam�tres de per�age Avion
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
        
'Ajout des nom de points s�lectionn�s et des diam�tres de per�age avion dans le formulaire
    FRM_DiamPercage.LB_SelTrous.ColumnCount = 2
    FRM_DiamPercage.LB_SelTrous.List = TranspositionTabl(Tab_Select_Points)
        
'v�rification que tous les UDF sont de m�me diam�tres
    DiamUdfSel = Tab_Select_Points(1, 0)
    FRM_DiamPercage.TBX_DiamTrouAvUDF = DiamUdfSel
    For i = 0 To UBound(Tab_Select_Points, 2)
        If Tab_Select_Points(1, i) <> DiamUdfSel Then
            MsgBox "Tous les UDF s�lectionn�es ne sont pas de m�me diam�tres !", vbCritical, "S�lection invalide"
            FRM_DiamPercage.LB_SelTrous.BackColor = RGB(255, 212, 255) 'Rouge
            Exit For
        End If
    Next
         
    Set GrilleTemp = Nothing
    Me.Show

End Sub

Private Sub CB_Lamage_Click()
    If Me.CB_Lamage Then
        Me.LB_DiamLamage.enabled = True
        Me.TBX_DiamLamage.enabled = True
    Else
        Me.LB_DiamLamage.enabled = False
        Me.TBX_DiamLamage.enabled = False
    End If
End Sub

Private Sub CBL_NumMachine_Change()
Dim NumMachine As String
    If Me.CBL_NumMachine <> "" Then
        NumMachine = Me.CBL_NumMachine
        'R�cup�re les info pour la machine choisie
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
Me.CBL_NumMachine.Clear
ClearTBX
If Me.RB_GrilleCC Then
 'Rempli la liste d�roulantes des num�ro machine
        NumMachineTemp = CollMachines.ListeMachinesCC
    For i = 0 To UBound(NumMachineTemp, 2)
        FRM_DiamPercage.CBL_NumMachine.AddItem (NumMachineTemp(0, i))
    Next
    'redimentione le cadre
Me.Fr_Machine.Height = 102
Me.Fr_Machine.Width = 276
    'Masquage des champs Vis Arretoire
    Me.TBX_DiamArret.Visible = False
    Me.LB_DiamArret.Visible = False
    Me.TBX_PosArret.Visible = False
    Me.LB_PosArret.Visible = False
    Me.TBX_ProfArret.Visible = False
    Me.LB_ProfArret.Visible = False
    Me.LB_ProfTaraud.Visible = False
    Me.TBX_ProfTaraud.Visible = False
    Me.LB_NBVis.Visible = False
    Me.CBL_NBVis.Visible = False
    'Choix Lamage
    Me.CB_Lamage.Visible = False
    Me.LB_DiamLamage.Visible = False
    Me.TBX_DiamLamage.Visible = False
    'Ref Vis et bague
    Me.LB_NumBague.Visible = False
    Me.TBX_NumBague.Visible = False
    Me.LB_NumVis.Visible = False
    Me.TBX_NumVis.Visible = False
End If
End Sub

Private Sub RB_GrilleVT_Click()
Dim i As Long
Me.CBL_NumMachine.Clear
ClearTBX
If Me.RB_GrilleVT Then
 'Rempli la liste d�roulantes des premiers terme des num�ro machine
    Dim NumMachineTemp()
        NumMachineTemp = CollMachines.ListeMachinesVT
    For i = 0 To UBound(NumMachineTemp, 2)
        FRM_DiamPercage.CBL_NumMachine.AddItem (NumMachineTemp(0, i))
    Next
'redimentione le cadre
Me.Fr_Machine.Height = 204
Me.Fr_Machine.Width = 390
'Affichage des champs Vis Arretoire
    Me.TBX_DiamArret.Visible = True
    Me.LB_DiamArret.Visible = True
    Me.TBX_PosArret.Visible = True
    Me.LB_PosArret.Visible = True
    Me.TBX_ProfArret.Visible = True
    Me.LB_ProfArret.Visible = True
    Me.LB_ProfTaraud.Visible = True
    Me.TBX_ProfTaraud.Visible = True
    Me.CB_Lamage.Visible = True
    Me.LB_NBVis.Visible = True
    Me.CBL_NBVis.Visible = True
    'Choix Lamage
    Me.CB_Lamage.Value = False
    Me.LB_DiamLamage.enabled = False
    Me.LB_DiamLamage.Visible = True
    Me.TBX_DiamLamage.enabled = False
    Me.TBX_DiamLamage.Visible = True
    'Ref Vis et bague
    Me.LB_NumBague.Visible = True
    Me.TBX_NumBague.Visible = True
    Me.LB_NumVis.Visible = True
    Me.TBX_NumVis.Visible = True
End If
End Sub

Private Sub TBX_DiamTrouAvion_Change()
'V�rifie que le diam�tre de per�age avion correspond aux diam�tre enregistr� dans les UDF
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
