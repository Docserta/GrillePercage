VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FRM_Plan 
   Caption         =   "Choix du Fichier"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7815
   OleObjectBlob   =   "FRM_Plan.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FRM_Plan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Btn_Annule_Click()
'Clic sur bouton Annuler
    Me.Hide
    Me.ChB_OkAnnule = False
End Sub

Private Sub Btn_Fic_Excel_Click()
'Documente les champs de la boite de dialogue à partir
'du fichier excel DSCGP
On Error GoTo Erreur
Dim DSCGP_EC As New c_DSCGP

    Me.TB_FicDSCGP = CATIA.FileSelectionBox("Selectionner le fichier DSCGP", "*.xls;*.xlsx", CatFileSelection)
    If Me.TB_FicDSCGP <> "" Then
            DSCGP_EC.VersionDscgp = 2
        DSCGP_EC.OpenDSCGP = Me.TB_FicDSCGP
        ValDscgp.NumGrilleAss = DSCGP_EC.NumGrille
        ValDscgp.NumGrilleAssSym = DSCGP_EC.NumGrilleSym
        ValDscgp.NumGrilleNue = DSCGP_EC.NumGrilleNue
        ValDscgp.NumGrilleNueSym = DSCGP_EC.NumGrilleSymNue
        ValDscgp.design = DSCGP_EC.DesignGrille
        ValDscgp.DesignSym = DSCGP_EC.DesignGrilleSym
        ValDscgp.NumEnvAvion = DSCGP_EC.EnvAvionCAO
        ValDscgp.Site = DSCGP_EC.Site
        ValDscgp.NumProgAvion = DSCGP_EC.NoProgAvion
        ValDscgp.CoteAvion = DSCGP_EC.CoteAvion
    
        'Coche la case "Sym" si grille sym
        If ValDscgp.NumGrilleAss <> "" And ValDscgp.NumGrilleAssSym <> "" Then
            Me.ChB_Sym.Value = True
        Else
            Me.ChB_Sym.Value = False
        End If
        'Affiche l'image check
        Me.img_check.Visible = True
        Me.img_uncheck.Visible = False
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
End Sub

Private Sub Btn_OK_Click()
'Cache la boite de dialogue et redonne la main au programme
    If Me.Cbx_FicaTraiter <> "" Then
        Me.Hide
        Me.ChB_OkAnnule = True
    Else
        Me.Cbx_FicaTraiter.BackColor = RGB(250, 0, 0)
        
    End If
End Sub




Private Sub Logo_eXcent_Click()
'Chargement de la boite eXcent
    Load Frm_eXcent
    Frm_eXcent.Show
    Unload Frm_eXcent
End Sub

Private Sub Rbt_Det_Change()
' Met a jour la liste des documents
    Dim mItem
    'Vide la liste
    With Me.Cbx_FicaTraiter
        While .ListCount > 0: .RemoveItem (0): Wend
    End With
    'Rempli la liste
    If Me.Rbt_Det Then
        For Each mItem In LstEnsDet("D")
            Me.Cbx_FicaTraiter.AddItem mItem
        Next
        Me.Cbx_FicGriNue.Visible = False
        Me.Lbl_FicGriNue.Visible = False
    End If
End Sub

Private Sub Rbt_Ens_Change()
' Met a jour la liste des documents
    Dim mItem
    'Vide la liste
    With Me.Cbx_FicaTraiter
        While .ListCount > 0: .RemoveItem (0): Wend
    End With
    If Me.Rbt_Ens Then
        'Rempli la liste
        For Each mItem In LstEnsDet("E")
            Me.Cbx_FicaTraiter.AddItem mItem
        Next
        For Each mItem In LstEnsDet("D")
            Me.Cbx_FicGriNue.AddItem mItem
        Next
        Me.Cbx_FicGriNue.Visible = True
        Me.Lbl_FicGriNue.Visible = True
    End If
End Sub

Private Sub RBt_VT_Change()
If Me.RBt_VT Then
    Me.ChB_Tronq.Enabled = True
Else
    Me.ChB_Tronq.Enabled = False
End If
End Sub


Private Sub UserForm_Initialize()
'Affiche l'image check
    Me.img_check.Visible = False
    Me.img_uncheck.Visible = True
'Ajout des formats à la liste
    Me.CBx_Format.AddItem "AIRBUS_FAL"
    'Me.CBx_Format.AddItem "AIRBUS_PREFAL"

'Active les boutons par défault
    Me.Rbt_Det = True
    Me.RBt_Horiz = True
    Me.RBt_CC = True
    Me.ChB_Tronq.Enabled = False
    Me.CBx_Format.Value = "AIRBUS_FAL"
End Sub

Private Function LstEnsDet(str As String) As Collection
'renvois la liste des Ensembles ou des parts de la collection des documents chargés
Dim mDoc As Document
Dim mdos As Documents
Dim mcol As New Collection

Set mDocs = CATIA.Documents
For Each mDoc In mDocs
    If InStr(1, mDoc.Name, ".CATProduct", vbTextCompare) And str = "E" Then
        mcol.Add mDoc.Name
    ElseIf InStr(1, mDoc.Name, ".CATPart", vbTextCompare) And str = "D" Then
        mcol.Add mDoc.Name
    End If
Next
Set LstEnsDet = mcol
Set mcol = Nothing
End Function
