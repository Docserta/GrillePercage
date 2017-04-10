VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm_Check3D 
   Caption         =   "Choix du Fichier"
   ClientHeight    =   9375
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7350
   OleObjectBlob   =   "Frm_Check3D.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Frm_Check3D"
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
'Sélectionne le fichier excel du DSCGP (Chemin+nom)
    Me.TB_FicDSCGP = CATIA.FileSelectionBox("Selectionner le fichier DSCGP", "*.xls;*.xlsx", CatFileSelectionModeOpen)
    
    'Recupère le nom seul du fichier DSCGP
    If Me.TB_FicDSCGP <> "" Then
        Me.Tbx_NoDSCGP = DecoupeSlash(Me.TB_FicDSCGP)
        Me.img_uncheck.Visible = False
        Me.img_check.Visible = True
    End If
End Sub

Private Sub Btn_OK_Click()
'Cache la boite de dialogue et redonne la main au programme
     If Me.TB_FicDSCGP = "" Then
        MsgBox "Pas de fichier excel sélectionné !"
    Else
        Me.Hide
        Me.ChB_OkAnnule = True
    End If
    
End Sub

Private Sub ChecB_Tous_Click()
'Active ou désactives toutes les cases a cocher
Dim colChbox As Controls
Dim cbx As Variant
Set colChbox = Frm_Check3D.Controls
    If Me.ChecB_Tous = True Then
        For Each cbx In colChbox
            If Left(cbx.Name, 3) = "ChB" Then
                cbx.Value = True
            End If
        Next
    Else
        For Each cbx In colChbox
            If Left(cbx.Name, 3) = "ChB" Then
                cbx.Value = False
            End If
        Next
    End If
End Sub

Private Sub Logo_eXcent_Click()
'Chargement de la boite eXcent
    Load Frm_eXcent
    Frm_eXcent.Tbx_Version = VMacro
    Frm_eXcent.Show

    Unload Frm_eXcent
End Sub

Private Sub UserForm_Initialize()
Me.img_check.Visible = False
Me.img_uncheck.Visible = True
End Sub
