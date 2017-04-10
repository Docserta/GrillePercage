VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FRM_NoMachine 
   Caption         =   "Numéro de machine"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7200
   OleObjectBlob   =   "FRM_NoMachine.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FRM_NoMachine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Private Sub Btn_Nav_Env_Click()
'Load FRM_Navigateur_Fic
'FRM_Navigateur_Fic.Show
''Me.TBX_FicEnv = Rep
'
'Unload FRM_Navigateur_Fic
'End Sub
'
'Private Sub Btn_Nav_RepSauv_Click()
'Load FRM_Navigateur_Rep
'FRM_Navigateur_Rep.Show
''Me.TBX_RepSave = Rep
'Unload FRM_Navigateur_Rep
'End Sub

Private Sub BtnAnnul_Click()
Me.Hide
Me.ChB_OkAnnule = False


End Sub

Private Sub BtnOK_Click()

Me.ChB_OkAnnule = True
Erreur = False

    If Me.TBX_NomAss = "" Then
        Me.TBX_NomAss.BackColor = RGB(255, 212, 255)
        Erreur = True
    Else
        Me.TBX_NomAss.BackColor = RGB(212, 255, 255)
    End If
If Not Erreur Then
    Me.Hide
End If
End Sub

Private Sub Logo_eXcent_Click()
'Chargement de la boite eXcent
    Load Frm_eXcent
    Frm_eXcent.Tbx_Version = VMacro
    Frm_eXcent.Show
    
    Unload Frm_eXcent
End Sub
