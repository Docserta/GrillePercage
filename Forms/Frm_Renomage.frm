VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm_Renomage 
   Caption         =   "Renommage des Lignes, PtA et PtB"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8025
   OleObjectBlob   =   "Frm_Renomage.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Frm_Renomage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BtnAnnul_Click()
Me.Hide
Me.ChB_OkAnnule = False

End Sub

Private Sub BtnOK_Click()

Me.ChB_OkAnnule = True
Erreur = False

If Not Erreur Then
    Me.Hide
End If
End Sub

Private Sub CB_Isol_Click()
    If CB_Isol.Value = True Then
        Me.Frame_NomPt.enabled = False
        FrameEnabled Me.Frame_NomPt, False
        Me.Frm_Ordre.enabled = False
        FrameEnabled Me.Frm_Ordre, False
        Me.FrameIsole.enabled = True
        FrameEnabled Me.FrameIsole, True
    Else
        Me.Frame_NomPt.enabled = True
        FrameEnabled Me.Frame_NomPt, True
        Me.Frm_Ordre.enabled = True
        FrameEnabled Me.Frm_Ordre, True
        Me.FrameIsole.enabled = False
        FrameEnabled Me.FrameIsole, False
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

Me.RbtNumNomFast = True
Me.Rbt_RefSTD = True
Me.FrameIsole.enabled = False
FrameEnabled Me.FrameIsole, False
End Sub

Private Sub FrameEnabled(nFrame, statut)
Dim FrmControl As Control

For Each FrmControl In nFrame.Controls
    FrmControl.enabled = statut
Next
End Sub
