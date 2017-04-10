VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm_CreationPtA 
   Caption         =   "Création des Ligne PtA et PtB"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8970
   OleObjectBlob   =   "Frm_CreationPtA.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Frm_CreationPtA"
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

Private Sub Logo_eXcent_Click()
'Chargement de la boite eXcent
    Load Frm_eXcent
    Frm_eXcent.Tbx_Version = VMacro
    Frm_eXcent.Show
    
    Unload Frm_eXcent
End Sub

Private Sub Rbt_RefLEgacy_Click()
If Me.Rbt_RefLEgacy Then
    Me.Frame_NomPt.enabled = False
    Me.RbtNumCommentStd.enabled = False
    Me.RbtNumNomStd.enabled = False
    Me.Frame_SelPts.enabled = False
    Me.Rbt_SelPts.enabled = False
    Me.Rbt_SelSetRef.enabled = False
End If
End Sub

Private Sub Rbt_RefSTD_Click()
If Me.Rbt_RefSTD Then
    Me.Frame_NomPt.enabled = True
    Me.RbtNumCommentStd.enabled = True
    Me.RbtNumNomStd.enabled = True
    Me.Frame_SelPts.enabled = True
    Me.Rbt_SelPts.enabled = True
    Me.Rbt_SelSetRef.enabled = True

End If

End Sub

Private Sub UserForm_Initialize()
Me.Rbt_RefSTD = True
Me.RbtNumNomStd = True
Me.Rbt_SelSetRef = True

End Sub
