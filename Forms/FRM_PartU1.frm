VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FRM_PartU1 
   Caption         =   "Choix du Fichier"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7260
   OleObjectBlob   =   "FRM_PartU1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FRM_PartU1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub Btn_Add_U1_Click()

Set coll_docs = CATIA.Documents
Dim partDoc As PartDocument
Set partDoc = coll_docs.Add("Part")
Me.TBX_NomU01 = "c:\temp\Part_U01.CATPart"
partDoc.SaveAs "c:\temp\Part_U01"

End Sub

Private Sub Btn_Annule_Click()
'Clic sur bouton Annuler
    Me.Hide
    Me.ChB_OkAnnule = False
End Sub

Private Sub Btn_Nav_Gril_Click()
Dim NomComplet As String
NomComplet = CATIA.FileSelectionBox("Selectionner la part de la grille nue.", "*.CATPart", CatFileSelection)
Me.TBX_NomGrille = NomComplet
End Sub

Private Sub Btn_Nav_U1_Click()
Dim NomComplet As String
NomComplet = CATIA.FileSelectionBox("Selectionner la part U01.", "*.CATPart", CatFileSelection)
Me.TBX_NomU01 = NomComplet
End Sub

Private Sub Btn_OK_Click()
'Cache la boite de dialogue et redonne la main au programme
    Me.Hide
    Me.ChB_OkAnnule = True

End Sub

'Private Sub Btn_SelGrille_Click()
''Sélection de la part de la grille
'Me.Hide
'Dim BSG_nomPart As String
''Selection de la grille
'BSG_nomPart = Select_PartGrille(3)
'Me.TBX_NomGrille = BSG_nomPart
'Me.Show
'
'End Sub
'
'Private Sub Btn_SelU01_Click()
''Sélection de la part U01
'Me.Hide
'Dim BSU_nomPart As String
''Selection de la part
'BSU_nomPart = Select_PartGrille(4)
'Me.TBX_NomU01 = BSU_nomPart
'Me.Show
'
'End Sub

Private Sub Logo_eXcent_Click()
'Chargement de la boite eXcent
    Load Frm_eXcent
    Frm_eXcent.Tbx_Version = VMacro
    Frm_eXcent.Show
    Unload Frm_eXcent
End Sub

