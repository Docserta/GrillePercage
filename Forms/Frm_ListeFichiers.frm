VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm_ListeFichiers 
   Caption         =   "Liste fichiers"
   ClientHeight    =   8880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8925
   OleObjectBlob   =   "Frm_ListeFichiers.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Frm_ListeFichiers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Btn_Annule_Click()
Me.Hide
Me.ChB_OkAnnule = False

End Sub

Private Sub Btn_Lancer_Click()
Me.Hide
Me.ChB_OkAnnule = True

End Sub

Private Sub Btn_MajDetromp_Click()
    Dim i As Integer
    Dim SelectedElement As Boolean
    For i = 0 To Frm_ListeFichiers.ListBox1.ListCount - 1
        'Boucle sur la liste des fichiers et test si le fichier est sélectionné
        If Frm_ListeFichiers.ListBox1.Selected(i) And Not SelectedElement Then
            Dim DscgpMaj As New c_DSCGP
            DscgpMaj.VersionDscgp = 2
            DscgpMaj.OpenDSCGP = CheminFicLot & Frm_ListeFichiers.ListBox1.List(i)
            Me.TBX_NomDtromp = DscgpMaj.NumPartDet
            SelectedElement = True
        End If
    Next i
    If Not SelectedElement Then
        MsgBox "Sélectionnez les fichiers DSCGP en cliquant sur le bouton Parcourir pour récupérer le nom de la part de détrompage.", vbInformation
    End If
End Sub

Private Sub Btn_Nav_Dest_Click()

    Me.TBX_FicDest = GetPath("dossier de destianation grille")

End Sub

Private Sub Btn_Nav_Detromp_Click()
Dim NomComplet As String
NomComplet = CATIA.FileSelectionBox("Selectionner le part de détrompage", "*.CATPart", CatFileSelection)
Me.TBX_NomDtromp = NomComplet

End Sub

Private Sub Btn_Nav_Env_Click()
    
    Dim NomComplet As String
    NomComplet = CATIA.FileSelectionBox("Selectionnez un des fichiers CATPART du répertoire à traiter.", "*.CATProduct", CatFileSelectionModeOpen)
    If NomComplet = "" Then Exit Sub 'on vérifie que qque chose a bien été selectionné
    Me.TBX_EnvAvion = NomComplet
    
End Sub

Private Sub Btn_Parcourir_Click()
'Recherche des fichiers pour les mettres dans la liste
'pour recuperer un chemin d'acces de fichier par selection graphique
    ListBox1.Clear
    Dim BP_NomComplet As String
    
    BP_NomComplet = CATIA.FileSelectionBox("Selectionnez un des fichiers CATPART du répertoire à traiter.", Me.Tbx_Extent, CatFileSelectionModeOpen)

    If BP_NomComplet = "" Then Exit Sub 'on vérifie que qque chose a bien été selectionné

    Dim ObjFichiers As File
    Set ObjFichiers = CATIA.FileSystem.GetFile(BP_NomComplet)
    Dim ObjRepertoire As Folder
    Set ObjRepertoire = ObjFichiers.ParentFolder
    CheminFicLot = ObjRepertoire.Path & "\"
    Dim ColecFichiers As Files
    Set ColecFichiers = ObjRepertoire.Files
    Dim NomFichier As String
    
    For j = 1 To ColecFichiers.Count
        NomFichier = ColecFichiers.Item(j).Name
        'test si c'est fichier du type souhaité
        If Right(NomFichier, Len(NomFichier) - InStr(1, NomFichier, ".", vbTextCompare)) = Right(Me.Tbx_Extent, Len(Me.Tbx_Extent) - InStr(1, Me.Tbx_Extent, ".", vbTextCompare)) Then
            'Ajout a la liste
            ListBox1.AddItem (NomFichier)
        End If
    Next j
End Sub

Private Sub Btn_ToutSel_Click()
'Tous selectionner
ListBox1.Visible = False
For intIndex = 0 To ListBox1.ListCount - 1
    ListBox1.Selected(intIndex) = True
Next
ListBox1.Visible = True
End Sub

Private Sub Btn_Trier_Click()
'Bouton de triage
    Nb_Item = ListBox1.ListCount
    Dim a As String
    Dim B As String
    For i = 1 To Nb_Item
        ind = Nb_Item - i
        a = ListBox1.List(Nb_Item - i)
        For j = 0 To Nb_Item - i
            B = ListBox1.List(j)
            If supp(a, B) Then
                a = B
                ind = j
            End If
        Next j
        'a ce moment la, a est le plus petit : on le place donc à la fin. Pour cela,
        'on le supprime, et on en cree un nouveau
        ListBox1.RemoveItem (ind)
        ListBox1.AddItem (a)
    Next i

End Sub

Private Sub CB_Catpartactif_Click()
If Me.CB_Catpartactif Then
    Me.Cadre_ListeFic.Enabled = False
    Me.Btn_Parcourir.Enabled = False
    Me.Btn_ToutSel.Enabled = False
    Me.Btn_Trier.Enabled = False
    Me.ListBox1.Visible = False
Else
    Me.Cadre_ListeFic.Enabled = True
    Me.Btn_Parcourir.Enabled = True
    Me.Btn_ToutSel.Enabled = True
    Me.Btn_Trier.Enabled = True
    Me.ListBox1.Visible = True
End If
End Sub

Private Sub ChB_Detromp_Click()
    If Me.ChB_Detromp Then
        Me.LB_Detromp.Enabled = True
        Me.TBX_NomDtromp.Enabled = True
        Me.Btn_MajDetromp.Visible = True
        'Me.Btn_Nav_Detromp.Enabled = True
    Else
        Me.LB_Detromp.Enabled = False
        Me.TBX_NomDtromp.Enabled = False
        Me.Btn_MajDetromp.Visible = False
        'Me.Btn_Nav_Detromp.Enabled = False
    End If
End Sub

Private Sub Logo_eXcent_Click()
'Chargement de la boite eXcent
    Load Frm_eXcent
    Frm_eXcent.Tbx_Version = VMacro
    Frm_eXcent.Show
    
    Unload Frm_eXcent
End Sub


Function supp(a, B) As Boolean
    supp = (a > B)
End Function


Private Sub UserForm_Initialize()
'Me.TBX_EnvAvion = "C:\CFR\Dropbox\Macros\Grilles-Percage\FichiersCAO\test\aile_avion.CATProduct"
'Me.TBX_FicDest = "c:\temp"
Me.LB_Detromp.Enabled = False
Me.TBX_NomDtromp.Enabled = False
Me.Btn_Nav_Detromp.Enabled = False
End Sub
