VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm_Navigateur_Rep 
   Caption         =   "Navigateur"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6105
   OleObjectBlob   =   "Frm_Navigateur_Rep.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Frm_Navigateur_Rep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Public Liste_Dsk_Nav, Liste_Rep_Nav, Liste_Fic_Nav As Variant
Public PremDisk_Nav As String
Public DskChoisi_Nav, RepChoisi_Nav, FicChoisi_Nav As String
Public DSK_Nav As Variant


Private Sub Btn_Annul_Click()
'Clic sur bouton Annuler
    Me.Hide
    Unload Frm_Navigateur_Rep
'Renvoi des valeurs vide
    Liste_Dsk = ""
    Liste_Rep = ""
    Liste_Fic = ""
    Dsk = ""
    Rep = ""
    Fic = ""
  
End Sub

Private Sub Btn_OK_Click()
'Clic sur bouton OK
    Me.Hide
'Renvoi les infos de disque, répertoire et fichiers
    'Liste_Dsk = Liste_Dsk_Nav
    'Liste_Rep = Liste_Rep_Nav
    Liste_Fic = Liste_Fic_Nav
    'Dsk = DskChoisi_Nav
    Rep = Me.ListDisqueForm
    'Fic = FicChoisi_Nav
     
End Sub

Private Sub ListDisqueForm_Change()
On Error GoTo Err_ListDisqueForm_Change
'Mise à jour de la liste des Répertoires en fonction du disque choisi
Dim LDFC_Dsk As String
Dim LDFTypeFic As String
LDFTypeFic = TypeFic
    LDFC_Dsk = Me.ListDisqueForm & "\"
    Liste_Rep_Nav = ListeRep(LDFC_Dsk)
    Construc_Liste

Quit_ListDisqueForm_Change:
    Exit Sub

Err_ListDisqueForm_Change:
    MsgBox Err.Number
    MsgBox Err.Description & "ListDisqueForm_Change"
    GoTo Quit_ListDisqueForm_Change

End Sub


Private Sub ListeRepForm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
'selectionne le repertoire double cliqué et met à jour les listes
On Error GoTo Err_ListDisqueForm_DblClick
Dim RepActuel_Nav As String
Dim RepCible_Nav As String
Dim TypFic_Nav As String
Liste_Fic_Nav = TypeFic
TypFic_Nav = TypeFic
RepActuel_Nav = Me.ListDisqueForm
    'teste si la liste des repertoire est vide
    If IsNull(Me.ListeRepForm) Then
        GoTo Quit_ListDisqueForm_DblClick
    Else
        RepChoisi_Nav = Me.ListeRepForm
    End If
    If RepChoisi_Nav = ".." Then
        If Len(RepActuel_Nav) > 3 Then
            RepCible_Nav = Enleve1Rep(RepActuel_Nav)
        Else
            GoTo Quit_ListDisqueForm_DblClick
        End If
    ElseIf RepChoisi_Nav = "." Then
        RepCible_Nav = Left(RepActuel_Nav, 3)
    Else
        RepCible_Nav = RepActuel_Nav & "\" & RepChoisi_Nav
        'concatène le nom du repertoire choisis avec le chemin présent
        'dans la liste des disques
    End If
        Me.ListDisqueForm = RepCible_Nav

        Liste_Rep_Nav = ListeRep(RepCible_Nav)
        Construc_Liste

Quit_ListDisqueForm_DblClick:
    Exit Sub

Err_ListDisqueForm_DblClick:
    'Nétoyage de la liste des Fichiers et des repertoires
    Efface_Liste
    MsgBox Err.Number
    MsgBox Err.Description & "Err_ListDisqueForm_DblClick"
    GoTo Quit_ListDisqueForm_DblClick
End Sub

Private Sub Logo_eXcent_Click()
'Chargement de la boite eXcent
    Load Frm_eXcent
    Frm_eXcent.Tbx_Version = VMacro
    Frm_eXcent.Show
    
    Unload Frm_eXcent
End Sub

Private Sub UserForm_Initialize()

    'Création de la liste des Disques
    Liste_Dsk_Nav = ListeDisque
    Dim TypFic_Nav As String
    TypFic_Nav = TypeFic
    'Rempli la liste du formulaire
    For Each DSK_Nav In Liste_Dsk_Nav
        Me.ListDisqueForm.AddItem (DSK_Nav)
    Next
    Me.ListDisqueForm.Value = Liste_Dsk_Nav(0)
    'Création de la liste des Répertoires
    PremDisk_Nav = Liste_Dsk_Nav(0) & "\"
    Liste_Rep_Nav = ListeRep(PremDisk_Nav)
    Construc_Liste
End Sub


Public Function ListeDisque() As Variant
'renvoi la liste des Disques et lecteurs réseau du poste de travail
Dim fsDsk, fdsk, D1dsk As Object
Dim LDList() As String
Set fsDsk = CreateObject("scripting.filesystemobject")
'Set fdsk = fsDsk.Drives
i = 0
    For Each D1dsk In fsDsk.Drives
        ReDim Preserve LDList(i)
        LDList(i) = D1dsk.DriveLetter & ":"
        i = i + 1
    Next
ListeDisque = LDList

End Function


Public Function ListeRep(LRChemin As String) As Variant
'renvoi la liste des répertoire contenu dans le repertoire passé en argument
On Error GoTo Err_ListeRep
Dim fsRep, fRep, lsRep, r1Rep As Object
Dim LRList() As String
Set fsRep = CreateObject("scripting.filesystemobject")
Set fRep = fsRep.GetFolder(LRChemin)
Set lsRep = fRep.SubFolders
    'Ajoute les répertoires . et ..
    ReDim Preserve LRList(0)
    LRList(0) = "."
    ReDim Preserve LRList(1)
    LRList(1) = ".."
    'Ajoute les sous repertoires
    i = 2
    For Each r1Rep In lsRep
        ReDim Preserve LRList(i)
        LRList(i) = r1Rep.Name
        i = i + 1
    Next
ListeRep = LRList

Quit_ListeRep:
    Exit Function

Err_ListeRep:
'Nétoyage de la liste des Fichiers et des repertoires
    Efface_Liste
    MsgBox Err.Number
    MsgBox Err.Description & "Err_ListeRep"
    GoTo Quit_ListeRep

End Function


Public Sub Efface_Liste()
Dim i, Fin As Integer
'Nétoyage de la liste des répertoires
    Fin = Me.ListeRepForm.ListCount - 1
    For i = Fin To 0 Step -1
        Me.ListeRepForm.RemoveItem (i)
    Next
End Sub
Public Sub Construc_Liste()
On Error GoTo Err_Construc_Liste
Dim Rep, Fic As Variant

'Nétoyage de la liste des répertoires et des fichiers
    Efface_Liste
    If Not (IsEmpty(Liste_Rep_Nav)) Then
        'Construction de la liste des répertoires
        For Each Rep In Liste_Rep_Nav
            Me.ListeRepForm.AddItem (Rep)
        Next
    End If

Quit_Construc_Liste:
    Exit Sub

Err_Construc_Liste:
    'MsgBox Err.Number
    'msgBox Err.Description & "Construc_Liste"
    GoTo Quit_Construc_Liste

End Sub


Public Function Enleve1Rep(E1R_Rep As String) As String
'Enleve un répertoire à la chaine
'détecte le rep avec le \

While Right(E1R_Rep, 1) <> "\"
    E1R_Rep = Left(E1R_Rep, Len(E1R_Rep) - 1)
Wend
Enleve1Rep = Left(E1R_Rep, Len(E1R_Rep) - 1)
End Function

