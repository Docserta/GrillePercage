VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Navigateur1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*********************************************************************
'* Classe : Navigateur
'*
'* Fonctions :  Navigateur de fichiers.
'*              Permet de récupérer les disques, les répertoires et les fichiers
'*
'*
'* Version : 1
'* Création :  CFR
'* Modification : 24/08/15
'*
'**********************************************************************

Option Explicit

'Listes du contenu des disques et des répertoires dans le formulaire
Private Liste_Dsk_Nav As Variant, Liste_Rep_Nav As Variant
Private PremDisk_Nav As String

'Disques, repertoires
Private DskChoisi_Nav As String, RepChoisi_Nav As String, FicChoisi_Nav As String
Public DSK_Nav As Variant

'Liste des disques et des répertoires sur le disque
Public Liste_Dsk As Variant, Liste_Rep As Variant
Public Dsk As Variant

Private Sub Class_initialize()
    'affichage de la barre de progression
    'Load Frm_Navigateur
    'Frm_Navigateur_Fic.Show vbModeless

End Sub

Public Function ListeDisque() As Variant
'renvoi la liste des Disques et lecteurs réseau du poste de travail
Dim fsDsk As Object, fdsk As Object, D1dsk As Object
Dim LDList() As String
Dim i As Long
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
Dim fsRep As Object, fRep As Object, lsRep As Object, r1Rep As Object
Dim LRList() As String
Dim i As Long
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
    'Efface_Liste
    MsgBox Err.Number
    MsgBox Err.Description & "Err_ListeRep"
    GoTo Quit_ListeRep

End Function


Private Sub Construc_Liste()
'Construit la liste des répertoires
On Error GoTo Err_Construc_Liste
Dim Rep As Variant, Fic As Variant

'Nétoyage de la liste des répertoires
    'Efface_Liste
    If Not (IsEmpty(Liste_Rep_Nav)) Then
        'Construction de la liste des répertoires
        For Each Rep In Liste_Rep_Nav
            Frm_Navigateur_Rep.ListeRepForm.AddItem (Rep)
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




Private Sub Class_Terminate()
'masque et décharge de la barre de progression
'    FRM_Progression.Hide
'    Unload FRM_Progression
    
End Sub


