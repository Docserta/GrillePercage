VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_ProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*********************************************************************
'* Classe : c_ProgressBar
'*
'* Fonctions :  Afficher et faire défiler une barre de progression avec
'*              différentes informations
'*
'*
'* Version : 1
'* Création :  CFR
'* Modification : 18/08/15
'*
'**********************************************************************

Private Sub Class_initialize()
    'affichage de la barre de progression
    Load Frm_Progression
    Frm_Progression.Show vbModeless

End Sub

Property Let Balayage(ValProgress As Long)
'Modifie la longueur et la position de la barre de progression
'Pour effectuer un balayage.
'Utilisé lorsque le traitement peu être long et qu'on ne peu pas calculer le nombre d'opérations
 Dim Lgbarre As Integer
 Dim posMin As Integer, posMax As Integer, DebBarre As Integer, FinBarre As Integer
 
    Lgbarre = 7
    posMin = 18
    posMax = 318
    While ValProgress > 100
        ValProgress = ValProgress - 100
    Wend
    If ValProgress < Lgbarre Then
        DebBarre = posMin
        FinBarre = ValProgress * 3
    ElseIf (ValProgress >= 100 - Lgbarre And ValProgress < 100) Then
        DebBarre = ValProgress * 3
        'FinBarre = (Lgbarre * 3 - (posMax - posMin - ValProgress * 3)) * 3
        FinBarre = (posMax - posMin - ValProgress * 3)
    ElseIf ValProgress >= Lgbarre And ValProgress < 100 - Lgbarre Then
        DebBarre = ValProgress * 3
        FinBarre = Lgbarre * 3
    End If
    
        Frm_Progression.Bar1.Left = DebBarre
        Frm_Progression.Bar1.Width = FinBarre
    
       
End Property

Property Let Progression(ValProgress As Long)
'Modifie la longueur de la barre de progression
    If ValProgress > 100 Then PBAvancement = 100
    If ValProgress < 1 Then ValProgress = 1
    Frm_Progression.Bar1.Width = ValProgress * 3

End Property

Property Let Titre(StrTitre As String)
'Change le texte de titre du formulaire
    Frm_Progression.Caption = StrTitre
End Property

Property Let Etape(StrEtape As String)
'Change le texte dans la barre de progression
    Frm_Progression.Lbl_Etape = StrEtape
End Property

Public Sub ProgressTitre(ValProgess As Long, StrTitre As String)
' Modifie la longueur de la barre de progression et change le titre en une seule fonction
    Titre = StrTitre
    Progression = ValProgess
End Sub

Public Sub ProgressEtape(ValProgess As Long, StrEtape As String)
' Modifie la longueur de la barre de progression et change le texte dans la barre de progression en une seule fonction
    Etape = StrEtape
    Progression = ValProgess
End Sub

Public Sub CalculProgression(PB_Etape As Long, PB_Nb_Etapes As Long, PB_Item As Long, PB_NbItem As Long, PB_Caption As String)
'gestion de l'avancement de la barre de progression pour une succession d'étapes contenant plusieurs item

' PB_Etape = etape en cours
' PB_Nb_Etapes = nombre d'étapes
' PB_Item = item en cours de l'étape en cours
' PB_NbItem = Nombre d'items dans l'étape en cour
' PB_Caption = Titre de la fenètre
'(((PB_Etape - 1) * (100 / PB_NB_Etapes)) + (((100 / PB_NB_Etapes) / PB_NbItem) * PB_Item))
Dim PB_Barre As Long
    Titre = PB_Caption
    If PB_Nb_Etapes = 0 Or PB_NbItem = 0 Then
        Exit Sub
    End If
    PB_Barre = (((PB_Etape - 1) * (100 / PB_Nb_Etapes)) + (((100 / PB_Nb_Etapes) / PB_NbItem) * PB_Item))
    Progression = PB_Barre
End Sub

Public Sub Cache()
'Masque la barre de progression
    Frm_Progression.Hide
End Sub

Public Sub Affiche()
'Affiche la barre de progression
    Frm_Progression.Show vbModeless
End Sub

Private Sub Class_Terminate()
'masque et décharge de la barre de progression
    Frm_Progression.Hide
    Unload Frm_Progression
    
End Sub

