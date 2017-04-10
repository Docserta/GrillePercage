Attribute VB_Name = "Y2_Check_DSCGP"
Option Explicit
Public GrAss As New GrilleAss
Public GrNue As New c_PartGrille

Sub catmain()
' *****************************************************************
'* Macro : Y1_Check_DSCGP
'*
'* Fonctions :  Compare les fichiers de grille au DSCGP
'*              V�rifie que les infos correspondent
'*              Propose la correction de certains param�tres
'*
'* Version : 9
'* Cr�ation :  CFR
' *
' * Cr�ation CFR le : 08/04/2016
' *
' *****************************************************************
 
'Log de l'utilisation de la macro
LogUtilMacro nPath, nFicLog, nMacro, "Y2_Check_DSCGP", VMacro
 
 'S�lection du Product grille assembl�e et du DSCGP
    Load FRM_CheckDscgp
    FRM_CheckDscgp.Show

'Clic sur Bouton "Annule" dans formulaire
    If Not (FRM_CheckDscgp.ChB_OkAnnule) Then
        End
    End If

    Unload FRM_CheckDscgp
    
End Sub

