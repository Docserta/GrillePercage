Attribute VB_Name = "C2_Rename_Set"
Option Explicit
'*********************************************************************
'* Macro : C1_Rename_Set
'*
'* Fonctions :  Macro de renommage des points set g�om�trique
'*              suite a la suppression des voyelles accentu�es et des espaces
'*
'* Version : 1
'* Cr�ation :  CFR
'* Modification :
'*
'**********************************************************************
Sub catmain()

'Log de l'utilisation de la macro
LogUtilMacro nPath, nFicLog, nMacro, "C2_Rename_Set", VMacro

Dim mHybridBods As HybridBodies, mHybridSubBods As HybridBodies
Dim mHybridBody As HybridBody, mHybridSubBody As HybridBody
Dim PartGriNue As PartDocument

'---------------------------
' Checker l'environnement
'---------------------------
  
    Err.Clear
    On Error Resume Next
    Set PartGriNue = CATIA.ActiveDocument
    If Err.Number <> 0 Then
        MsgBox "Le document de la fen�tre courante n'est pas un CATPart !", vbCritical, "Environnement incorrect"
        End
    End If
    On Error GoTo 0
    
    Set mHybridBods = PartGriNue.part.HybridBodies
    If mHybridBods.Count > 0 Then
        For Each mHybridBody In mHybridBods
            Select Case mHybridBody.Name
                Case "r�f�rences externes isol�es"
                    mHybridBody.Name = nHBRefExtIsol
                Case "travail"
                    Set mHybridSubBods = mHybridBody.HybridBodies
                    If mHybridSubBods.Count > 0 Then
                        For Each mHybridSubBody In mHybridSubBods
                            Select Case mHybridSubBody.Name
                                Case "geometrie de reference"
                                    mHybridSubBody.Name = nHBGeoRef
                                Case "draft feet"
                                    mHybridSubBody.Name = nHBDrFeet
                                Case "draft pinules"
                                    mHybridSubBody.Name = nHBDrPin
                                Case "draft gravures"
                                    mHybridSubBody.Name = nHBDrGrav
                            End Select
                        Next
                    End If
            End Select
        Next
    End If
MsgBox "Renommage des set termin�", vbInformation

End Sub
