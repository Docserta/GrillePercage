Attribute VB_Name = "Y1_Check_3D"

Option Explicit
Public ObjResultCheck3D As Object
Public Col_Parts As c_cKParts
Public WriteMesures As Boolean

' *****************************************************************
'* Macro : Y1_Check_3D
'*
'* Fonctions :  Effectue les Checks CTD
'*
'*
'* Version : 9
'* Création :  CFR
' *
' * Création CFR le : 09/05/2016
' *
' *****************************************************************

Sub CATMain()

Dim Doc_AssGen As ProductDocument
Dim ProdAssGen As Product
Dim ColProdAss As Products
Dim GrilleASs_EC As Product
Dim CK3d As New check3d
Dim FicRes As String
Dim testDSCGP As c_DSCGP

'Log de l'utilisation de la macro
LogUtilMacro nPath, nFicLog, nMacro, "Y1_Check_3D", VMacro

    '-------------------------
    ' Check de l'environnement
    '-------------------------
Set coll_docs = CATIA.Documents
    If Not Check_GrilleAss() Then
        Exit Sub
    End If
        
'Ouvre la boite de dlg "Frm_check3d"
    Load Frm_Check3D
    
'Recupération du Product général
    Set Doc_AssGen = CATIA.ActiveDocument
    Set ProdAssGen = Doc_AssGen.Product
    
'Recupération de la liste des Griles ass dans le l'assemblage
    Set ColProdAss = ProdAssGen.Products
'#
'# Attention si un product Grille assemblé est actif et non un product Lot => message d'erreur
'#

    If ColProdAss.Count > 0 Then
        For Each GrilleASs_EC In ColProdAss
            On Error Resume Next
            Frm_Check3D.CBL_GrilleAss.AddItem GrilleASs_EC.PartNumber
            If Err.Number <> 0 Then
                Err.Clear
            End If
            On Error GoTo 0
        Next
    End If
    Frm_Check3D.Tbx_NoGrille = ProdAssGen.Name
    Frm_Check3D.Show

'Sort du programme si click sur bouton Annuler
    If Not (Frm_Check3D.ChB_OkAnnule) Then
        Unload Frm_Check3D
        Exit Sub
    End If
    Frm_Check3D.Hide
    
'collecte de la liste des checks
    Set ObjResultCheck3D = New c_itemChecks3D
    InitColResul

'Déclaration du nom du fichier excel du DSCGP
    If Frm_Check3D.TB_FicDSCGP <> "" Then
        'Vérification que le product général correspond bien à un lot de grilles
        Set testDSCGP = New c_DSCGP
        testDSCGP.VersionDscgp = 2
        testDSCGP.OpenDSCGP = Frm_Check3D.TB_FicDSCGP
        If Doc_AssGen.Product.PartNumber <> testDSCGP.NumduLot Then
            MsgBox "Le nom du Product général ne correspond pas à un lot de grilles!" & Chr(110) & "Lancez la macro sur le CATProduct du lot", vbCritical, " Environnement incorrect"
            End
        Else
            Set testDSCGP = Nothing
        End If
        CK3d.VersDscgp = 2
        CK3d.NomDscgp = Frm_Check3D.TB_FicDSCGP
    Else
        MsgBox "Pas de fichier excel sélectionné !"
        End
    End If
    
'Déclaration de la grille assemblée choisie
    CK3d.NomDocGrilleAss = Frm_Check3D.CBL_GrilleAss & ".CATProduct"

    If Frm_Check3D.ChB_F Then
        ObjResultCheck3D.WriteItem "F-01", CK3d.CK_F01.Check, CK3d.CK_F01.statut, CK3d.CK_F01.Comment
        ObjResultCheck3D.WriteItem "F-02", CK3d.CK_F02.Check, CK3d.CK_F02.statut, CK3d.CK_F02.Comment
        ObjResultCheck3D.WriteItem "F-03", CK3d.CK_F03.Check, CK3d.CK_F03.statut, CK3d.CK_F03.Comment
        ObjResultCheck3D.WriteItem "F-04", CK3d.CK_F04.Check, CK3d.CK_F04.statut, CK3d.CK_F04.Comment
        ObjResultCheck3D.WriteItem "F-05", CK3d.CK_F05.Check, CK3d.CK_F05.statut, CK3d.CK_F05.Comment
    End If
    If Frm_Check3D.ChB_G Then
        ObjResultCheck3D.WriteItem "G-01", CK3d.CK_G01.Check, CK3d.CK_G01.statut, CK3d.CK_G01.Comment
        ObjResultCheck3D.WriteItem "G-02", CK3d.CK_G02.Check, CK3d.CK_G02.statut, CK3d.CK_G02.Comment
        ObjResultCheck3D.WriteItem "G-03", CK3d.CK_G03.Check, CK3d.CK_G03.statut, CK3d.CK_G03.Comment
        ObjResultCheck3D.WriteItem "G-04", CK3d.CK_G04.Check, CK3d.CK_G04.statut, CK3d.CK_G04.Comment
        ObjResultCheck3D.WriteItem "G-06", CK3d.CK_G06.Check, CK3d.CK_G06.statut, CK3d.CK_G06.Comment
    End If
    If Frm_Check3D.ChB_H Then
        ObjResultCheck3D.WriteItem "H-01", CK3d.CK_H01.Check, CK3d.CK_H01.statut, CK3d.CK_H01.Comment
        ObjResultCheck3D.WriteItem "H-02", CK3d.CK_H02.Check, CK3d.CK_H02.statut, CK3d.CK_H02.Comment
        ObjResultCheck3D.WriteItem "H-03", CK3d.CK_H03.Check, CK3d.CK_H03.statut, CK3d.CK_H03.Comment
        ObjResultCheck3D.WriteItem "H-04", CK3d.CK_H04.Check, CK3d.CK_H04.statut, CK3d.CK_H04.Comment
        ObjResultCheck3D.WriteItem "H-05", CK3d.CK_H05.Check, CK3d.CK_H05.statut, CK3d.CK_H05.Comment
        ObjResultCheck3D.WriteItem "H-06", CK3d.CK_H06.Check, CK3d.CK_H06.statut, CK3d.CK_H06.Comment
        ObjResultCheck3D.WriteItem "H-07", CK3d.CK_H07.Check, CK3d.CK_H07.statut, CK3d.CK_H07.Comment
        ObjResultCheck3D.WriteItem "H-08", CK3d.CK_H08.Check, CK3d.CK_H08.statut, CK3d.CK_H08.Comment
    End If
    If Frm_Check3D.ChB_I Then
        ObjResultCheck3D.WriteItem "I-01", CK3d.CK_I01.Check, CK3d.CK_I01.statut, CK3d.CK_I01.Comment
        ObjResultCheck3D.WriteItem "I-02", CK3d.CK_I02.Check, CK3d.CK_I02.statut, CK3d.CK_I02.Comment
        ObjResultCheck3D.WriteItem "I-03", CK3d.CK_I03.Check, CK3d.CK_I03.statut, CK3d.CK_I03.Comment
        ObjResultCheck3D.WriteItem "I-04", CK3d.CK_I04.Check, CK3d.CK_I04.statut, CK3d.CK_I04.Comment
    End If
    If Frm_Check3D.ChB_J Then
        WriteMesures = True
        ObjResultCheck3D.WriteItem "J-01", CK3d.CK_J01.Check, CK3d.CK_J01.statut, CK3d.CK_J01.Comment
    End If
    If Frm_Check3D.ChB_K Then
    
    End If
    If Frm_Check3D.ChB_L Then
    
    End If
    If Frm_Check3D.ChB_M Then
    
    End If
    If Frm_Check3D.ChB_M2 Then
    
    End If
    If Frm_Check3D.ChB_M3 Then
    
    End If
    If Frm_Check3D.ChB_N Then
        ObjResultCheck3D.WriteItem "N-01", CK3d.CK_N01.Check, CK3d.CK_N01.statut, CK3d.CK_N01.Comment
        ObjResultCheck3D.WriteItem "N-02", CK3d.CK_N02.Check, CK3d.CK_N02.statut, CK3d.CK_N02.Comment
        ObjResultCheck3D.WriteItem "N-04", CK3d.CK_N04.Check, CK3d.CK_N04.statut, CK3d.CK_N04.Comment
        ObjResultCheck3D.WriteItem "N-06", CK3d.CK_N06.Check, CK3d.CK_N06.statut, CK3d.CK_N06.Comment
    End If
    
    If Frm_Check3D.ChB_O Then
        ObjResultCheck3D.WriteItem "O-06", CK3d.CK_O06.Check, CK3d.CK_O06.statut, CK3d.CK_O06.Comment
    End If
    
     If Frm_Check3D.ChB_P Then
    
    End If
    
    If Frm_Check3D.ChB_Q Then
    
    End If
    
    If Frm_Check3D.ChB_R Then
    
    End If
    
    FicRes = EcritResult(Frm_Check3D.CBL_GrilleAss)
    Unload Frm_Check3D
    Set ObjResultCheck3D = Nothing
    
    MsgBox "Fin du check 3D. Le fichier de résultat est enregistré sous : " & FicRes
    
    'Libération des classes
    Set CK3d = Nothing
    
End Sub

Public Sub InitColResul()
'Construit la collection de item a controler à partir du template excel
'ObjResultCheck3D = collection des items de controle
'nFicXl = Nom du fichier excel si reprise d'un check préalablement fait
Dim CheminTemplate As String
    CheminTemplate = Get_Active_CATVBA_Path & "\" & NomTemplateCheck3D
   
Dim objexcel
Dim objWorkBook
Dim objWS
    Set objexcel = CreateObject("EXCEL.APPLICATION")
    Set objWorkBook = objexcel.Workbooks.Open(CStr(CheminTemplate))
    Set objWS = objWorkBook.Sheets("Checks")
    Dim Ltemp As Long
        Ltemp = PremLigSheetCk '1ere ligne des Checks
    Dim ChkFait As Boolean
    Dim Cote As String, statut As String, Criticite As String, Comment As String
    Cote = objWS.cells(Ltemp, c_Cote).Value
    Do While Cote <> ""
            If objWS.cells(Ltemp, c_Check).Value = "" Then ChkFait = False Else ChkFait = True
            Cote = objWS.cells(Ltemp, c_Cote).Value
            If IsEmpty(objWS.cells(Ltemp, c_Statut).Value) Then statut = "" Else statut = objWS.cells(Ltemp, c_Statut).Value
            If IsEmpty(objWS.cells(Ltemp, c_Comment).Value) Then Comment = "" Else Comment = objWS.cells(Ltemp, c_Comment).Value
            'Incrémentation de la collection
            ObjResultCheck3D.Add Cote, ChkFait, statut, Comment
            Ltemp = Ltemp + 1
        Loop

objWorkBook.Close
Set objexcel = Nothing

End Sub

Public Function EcritResult(ByVal NomFic As String) As String
'Ecrit le résultat des checks dans un fichier excel
Dim NomFicRes As String
    NomFicRes = "RC3D-" & NomFic & ".xls"
Dim CheminTemplate As String
    CheminTemplate = Get_Active_CATVBA_Path & "\" & NomTemplateCheck3D

'verifie si un fichier de rapport est déja présent et l'efface
    If Not (EffaceFicNom(CheminDestRapport, NomFicRes)) Then
        End
    End If

Dim objexcel
Dim objWorkBook
Dim objWS
    Set objexcel = CreateObject("EXCEL.APPLICATION")
    Set objWorkBook = objexcel.Workbooks.Open(CStr(CheminTemplate))
    Set objWS = objWorkBook.Sheets("Checks")
    Dim Ltemp As Long
        Ltemp = PremLigSheetCk '1ere ligne des Checks
    Dim ChkFait As Boolean
    Dim Cote As String
    Dim check_EC As c_itemCheck3D
    Cote = objWS.cells(Ltemp, c_Cote).Value
    Do While Cote <> ""
        Set check_EC = ObjResultCheck3D.Item(Cote)
        If check_EC.Check Then objWS.cells(Ltemp, c_Check).Value = "Checked" Else objWS.cells(Ltemp, c_Check).Value = "Not checked"
        objWS.cells(Ltemp, c_Statut).Value = check_EC.statut
        objWS.cells(Ltemp, c_Comment).Value = check_EC.Comment
            
        Ltemp = Ltemp + 1
        Cote = objWS.cells(Ltemp, c_Cote).Value
    Loop
    
'Mise en forme
    objWS.range(Cel_NumGrille).Value = NomFic
    objWS.range(Cel_Date).Value = "'" & Date
    objWS.range(Cel_Controleur).Value = ReturnUserName
    
'Ecriture des mesures
    If WriteMesures Then
        Set objWS = objWorkBook.Sheets(nSheetMes)
        WriteMesure objWS
    End If

'Sauvegarde et fermeture du fichier excel
objexcel.DisplayAlerts = False
objWorkBook.SaveAs CheminDestRapport & NomFicRes, True
EcritResult = CheminDestRapport & NomFicRes

objWorkBook.Close
objexcel.DisplayAlerts = True
Set objexcel = Nothing
End Function

Private Sub WriteMesure(tWS)
'Ecrit les mesures dans l'onglet "mesures"
Dim tMes As c_mesure
Set tMes = New c_mesure
Dim i As Long
Dim Ltemp As Long
Dim Plage As String
    Ltemp = PremLigSheetMes
    For i = 1 To col_Mes.Count
        Set tMes = col_Mes.Item(i)
        tWS.cells(Ltemp, c_nomFas).Value = tMes.nom
        tWS.cells(Ltemp, c_Xr).Value = tMes.X
        tWS.cells(Ltemp, c_Yr).Value = tMes.Y
        tWS.cells(Ltemp, c_Zr).Value = tMes.Z
        tWS.cells(Ltemp, c_Xe).Value = tMes.Xe
        tWS.cells(Ltemp, c_Ye).Value = tMes.Ye
        tWS.cells(Ltemp, c_Ze).Value = tMes.Ze
        tWS.cells(Ltemp, c_Xec).Value = Round(Abs(tMes.X - tMes.Xe), 3)
        tWS.cells(Ltemp, c_Yec).Value = Round(Abs(tMes.Y - tMes.Ye), 3)
        tWS.cells(Ltemp, c_Zec).Value = Round(Abs(tMes.Z - tMes.Ze), 3)
        'Les Points A
        tWS.cells(Ltemp, c_nomPtA).Value = tMes.NomPtA
        If InStr(1, tMes.NomPtA, tMes.nom, vbTextCompare) = 0 Then
            tWS.cells(Ltemp, c_nomPtA).Font.colorindex = 3
        End If
        tWS.cells(Ltemp, c_PtAX).Value = Round(tMes.PtAX, 3)
        tWS.cells(Ltemp, c_PtAY).Value = Round(tMes.PtAY, 3)
        tWS.cells(Ltemp, c_PtAZ).Value = Round(tMes.PtAZ, 3)
'        tWS.cells(Ltemp, c_PtAX).Value = Round(Abs(tMes.Xe - tMes.PtAX), 3)
'        tWS.cells(Ltemp, c_PtAY).Value = Round(Abs(tMes.Ye - tMes.PtAY), 3)
'        tWS.cells(Ltemp, c_PtAZ).Value = Round(Abs(tMes.Ze - tMes.PtAZ), 3)
        
        Ltemp = Ltemp + 1
    Next
    'Mise en forme conditionnelle. Si l'écart des coordonnées réelles des fasteners est suppérieur à 0.002
        Plage = NumCar(c_Xec) & PremLigSheetMes & ":" & NumCar(c_Zec) & Ltemp
        With tWS.range(Plage)
        .formatconditions.Delete
        .formatconditions.Add xLCellValue, xLGreater, "0,002"
        .formatconditions(1).Font.colorindex = 3
        End With
'    'Mise en forme conditionnelle. Si l'écart des coordonnées des pta A de 0.002 aux coordonnées des Fasteners
'        Plage = NumCar(c_PtAX) & PremLigSheetMes & ":" & NumCar(c_PtAZ) & Ltemp
'        With tWS.range(Plage)
'        .formatconditions.Delete
'        .formatconditions.Add xLCellValue, xLGreater, "0,002"
'        .formatconditions(1).Font.colorindex = 3
'        End With
    
'Libération des classes
Set tMes = Nothing

End Sub
