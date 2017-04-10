Attribute VB_Name = "Y3_Check_Part"
Option Explicit

Sub catmain()
' *****************************************************************
'* Macro : Y1_Check_Part
'*
'* Fonctions :  Check la structure du catpart actif
'*              V�rifie que tous les �l�mnets standart sont pr�sents (Set g�om�trique, surfaces, pts etc.
'*
'* Version : 8
'* Cr�ation :  CFR
' *
' * Cr�ation CFR le : 21/06/2015
' * Modification le : 24/06/16
' *                   remplacement du tableau Coll_RefExIsol par la classe c_Fasteners
' * Modification le : 01/12/16
' *                   ajout mise en forme conditionnelle sur faux points B
' *****************************************************************

'Log de l'utilisation de la macro
LogUtilMacro nPath, nFicLog, nMacro, "Y3_Check_Part", VMacro
'---------------------------
' Checker l'environnement
'---------------------------
    Dim instance_catpart_grille_nue As PartDocument
    Err.Clear
    On Error Resume Next
    Set instance_catpart_grille_nue = CATIA.ActiveDocument
    If Err.Number <> 0 Then
        MsgBox "Le document de la fen�tre courante n'est pas un CATPart !", vbCritical, "Environnement incorrect"
        End
    End If
    On Error GoTo 0
    
    Dim GrilleActive As New c_PartGrille
    Dim tLisfast As c_Fasteners
    Set tLisfast = GrilleActive.Fasteners
    Dim tFast As c_Fastener
    Set tFast = New c_Fastener
    
    Dim NomRapportCheck As String
        NomRapportCheck = "Check_" & GrilleActive.nom & ".xlsx"
    Dim LigneEC As Long
        LigneEC = 1
    Dim i As Long
    Dim NB_Ligne As Long
    Dim XlsFormule As String
    
'Ouverture de la trame Excel
    Dim objExcelCheck
    Set objExcelCheck = CreateObject("EXCEL.APPLICATION")

    Dim objWorkSheet

    objExcelCheck.WindowState = 1
    objExcelCheck.Visible = True
    objExcelCheck.Workbooks.Add

    Set objWorkSheet = objExcelCheck.worksheets.Item(1)

'verifie si un fichier de rapport est d�ja pr�sent et l'efface
    If Not (EffaceFicNom(CheminDestRapport, NomRapportCheck)) Then
        End
    End If
    
 'Ent�te
    objWorkSheet.range("A" & LigneEC) = "Check des �lements constitutif de la grille"
    LigneEC = LigneEC + 2
    
 'Test du Format du Nom de fichier
 
 'Test des Sets g�om�triques
    objWorkSheet.range("A" & LigneEC) = "Test l'existance des Set g�om�triques :"
    LigneEC = LigneEC + 1
    
    objWorkSheet.range("A" & LigneEC) = "Set g�om�trique R�f�rence externes isol�e"
    If GrilleActive.Exist_HB(nHBRefExtIsol) Then EcritCheck "B" & LigneEC, "OK", objWorkSheet

    LigneEC = LigneEC + 1
    
    objWorkSheet.range("A" & LigneEC) = "Set g�om�trique detrompage"
    If GrilleActive.Exist_HB(nHBSDetr) Then EcritCheck "B" & LigneEC, "OK", objWorkSheet Else EcritCheck "B" & LigneEC, "KO", objWorkSheet
 
    LigneEC = LigneEC + 1
    
    objWorkSheet.range("A" & LigneEC) = "Set g�om�trique draft feet"
    If GrilleActive.Exist_HB(nHBDrFeet) Then EcritCheck "B" & LigneEC, "OK", objWorkSheet Else EcritCheck "B" & LigneEC, "KO", objWorkSheet
 
    LigneEC = LigneEC + 1
    
    objWorkSheet.range("A" & LigneEC) = "Set g�om�trique draft gravures"
    If GrilleActive.Exist_HB(nHBDrGrav) Then EcritCheck "B" & LigneEC, "OK", objWorkSheet Else EcritCheck "B" & LigneEC, "KO", objWorkSheet
 
    LigneEC = LigneEC + 1
    
    objWorkSheet.range("A" & LigneEC) = "Set g�om�trique draft pinules "
    If GrilleActive.Exist_HB(nHBDrPin) Then EcritCheck "B" & LigneEC, "OK", objWorkSheet Else EcritCheck "B" & LigneEC, "KO", objWorkSheet
 
    LigneEC = LigneEC + 1
    
    objWorkSheet.range("A" & LigneEC) = "Set g�om�trique feet"
    If GrilleActive.Exist_HB(nHBFeet) Then EcritCheck "B" & LigneEC, "OK", objWorkSheet Else EcritCheck "B" & LigneEC, "KO", objWorkSheet
 
    LigneEC = LigneEC + 1
    
    objWorkSheet.range("A" & LigneEC) = "Set g�om�trique geometrie de reference "
    If GrilleActive.Exist_HB(nHBGeoRef) Then EcritCheck "B" & LigneEC, "OK", objWorkSheet Else EcritCheck "B" & LigneEC, "KO", objWorkSheet
 
    LigneEC = LigneEC + 1
    
    objWorkSheet.range("A" & LigneEC) = "Set g�om�trique gravures"
    If GrilleActive.Exist_HB(nHBGrav) Then EcritCheck "B" & LigneEC, "OK", objWorkSheet Else EcritCheck "B" & LigneEC, "KO", objWorkSheet
 
    LigneEC = LigneEC + 1
    
    objWorkSheet.range("A" & LigneEC) = "Set g�om�trique pinules"
    If GrilleActive.Exist_HB(nHBPin) Then EcritCheck "B" & LigneEC, "OK", objWorkSheet Else EcritCheck "B" & LigneEC, "KO", objWorkSheet
 
    LigneEC = LigneEC + 1
    
    objWorkSheet.range("A" & LigneEC) = "Set g�om�trique pointsA "
    If GrilleActive.Exist_HB(nHBPtA) Then EcritCheck "B" & LigneEC, "OK", objWorkSheet Else EcritCheck "B" & LigneEC, "KO", objWorkSheet
 
    LigneEC = LigneEC + 1
    
    objWorkSheet.range("A" & LigneEC) = "Set g�om�trique pointsB "
    If GrilleActive.Exist_HB(nHBPtB) Then EcritCheck "B" & LigneEC, "OK", objWorkSheet Else EcritCheck "B" & LigneEC, "KO", objWorkSheet
 
    LigneEC = LigneEC + 1
    
    objWorkSheet.range("A" & LigneEC) = "Set g�om�trique points_de_construction"
    If GrilleActive.Exist_HB(nHBPtConst) Then EcritCheck "B" & LigneEC, "OK", objWorkSheet Else EcritCheck "B" & LigneEC, "KO", objWorkSheet
    
    LigneEC = LigneEC + 1
    
    objWorkSheet.range("A" & LigneEC) = "Set g�om�trique r�f�rences externes isol�es"
    If GrilleActive.Exist_HB(nHBRefExtIsol) Then EcritCheck "B" & LigneEC, "OK", objWorkSheet Else EcritCheck "B" & LigneEC, "KO", objWorkSheet
    
    LigneEC = LigneEC + 1
    
    objWorkSheet.range("A" & LigneEC) = "Set g�om�trique std "
    If GrilleActive.Exist_HB(nHBStd) Then EcritCheck "B" & LigneEC, "OK", objWorkSheet Else EcritCheck "B" & LigneEC, "KO", objWorkSheet
    
    LigneEC = LigneEC + 1
    
    objWorkSheet.range("A" & LigneEC) = "Set g�om�trique surf0"
    If GrilleActive.Exist_HB(nSurf0) Then EcritCheck "B" & LigneEC, "OK", objWorkSheet Else EcritCheck "B" & LigneEC, "KO", objWorkSheet
    
    LigneEC = LigneEC + 1
    
    objWorkSheet.range("A" & LigneEC) = "Set g�om�trique surf100 "
    If GrilleActive.Exist_HB(nHBS100) Then EcritCheck "B" & LigneEC, "OK", objWorkSheet Else EcritCheck "B" & LigneEC, "KO", objWorkSheet
    
    LigneEC = LigneEC + 1
    
    objWorkSheet.range("A" & LigneEC) = "Set g�om�trique travail"
    If GrilleActive.Exist_HB(nHBTrav) Then EcritCheck "B" & LigneEC, "OK", objWorkSheet Else EcritCheck "B" & LigneEC, "KO", objWorkSheet
    
    LigneEC = LigneEC + 2
    
'Test des �l�ments g�om�trique

    objWorkSheet.range("A" & LigneEC) = "Test l'existance des �l�ments g�om�triques :"
    LigneEC = LigneEC + 1
    
    objWorkSheet.range("A" & LigneEC) = "Orientation Grille"
    If GrilleActive.Exist_OrientationGrille Then EcritCheck "B" & LigneEC, "OK", objWorkSheet Else EcritCheck "B" & LigneEC, "KO", objWorkSheet
    
    LigneEC = LigneEC + 1
    
    objWorkSheet.range("A" & LigneEC) = nSurf0
    If GrilleActive.Exist_Shape(nSurf0, nHBS0) Then EcritCheck "B" & LigneEC, "OK", objWorkSheet Else EcritCheck "B" & LigneEC, "KO", objWorkSheet
    
    LigneEC = LigneEC + 1

    objWorkSheet.range("A" & LigneEC) = nSurf100
    If GrilleActive.Exist_Shape(nSurf100, nHBS100) Then EcritCheck "B" & LigneEC, "OK", objWorkSheet Else EcritCheck "B" & LigneEC, "KO", objWorkSheet
    
    LigneEC = LigneEC + 1
    
    objWorkSheet.range("A" & LigneEC) = nSurfInf
    If GrilleActive.Exist_Shape(nSurfInf, nHBTrav) Then EcritCheck "B" & LigneEC, "OK", objWorkSheet Else EcritCheck "B" & LigneEC, "KO", objWorkSheet
    
    LigneEC = LigneEC + 1
    
    objWorkSheet.range("A" & LigneEC) = nSurfSup
    If GrilleActive.Exist_Shape(nSurfSup, nHBTrav) Then EcritCheck "B" & LigneEC, "OK", objWorkSheet Else EcritCheck "B" & LigneEC, "KO", objWorkSheet
    
    LigneEC = LigneEC + 2
    
'Test de Corespondance entre les faux point A et les coordonn�es des UDF
    
    objWorkSheet.range("A" & LigneEC) = "Test de Corespondance entre les faux point A et les coordonn�es des UDF :"
    LigneEC = LigneEC + 1
    objWorkSheet.range("A" & LigneEC) = "Nom UDF"
    objWorkSheet.range("B" & LigneEC) = "X"
    objWorkSheet.range("C" & LigneEC) = "Y"
    objWorkSheet.range("D" & LigneEC) = "Z"
    objWorkSheet.range("E" & LigneEC) = "Faut Pt A"
    objWorkSheet.range("F" & LigneEC) = "X"
    objWorkSheet.range("G" & LigneEC) = "Y"
    objWorkSheet.range("H" & LigneEC) = "Z"
    objWorkSheet.range("I" & LigneEC) = "Ecart en X"
    objWorkSheet.range("J" & LigneEC) = "Ecart en Y"
    objWorkSheet.range("K" & LigneEC) = "Ecart en Z"
    LigneEC = LigneEC + 1

    'Coordonn�es des UDF
    NB_Ligne = 0
    For i = 1 To tLisfast.Count
        Set tFast = tLisfast.Item(i)
        objWorkSheet.range("A" & LigneEC) = tFast.nom
        objWorkSheet.range("B" & LigneEC) = CDbl(tFast.Xe)
        objWorkSheet.range("C" & LigneEC) = CDbl(tFast.Ye)
        objWorkSheet.range("D" & LigneEC) = CDbl(tFast.Ze)
        NB_Ligne = NB_Ligne + 1
        LigneEC = LigneEC + 1
    Next

    'Coordonn�es des faux Pt A
    LigneEC = LigneEC - NB_Ligne
    NB_Ligne = 0
    If GrilleActive.Exist_HB(nHBPtConst) Then
        For i = 1 To GrilleActive.Hb(nHBPtConst).HybridShapes.Count
            If Left(GrilleActive.Hb(nHBPtConst).HybridShapes.Item(i).Name, 6) = "faux A" Then
                objWorkSheet.range("E" & LigneEC) = GrilleActive.Hb(nHBPtConst).HybridShapes.Item(i).Name
                objWorkSheet.range("F" & LigneEC) = CDbl(GrilleActive.Hb(nHBPtConst).HybridShapes.Item(i).X.Value)
                objWorkSheet.range("G" & LigneEC) = CDbl(GrilleActive.Hb(nHBPtConst).HybridShapes.Item(i).Y.Value)
                objWorkSheet.range("H" & LigneEC) = CDbl(GrilleActive.Hb(nHBPtConst).HybridShapes.Item(i).Z.Value)
                NB_Ligne = NB_Ligne + 1
                LigneEC = LigneEC + 1
            End If
        Next
    End If
    
    'Test de l'ecart
    LigneEC = LigneEC - NB_Ligne

    For i = 1 To NB_Ligne
        XlsFormule = "=SI(ABS(B" & LigneEC & "-F" & LigneEC & ")<0,001 ; " & Chr(34) & "OK" & Chr(34) & " ; " & Chr(34) & "KO" & Chr(34) & ")"
        objWorkSheet.range("I" & LigneEC).formulalocal = XlsFormule
        XlsFormule = "=SI(ABS(C" & LigneEC & "-G" & LigneEC & ")<0,001 ; " & Chr(34) & "OK" & Chr(34) & " ; " & Chr(34) & "KO" & Chr(34) & ")"
        objWorkSheet.range("J" & LigneEC).formulalocal = XlsFormule
        XlsFormule = "=SI(ABS(D" & LigneEC & "-H" & LigneEC & ")<0,001 ; " & Chr(34) & "OK" & Chr(34) & " ; " & Chr(34) & "KO" & Chr(34) & ")"
        objWorkSheet.range("K" & LigneEC).formulalocal = XlsFormule
        LigneEC = LigneEC + 1
    Next

    'Mise en forme conditionnelle. Si l'�cart est supp�rieur � 0.2, le texte passe en rouge
    With objWorkSheet.range("I" & LigneEC - NB_Ligne & ":K" & LigneEC)
        .formatconditions.Delete
        .formatconditions.Add Type:=xLTextString, String:="KO", TextOperator:=xlContains
        .formatconditions.Add Type:=xLTextString, String:="OK", TextOperator:=xlContains
        .formatconditions(1).Font.colorindex = 3
        .formatconditions(2).Font.colorindex = 4
    End With
    LigneEC = LigneEC + 1
    
'Test de Corespondance entre les faux point B et les coordonn�es des UDF
    objWorkSheet.range("A" & LigneEC) = "Test de Corespondance entre les faux point B et les coordonn�es des UDF :"
    LigneEC = LigneEC + 1
    objWorkSheet.range("A" & LigneEC) = "Nom UDF"
    objWorkSheet.range("B" & LigneEC) = "XDir"
    objWorkSheet.range("C" & LigneEC) = "YDir"
    objWorkSheet.range("D" & LigneEC) = "ZDir"
    objWorkSheet.range("E" & LigneEC) = "Faut Pt B"
    objWorkSheet.range("F" & LigneEC) = "X"
    objWorkSheet.range("G" & LigneEC) = "Y"
    objWorkSheet.range("H" & LigneEC) = "Z"
    objWorkSheet.range("I" & LigneEC) = "Ecart en X"
    objWorkSheet.range("J" & LigneEC) = "Ecart en Y"
    objWorkSheet.range("K" & LigneEC) = "Ecart en Z"
    LigneEC = LigneEC + 1
    NB_Ligne = 0
    
    'Coordonn�es des UDF
    For i = 1 To tLisfast.Count
        Set tFast = tLisfast.Item(1)
        objWorkSheet.range("A" & LigneEC) = tFast.nom
        objWorkSheet.range("B" & LigneEC) = CDbl(tFast.Xe + 100 * tFast.Xdir)
        objWorkSheet.range("C" & LigneEC) = CDbl(tFast.Ye + 100 * tFast.Ydir)
        objWorkSheet.range("D" & LigneEC) = CDbl(tFast.Ze + 100 * tFast.Zdir)
        NB_Ligne = NB_Ligne + 1
        LigneEC = LigneEC + 1
    Next
    LigneEC = LigneEC - NB_Ligne
    
    'Coordonn�es des faux Pt A
    If GrilleActive.Exist_HB(nHBPtConst) Then
        For i = 1 To GrilleActive.Hb(nHBPtConst).HybridShapes.Count
            If Left(GrilleActive.Hb(nHBPtConst).HybridShapes.Item(i).Name, 6) = "faux B" Then
                objWorkSheet.range("E" & LigneEC) = GrilleActive.Hb(nHBPtConst).HybridShapes.Item(i).Name
                objWorkSheet.range("F" & LigneEC) = GrilleActive.Hb(nHBPtConst).HybridShapes.Item(i).X.Value
                objWorkSheet.range("G" & LigneEC) = GrilleActive.Hb(nHBPtConst).HybridShapes.Item(i).Y.Value
                objWorkSheet.range("H" & LigneEC) = GrilleActive.Hb(nHBPtConst).HybridShapes.Item(i).Z.Value
                LigneEC = LigneEC + 1
            End If
        Next
    End If
        
    'Test de l'ecart
    LigneEC = LigneEC - NB_Ligne
    'XlsFormule = "=SI(B" & LigneEC & "-F" & LigneEC & "<0,001 ; 1 ; 0 )"
    For i = 1 To NB_Ligne
        XlsFormule = "=SI(B" & LigneEC & "-F" & LigneEC & "<0,001 ; " & Chr(34) & "OK" & Chr(34) & " ; " & Chr(34) & "KO" & Chr(34) & ")"
        objWorkSheet.range("I" & LigneEC).formulalocal = XlsFormule
        XlsFormule = "=SI(C" & LigneEC & "-G" & LigneEC & "<0,001 ; " & Chr(34) & "OK" & Chr(34) & " ; " & Chr(34) & "KO" & Chr(34) & ")"
        objWorkSheet.range("J" & LigneEC).formulalocal = XlsFormule
        XlsFormule = "=SI(D" & LigneEC & "-H" & LigneEC & "<0,001 ; " & Chr(34) & "OK" & Chr(34) & " ; " & Chr(34) & "KO" & Chr(34) & ")"
        objWorkSheet.range("K" & LigneEC).formulalocal = XlsFormule
        LigneEC = LigneEC + 1
    Next
    
    'Mise en forme conditionnelle. Si l'�cart est supp�rieur � 0.2, le texte passe en rouge
    With objWorkSheet.range("I" & LigneEC - NB_Ligne & ":K" & LigneEC)
        .formatconditions.Delete
        .formatconditions.Add Type:=xLTextString, String:="KO", TextOperator:=xlContains
        .formatconditions.Add Type:=xLTextString, String:="OK", TextOperator:=xlContains
        .formatconditions(1).Font.colorindex = 3
        .formatconditions(2).Font.colorindex = 4
    End With
    LigneEC = LigneEC + 1
    
'objExcelCheck.SaveAs CheminDestRapport & NomRapportCheck

End Sub

Private Sub EcritCheck(EC_Cell, EC_Check, EC_WorkSheet)
' Ecrit le r�sultat du check et colore la cellule
' si EC_Check = "OK", ecrit "OK" et colore la cellule en vert
' sinon ecrit "KO", et colore la cellule en rouge
If EC_Check = "OK" Then
    EC_WorkSheet.range(EC_Cell) = "OK"
    FormText EC_WorkSheet, EC_Cell, "vert"
Else
    EC_WorkSheet.range(EC_Cell) = "KO"
    FormText EC_WorkSheet, EC_Cell, "rouge"
End If

End Sub
