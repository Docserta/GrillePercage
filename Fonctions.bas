Attribute VB_Name = "Fonctions"
Option Explicit

Public Sub LogUtilMacro(ByVal mPath As String, ByVal mFic As String, ByVal mMacro As String, ByVal mModule As String, ByVal mVersion As String)
'Log l'utilisation de la macro
'Ecrit une ligne dans un fichier de log sur le serveur
'mPath = localisation du fichier de log ("\\serveur\partage")
'mFic = Nom du fichier de log ("logUtilMacro.txt")
'mMacro = nom de la macro ("NomGSE")
'mVersion = Version de la macro ("version 9.1.4")
'mModule = Nom du module ("_Info_Outillage")

Dim mDate As String
Dim mUser As String
Dim nFicLog As String
Dim nLigLog As String
Const ForWriting = 2, ForAppending = 8

    mDate = Date & " " & Time()
    mUser = ReturnUserName()
    nFicLog = mPath & "\" & mFic

    nLigLog = mDate & ";" & mUser & ";" & mMacro & ";" & mModule & ";" & mVersion

    Dim fs, f
    Set fs = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    Set f = fs.GetFile(nFicLog)
    If Err.Number <> 0 Then
        Set f = fs.opentextfile(nFicLog, ForWriting, 1)
    Else
        Set f = fs.opentextfile(nFicLog, ForAppending, 1)
    End If
    
    f.Writeline nLigLog
    f.Close
    On Error GoTo 0
    
End Sub


Function ReturnUserName() As String 'extrait d'un code de Paul, Dave Peterson Exelabo
'Renvoi le user name de l'utilisateur de la station
'fonctionne avec la fonction GetUserName dans l'entète de déclaration
    Dim Buffer As String * 256
    Dim BuffLen As Long
    BuffLen = 256
    If GetUserName(Buffer, BuffLen) Then _
    ReturnUserName = Left(Buffer, BuffLen - 1)
End Function

Public Function Get_Active_CATVBA_Path() As String
Dim APC_Obj As New MSAPC.Apc
Dim TempName As String
Dim i As Long
   TempName = APC_Obj.VBE.ActiveVBProject.FileName
   For i = Len(TempName) To 1 Step -1
        If Mid(TempName, i, 1) = "\" Then
            TempName = Left(TempName, i)
            Exit For
        End If
   Next
   Get_Active_CATVBA_Path = TempName
End Function

Public Function GetDimSheet(mPaperSize As CatPaperSize) As Pos2D
    Dim tDimSheet As Pos2D
    Select Case mPaperSize
        Case catPaperA0
            tDimSheet.X = 1189
            tDimSheet.Y = 841
        Case catPaperA1
            tDimSheet.X = 841
            tDimSheet.Y = 594
        Case catPaperA2
            tDimSheet.X = 594
            tDimSheet.Y = 420
        Case catPaperA4
            tDimSheet.X = 420
            tDimSheet.Y = 297
    End Select
    GetDimSheet = tDimSheet
End Function


Public Function InsLineSpace(str As String) As String
'Remplace les saut de ligne chr(10) par 2 saut de lignes
    InsLineSpace = Replace(str, Chr(10), Chr(10) & Chr(10), 1)
    
End Function

Public Function Split_txt(str As String, pos) As String
'Coupe une chaine de carratères séparée par des points virgule pour récupérer la xieme partie
'str chaine avec PointVirgule de séparation
'Pos partie a récupérer
Dim i As Integer
    i = 1
Dim str_Temp As String
    str_Temp = str
    While i < pos
        str_Temp = Right(str_Temp, Len(str_Temp) - InStr(1, str_Temp, ";", vbTextCompare))
        i = i + 1
    Wend
    If InStr(1, str_Temp, ";", vbTextCompare) > 0 Then
        Split_txt = Left(str_Temp, InStr(1, str_Temp, ";", vbTextCompare) - 1)
    Else 'Dernier terme, il n'y as plus de ";"
        Split_txt = str_Temp
    End If
End Function

Public Function StringtoTab(strg As String) As String()
'Converti un chaine de carractères contenant des "CHR(10)" en tableau
Dim TempTab() As String
ReDim TempTab(0)
TempTab(0) = ""
Dim TempStr As String
Dim i As Long, nblig As Integer
nblig = 0
    For i = 1 To Len(strg)
        If Mid(strg, i, 1) = Chr(10) Then
            ReDim Preserve TempTab(nblig)
            TempTab(nblig) = TempStr
            TempStr = ""
            nblig = nblig + 1
        Else
            TempStr = TempStr & Mid(strg, i, 1)
        End If
    Next i
    StringtoTab = TempTab
End Function

Public Function TabtoString(tabl()) As String
'Converti un tableau en chaine de carractère avec des chr(10) apres chaque ligne
Dim i As Integer
Dim TempStr As String
    For i = 0 To UBound(tabl)
        TempStr = TempStr & tabl(i) & Chr(10)
    Next i
    TabtoString = TempStr
End Function

Public Function TestParamExist(mParams As Parameters, NomParam As String) As String
'test si le paramètre passé en argument existe dans le part.
'si oui renvoi sa valeur,
'sinon la crée et lui affecte la valeur XX
Dim mParam As StrParam
On Error Resume Next
    Set mParam = mParams.Item(NomParam)

    If (Err.Number <> 0) Then
        ' Le paramètre n'existe pas
        Err.Clear
        Set mParam = mParams.CreateString(NomParam, "XX")
    End If
    TestParamExist = mParam.Value
End Function

Public Sub CreateParamExistString(CPE_Parametres As Parameters, CPE_NomParam As String, CPE_ValueParam As String)
'Test si le paramètre de type string passé en argument existe dans la collection de paramètres
'Si oui remplace sa valeur par la valeur CPE_ValueParam
'Si non Crée le paramètre et lui affecte la valeur CPE_ValueParam
Dim CPE_Param As StrParam
On Error Resume Next
    Set CPE_Param = CPE_Parametres.Item(CPE_NomParam)
    
    If Err.Number <> 0 Then
        ' Le paramètre n'existe pas
        Err.Clear
        Set CPE_Param = CPE_Parametres.CreateString(CPE_NomParam, CPE_ValueParam)
    Else
        CPE_Param.Value = CPE_ValueParam
    End If
    
End Sub

Public Sub CreateParamExistDimension(mParams As Parameters, NomParam As String, vDbl As Double, vType As String)
'Test si le paramètre de type dimension passé en argument existe dans la collection de paramètres
'Si oui remplace sa valeur par la valeur vDbl
'Si non Crée le paramètre et lui affecte la valeur vDbl
'vType = "LENGTH" ou "ANGLE"
Dim mParam As Dimension
On Error Resume Next
    Set mParam = mParams.Item(NomParam)
    
    If Err.Number <> 0 Then
        ' Le paramètre n'existe pas
        Err.Clear
        Set mParam = mParams.CreateDimension(NomParam, vType, vDbl)
    Else
        mParam.Value = vDbl
    End If
    
End Sub

Public Function TranspositionTabl(mTable() As String) As String()
'Transposition des lignes et des colonnes du tableau
    Dim mTableTemp() As String
    Dim i As Long, j As Long
    ReDim mTableTemp(UBound(mTable, 2), UBound(mTable, 1))
    For i = 0 To UBound(mTable, 2)
        For j = 0 To UBound(mTable, 1)
            mTableTemp(i, j) = mTable(j, i)
        Next
    Next
    TranspositionTabl = mTableTemp
End Function

Public Function AjoutPart(ProdPere As Product, n_Part As String) As Product
'Ajoute une Part dans le product passé en argument
'ProdPere = product dans lesquel le nouveau part sera créé
'n_Part = nom de la nouvelle part

    Dim ProdPere_prods As Products
    Set ProdPere_prods = ProdPere.Products
       
    Set coll_docs = CATIA.Documents
    
    Dim ProdNewPart As Product
    Set ProdNewPart = ProdPere_prods.AddNewComponent("Part", n_Part)
    Set AjoutPart = ProdNewPart
End Function

Public Sub fixePart(n_Prod As String, n_Part As String)
'Fixe une part dans un product
'n_Prod = nom du document du product dans lequel un fixe le part
'n_Part = nom du document du part fixé

    Set coll_docs = CATIA.Documents
    Dim ProdDoc As ProductDocument
    Set ProdDoc = coll_docs.Item(n_Prod & ".CATProduct")
    Dim Prod As Product
    Set Prod = ProdDoc.Product
    Dim Prod_Ctrst As Constraints
    Set Prod_Ctrst = Prod.Connections("CATIAConstraints")
    Dim Prod_Ref As Reference
    Set Prod_Ref = Prod.CreateReferenceFromName(n_Prod & "/" & n_Part & ".1/!" & n_Prod & "/" & n_Part & ".1/")
    Dim Prod_Fix As Constraint
    Set Prod_Fix = Prod_Ctrst.AddMonoEltCst(catCstTypeReference, Prod_Ref)
    Prod_Fix.Name = "Fixe." & CStr(n_Part)

End Sub

Public Sub FixePart2(n_Pere As String, n_Fils As String)
'Fixe une part dans un product
'n_Pere = nom du document du product dans lequel un fixe le part
'n_fils = nom du document du part (ou product) a Fixer

    Set coll_docs = CATIA.Documents
    Dim ProdDoc As ProductDocument
    Set ProdDoc = coll_docs.Item(n_Pere & ".CATProduct")
    Dim Prod As Product
    Set Prod = ProdDoc.Product
    Dim Prod_Ctrst As Constraints
    Set Prod_Ctrst = Prod.Connections("CATIAConstraints")
    Dim Prod_Ref As Reference
    '61500V53317053020-00000X01/61500V53317055-00000Y01.1/!61500V53317053020-00000X01/61500V53317055-00000Y01.1
    Set Prod_Ref = Prod.CreateReferenceFromName(n_Pere & "/" & n_Fils & "/!" & n_Pere & "/" & n_Fils & "/")
    Dim Prod_Fix As Constraint
    Set Prod_Fix = Prod_Ctrst.AddMonoEltCst(catCstTypeReference, Prod_Ref)
    Prod_Fix.Name = "Fixe." & CStr(n_Fils)

End Sub

Public Sub FixeProdNoeud(n_Pere As String, n_Fils As String)
'Fixe un  product dans le Noeud environnement
'n_Pere = nom du product contenant le Noeud Environneemnt
'n_fils = nom du document du par (ou product) a fixer
    Dim Prd_Env As Product
    Set Prd_Env = CATIA.ActiveDocument.GetItem("env")

    'Dim prdRefProduct As Product
    'Set prdRefProduct = Prd_Env.ReferenceProduct
    
    Dim EnvContraints As Constraints
    'Set EnvContraints = prdRefProduct.Connections("CATIAConstraints")
    Set EnvContraints = Prd_Env.Connections("CATIAConstraints")
    
    Dim EnvRef As Reference
    Dim NameRef As String
    'NameRef = n_Pere & "/env.1/" & n_Fils & "/!" & n_Pere & "/env.1/" & n_Fils
    NameRef = "env/" & n_Fils & "/!env/" & n_Fils & "/"
    'Set EnvRef = prdRefProduct.CreateReferenceFromName(CStr(NameRef))
    Set EnvRef = Prd_Env.CreateReferenceFromName(CStr(NameRef))

    Dim EnvFix As Constraint
    Set EnvFix = EnvContraints.AddMonoEltCst(catCstTypeReference, EnvRef)
    EnvFix.Name = "Fixe." & CStr(n_Fils)
End Sub

Public Sub AjoutGrille(ProdPere As Product, n_GrillAss As String, DesGriAss As String, n_GrillNue As String, n_Ass As String, n_Part_U01 As String)
'Ajoute une grille Asemblée et ces composants (set géométriques et propriètès)
'ProdPere = product dans lesquel la nouvelle grille sera créee
'n_GrillAss = Nom de la grille assemblée
'desGriAss = designation de la grille assemblèe
'n_GrillNue = Nom de la grille nue
'n_Ass = Nom de l'assemblage (N° du lot)
'n_Part_U01 =  Nom de la part U01. si n_Part_U01 = "" => pas de part U01

    Dim ProdPere_prods As Products
    Set ProdPere_prods = ProdPere.Products
   
    Set coll_docs = CATIA.Documents
'    Dim objSelection As Selection
    
'Création du product Grille Assemblée
    Dim ProdGrilleAss As Product
    Set ProdGrilleAss = ProdPere_prods.AddNewComponent("Product", n_GrillAss)
    'Ajout des attributs
    Ajout_Proprietes ProdGrilleAss.ReferenceProduct, False, DesGriAss
      
'Création du Part Grille nue
    Dim ProdGrilleNue As Product
    Set ProdGrilleNue = AjoutPart(ProdGrilleAss, n_GrillNue)
    Dim GrilleNueDoc As PartDocument
    Set GrilleNueDoc = coll_docs.Item(n_GrillNue & ".CATPart")

'Fixe le part Grille nue
    FixePart2 n_GrillAss, ProdGrilleNue.Name
    
'Création des set géométriques dans le part Grille nue
    AjoutSet GrilleNueDoc, False
    Ajout_Proprietes ProdGrilleNue, True, DesGriAss
    
    'Ajout de la part U01
    If n_Part_U01 <> "" Then
        'Création du Part U01
        Dim ProdPartU01 As Product
        Set ProdPartU01 = AjoutPart(ProdGrilleAss, n_Part_U01)
        Dim PartU01Doc As PartDocument
        Set PartU01Doc = coll_docs.Item(n_Part_U01 & ".CATPart")
    
        'Fixe le part Grille nue
        FixePart2 n_GrillAss, ProdPartU01.Name
        'Ajout des set géométriques de la Part U1
        AjoutSet PartU01Doc, True

    End If
                      
    'PartGrilleNuePart.Update

End Sub

Public Sub Ajout_Proprietes(mProd, mPart, design As String)
'Ajout des paramètres dans le fichier d'assemblage
'mProd = Produit
'mPart = true si c'est une Catpart, False si c'est un Catproduct

Dim mParams As Parameters
Set mParams = mProd.ReferenceProduct.UserRefProperties
Dim ParamAdd As StrParam

    mProd.Source = catProductMade
    Set ParamAdd = mParams.CreateString("THICKNESS/DIAMETER", "")
    Set ParamAdd = mParams.CreateString("LENGTH", "")
    Set ParamAdd = mParams.CreateString("WIDTH", "")
    Set ParamAdd = mParams.CreateString("MASS", "")

    If mPart Then ' si c'est la part grille nue
        Set ParamAdd = mParams.CreateString(nPrmMaterial, ValDscgp.Mat)
        Set ParamAdd = mParams.CreateString(nPrmRecogn, vRecogn)
        mProd.DescriptionRef = vDescGNue
        Set ParamAdd = mParams.CreateString(nPrmObserv, ValDscgp.Observ)
        Set ParamAdd = mParams.CreateString(nPrmDtempl, ValDscgp.Dtemplate)
        'ces 2 attributs servent à documenter le Proces Verbal (Macro Z5_xxxxx)
        Set ParamAdd = mParams.CreateString(nPrmNumout, ValDscgp.Numout)
        Set ParamAdd = mParams.CreateString(nPrmDesign, ValDscgp.design)
        Set ParamAdd = mParams.CreateString(nPrmExempl, ValDscgp.Exemplaire)
    Else ' sinon c'est la grille ass
        Set ParamAdd = mParams.CreateString(nPrmMaterial, "")
        Set ParamAdd = mParams.CreateString(nPrmRecogn, "")
        mProd.DescriptionRef = design
        Set ParamAdd = mParams.CreateString(nPrmObserv, "")
        Set ParamAdd = mParams.CreateString(nPrmDtempl, "")
    End If
    
    Set ParamAdd = mParams.CreateString(nPrmPiecPer, ValDscgp.NumPiecesPerc)
    Set ParamAdd = mParams.CreateString(nPrmSite, ValDscgp.Site)
    Set ParamAdd = mParams.CreateString(nPrmProgAv, ValDscgp.NumProgAvion)
End Sub

Public Function IsLoadPart(coll_docs, NomPart As String) As Boolean
'Vérifie si le le nom du part passé en argument est chargé
'Coll_docs = collection des documenrs chargés
'NomPart = nom de la part recherchée format : "xxxxxxxxxxxxx.CATPart"
Dim ParTemp As PartDocument
    On Error Resume Next
    Set ParTemp = coll_docs.Item(NomPart)
    If Err.Number <> 0 Then
        Err.Clear
        IsLoadPart = False
    Else
        IsLoadPart = True
    End If
End Function
Public Function Select_PartGrille(vInt As Integer) As String
'Demande à l'utilisateur de sélectionner un product correspondant à la grille nue
'vInt = 1 pour Gauche et 2 pour Droite
Dim varfilter(0) As Variant
Dim objSel As Selection
Dim objSelLB As Object
Dim strReturn As String
Dim strMsg As String
    varfilter(0) = "Part"
    Set objSel = CATIA.ActiveDocument.Selection
    Set objSelLB = objSel
    Select Case vInt
        Case 1
            varfilter(0) = "Part"
            strMsg = "Selectionnez la grille Droite"
        Case 2
            varfilter(0) = "Part"
            strMsg = "Selectionnez la grille Gauche"
        Case 3
            varfilter(0) = "Part"
            strMsg = "Selectionnez la grille"
        Case 4
            varfilter(0) = "Part"
            strMsg = "Selectionnez la part U1"
        Case 5
            varfilter(0) = "Product"
            strMsg = "Selectionnez le product grille ass"
    End Select
    objSel.Clear
    strReturn = objSelLB.SelectElement2(varfilter, strMsg, False)
    
    If ((strReturn = "Cancel") Or (strReturn = "Undo")) Then
        Select_PartGrille = ""
    Else
        'Objet sélectionné dans l'arbre
        Select_PartGrille = objSel.Item2(1).Value.Name
    End If
'objSel.Clear
End Function

Public Sub SelectPTA(GrilleActive)
'Renvois une sélection des points a traiter

Dim tab_selection(0)
    tab_selection(0) = "HybridShape"
Dim Retour_Selection As String
    Retour_Selection = ""
Dim MsgSel As String
    MsgSel = "Sélectionnez les UDF dans la fenètre graphique ou dans le set géométrique Ref externe isolées"
    
    Retour_Selection = GrilleActive.GrilleSelection.SelectElement3(tab_selection, MsgSel, True, CATMultiSelTriggWhenUserValidatesSelection, False)

End Sub

Public Function Check_partActif() As Boolean
'Vérifie que le document actif est un Catpart

    Dim ActiveDoc As Document
    On Error Resume Next
    Set ActiveDoc = CATIA.ActiveDocument
    Dim ActivePart As part
    Set ActivePart = ActiveDoc.part
    If (Err.Number <> 0) Then ' pas de docactif de type part
        Err.Clear
        Check_partActif = False
    Else
        Check_partActif = True
    End If
    
End Function
Public Function Check_GrilleAss() As Boolean

'Vérifie que le document actif est un product de grille assemblée
'et qu'il contiens un part grille nue

'Dim instance_catpart_grille_nue As Product
Dim GrilleAss As ProductDocument
'Dim Product_GrilleAss As Product
'Dim i As Long

    Check_GrilleAss = True

    Err.Clear
    On Error Resume Next
    Set GrilleAss = CATIA.ActiveDocument
    If Err.Number <> 0 Then
        MsgBox "Le document de la fenêtre courante n'est pas un CATProduct !", vbCritical, "Environnement incorrect"
        GoTo Erreur_sortie
'    Else
'        Set Product_GrilleAss = GrilleAss.Product
    End If

'' On vérifie que qu'il y a une pièce dans l'assemblage
'    If Product_GrilleAss.Products.Count = 0 Then
'        MsgBox "Le CATProduct est vide !" & vbCrLf & "Veuillez ouvrir un product de grille Assemblée contenant une grille nue !", vbCritical, "Environnement incorrect"
'        GoTo Erreur_sortie
'    End If
'
'' Recherche de la grille nue
'    For i = 1 To Product_GrilleAss.Products.Count
'        If Left(Product_GrilleAss.Products.Item(i).Name, 10) = Left(GrilleAss.Name, 10) Then
'            Set instance_catpart_grille_nue = Product_GrilleAss.Products.Item(i)
'            Check_GrilleAss = True
'            Exit For
'        End If
'    Next

GoTo Fin
Erreur_sortie:
Check_GrilleAss = False
Fin:
End Function


Function Check_PtExist(HBody, PtName As String) As Boolean
'Vérifie si les points existe deja dans le set géométrique passé en argument, empeche de recreer des points deja existants,
'elle est appele lors de la creation des points A, B
'hBody set géographique dans lequel on recherche le point
'PtName Nom du point recherché

Dim CPE_hybridShapes As HybridShapes
Set CPE_hybridShapes = HBody.HybridShapes
Dim Point_Name As String
Dim k As Integer
Dim CPE_Present As Boolean
    CPE_Present = False
    
    For k = 1 To CPE_hybridShapes.Count
        If (CPE_hybridShapes.Count > 0) Then
            Point_Name = CPE_hybridShapes.Item(k).Name
            If (StrComp(Point_Name, PtName, vbTextCompare) = 0) Then
                CPE_Present = True 'le point existe deja
                Exit For
            End If
        End If
    Next
    Check_PtExist = CPE_Present
End Function

Function Create_PtCoord(Xe As Double, Ye As Double, Ze As Double, PtName As String, GrilleActive) As HybridShapePointCoord
'Création d'un point au coordonnée Xe, Ye, Ze de nom PtName
'dans le set la part "Points de construction" de la grille
'Si le point existe, renvoi le point, sinon le crée

Dim hybridShapePointCoord1 As HybridShapePointCoord
Dim PtExist As Boolean
    PtExist = False
Dim k As Integer
Dim ExistPtName As String

'Vérification si le point existe
Dim CPE_hybridShapes As HybridShapes
Set CPE_hybridShapes = GrilleActive.Hb(nHBPtConst).HybridShapes

    For k = 1 To CPE_hybridShapes.Count
        If (CPE_hybridShapes.Count > 0) Then
            ExistPtName = CPE_hybridShapes.Item(k).Name
            If (StrComp(PtName, ExistPtName, vbTextCompare) = 0) Then
                PtExist = True 'le point existe deja
                Set hybridShapePointCoord1 = CPE_hybridShapes.Item(k)
                Exit For
            End If
        End If
    Next
    If Not PtExist Then
        Set hybridShapePointCoord1 = GrilleActive.HShapeFactory.AddNewPointCoord(Xe, Ye, Ze)
        GrilleActive.Hb(nHBPtConst).AppendHybridShape hybridShapePointCoord1
        hybridShapePointCoord1.Name = PtName
    End If

    Set Create_PtCoord = hybridShapePointCoord1
End Function

Function Create_Line_PtPt(Pt1, Pt2, ByRef GrilleActive, PtName) As Boolean
'Crée une ligne entre les pts Pt1 et Pt2
'Active le set géométrique "std" dans le part actif
Dim CLPP_hybridShapeLinePtPt As HybridShapeLinePtPt
    GrilleActive.PartGrille.InWorkObject = GrilleActive.Hb(nHBStd)
    Set CLPP_hybridShapeLinePtPt = GrilleActive.HShapeFactory.AddNewLinePtPtExtended(Pt1, Pt2, 110, 110)
    GrilleActive.Hb(nHBStd).AppendHybridShape CLPP_hybridShapeLinePtPt
    CLPP_hybridShapeLinePtPt.Name = PtName
    GrilleActive.PartGrille.InWorkObject = CLPP_hybridShapeLinePtPt
    Create_Line_PtPt = True
End Function

Public Function ChangeSingle(str) As Single
'Converti la valeur passée en argument en type single
'Remplace le point par une virgule ou l'inverse (fonction du paramètre régional Windows)
On Error Resume Next
Dim CS_temp As Single
    CS_temp = CSng(str)
    If (Err.Number <> 0) Then
        Err.Clear
        If InStr(str, ".") <> 0 Then
            ChangeSingle = Replace(str, ".", ",")
        ElseIf InStr(str, ",") <> 0 Then
            ChangeSingle = Replace(str, ",", ".")
        End If
    Else
        ChangeSingle = CS_temp
    End If
End Function

Function IsUpdatable(mPart, mObj As Variant) As Boolean
On Error Resume Next
    mPart.UpdateObject mObj
    If (Err.Number <> 0) Then
        Err.Clear
        IsUpdatable = False
    Else
        IsUpdatable = True
    End If
End Function

Public Function EffaceFicNom(mFolder, FicNom) As Boolean
'Effacement d'un fichier de excel pré-existant
On Error GoTo Err_EffaceFicNom
Dim EF_FS, EF_Fold, EF_Files, EF_File
    Set EF_FS = CreateObject("Scripting.FileSystemObject")
    Set EF_Fold = EF_FS.GetFolder(mFolder)
    Set EF_Files = EF_Fold.Files
    For Each EF_File In EF_Files
        If EF_File.Name = FicNom Then
            EF_FS.DeleteFile (CStr(mFolder & "\" & FicNom))
        End If
    Next
    EffaceFicNom = True
    GoTo Quit_EffaceFicNom

Err_EffaceFicNom:
    MsgBox "Il est possible que le fichier de rapport soit encore ouvert dans Excel. Veuillez le fermer et relancer la macro.", vbCritical, "erreur"
EffaceFicNom = False
Quit_EffaceFicNom:
End Function

Public Function SignePlusMoins(V1, V2) As Double
'Renvoi -1 si V1 est inférieur à V2
'Renvoi +1 si V1 est suppérieur ou egal à V2
    If V1 >= V2 Then
        SignePlusMoins = 1
    Else
        SignePlusMoins = -1
    End If
End Function

Public Function AddLigneReport(str, mTab) As String()
'Ajoute la ligne str au tableau mTab et renvois le tableau
    ReDim Preserve mTab(UBound(mTab, 1) + 1)
    mTab(UBound(mTab, 1)) = str
    AddLigneReport = mTab
End Function

Public Function DecoupeSlash(str As String) As String
'Recupère la partie finale d'une string apres le dernier "\"
Do While InStr(1, str, "\", vbTextCompare) > 0
    str = Right(str, Len(str) - InStr(1, str, "\", vbTextCompare))
Loop
DecoupeSlash = str
End Function

Public Sub BorduresCell(wSheet, cel1, cel2)
'Trace une bordure autour des cellules de la plage cel1:cel2

    wSheet.range(cel1, cel2).Borders.LineStyle = 1
    wSheet.range(cel1, cel2).Borders.Weight = xLMoyen
End Sub

Public Sub CouleurCell(wSheet, cel1, cel2, Coul As String)
'Colorie la plage de cellules cel1:cel2 dans la couleur passée en argument
Dim lgCol As Long
    If Coul = "gris" Then
        lgCol = 11842740 'Gris
    ElseIf Coul = "jaune" Then
        lgCol = 65535 'Jaune
    ElseIf Coul = "vert" Then
        lgCol = 5296274 'Vert
    Else
        lgCol = 0
    End If
    
    With wSheet.range(cel1, cel2).Interior
        .Color = lgCol
    End With
End Sub

Public Sub FormText(wSheet, Cell, mtype As String)
'Change le format du texte de  la cellule passée en argument
Dim FT_Color As Long
Dim FT_Size As Long
If mtype = "Titre" Then
    FT_Size = 13
    FT_Color = 0
ElseIf mtype = "vert" Then
    FT_Size = 10
    FT_Color = -11489280
ElseIf mtype = "rouge" Then
    FT_Size = 10
    FT_Color = -16776961
Else
    FT_Size = 10
    FT_Color = 0
End If
    
With wSheet.range(Cell).Font
    .Name = "Arial"
    .Size = FT_Size
    .Color = FT_Color
End With
With wSheet.range(Cell).cells
    .HorizontalAlignment = xLCenter
    .VerticalAlignment = xLHaut
    .WrapText = True
End With
End Sub

Public Function CorrigeDFS() As String
'Corrige une erreur du DFS. a savoir que sur le site de Xsn,
'La Bibli est nommé : W:\50-PRJ Grilles
'alors que pour les autres sites, elle est nommé :W:\50 - PRJ Grilles
Dim FileSystem As New Navigateur1
Dim ListRepBibli
Dim RepBibli

ListRepBibli = FileSystem.ListeRep("W:\")
For Each RepBibli In ListRepBibli
    If Left(RepBibli, 2) = "50" Then
        CorrigeDFS = "W:\" & RepBibli
    End If
Next

End Function

Public Sub WriteLog(Contenu, DestPath, NomFicLog)
'Ecriture du log
'Ecrit dans un fichier texte le contenu du tableau passé en argument
'Contenu = tableau de string 2 dimensions
'DestPath = path de destination du fichier log
'NomFicLog = Nom du fichier de log
Dim cfdate As String
    cfdate = Date & Time()
    cfdate = Replace(cfdate, "/", "")
    cfdate = Replace(cfdate, ":", "")
NomFicLog = DestPath & NomFicLog & "_" & cfdate & "_log.txt"

Dim i As Integer, j As Integer
        
Dim LigEncours As String
Dim fs, f
        Set fs = CreateObject("scripting.filesystemobject")
        Set f = fs.CreateTextFile(NomFicLog, True)
        For i = 0 To UBound(Contenu)
                LigEncours = Contenu(i)
            f.Writeline (LigEncours)
        Next
        f.Close
End Sub

Public Function FileExist(StrFile As String) As Boolean
'Teste si le répertoire existe et revois vrai ou faux
'StrFile = nom du chemin complet jusqu'au répertoire a tester ex "c:\temp\test"
Dim fs, f
    Set fs = CreateObject("scripting.filesystemobject")
    On Error Resume Next
    Set f = fs.GetFile(StrFile)
    If Err.Number <> 0 Then
        Err.Clear
        FileExist = False
    Else
        FileExist = True
    End If
End Function


Public Function FldExist(Fld As String) As Boolean
'Teste si le répertoire existe et revois vrai ou faux
'Fld = nom du chemin complet jusqu'au répertoire a tester ex "c:\temp\test"
Dim fs, fd
    Set fs = CreateObject("scripting.filesystemobject")
    On Error Resume Next
    Set fd = fs.GetFolder(Fld)
    If Err.Number <> 0 Then
        Err.Clear
        FldExist = False
    Else
        FldExist = True
    End If
End Function

Function GetPath(Titre As String) As String
    Dim ObjShell As Object, ObjFolder As Object
    Set ObjShell = CreateObject("shell.Application")
    Set ObjFolder = ObjShell.BrowseForFolder(0, Titre, 0)
    If (Not ObjFolder Is Nothing) Then
        GetPath = ObjFolder.Items.Item.Path
    End If
    Set ObjFolder = Nothing
    Set ObjShell = Nothing
End Function

Public Function CreatFld(Fld As String, Path As String) As Boolean
'Crée un répertoire dans le dossier passé en argument
'Path = repertoire dans lequel doit être créé le nouveau répertoire
'Fld = Nom du répertoire a créer
Dim fs, fd
    Set fs = CreateObject("scripting.filesystemobject")
    On Error Resume Next
    Set fd = fs.CreateFolder(Path & "\" & Fld)
    If Err.Number <> 0 Then
        Err.Clear
        CreatFld = False
    Else
        CreatFld = True
    End If
End Function

Public Sub AjoutSet(mDoc As PartDocument, isU1 As Boolean)
'Création des set géométriques dans le part
'mDoc = part document
'isU1 = true si c'est une part de controle (U1) seuls certains set sont créé dans cette parts
    Dim mPart As part
    Set mPart = mDoc.part
    Dim mBodies As HybridBodies, mSubBodies As HybridBodies
    Set mBodies = mPart.HybridBodies
    'Set Géométrique Niveau 1
    Dim mBody As HybridBody
    If (Not isU1) Then
        Set mBody = mBodies.Add()
        mBody.Name = nHBRefExtIsol
    End If
    Set mBody = mBodies.Add()
    mBody.Name = nHBS0
    Set mBody = mBodies.Add()
    mBody.Name = nHBPin
    Set mBody = mBodies.Add()
    mBody.Name = nHBFeet
    Set mBody = mBodies.Add()
    mBody.Name = nHBPtA
    Set mBody = mBodies.Add()
    mBody.Name = nHBStd
    Set mBody = mBodies.Add()
    mBody.Name = nHBS100
    Set mBody = mBodies.Add()
    mBody.Name = nHBPtB
    If (Not isU1) Then
        Set mBody = mBodies.Add()
        mBody.Name = nHBGrav
        Set mBody = mBodies.Add()
        mBody.Name = nHBPtConst
        Set mBody = mBodies.Add()
        mBody.Name = nHBTrav
        'Set Géométrique Niveau 2
        Dim mSubBody As HybridBody
        Set mSubBodies = mPart.HybridBodies.Item(nHBTrav)
        Set mSubBody = mSubBodies.Add()
        mSubBody.Name = nHBGeoRef
        Set mSubBody = mSubBodies.Add()
        mSubBody.Name = nHBSDetr
        Set mSubBody = mSubBodies.Add()
        mSubBody.Name = nHBDrFeet
        Set mSubBody = mSubBodies.Add()
        mSubBody.Name = nHBDrPin
        Set mSubBody = mSubBodies.Add()
        mSubBody.Name = nHBDrGrav
    End If
    
    mPart.Update
End Sub

Public Sub Ajout1Set(mDoc As PartDocument, NomSet As String)
'Création d'un set géométrique dans le part
'mDoc = part document
'NomSet = Nom du set a créer
    Dim mPart As part
    Set mPart = mDoc.part
    Dim mBodies As HybridBodies
    Set mBodies = mPart.HybridBodies
    Dim mBody As HybridBody
    Set mBody = mBodies.Add()
    mBody.Name = NomSet

End Sub

 Public Function Interv(ByVal str As String, pos As Integer, nbcar As Integer, inf As Integer, sup As Integer) As Boolean
'Renvoi true si la valeur de str (convertie en string) est comprise entre "inf" et "sup"
Dim IntTemp As Long
On Error Resume Next
    IntTemp = CLng(Mid(str, pos, nbcar))
    If Err.Number <> 0 Then
        Err.Clear
        Interv = False
        On Error GoTo 0
    Else
        If IntTemp >= inf And IntTemp <= sup Then
            Interv = True
        Else
            Interv = False
        End If
    End If
End Function

Public Function NumCar(num As Integer) As String
'Converti un chiffre en lettre
'1 = A, 2 = B etc
'Attention la numérotation de Array commence à 0 d'ou le double A dans la liste
Dim ListCar
If num > 78 Then ' a changer si on ajoute des colonnes a la liste Array
    num = 1
End If
ListCar = Array("A", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", _
                "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ", _
                "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ", "BK", "BL", "BM", "BN", "BO", "BP", "BQ", "BR", "BS", "BT", "BU", "BV", "BW", "BX", "BY", "BZ")
NumCar = ListCar(num)
End Function

Public Function MatBody(mPart As part, mBody As Body) As Material
'Renvoi la matière appliquée au part body
    
    Dim oManager As MaterialManager
    Dim oAppliedMaterial As Material
    Set oManager = mPart.GetItem("CATMatManagerVBExt")

    oManager.GetMaterialOnBody mBody, oAppliedMaterial
    Set MatBody = oAppliedMaterial
End Function

Public Function Matpart(mPart As part) As Material
'Renvoi la matière appliquée au part
    
    Dim oManager As MaterialManager
    Dim oAppliedMaterial As Material
    Set oManager = mPart.GetItem("CATMatManagerVBExt")

    oManager.GetMaterialOnPart mPart, oAppliedMaterial
    Set Matpart = oAppliedMaterial
End Function

Public Function MassPart(mProd As Product) As Double
'Renvoi la masse du product du part
    MassPart = 0
    Dim oMass As Double
    Dim oInertia As Inertia
    Set oInertia = mProd.GetTechnologicalObject("Inertia")
    oMass = oInertia.Mass
    MassPart = oMass
End Function

Public Function DistMat(CoPt1 As c_Coord, CoPt2 As c_Coord) As Double
'Renvoi la distance entre deux points de coordonées x, y, z
Dim Dx As Double, DY As Double, Dz As Double
    Dx = (CoPt1.X - CoPt2.X) ^ 2
    DY = (CoPt1.Y - CoPt2.Y) ^ 2
    DY = (CoPt1.Z - CoPt2.Z) ^ 2
    DistMat = Round(Sqr(Dx + DY + DY), 3)
End Function

Public Function CoordPt(tpart As part, tHSEC As HybridShape) As c_Coord
'Renvoi la coordonnées X, Y ou Z mesurée du point passé en argument
Dim spa_workbench As SPAWorkbench
Dim mMes 'as Measurable
Dim mRef As Reference
Dim Pt(2)
Dim tCoordPt As c_Coord
Set tCoordPt = New c_Coord
    Set spa_workbench = CATIA.ActiveDocument.GetWorkbench("SPAWorkbench")
    Set mRef = tpart.CreateReferenceFromObject(tHSEC)
    Set mMes = spa_workbench.GetMeasurable(mRef)
    mMes.GetPoint Pt
    With tCoordPt
        .X = Pt(0)
        .Y = Pt(1)
        .Z = Pt(2)
    End With
    Set CoordPt = tCoordPt
End Function

Public Function ReadXlsBagues(ficXls As String) As c_DefBagues
'Construit la collection des bagues spécifiques
Dim oBague As c_DefBague
Dim oBagues As c_DefBagues
Dim objexcel
Dim objWorkBook
Dim objWorkSheet
Dim LigEC As Long

    'initialisation des classes
    Set oBagues = New c_DefBagues
    Set objexcel = CreateObject("EXCEL.APPLICATION")
    Set objWorkBook = objexcel.Workbooks.Open(ficXls, True, True)
    Set objWorkSheet = objWorkBook.Sheets("Catalogue PRS01")
       LigEC = 2
    'lecture du fichier excel
    While objWorkSheet.cells(LigEC, 1) <> ""
        Set oBague = New c_DefBague
        oBague.Ref = objWorkSheet.cells(LigEC, 1) ' Reference
        oBague.DNom = objWorkSheet.cells(LigEC, 2) ' Diam Nominal
        oBague.Mat = objWorkSheet.cells(LigEC, 3) ' Matière
        oBague.D1 = objWorkSheet.cells(LigEC, 4) 'D1
        oBague.D2 = objWorkSheet.cells(LigEC, 5) 'D2
        oBague.D3 = objWorkSheet.cells(LigEC, 6) 'D3
        oBague.L1 = objWorkSheet.cells(LigEC, 7) 'L1
        oBague.L2 = objWorkSheet.cells(LigEC, 8) 'L2
        oBague.NomFic = objWorkSheet.cells(LigEC, 9) 'Nom du Catpart
        LigEC = LigEC + 1
        oBagues.Add oBague.Ref, oBague.DNom, oBague.Mat, oBague.D1, oBague.D2, oBague.D3, oBague.L1, oBague.L2, oBague.NomFic
        Set oBague = Nothing
    Wend

    Set ReadXlsBagues = oBagues

'libération des classes
Set oBague = Nothing
Set oBagues = Nothing
Set objexcel = Nothing

End Function
