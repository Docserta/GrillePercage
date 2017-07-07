Attribute VB_Name = "A2_CreationLot"

Option Explicit
'*********************************************************************
'* Macro : A2_CreationLot
'*
'* Fonctions :  Création des ensembles product grille assemblée et part grille
' *             Pour un lot complet defini par la liste des DSCGP contenu dans un répertoire
'*              Crée le product général, le product grille ass, la part grille nue
'*              Sélectionne et importe l'environnement avion
'*              Crée les set géométriques et les contrainte de fixation
'*              Ajoute des attributs provenant d'un fichier excel
'*
'* Version : 9
'* Création :  CFR
'* Modification : 26/02/16
'*
'**********************************************************************


Sub CATMain()
On Error GoTo Err_CreationLot

'Log de l'utilisation de la macro
LogUtilMacro nPath, nFicLog, nMacro, "A2_CreationLot", VMacro

Dim ListeDscgp() As String
    
Dim IndSelection As Long 'indice qui permet de parcourir la listbox de Frm_Donnees
Dim i As Integer, j As Integer
Dim DSCGP_EC As c_DSCGP
Dim RepCible As String 'Chemin du dossier dans lequel on crée les répertoire
Dim RepCibleLot As String 'Chemin du dossier du lot de grilles
Dim RepCibleGriAss  As String 'Chemin du dossier de la grille assemblée en cours de création (sous répertoire de RepCibleLot)
Dim RepEnvAvion As String ' Chemin + nom de l'environnement avion
Dim nRepGriAss As String 'Nom du répertoire de la grille en cours de création (sous répertoire de RepCibleLot)
Dim DocGrilleAss As ProductDocument 'Prduct de la grille assemblée créé pour sauvegarde
Dim LstGrilleLot() As String 'Liste des grilles assemblées pour assemblage du lot final
Dim ProdGrilleAss As Product

Dim cas As Integer
Dim nGriAss1 As String, _
    nGriNue1 As String, _
    nGriAss2 As String, _
    nGriNue2 As String
Dim DesignGri1 As String, DesignGri2 As String

'Dim nPreGriNue1 As String, _
'    nPreGriNue2 As String

IndSelection = 0
ReDim ListeDscgp(0)
ReDim Preserve ReportLog(0)
ReDim LstGrilleLot(0)

    'Chargement et configuration du formulaire
    Load Frm_ListeFichiers
    Frm_ListeFichiers.Caption = "Création par lot"
    Frm_ListeFichiers.Tbx_Extent = "*.xls"
    Frm_ListeFichiers.CB_Catpartactif.Visible = False
    Frm_ListeFichiers.Lbl_Extent.Visible = False
    Frm_ListeFichiers.Show
    
 'Sort du programme si click sur bouton Annuler du formulaire
    If Not (Frm_ListeFichiers.ChB_OkAnnule) Then
        Unload Frm_ListeFichiers
        Exit Sub
    End If
'Collecte des infos du formulaire
RepCible = Frm_ListeFichiers.TBX_FicDest
RepEnvAvion = Frm_ListeFichiers.TBX_EnvAvion
   
   'Initialisation du log
   ReDim Preserve ReportLog(UBound(ReportLog) + 4)
   ReportLog(UBound(ReportLog) - 4) = "##########################################################"
   ReportLog(UBound(ReportLog) - 3) = " Création du lot de grille : " & CheminFicLot
   ReportLog(UBound(ReportLog) - 2) = "Dossier crée par : " & ReturnUserName & " Le : " & Date
   ReportLog(UBound(ReportLog) - 1) = "##########################################################"
   
   'Collecte de la liste de fichiers DSCGP a traiter
    ReportLog(UBound(ReportLog)) = "Collecte de la liste de fichiers DSCGP a traiter"
    For i = 0 To Frm_ListeFichiers.ListBox1.ListCount - 1
        'Boucle sur la liste des fichiers et test si le fichier est sélectionné
        If Frm_ListeFichiers.ListBox1.Selected(i) Then
            'Maj Log
            ReDim Preserve ListeDscgp(UBound(ListeDscgp) + 1)
            ReDim Preserve ReportLog(UBound(ReportLog) + 1)
            ListeDscgp(UBound(ListeDscgp)) = Frm_ListeFichiers.ListBox1.List(i)
            ReportLog(UBound(ReportLog)) = ListeDscgp(UBound(ListeDscgp))
            IndSelection = IndSelection + 1
        End If
    Next i
    
    'Sortie si pas de fichier sélectionné
    If IndSelection = 0 Then
        MsgBox "Pas de fichier sélectionné!", vbInformation, "Pas de fichier sélectionné"
        Exit Sub
    End If
    
    Set DSCGP_EC = New c_DSCGP
    For i = 1 To UBound(ListeDscgp, 1) 'dans le tableau l'index 0 est vide
        
        DSCGP_EC.VersionDscgp = 2
        DSCGP_EC.OpenDSCGP = CheminFicLot & ListeDscgp(i)
        
        'Création du dossier du lot de grilles a la première grille
        If DSCGP_EC.NumLot = "" Then
            MsgBox "Le Numero du lot n'est pas renseigné dans le DSCGP : " & ListeDscgp(i), vbCritical, "Erreur de DSCGP"
            End
        End If
        If Not (FldExist(RepCible & "\" & DSCGP_EC.NumLot)) Then
            RepCibleLot = RepCible & "\" & DSCGP_EC.NumLot
            If CreatFld(DSCGP_EC.NumLot, RepCible) Then
                'Maj Log
                ReDim Preserve ReportLog(UBound(ReportLog) + 2)
                ReportLog(UBound(ReportLog) - 1) = "----------------------"
                ReportLog(UBound(ReportLog)) = "Création du répertoire : " & DSCGP_EC.NumLot & " du lot de grille dans : " & RepCible
            Else 'erreur de création de dossier
                ReDim Preserve ReportLog(UBound(ReportLog) + 1)
                'Maj Log
                ReportLog(UBound(ReportLog)) = "Echeck de la création du répertoire ( " & DSCGP_EC.NumduLot & " )du lot de grille dans : " & RepCibleLot
                WriteLog ReportLog, RepCibleLot & "\", "MacroCreationLot"
            End If
            'Ajout de l'environnement à la liste des fichiers a remonter dans l'assemblage général
            LstGrilleLot(0) = RepEnvAvion
        End If
        
        
        'Création du dossier de la grille assemblée
        'Maj Log
        ReDim Preserve ReportLog(UBound(ReportLog) + 3)
        ReportLog(UBound(ReportLog) - 1) = "Traitement de la grille du DSCGP : " & ListeDscgp(i)
        ReportLog(UBound(ReportLog)) = "----------------------------------------------------"
                              
        'Collecte des paramètres
        ValDscgp.CoteAvion = DSCGP_EC.CoteAvion
        ValDscgp.Mat = DSCGP_EC.MatGrille
        ValDscgp.Observ = DSCGP_EC.Observations
        ValDscgp.Dtemplate = DSCGP_EC.Dtemplate
        ValDscgp.Numout = DSCGP_EC.NumOutillage
        ValDscgp.Exemplaire = DSCGP_EC.Exemplaire
        ValDscgp.NumPiecesPerc = DSCGP_EC.PiecesPercees
        ValDscgp.Site = DSCGP_EC.Site
        ValDscgp.NumProgAvion = DSCGP_EC.NoProgAvion
        
        nGriAss1 = DSCGP_EC.NumGrille
        nGriNue1 = DSCGP_EC.NumGrilleNue
        DesignGri1 = DSCGP_EC.DesignGrille
        nGriAss2 = DSCGP_EC.NumGrilleSym
        nGriNue2 = DSCGP_EC.NumGrilleSymNue
        DesignGri2 = DSCGP_EC.DesignGrilleSym
        nRepGriAss = DSCGP_EC.NumRadGrille
        RepCibleGriAss = RepCibleLot & "\" & DSCGP_EC.NumRadGrille
        
        'Calcul des cas
        Select Case ValDscgp.CoteAvion
            Case "GAUCHE"
                If DSCGP_EC.NumGrille = "" Then
                    cas = 0 'Erreur N° grille vide
                ElseIf DSCGP_EC.NumGrille <> "" And DSCGP_EC.NumGrilleSym = "" Then
                    cas = 1 'Cas =1 => Grille gauche seule
                    nGriAss2 = ""
                    nGriNue2 = ""
                    DesignGri2 = ""
                End If
            Case "DROIT"
                If DSCGP_EC.NumGrille = "" Then
                    cas = 0 'Erreur N° grille vide
                ElseIf DSCGP_EC.NumGrille <> "" And DSCGP_EC.NumGrilleSym = "" Then
                    cas = 3 'Cas =3 => Grille droite seule
                    nGriAss2 = ""
                    nGriNue2 = ""
                    DesignGri2 = ""
                ElseIf DSCGP_EC.NumGrille <> "" And DSCGP_EC.NumGrilleSym <> "" Then
                    cas = 4  'Cas =4 => Grille droite + sym gauche
                    'Inversion des nom de grille
                    nGriAss1 = DSCGP_EC.NumGrilleSym
                    nGriNue1 = DSCGP_EC.NumGrilleSymNue
                    DesignGri1 = DSCGP_EC.DesignGrilleSym
                    nGriAss2 = DSCGP_EC.NumGrille
                    nGriNue2 = DSCGP_EC.NumGrilleNue
                    DesignGri2 = DSCGP_EC.DesignGrille
                End If
            Case "CENTRE"
                If DSCGP_EC.NumGrille = "" Then
                    cas = 0 'Erreur N° grille vide
                Else
                    cas = 5  'Cas =5 => Grille gauche seule
                    nGriAss2 = ""
                    nGriNue2 = ""
                    DesignGri2 = ""
                End If
            Case Else
                cas = 0 'Erreur c0té avion
        End Select
        
        If cas = 0 Then
            ReDim Preserve ReportLog(UBound(ReportLog) + 1)
            ReportLog(UBound(ReportLog)) = "     Ereure détectée dans le DSCGP : " & ListeDscgp(i) & " les numéros des grilles ou le du coté de comnception ne sont pas cohérent "
        Else
            If Not (FldExist(RepCibleGriAss)) Then
                'Maj Log
                ReDim Preserve ReportLog(UBound(ReportLog) + 4)
                ReportLog(UBound(ReportLog) - 4) = "     N° Grille ass1 : " & nGriAss1
                ReportLog(UBound(ReportLog) - 3) = "     N° Grille nue1 : " & nGriNue1
                ReportLog(UBound(ReportLog) - 2) = "     N° Grille ass2 : " & nGriAss2
                ReportLog(UBound(ReportLog) - 1) = "     N° Grille nue2 : " & nGriNue2
                'Création du répertoire de la grille assemblée
                If CreatFld(nRepGriAss, RepCibleLot) Then
                    ReDim Preserve ReportLog(UBound(ReportLog) + 1)
                    ReportLog(UBound(ReportLog)) = "     Création du répertoire ( " & nRepGriAss & " )de la grille assemblée dans : " & RepCibleLot
                    
                    'Création du product grille assemblée
                    'Maj Log
                    ReDim Preserve ReportLog(UBound(ReportLog) + 2)
                    ReportLog(UBound(ReportLog) - 2) = "     Création de la grille principale : " & nGriAss1
                    ReportLog(UBound(ReportLog) - 1) = "        Contenant la grille Nue : " & nGriNue1
                    
                    CreateCAO DSCGP_EC.NumduLot, nGriAss1, DesignGri1, nGriNue1, RepCibleLot, RepEnvAvion, DSCGP_EC.NumPartU01, Frm_ListeFichiers.TBX_NomDtromp
                    
                    ReDim Preserve ReportLog(UBound(ReportLog) + 1)
                    
                    Set DocGrilleAss = CATIA.ActiveDocument
                    Set ProdGrilleAss = DocGrilleAss.Product
                    
                    'Création de la grille sym
                    If nGriAss2 <> "" And nGriNue2 <> "" Then
                        ValDscgp.design = DSCGP_EC.DesignGrilleSym
                        'Maj Log
                        ReDim Preserve ReportLog(UBound(ReportLog) + 2)
                        ReportLog(UBound(ReportLog) - 2) = "     Ajout de la grille Sym : " & nGriAss2
                        ReportLog(UBound(ReportLog) - 1) = "        Contenant la grille Nue : " & nGriNue2
                        AjoutGrille ProdGrilleAss, nGriAss2, DesignGri2, nGriNue2, DSCGP_EC.NumduLot, DSCGP_EC.NumPartU01Sym
                        'Fixe le product Grille Assemblée
                        For j = 1 To ProdGrilleAss.Products.Count
                            If InStr(1, ProdGrilleAss.Products.Item(j).Name, nGriAss2, vbTextCompare) <> 0 Then
                                'Maj Log
                                ReDim Preserve ReportLog(UBound(ReportLog) + 2)
                                ReportLog(UBound(ReportLog) - 2) = "     Ajout des contraintes"
                                FixePart2 DSCGP_EC.NumduLot, ProdGrilleAss.Products.Item(j).Name
                            End If
                        Next j
                    End If
                    
                    'sauvegarde de la grille assemblée
                    CATIA.DisplayFileAlerts = False
                    DocGrilleAss.SaveAs (RepCibleGriAss & "\" & DSCGP_EC.NumduLot)
                    CATIA.DisplayFileAlerts = True
                    ReDim Preserve ReportLog(UBound(ReportLog) + 1)
                    ReportLog(UBound(ReportLog)) = "     Sauvegarde de la grille assemblée  : " & DSCGP_EC.NumduLot
                    DocGrilleAss.Close
                
                    'Ajout de la grille a la liste du lot
                    If LstGrilleLot(0) <> "" Then 'on ne redimentionne le tableau que si on a déja rempli l'index 0
                        ReDim Preserve LstGrilleLot(UBound(LstGrilleLot) + 1)
                    End If
                    LstGrilleLot(UBound(LstGrilleLot)) = RepCibleGriAss & "\" & nGriAss1 & ".CATProduct"
                    'Ajout de la grille sym a la liste du lot
                    If nGriAss2 <> "" And nGriNue2 <> "" Then
                        ReDim Preserve LstGrilleLot(UBound(LstGrilleLot) + 1)
                        LstGrilleLot(UBound(LstGrilleLot)) = RepCibleGriAss & "\" & nGriAss2 & ".CATProduct"
                    End If
                                 
                Else 'erreur de création de dossier
                    'Maj Log
                    ReDim Preserve ReportLog(UBound(ReportLog) + 1)
                    ReportLog(UBound(ReportLog)) = "Erreur de création du répertoire : " & RepCibleLot & ListeDscgp(i)
                    WriteLog ReportLog, RepCibleLot & "\", "MacroCreationLot"
                End If
            Else
                MsgBox "Ce répertoire (" & RepCibleGriAss & "\" & nRepGriAss & ") existe déjà. Effacez le ou vérifiez le chemin de sauvegarde du lot de grilles.", vbCritical, "Erreur"
                End
            End If
        End If
    Next i
    
    'Création du product du lot
    If LstGrilleLot(0) <> "" Then
        Assemblage_Lot LstGrilleLot, RepCibleLot & "\", DSCGP_EC.NumduLot & ".CATProduct", Frm_ListeFichiers.TBX_NomDtromp
        ReDim Preserve ReportLog(UBound(ReportLog) + 2)
        ReportLog(UBound(ReportLog)) = "Sauvegarde du fichier de remontage du lot de grille : " & DSCGP_EC.NumduLot & " dans : " & RepCibleLot
    End If
    
GoTo FinOK

Err_CreationLot:
    ReDim Preserve ReportLog(UBound(ReportLog) + 3)
    ReportLog(UBound(ReportLog) - 2) = "##############################################"
    ReportLog(UBound(ReportLog) - 1) = "# Erreure détectée lors de la création du lot #"
    ReportLog(UBound(ReportLog)) = "##############################################"
    WriteLog ReportLog, RepCibleLot & "\", "MacroCreationLot"
    MsgBox "Erreur détectée, vérifier les DSCGP !", vbInformation, "Fin de traitement"
    GoTo FinKO
    
FinOK:
    ReDim Preserve ReportLog(UBound(ReportLog) + 2)
    ReportLog(UBound(ReportLog)) = "Fin de création du lot de grille pas d'erreur détectée"
    WriteLog ReportLog, RepCibleLot & "\", "MacroCreationLot"
    MsgBox "Fin de traitement des fichiers.", vbInformation, "Fin de traitement"
    
FinKO:
    
    Unload Frm_ListeFichiers

End Sub

Private Sub Assemblage_Lot(LstGrilleLot As Variant, ChemCible As String, NomCible As String, NomDet As String)
'remonte l'ensemble des gilles assemblées dans un Catproduct
'List_Grille = liste des Catproduct a remonter sous la forme "Lect:\path\Nomfichier.extension"
'ChemCible = Nom du fichier Catproduct du catproduct "lot de grilles" sous la forme "Lect:\path\Nomfichier.extension"

Set coll_docs = CATIA.Documents
Dim LotDoc As ProductDocument
Dim ProdLot As Product
Dim Lotprods As Products
Dim LotsProdVar As Variant
Dim LotDocProds As Products
Dim arrayofvariant()

Dim Nb As Integer, i As Integer

Set LotDoc = coll_docs.Add("Product")
Set ProdLot = LotDoc.Product
Set Lotprods = ProdLot.Products
Set LotsProdVar = Lotprods

    'Inserre le product environnement
    Création_Noeud ProdLot, CStr(LstGrilleLot(0)), NomDet
    Set LotDocProds = LotDoc.Product.Products
    LotDocProds.Item(1).Name = "Env.1"

    'reconstruction de la liste dans un variant (variant/variant)
    'la fonction AddComponentsFromFiles ne marche que comme cela
    Nb = UBound(LstGrilleLot) - 1 'L'index 0 correspond au fichier d'environnement
    
    ReDim arrayofvariant(Nb)
    For i = 0 To Nb
        arrayofvariant(i) = LstGrilleLot(i + 1)
    Next i
    LotsProdVar.AddComponentsFromFiles arrayofvariant, "All"
    LotDoc.Product.PartNumber = Left(NomCible, InStr(1, NomCible, ".CATProduct") - 1)
    LotDoc.SaveAs ChemCible & NomCible
    
    'fixe les products dans l'assemblage
    For i = 1 To LotDocProds.Count
        FixePart2 Left(LotDoc.Name, InStr(1, LotDoc.Name, ".CATProduct") - 1), LotDocProds.Item(i).Name
    Next

End Sub

Private Function IncNoGrinue(Noprec As String, sysnum As Integer) As String
'Incrémente les numéro de grille nue en fonction du system de numérotation
Dim tmpDigit As Integer
    
    On Error Resume Next
    Select Case sysnum
        Case 1, 11
            tmpDigit = CInt(Mid(Noprec, 16, 1)) + 2
        Case 2, 3, 12, 13
            tmpDigit = CInt(Right(Noprec, 2)) + 2
    End Select
    
    If Err.Number <> 0 Then
        IncNoGrinue = Noprec & "err_increm"
    Else
        Select Case sysnum
            Case 1, 11
                IncNoGrinue = Left(Noprec, 15) & tmpDigit & Right(Noprec, 3)
            Case 2, 3, 12, 13
                IncNoGrinue = Left(Noprec, Len(Noprec) - 2) & tmpDigit
        End Select
    End If


End Function
