Attribute VB_Name = "A_CreationAssemblage"
Option Explicit
'*********************************************************************
'* Macro : A_CreationAssemblage
'*
'* Fonctions :  Création d'un Product de grille
'*              Crée le product général, le product grille ass, la part grille nue
'*              Sélectionne et importe l'environnement avion
'*              Crée les set géométriques et les contrainte de fixation
'*              Ajoute des attributs provenant d'un fichier excel
'*
'* Version : 9
'* Création :  CFR
'* Modification : 15/04/14
'* Modification : 10/03/16
'*                decoupage de la partie création des fichiers CAO dans un autre procédure pour pouvoir l'appeller dans la macro de création par lot
'*                externalisation de la création d'une nouvelle part dans une fonction externe
'*                externalisation de l'ajout des contraine de Fixité dans une procédure externe
'*                Ajout Part de détrompage
'* Modification : 18/12/16 Ajout de variables reprenant les infos du formulaire pour faciliter la lecture du code
'*
'**********************************************************************

Sub catmain()

'Log de l'utilisation de la macro
LogUtilMacro nPath, nFicLog, nMacro, "A_CreationAssemblage", VMacro

Dim i As Integer
Dim ProdLotGrille As Product
Dim nIstanceGrillesym As String
Dim nPartU1 As String, nPartU1sym As String, NomPartDet As String
Dim nLot As String, nGriAss As String, nGriAssSym As String, nGriNue As String, nGriNueSym As String
Dim DesGriAss As String, DesgriAssSym As String
Dim nPathSave As String, nEnv As String
Dim cas As Integer

'Ouvre la boite de dlg "Frm_DonnEntre"
    Load Frm_DonnEntre
    Frm_DonnEntre.Show
 
'Sort du programme si click sur bouton Annuler dans FRM_DonnEntre
    If Not (Frm_DonnEntre.ChB_OkAnnule) Then
        Unload Frm_DonnEntre
        Exit Sub
    End If
    
    'Stockage des infos du formulaire
    nLot = Frm_DonnEntre.TBX_NomAss
    nPathSave = Frm_DonnEntre.TBX_RepSave
    nEnv = Frm_DonnEntre.TBX_FicEnv
    nGriAss = Frm_DonnEntre.TBX_NomGriAss
    nGriAssSym = Frm_DonnEntre.TBX_NomGriAssSym
    nGriNue = Frm_DonnEntre.TBX_NomGriNue
    nGriNueSym = Frm_DonnEntre.TBX_NomGriNueSym
    DesGriAss = ValDscgp.design
    DesgriAssSym = ValDscgp.DesignSym
                
    If Frm_DonnEntre.ChB_U01 Then   '="" si pas de part U1 cochée dans le formulaire (même si son N° est documenté dans le DSCGP)
        nPartU1 = Frm_DonnEntre.TBX_NomU01
        nPartU1sym = Frm_DonnEntre.TBX_NomU01Sym
    End If
    
    If Frm_DonnEntre.ChB_Detromp Then
        NomPartDet = Frm_DonnEntre.TBX_NomDtromp
    Else
        NomPartDet = ""
    End If
    
    'Calcul des cas
    If ValDscgp.CoteAvion = "GAUCHE" Then
        If Frm_DonnEntre.TBX_NomGriAss <> "" And Frm_DonnEntre.TBX_NomGriAssSym = "" Then
            cas = 1 'Cas =1 => Grille gauche seule
        ElseIf Frm_DonnEntre.TBX_NomGriAss <> "" And Frm_DonnEntre.TBX_NomGriAssSym <> "" Then
            cas = 2  'Cas =2 => Grille gauche + sym droit
        End If
    ElseIf ValDscgp.CoteAvion = "DROIT" Then
        If Frm_DonnEntre.TBX_NomGriAss <> "" And Frm_DonnEntre.TBX_NomGriAssSym = "" Then
            cas = 3 'Cas =3 => Grille droite seule
        ElseIf Frm_DonnEntre.TBX_NomGriAss <> "" And Frm_DonnEntre.TBX_NomGriAssSym <> "" Then
            cas = 4  'Cas =4 => Grille droite + sym gauche
            'Inversion des nom de grille
            nGriAss = Frm_DonnEntre.TBX_NomGriAssSym
            nGriAssSym = Frm_DonnEntre.TBX_NomGriAss
            nGriNue = Frm_DonnEntre.TBX_NomGriNueSym
            nGriNueSym = Frm_DonnEntre.TBX_NomGriNue
            nPartU1 = Frm_DonnEntre.TBX_NomU01Sym
            nPartU1sym = Frm_DonnEntre.TBX_NomU01
            DesGriAss = ValDscgp.DesignSym
            DesgriAssSym = ValDscgp.design
        End If
    ElseIf ValDscgp.CoteAvion = "CENTRE" Then
        cas = 5  'Cas =5 => Grille gauche seule
    End If
    Unload Frm_DonnEntre

    Select Case cas
        Case 1, 3, 5
            'Création de la grille principale
            CreateCAO nLot, nGriAss, DesGriAss, nGriNue, nPathSave, nEnv, nPartU1, NomPartDet
        Case 2, 4
            'Création de la grille principale
            CreateCAO nLot, nGriAss, DesGriAss, nGriNue, nPathSave, nEnv, nPartU1, NomPartDet
            'Ajout de la grille sym dans le product du lot
            Set coll_docs = CATIA.Documents
            For i = 1 To coll_docs.Count
                If InStr(1, coll_docs.Item(i).Name, nLot, vbTextCompare) <> 0 Then
                    Set ProdLotGrille = coll_docs.Item(i).Product
                    AjoutGrille ProdLotGrille, nGriAssSym, DesgriAssSym, nGriNueSym, nLot, nPartU1sym
                End If
             Next i
            'Fixe le product Grille Assemblée sym dans le lot
             For i = 1 To coll_docs.Count
                If InStr(1, coll_docs.Item(i).Name, nGriAssSym, vbTextCompare) <> 0 Then
                    nIstanceGrillesym = Left(coll_docs.Item(i).Name, InStr(1, coll_docs.Item(i).Name, ".") - 1) & ".1"
                    FixePart2 nLot, nIstanceGrillesym
                    Exit For
                End If
            Next i
            'End If
    End Select
    
End Sub

Public Sub CreateCAO(Nom_Ass As String, nGrilleAss As String, DesGriAss As String, nGrilleNue As String, RepGrille As String, fEnv As String, nPartU1 As String, nPartDet As String)
'Création des fichiers CAO de la grille
'Nom_Ass = Nom de l'assemblage (N° du lot)
'nGrilleAss = N° de la grille assemblée
'desGriAss = designation de la grille assemblèe
'nGrilleNue = N° de la grille nue
'RepGrille = repertoire de sauvegarde des fichiers CAO de la grille
'fEnv = répertoire et Nom du fichier d'environnement avion
'nPartU1 = nom de la part U1
'nPartDet = Nom de la part de détrompage
Dim i As Integer
Dim nDocAss As ProductDocument
Dim ProdAssGen As Product
Dim AssGenProds As Products
Dim mParams As Parameters
Dim ParamAdd As StrParam

'Création du Product Assemblage
    Set coll_docs = CATIA.Documents
    Set nDocAss = coll_docs.Add("Product")
    Set ProdAssGen = nDocAss.Product
    ProdAssGen.PartNumber = Nom_Ass
    Set AssGenProds = ProdAssGen.Products

'Sauvegarde des Infos de la boite Dialogue Donnée d'entrée
'dans des paramètres enregistrés dans le fichier d'assemblage
    
    Set mParams = ProdAssGen.UserRefProperties
    Set ParamAdd = mParams.CreateString("Param_Assembl", Nom_Ass)
    Set ParamAdd = mParams.CreateString("param_GrillAss", nGrilleAss)
    Set ParamAdd = mParams.CreateString("Param_GrillNue", nGrilleNue)
    Set ParamAdd = mParams.CreateString("Param_FicEnvAvion", fEnv)
    Set ParamAdd = mParams.CreateString("Param_RepSauv", RepGrille)

'Création du Noeud Environnement
    Création_Noeud ProdAssGen, fEnv, nPartDet

'Fixe le Noeud Environnement
    FixePart2 Nom_Ass, "env.1"

'Ajoute un product grille Assemblée
    AjoutGrille ProdAssGen, nGrilleAss, DesGriAss, nGrilleNue, nGrilleAss, nPartU1

'Fixe le product Grille Assemblée
    For i = 1 To nDocAss.Product.Products.Count
        If InStr(1, nDocAss.Product.Products.Item(i).Name, nGrilleAss, vbTextCompare) <> 0 Then
            FixePart2 Nom_Ass, nDocAss.Product.Products.Item(i).Name
        End If
    Next
    
'nDocAss.Activate
End Sub


Public Sub Création_Noeud(AssGenProd As Product, Fich_env As String, Fich_Det As String)
'Création du noeud environnement avec imports et fixation du fichier de l'environnement
'AssGenProd = Product de l'ensemnle général
'Fich_env = Nom du fichier de l'environnement sous la forme "lecteur:\rep\nomfichier.ext"
'Fic_Det = Nom du fichier de la part de détrompage sous la forme "nomfichier.ext"
    
    Dim coll_prods As Products
    Dim NoeudEnvProd As Product
    Dim NoeudEnvProds As Products
    Dim arrayofvariant(0)
    Dim VNoeudEnvProds  As Variant
    Dim Nom_Ass As String
    Dim EnvProd As Product
    Dim EnvAvion As Product
    
    Set coll_prods = AssGenProd.Products
    Set NoeudEnvProd = coll_prods.AddNewProduct("env")
        NoeudEnvProd.Name = "env.1"
    Set NoeudEnvProds = NoeudEnvProd.Products
        Nom_Ass = AssGenProd.Name
    
'Insertion du product environnement
    arrayofvariant(0) = Fich_env
    Set VNoeudEnvProds = NoeudEnvProds
    VNoeudEnvProds.AddComponentsFromFiles arrayofvariant, "All"

'Fixe le Product Environnement dans le Noeud Env
    Set EnvProd = VNoeudEnvProds.Parent
    Set EnvAvion = EnvProd.Products.Item(1)
    FixeProdNoeud Nom_Ass, EnvAvion.Name

'Création de la part de détrompage
    If Fich_Det <> "" Then
        AjoutPart NoeudEnvProd, Fich_Det
        
        'Fixation de la part de détrompage
        Dim PartDet As Product
        Set PartDet = EnvProd.Products.Item(2)
        FixeProdNoeud Nom_Ass, PartDet.Name
    End If

End Sub
