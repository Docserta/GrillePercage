Attribute VB_Name = "J_ImportBagues"
Option Explicit


'*********************************************************************
'* Macro : J_ImportBagues
'*
'* Fonctions :  Import des bagues et vis arrétoirs dans le product Grille
'*
'*
'* Version 8
'* Création : CFR 13/06/2015
'* Modification :
'*
'*
'*
'**********************************************************************
Sub catmain()
Dim GrilleNueActive As New c_PartGrille
Dim GrilleAssActive As New GrilleAss
Dim GrilleActive As c_PartGrille
Dim TabCompImport() As String ' Table des Composants a importer
Dim HBShape_Std_EC As HybridShape 'STD en cours
Dim Std_Parameters As Parameters 'Collection des paramètre de la droite STD en cours
Dim Std_ParamEC As Parameter
Dim std_paramEC_NBVis As Parameter
Dim i As Long, j As Long
Dim Barre As ProgressBarre

'Log de l'utilisation de la macro
LogUtilMacro nPath, nFicLog, nMacro, "J_ImportBagues", VMacro

    '-------------------------
    ' Check de l'environnement
    '-------------------------
    Set coll_docs = CATIA.Documents
    If Not Check_GrilleAss() Then
        Exit Sub
    End If

    If GrilleAssActive.Exist_PartGrilleNue Then
        Set GrilleActive = New c_PartGrille
    Else
        MsgBox "Pas de grille nue détectée dans le product", vbCritical, "Erreur d'environnement"
        End
    End If

    'Test le chemin de la bibli des composants
    CheminBibliComposants = CorrigeDFS()

'Récupération des info dans la grille nue
ReDim TabCompImport(3, 0)
'TabCompImport(0, x) = Type (bague, vis, agrafe
'TabCompImport(1, x) = Nom du STD
'TabCompImport(2, x) = Nom du composant a importer
'TabCompImport(3, x) = Ref du repère

j = 1

'Recherche du document de la grille nue a partir de son nom dans la collection des documents ouverts
    For i = 1 To coll_docs.Count
        If coll_docs.Item(i).Name = GrilleAssActive.GrilleNueNom & ".CATPart" Then
            GrilleNueActive.PG_partDocGrille = coll_docs.Item(i)
'            GrilleNueActive.PG_partDocGrille = coll_docs.GetItem(GrilleAssActive.GrilleNueNom & ".CATPart")
            Exit For
        End If
    Next

 'Test l'existance de la ligne d'orientation grille
    If Not GrilleNueActive.Exist_OrientationGrille Then
        MsgBox "La ligne d'orientation grille est manquante ou mal orthographiée", vbCritical, "Eléments manquants"
        End
    End If
    
 'Test l'existance du set géométrique STD
    If Not GrilleNueActive.Exist_HB(nHBStd) Then
        MsgBox "Le set géométrique : " & (nHBStd) & " est manquant ou mal orthographié", vbCritical, "Eléments manquants"
        End
    End If

'affichage de la barre de progression
    Set Barre = New ProgressBarre
    Barre.ProgressTitre 1, " Import des composants, veuillez patienter."
    
 'Construction du tableau des composants a importer
    For i = 1 To GrilleNueActive.Hb(nHBStd).HybridShapes.Count
        Set HBShape_Std_EC = GrilleNueActive.Hb(nHBStd).HybridShapes.Item(i)
        'Récupération des paramètres sur le STD en cours
        Set Std_Parameters = GrilleNueActive.PartGrille.Parameters.SubList(HBShape_Std_EC, True)
        
        For Each Std_ParamEC In Std_Parameters
        
        'Documente le tableau pour les bagues
            If InStr(1, Std_ParamEC.Name, "NumBague", vbTextCompare) > 0 Then
                ReDim Preserve TabCompImport(3, j)
                TabCompImport(0, j) = "Bague"
                TabCompImport(1, j) = HBShape_Std_EC.Name
                TabCompImport(2, j) = Std_ParamEC.Value & ".CATPart"
                TabCompImport(3, j) = GrilleAssActive.Numero & "/" & GrilleAssActive.Produits.Item(1).Name & "/!" & "RepAss_BagueA" & Right(HBShape_Std_EC.Name, Len(HBShape_Std_EC.Name) - InStr(1, HBShape_Std_EC.Name, ".", vbTextCompare))
                j = j + 1
            ElseIf InStr(1, Std_ParamEC.Name, "NumVisArretoir", vbTextCompare) > 0 Then
            'Documente le tableau pour les Vis Arretoirs
                ReDim Preserve TabCompImport(3, j)
                TabCompImport(0, j) = "visArretoir"
                TabCompImport(1, j) = HBShape_Std_EC.Name
                TabCompImport(2, j) = Std_ParamEC.Value & ".CATPart"
                TabCompImport(3, j) = GrilleAssActive.Numero & "/" & GrilleAssActive.Produits.Item(1).Name & "/!" & "RepAss_VisArretoir1A" & Right(HBShape_Std_EC.Name, Len(HBShape_Std_EC.Name) - InStr(1, HBShape_Std_EC.Name, ".", vbTextCompare))
                j = j + 1
                'Si c'est une double vis arretoir on ajoute une ligne au tableau
                For Each std_paramEC_NBVis In Std_Parameters
                    If InStr(1, std_paramEC_NBVis.Name, "NbVisArretoir", vbTextCompare) > 0 Then
                        If UCase(std_paramEC_NBVis.Value) = "DOUBLE" Then
                            ReDim Preserve TabCompImport(3, j)
                            TabCompImport(0, j) = "visArretoir"
                            TabCompImport(1, j) = HBShape_Std_EC.Name
                            TabCompImport(2, j) = Std_ParamEC.Value & ".CATPart"
                            TabCompImport(3, j) = GrilleAssActive.Numero & "/" & GrilleAssActive.Produits.Item(1).Name & "/!" & "RepAss_VisArretoir2A" & Right(HBShape_Std_EC.Name, Len(HBShape_Std_EC.Name) - InStr(1, HBShape_Std_EC.Name, ".", vbTextCompare))
                            j = j + 1
                        End If
                    End If
                Next
            ElseIf InStr(1, Std_ParamEC.Name, "NoAgrafe", vbTextCompare) > 0 Then
            'Documente le tableau pour les agrafes
                ReDim Preserve TabCompImport(3, j)
                TabCompImport(0, j) = "Agrafe"
                TabCompImport(1, j) = HBShape_Std_EC.Name
                TabCompImport(2, j) = Std_ParamEC.Value & ".CATProduct"
                TabCompImport(3, j) = GrilleAssActive.Numero & "/" & GrilleAssActive.Produits.Item(1).Name & "/!" & "RepAss_AgrafeA" & Right(HBShape_Std_EC.Name, Len(HBShape_Std_EC.Name) - InStr(1, HBShape_Std_EC.Name, ".", vbTextCompare))
                j = j + 1
            End If
        Next
    Next
    
'Instanciation des Composants
    If UBound(TabCompImport, 2) > 0 Then 'La ligne 0 est vide, elle sert a éviter l'erreure ubound sur un tableau vide
        For j = 1 To UBound(TabCompImport, 2)
        'Mise a jour de la barre de progression
            Barre.ProgressTitre 100 / UBound(TabCompImport, 2) * j, " Import du composant, " & TabCompImport(2, j) & " veuillez patienter."
            InstanciationComposant GrilleAssActive, TabCompImport, j
        Next
    End If
    
'Libération des classes
    Set GrilleAssActive = Nothing
    Set GrilleActive = Nothing
    Set Barre = Nothing

    
End Sub

Public Sub InstanciationComposant(GrilleAssActive, IC_Composant() As String, IndexTab)

' Entrées :
' IC_Composant()    :   IC_Composant(0,IndexTab) Type de composant
'                       IC_Composant(1,IndexTab) Nom du STD
'                       IC_Composant(2,IndexTab) Nom du composant a importer
'                       IC_Composant(3,IndexTab) Nom de la Reférence cible (triedre de la grille)

Dim CollInstances
Dim NomRepComp As String
Dim NomInstanceComp As String
Dim arrayOfVariantOfBSTR1(0)

'Construction du chemin d'acces aux composants
    Select Case UCase(IC_Composant(0, IndexTab))
        Case "BAGUE"
            arrayOfVariantOfBSTR1(0) = CheminBibliComposants & "\" & ComplementCheminBibliComposants & RepBagues & "\" & IC_Composant(2, IndexTab)
        Case "VISARRETOIR"
            arrayOfVariantOfBSTR1(0) = CheminBibliComposants & "\" & ComplementCheminBibliComposants & RepVis & "\" & IC_Composant(2, IndexTab)
        Case "AGRAFE"
            arrayOfVariantOfBSTR1(0) = CheminBibliComposants & "\" & ComplementCheminBibliComposants & RepAgrafes & "\" & IC_Composant(2, IndexTab)
    End Select

'Test d'éxistance du fichier du composant
    If Dir(arrayOfVariantOfBSTR1(0), vbNormal) = "" Then
        MsgBox "Le composant : " & IC_Composant(2, IndexTab) & " est introuvable. Ajouter le à la bibliothèque des composants.", vbInformation
        Exit Sub
    End If

' Ajout de la pièce dans le CATProduct (instanciation)
    Set CollInstances = GrilleAssActive.Produits
    CollInstances.AddComponentsFromFiles arrayOfVariantOfBSTR1, "All"

' Récupération du nom d'instance
    NomInstanceComp = CollInstances.Item(CollInstances.Count).Name

' Création des références ex '"T000823666-00000Y01/agrafe1.1/CLD-ST00095127.1/!Absolute Axis System"
    Dim CollDocuments As Documents
    Set CollDocuments = CATIA.Documents
    Dim DocInstancePart As PartDocument
    Dim DocInstanceProd As ProductDocument
    Dim InstanceProduit As Product
    Dim PartInstance As part
    Dim NomRefSource As String, NomRefCible As String
                                NomRefCible = IC_Composant(3, IndexTab)
    Dim InstCLDAbsent As Boolean
        InstCLDAbsent = True
    'si c'est un part
    If InStr(1, IC_Composant(2, IndexTab), "CATPart", vbTextCompare) Then
        Set DocInstancePart = CollDocuments.Item(CStr(IC_Composant(2, IndexTab)))
        Set PartInstance = DocInstancePart.part
        'Chemin de la ref source
        NomRefSource = GrilleAssActive.Numero & "/" & NomInstanceComp
        
    'si c'est un Product
    ElseIf InStr(1, IC_Composant(2, IndexTab), "CATProduct", vbTextCompare) Then
        Set DocInstanceProd = CollDocuments.Item(CStr(IC_Composant(2, IndexTab)))
        For Each InstanceProduit In DocInstanceProd.Product.Products
            If Left(InstanceProduit.Name, 3) = "CLD" Then
                Set DocInstancePart = CollDocuments.Item(InstanceProduit.PartNumber & ".CATPart")
                Set PartInstance = DocInstancePart.part
                InstCLDAbsent = False
                'Chemin de la ref source
                NomRefSource = "GrilleAssActive.Numero/" & NomInstanceComp & "/" & InstanceProduit.Name
            End If

        Next
            If InstCLDAbsent Then
                MsgBox "Pas de part CLD trouvé dans le product " & IC_Composant(2, IndexTab) & " !", vbCritical, "Element manquant"
                End
            End If
    End If
     
'Recherche du triedre d'import dans le composant
'S'il n'y a un trièdre "rep_ass" , c'est le bon sinon on prend le premier
    Dim AxisInstance As AxisSystems
    Set AxisInstance = PartInstance.AxisSystems
    Dim AxeImportInstance As AxisSystem
    Dim i As Long
    
    Set AxeImportInstance = AxisInstance.Item(1)
    For i = 1 To AxisInstance.Count
        If AxisInstance.Item(i).Name = "rep_ass" Then
            Set AxeImportInstance = AxisInstance.Item(i)
            Exit For
        End If
    Next
     
    Dim NomAxis As String
        NomAxis = AxeImportInstance.Name
        
'Ajout du nom du triedre au chemin du nom de la refsource
    NomRefSource = NomRefSource & "/!" & NomAxis

' Calcul des références des deux composants
    Dim RefSource As Reference
    Set RefSource = GrilleAssActive.Produit.CreateReferenceFromName(NomRefSource)
    Dim RefCible As Reference
    Set RefCible = GrilleAssActive.Produit.CreateReferenceFromName(NomRefCible)

' Création de la contrainte de coincidence
    Dim ContrainteAssComposant As Constraint
    Set ContrainteAssComposant = GrilleAssActive.Contraintes.AddBiEltCst(catCstTypeOn, RefSource, RefCible)
    ContrainteAssComposant.Name = "Coincidence-" & NomInstanceComp & "_" & IC_Composant(1, IndexTab)

' Mise àjour de l'assemblage
    GrilleAssActive.Produit.Update

End Sub
