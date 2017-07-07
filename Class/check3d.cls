VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "check3d"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*********************************************************************
'* Classe : Check3D
'*
'* Fonctions :  checks de controle 3D
'*
'*
'* Version : 9
'* Création :  CFR - 26/04/16
'* Modification :
'*
'*********************************************************************
Option Explicit
Private cItemCk As c_itemCheck3D
Private strNomDscgp As String
Private intVersDscgp As Integer
Private DSCGP_EC As c_DSCGP
Private Lot_EC As GrilleLot
Private NomGrilleAss As String
Private GrilleASs_EC As GrilleAss
Private cGriNues As c_cKParts
Private cConstrs As c_Contraintes
Private i As Long
Private j As Long

'
'===========================
' Initialisation de l'objet
'===========================
'
Private Sub Class_initialize()
'Initialisation des sous classes
Set DSCGP_EC = New c_DSCGP
Set Lot_EC = New GrilleLot
Set GrilleASs_EC = New GrilleAss
Set cGriNues = New c_cKParts
Set cConstrs = New c_Contraintes
Set col_Mes = New c_mesures

End Sub

Property Let VersDscgp(ByVal ent As Integer)
'Récupère la version du DSCGP
    intVersDscgp = ent
    DSCGP_EC.VersionDscgp = intVersDscgp
End Property

Property Get VersDscgp() As Integer
'Renvoi la version du DSCGP
    VersDscgp = intVersDscgp
End Property

Property Let NomDscgp(ByVal str As String)
'Récupère le nom du DSCGP
    strNomDscgp = str
    DSCGP_EC.OpenDSCGP = strNomDscgp
End Property

Property Get NomDscgp() As String
'Renvoi le nom du DSCGP
    NomDscgp = strNomDscgp
End Property

'Affecte la grille choisie à la classe GrilleAss
Property Let NomDocGrilleAss(ByVal str As String)
    NomGrilleAss = str
    
    GrilleASs_EC.ProductDocGrille = getDocNom(NomGrilleAss)
    Dim tDocPart As PartDocument
    Dim tpart As part
    Dim tProp As Parameter
    Dim tShow As CatVisPropertyShow
    Dim tVisProp As VisPropertySet
    Dim tSel As Selection
        Set tSel = CATIA.ActiveDocument.Selection
    Dim tBod As HybridBody, SubTBod As HybridBody
    Dim cParams As c_Properties
    Dim cHbodies As c_Hbodies
    Dim cFasteners As c_Fasteners

    Dim tMaterial As Material
    Dim tMasse   As Double
    Dim PartinList As String
    
    'Construit la collection des parts de la grille Ass
    Dim Prod As Product
    On Error Resume Next
    For Each Prod In GrilleASs_EC.Produits
        
        Set tDocPart = coll_docs.Item(Prod.PartNumber & ".CATPart")
        Set tpart = tDocPart.part
        'Test si c'est une grille ou un élément du commerce
        If EstGriNue(Prod.PartNumber, DSCGP_EC.SystemNum) Then
            'Recherche si le part est déja présent dans la collection
            PartinList = cGriNues.Item(CStr(Prod.PartNumber)).nom
            If PartinList = "" Then
                'Le part n'existe pas dans la collection
                'Construit la collection des propriètès
                Set cParams = New c_Properties
                For Each tProp In tDocPart.Product.UserRefProperties
                    cParams.Add DecoupeSlash(tProp.Name), tProp.Value
                Next
    
                'Construit la colection des HybridBodies
                Set cHbodies = New c_Hbodies
                For Each tBod In tpart.HybridBodies
                    tSel.Clear
                    'Recupère le statut Show/noShow de l'oblet
                    tSel.Add tBod
                    Set tVisProp = tSel.VisProperties
                    tVisProp.GetShow tShow
                    cHbodies.Add tBod.Name, tShow, tBod
                Next
                
                'construit la collection des mesures
                Set cFasteners = LstFasteners(tpart)
                LstMesures tpart, cFasteners

                'récupération de la masse du part
                tMasse = MassPart(tDocPart.Product)
                'recupération de la matière du part
                Set tMaterial = MatBody(tpart, tpart.MainBody)
                
                cGriNues.Add Prod.PartNumber, tpart, Prod.Source, Prod.DescriptionRef, tMasse, tMaterial, cParams, cHbodies, cFasteners
                
            End If
        End If
        Set cFasteners = Nothing
        Set cParams = Nothing
        Set cHbodies = Nothing
    Next
    On Error GoTo 0
End Property

Property Get NomDocGrilleAss() As String
    NomDocGrilleAss = NomGrilleAss
End Property

'################################
'# Les checks
'################################

Property Get CK_F01() As c_itemCheck3D
'L'assemblage général porte le n° de lot du DSCGP
    Set cItemCk = New c_itemCheck3D
        cItemCk.Cote = "F-01"
        cItemCk.Check = True
    If Left(Lot_EC.NoLot, Len(DSCGP_EC.NumLot)) = DSCGP_EC.NumLot Then
        cItemCk.statut = "OK"
        cItemCk.Comment = "L'assemblage général porte le n° de lot du DSCGP"
    Else
        cItemCk.statut = "KO"
        cItemCk.Comment = "N° de l'assemblage : " & Lot_EC.NoLot & " Numéro du lot dans le DSCGP : " & DSCGP_EC.NumduLot
    End If
         
    'Renvoi l'objet crée
    Set CK_F01 = cItemCk
    'Liberation des classes
    Set cItemCk = Nothing

End Property

Property Get CK_F02() As c_itemCheck3D
'La numérotation de l'assemblage général est en accord avec le programme avion et le site client
    Set cItemCk = New c_itemCheck3D
        cItemCk.Cote = "F-02"
        cItemCk.Check = True
    If Lot_EC.NoLot = DSCGP_EC.NumduLot Then
        cItemCk.statut = "OK"
        cItemCk.Comment = "La numérotation de l'assemblage général est en accord avec le programme avion et le site client"
    Else
        cItemCk.statut = "KO"
        cItemCk.Comment = "N° de l'assemblage : " & Lot_EC.NoLot & " Numéro du lot prévu dans le DSCGP : " & DSCGP_EC.NumduLot
    End If
         
    'Renvoi l'objet crée
    Set CK_F02 = cItemCk
    'Liberation des classes
    Set cItemCk = Nothing
        
End Property

Property Get CK_F03() As c_itemCheck3D
'Le nom de l'assemblage de la grille est en accord avec le programme avion et le site client
On Error GoTo Erreur
    Set cItemCk = New c_itemCheck3D
        cItemCk.Cote = "F-03"
        cItemCk.Check = True
    If GrilleASs_EC.Numero = DSCGP_EC.NumGrille Then
        cItemCk.statut = "OK"
        cItemCk.Comment = "Le nom de l'assemblage de la grille est en accord avec le programme avion et le site client"
    Else
        cItemCk.statut = "KO"
        cItemCk.Comment = "N° de la grille assemblée : " & GrilleASs_EC.Numero & " Numéro de la grille assemblée prévue dans le DSCGP : " & DSCGP_EC.NumGrille
    End If
    GoTo Fin

Erreur:
    cItemCk.statut = "KO"
    If Err.Number > vbObjectError + 512 Then
        cItemCk.Comment = Err.Description
    Else
        cItemCk.Comment = "Erreur dans la récupération du nom de l'assemblage"
    End If
Fin:
    'Renvoi l'objet crée
    Set CK_F03 = cItemCk
    'Liberation des classes
    Set cItemCk = Nothing
        
End Property

Property Get CK_F04() As c_itemCheck3D
'Le nom du CatPart de la grille nue est en accord avec le programme avion et le site client
On Error GoTo Erreur
    Set cItemCk = New c_itemCheck3D
        cItemCk.Cote = "F-04"
        cItemCk.Check = True
    Dim part_ec As c_ckPart
    Set part_ec = New c_ckPart
    For i = 1 To cGriNues.Count
        Set part_ec = cGriNues.Item(i)
        If NumGriOK(part_ec.nom, DSCGP_EC.SystemNum, DSCGP_EC.NumLot) Then
            cItemCk.statut = "OK"
            cItemCk.Comment = "Le nom du CatPart de la grille nue est en accord avec le programme avion et le site client"
        Else
            cItemCk.Comment = cItemCk.Comment & "Part : " & part_ec.nom & " - Numérotation de la grille incorrect " & Chr(10)
            cItemCk.statut = "KO"
        End If
     Next i
    GoTo Fin
Erreur:
    cItemCk.statut = "KO"
    If Err.Number > vbObjectError + 512 Then
        cItemCk.Comment = Err.Description
    Else
        cItemCk.Comment = "Erreur dans la récupération du nom de l'assemblage"
    End If
Fin:
    'Renvoi l'objet crée
    Set CK_F04 = cItemCk
    'Liberation des classes
    Set cItemCk = Nothing
    Set part_ec = Nothing
        
End Property

Property Get CK_F05() As c_itemCheck3D
'L'assemblage général contient un composant nommé "env"
    Set cItemCk = New c_itemCheck3D
        cItemCk.Cote = "F-05"
        cItemCk.Check = True
    If Lot_EC.Exist_NoeudEnv Then
        cItemCk.statut = "OK"
        cItemCk.Comment = "Noeud Environnement trouvé"
    Else
        cItemCk.statut = "KO"
        cItemCk.Comment = "Noeud Environnement non trouvé"
    End If

    'Renvoi l'objet crée
    Set CK_F05 = cItemCk
    'Liberation des classes
    Set cItemCk = Nothing
    
End Property

Property Get CK_G01() As c_itemCheck3D
'L'environnement avion est fixé dans le composant "env"
    Set cItemCk = New c_itemCheck3D
        cItemCk.Cote = "G-01"
        cItemCk.Check = True
        cItemCk.statut = "OK"
    Dim tConstr_EC As c_Contrainte
    Set tConstr_EC = New c_Contrainte
    If Lot_EC.Exist_NoeudEnv Then
        On Error GoTo Erreur
        Set cConstrs = Lot_EC.Col_ConstEnv
        For i = 1 To cConstrs.Count
            Set tConstr_EC = cConstrs.Item(i)
            If tConstr_EC.RefType <> catCstRefTypeFixInSpace Then
                cItemCk.statut = "KO"
                cItemCk.Comment = cItemCk.Comment & tConstr_EC.nom & "n'est pas une fixité"
            End If
            If tConstr_EC.statut <> catCstStatusOK Then
                cItemCk.statut = "KO"
                cItemCk.Comment = cItemCk.Comment & tConstr_EC.nom & "n'est pas OK"
            End If
        Next i
    Else
        cItemCk.statut = "KO"
        cItemCk.Comment = "Noeud Environnement non trouvé"
    End If
    GoTo Fin

Erreur:
    If Err.Number = vbObjectError + 540 Then
        cItemCk.statut = "A vérifier"
        cItemCk.Comment = Err.Description
    End If
Fin:
    'Renvoi l'objet crée
    Set CK_G01 = cItemCk
    'Liberation des classes
    Set tConstr_EC = Nothing
    Set cConstrs = Nothing
    Set cItemCk = Nothing
    i = 0
End Property

Property Get CK_G02() As c_itemCheck3D
'Le composant "env" et les assemblages des grilles sont fixés dans l'assemblage général
    Set cItemCk = New c_itemCheck3D
        cItemCk.Cote = "G-02"
        cItemCk.Check = True
        cItemCk.statut = "OK"
    Dim tConstr_EC As c_Contrainte
    Set tConstr_EC = New c_Contrainte
    
'#actuellement la macro verifie l'état des contraintes
'# il faudrait vérifier que les liens des contraintes concerne bien l'Env

    If Lot_EC.Exist_NoeudEnv Then
        Set cConstrs = Lot_EC.Col_ConstLot
        For i = 1 To cConstrs.Count
            Set tConstr_EC = cConstrs.Item(i)
            If tConstr_EC.RefType <> catCstRefTypeFixInSpace Then
                cItemCk.statut = "KO"
                cItemCk.Comment = cItemCk.Comment & " - " & tConstr_EC.nom & "n'est pas une fixité" & Chr(10)
            End If
            If tConstr_EC.statut <> catCstStatusOK Then
                cItemCk.statut = "KO"
                cItemCk.Comment = cItemCk.Comment & " - " & tConstr_EC.nom & "n'est pas OK" & Chr(10)
            End If
        Next i
    Else
        cItemCk.statut = "KO"
        cItemCk.Comment = "Noeud Environnement non trouvé"
    End If
    If cItemCk.statut = "KO" Then
        cItemCk.Comment = "Les contraintes suivantes sont a vérifier : " & Chr(10) & cItemCk.Comment
    End If
    'Renvoi l'objet crée
    Set CK_G02 = cItemCk
    'Liberation des classes
    Set tConstr_EC = Nothing
    Set cConstrs = Nothing
    Set cItemCk = Nothing
    i = 0
End Property

Property Get CK_G03() As c_itemCheck3D
'Le CatPart de la grille nue est fixé dans l'assemblage de la grille
    Set cItemCk = New c_itemCheck3D
        cItemCk.Cote = "G-03"
        cItemCk.Check = True
        cItemCk.statut = "OK"
'    Dim CK_colParts As c_cKParts ' Collection de parts du product grille ass
'    Set CK_colParts = New c_cKParts
    Dim tConstr_EC As c_Contrainte
    Set tConstr_EC = New c_Contrainte

'#actuellement la macro verifie l'état des contraintes
'# il faudrait vérifier que les liens des contraintes concerne bien les parts nue

    Set cConstrs = GrilleASs_EC.Col_Const
    For i = 1 To cConstrs.Count
        Set tConstr_EC = cConstrs.Item(i)
        If tConstr_EC.RefType <> catCstRefTypeFixInSpace Then
            cItemCk.statut = "KO"
            cItemCk.Comment = cItemCk.Comment & " - " & tConstr_EC.nom & "n'est pas une fixité" & Chr(10)
        End If
        If tConstr_EC.statut <> catCstStatusOK Then
            cItemCk.statut = "KO"
            cItemCk.Comment = cItemCk.Comment & " - " & tConstr_EC.nom & "n'est pas OK" & Chr(10)
        End If
    Next i
    If cItemCk.statut = "KO" Then
        cItemCk.Comment = "Les contraintes suivantes sont a vérifier : " & Chr(10) & cItemCk.Comment
    End If
    'Renvoi l'objet crée
    Set CK_G03 = cItemCk
    'Liberation des classes
    Set cConstrs = Nothing
'    Set CK_colParts = Nothing
    Set cItemCk = Nothing

End Property

Property Get CK_G04() As c_itemCheck3D
'L'ensemble des élements de la grille sont correctement contraints sur la grille (/!\ états de contraintes)
    Set cItemCk = New c_itemCheck3D
        cItemCk.Cote = "G-04"
        cItemCk.Check = True
        cItemCk.statut = "OK"

    Dim tConstr_EC As c_Contrainte
    Set tConstr_EC = New c_Contrainte
   
    Set cConstrs = GrilleASs_EC.Col_Const
    For i = 1 To cConstrs.Count
        Set tConstr_EC = cConstrs.Item(i)
        If tConstr_EC.statut <> catCstStatusOK Then
            cItemCk.statut = "KO"
            cItemCk.Comment = cItemCk.Comment & tConstr_EC.nom & "n'est pas OK" & Chr(10)
        End If

        Set tConstr_EC = Nothing
    Next i
    If cItemCk.statut = "KO" Then
        cItemCk.Comment = "Les contraintes suivantes sont a vérifier : " & Chr(10) & cItemCk.Comment
    End If
    'Renvoi l'objet crée
    Set CK_G04 = cItemCk
    'Liberation des classes
    Set cItemCk = Nothing
    Set cConstrs = Nothing
    
End Property

Property Get CK_G06() As c_itemCheck3D
'Les contraintes de fixations et de coïncidences sont en NO SHOW
    Set cItemCk = New c_itemCheck3D
        cItemCk.Cote = "G-06"
        cItemCk.Check = True
        cItemCk.statut = "OK"
    Dim Tconst As Constraint
    Dim tShow As CatVisPropertyShow
    Dim tVisProp As VisPropertySet
    Dim tSel As Selection
    
    Set tSel = GrilleASs_EC.ProductDocGrille.Selection
    For Each Tconst In GrilleASs_EC.Contraintes
        tSel.Clear
        tSel.Add Tconst
        Set tVisProp = tSel.VisProperties
            tVisProp.GetShow tShow
        If tShow = catVisPropertyShowAttr Then
            cItemCk.statut = "KO"
            cItemCk.Comment = cItemCk.Comment & Tconst.Name & Chr(10)
        End If
    Next
    If cItemCk.statut = "KO" Then
        cItemCk.Comment = GrilleASs_EC.Numero & " - les contraites suivantes ne sont pas dans le 'NoShow' : " & Chr(10) & cItemCk.Comment
    End If
    'Renvoi l'objet crée
    Set CK_G06 = cItemCk
    'Liberation des classes
    Set cItemCk = Nothing
    tSel.Clear
End Property

Property Get CK_H01() As c_itemCheck3D
'Les champs suivants sont présents dans les propriétés de la Part :
'       material, thickness/diameter, length, width, mass, observation, recognition, Dtemplate
    Dim ResTemp As String
    Set cItemCk = New c_itemCheck3D
        cItemCk.Cote = "H-01"
        cItemCk.Check = True
        cItemCk.statut = "OK"

    Dim part_ec As c_ckPart
        Set part_ec = New c_ckPart
        Dim coll_PropertEC As c_Properties
        Set coll_PropertEC = New c_Properties
    For i = 1 To cGriNues.Count
        Set part_ec = cGriNues.Item(i)
        Set coll_PropertEC = part_ec.xProperties
        ResTemp = AllPropExist(coll_PropertEC, part_ec.nom)
        If ResTemp = "OK" Then
            cItemCk.statut = "OK"
        Else
            cItemCk.Comment = cItemCk.Comment & ResTemp & Chr(10)
            cItemCk.statut = "KO"
        End If
    Next i

    'Renvoi l'objet crée
    Set CK_H01 = cItemCk
    'Liberation des classes
    Set cItemCk = Nothing
    Set part_ec = Nothing
    Set coll_PropertEC = Nothing
End Property

Property Get CK_H02() As c_itemCheck3D
'Le champ "SOURCE" est renseigné "FABRIQUE"/"MADE" ou "ACHETE"/"BOUGHT"
    Set cItemCk = New c_itemCheck3D
        cItemCk.Cote = "H-02"
        cItemCk.Check = True
        cItemCk.statut = "OK"
    Dim part_ec As c_ckPart
        Set part_ec = New c_ckPart
    For i = 1 To cGriNues.Count
        Set part_ec = cGriNues.Item(i)
        If part_ec.chpSource = 1 Then
            cItemCk.Comment = cItemCk.Comment & part_ec.nom & " - Champs source = 'Inconnu'" & Chr(10)
            cItemCk.statut = "KO"
        Else
            cItemCk.Comment = cItemCk.Comment & part_ec.nom & " - Champs source = " & Choose(part_ec.chpSource, "Inconnu", "FABRIQUE", "ACHETE") & Chr(10)
        End If
    Next

    'Renvoi l'objet crée
    Set CK_H02 = cItemCk
    'Liberation des classes
    Set cItemCk = Nothing
    Set part_ec = Nothing
    i = 0
End Property

Property Get CK_H03() As c_itemCheck3D
'Le champ "DESCRIPTION" est rempli ("UNEQUIPPED TEMPLATE" pour la grille nue)
    Set cItemCk = New c_itemCheck3D
        cItemCk.Cote = "H-03"
        cItemCk.Check = True
        cItemCk.statut = "OK"
    Dim part_ec As c_ckPart
        Set part_ec = New c_ckPart
    For i = 1 To cGriNues.Count
        Set part_ec = cGriNues.Item(i)
        
        If part_ec.chpDescr = vDescGNue Or part_ec.chpDescr = vDescFloat Then
            cItemCk.Comment = cItemCk.Comment & part_ec.nom & " - Champs DESCRIPTION = " & part_ec.chpDescr & Chr(10)
        Else
            cItemCk.Comment = cItemCk.Comment & part_ec.nom & " - Champs DESCRIPTION = " & part_ec.chpDescr & Chr(10)
            cItemCk.statut = "KO"
        End If
    Next

    'Renvoi l'objet crée
    Set CK_H03 = cItemCk
    'Liberation des classes
    Set cItemCk = Nothing
    Set part_ec = Nothing
    i = 0
End Property

Property Get CK_H04() As c_itemCheck3D
'La matière est renseignée dans le champ "MATERIAL" et le traitement est renseigné dans le champ "OBSERVATION"
'le cas échéant. La  matière est appliquée à la CATPart et pas au corps de pièce.
    Set cItemCk = New c_itemCheck3D
        cItemCk.Cote = "H-04"
        cItemCk.Check = True
        cItemCk.statut = "OK"
    Dim part_ec As c_ckPart
        Set part_ec = New c_ckPart
    Dim Params_EC As c_Properties
    Dim Param_EC As c_Property
    For i = 1 To cGriNues.Count
        Set part_ec = cGriNues.Item(i)
        On Error Resume Next
        Set Params_EC = part_ec.xProperties
        cItemCk.Comment = cItemCk.Comment & part_ec.nom & Chr(10)
        If Err.Number <> 0 Then
            Err.Clear
        Else
            'Matiere
            Set Param_EC = Params_EC.Item(nPrmMaterial)
            cItemCk.Comment = cItemCk.Comment & "    paramètre MATERIAL sur part = " & Param_EC.Valeur & Chr(10) _
                                                & "         valeur DSCGP = " & DSCGP_EC.MatGrille & Chr(10)
            If Param_EC.Valeur <> DSCGP_EC.MatGrille Then
                cItemCk.statut = "KO"
            End If
            'Traitement Affiche les valeurs sans générer d'érreur
            Set Param_EC = Params_EC.Item("OBSERVATION")
            cItemCk.Comment = cItemCk.Comment & "    paramètre OBSERVATION sur part = " & Param_EC.Valeur & Chr(10) _
                                                & "         valeur DSCGP = " & DSCGP_EC.Observations & Chr(10)
        End If
        On Error GoTo 0
    Next i

    'Renvoi l'objet crée
    Set CK_H04 = cItemCk
    'Liberation des classes
    Set Param_EC = Nothing
    Set Params_EC = Nothing
    Set part_ec = Nothing
    Set cItemCk = Nothing

End Property

Property Get CK_H05() As c_itemCheck3D
'Matière de la grille en accord avec le DSCGP
    Set cItemCk = New c_itemCheck3D
        cItemCk.Cote = "H-05"
        cItemCk.Check = True
        cItemCk.statut = "OK"

    Dim part_ec As c_ckPart
        Set part_ec = New c_ckPart
    Dim Params_EC As c_Properties
    Dim Param_EC As c_Property
    For i = 1 To cGriNues.Count
        Set part_ec = cGriNues.Item(i)
        cItemCk.Comment = cItemCk.Comment & part_ec.nom & Chr(10)
        'Matiere
        cItemCk.Comment = cItemCk.Comment & "    Matière sur part = " & part_ec.NomMatiere & Chr(10) _
                                            & "         valeur DSCGP = " & DSCGP_EC.MatGrille & Chr(10)
        If part_ec.NomMatiere <> DSCGP_EC.MatGrille Then
            cItemCk.statut = "KO"
        End If
    Next i

    'Renvoi l'objet crée
    Set CK_H05 = cItemCk
    'Liberation des classes
    Set part_ec = Nothing
    Set cItemCk = Nothing

End Property

Property Get CK_H06() As c_itemCheck3D
'Propriétés des CATPart - Les autres champs sont correctement remplis (observation, recognition) (Dtemplate pour la grille nue)
    Set cItemCk = New c_itemCheck3D
        cItemCk.Cote = "H-06"
        cItemCk.Check = True
        cItemCk.statut = "OK"
    Dim part_ec As c_ckPart
        Set part_ec = New c_ckPart
    Dim Params_EC As c_Properties
    Dim Param_EC As c_Property
    For i = 1 To cGriNues.Count
        Set part_ec = cGriNues.Item(i)
        cItemCk.Comment = cItemCk.Comment & part_ec.nom & Chr(10)
        On Error Resume Next
        Set Params_EC = part_ec.xProperties
        If Err.Number <> 0 Then
            Err.Clear
            cItemCk.Comment = cItemCk.Comment & " Les paramètres (OBSERVATIONS, RECOGNITION ou DTEMPLATE) sont absents" & Chr(10)
        Else
            'param "OBSERVATIONS"
            Set Param_EC = Params_EC.Item(nPrmObserv)
            cItemCk.Comment = cItemCk.Comment & "    paramètre OBSERVATIONS sur part = " & Param_EC.Valeur & Chr(10)
            If Param_EC.Valeur <> DSCGP_EC.Observations Then
                        cItemCk.Comment = cItemCk.Comment & "    paramètre OBSERVATIONS sur part = " & Param_EC.Valeur & Chr(10) _
                                                & "         valeur DSCGP = " & DSCGP_EC.Observations & Chr(10)
                cItemCk.statut = "KO"
            End If
            'param "RECOGNITION"
            Set Param_EC = Params_EC.Item(nPrmRecogn)
            If Param_EC.Valeur = vRecogn Then
                cItemCk.Comment = cItemCk.Comment & "    paramètre RECOGNITION sur part = " & Param_EC.Valeur & Chr(10)
            Else
                cItemCk.Comment = cItemCk.Comment & "    paramètre RECOGNITION sur part = " & Param_EC.Valeur & Chr(10) _
                                                    & "         au lieu de : " & vRecogn & Chr(10)
                cItemCk.statut = "KO"
            End If
            'param "DTEMPLATE"
            Set Param_EC = Params_EC.Item(nPrmDtempl)
            cItemCk.Comment = cItemCk.Comment & "    paramètre DTEMPLATE sur part = " & Param_EC.Valeur & Chr(10)
            If Param_EC.Valeur <> DSCGP_EC.Dtemplate Then
                        cItemCk.Comment = cItemCk.Comment & "    paramètre DTEMPLATE sur part = " & Param_EC.Valeur & Chr(10) _
                                                & "         valeur DSCGP = " & DSCGP_EC.Dtemplate & Chr(10)
                cItemCk.statut = "KO"
            End If
        End If
        On Error GoTo 0
    Next

    'Renvoi l'objet crée
    Set CK_H06 = cItemCk
    'Liberation des classes
    Set Param_EC = Nothing
    Set Params_EC = Nothing
    Set part_ec = Nothing
    Set cItemCk = Nothing
    i = 0
End Property

Property Get CK_H07() As c_itemCheck3D
'La masse est renseignée dans le champ "MASS". La masse mesurée est cohérente à ±0,25kg
    Set cItemCk = New c_itemCheck3D
        cItemCk.Cote = "H-07"
        cItemCk.Check = True
        cItemCk.statut = "OK"
    Dim part_ec As c_ckPart
        Set part_ec = New c_ckPart
    Dim Params_EC As c_Properties
    Dim Param_EC As c_Property
    Dim ValMasse As Double
    For i = 1 To cGriNues.Count
        Set part_ec = cGriNues.Item(i)
        On Error Resume Next
        Set Params_EC = part_ec.xProperties
        If Err.Number <> 0 Then
            Err.Clear
        Else
            ValMasse = ConvMasse(Params_EC.Item("MASS").Valeur)
            If Params_EC.Item("MASS").Valeur = "" Then
                cItemCk.Comment = cItemCk.Comment & part_ec.nom & Chr(10)
                cItemCk.Comment = cItemCk.Comment & "    Paramètre 'MASS' du Part non/mal renseigné : " & Params_EC.Item("MASS").Valeur & Chr(10)
            ElseIf Abs(ValMasse - part_ec.Masse) > 0.2 Then
                cItemCk.Comment = cItemCk.Comment & part_ec.nom & Chr(10)
                cItemCk.Comment = cItemCk.Comment & "    Paramètre 'MASS' du Part = " & Params_EC.Item("MASS").Valeur & Chr(10)
                cItemCk.Comment = cItemCk.Comment & "    Masse mesurée du part = " & Round(part_ec.Masse, 3) & Chr(10)
                cItemCk.Comment = cItemCk.Comment & "    Ecart = " & Round(Abs(ValMasse - part_ec.Masse), 3) & Chr(10)
                cItemCk.statut = "KO"
            End If
    
        End If
        On Error GoTo 0
    Next

    'Renvoi l'objet crée
     Set CK_H07 = cItemCk
    'Liberation des classes
    Set Param_EC = Nothing
    Set Params_EC = Nothing
    Set part_ec = Nothing
    Set cItemCk = Nothing

End Property

Property Get CK_H08() As c_itemCheck3D
'Les champs "LENGTH", "WIDTH", "THICKNESS/DIAMETER" sont laissés vides
    Set cItemCk = New c_itemCheck3D
        cItemCk.Cote = "H-08"
        cItemCk.Check = True
        cItemCk.statut = "OK"

    Dim part_ec As c_ckPart
        Set part_ec = New c_ckPart
    Dim Params_EC As c_Properties
    Dim Param_EC As c_Property
    Dim tComment As String
    For i = 1 To cGriNues.Count
        tComment = ""
        Set part_ec = cGriNues.Item(i)
        On Error Resume Next
        Set Params_EC = part_ec.xProperties
        If Err.Number <> 0 Then
            Err.Clear
        Else
            If Params_EC.Item("LENGTH").Valeur <> "" Then
                tComment = tComment & "    - LENGTH = " & Params_EC.Item("LENGTH").Valeur & Chr(10)
                cItemCk.statut = "KO"
            End If
    
            If Params_EC.Item("WIDTH").Valeur <> "" Then
                tComment = tComment & "    - WIDTH = " & Params_EC.Item("WIDTH").Valeur & Chr(10)
                cItemCk.statut = "KO"
            End If
            
            If Params_EC.Item("THICKNESS/DIAMETER").Valeur <> "" Then
                tComment = tComment & "    - THICKNESS/DIAMETER = " & Params_EC.Item("THICKNESS/DIAMETER").Valeur & Chr(10)
                cItemCk.statut = "KO"
            End If
            If cItemCk.statut = "KO" Then
                cItemCk.Comment = cItemCk.Comment & part_ec.nom & " Paramètres non vide :" & Chr(10) & tComment
            End If
        End If
        On Error GoTo 0
    Next

    'Renvoi l'objet crée
    Set CK_H08 = cItemCk
    'Liberation des classes
    Set Param_EC = Nothing
    Set Params_EC = Nothing
    Set part_ec = Nothing
    Set cItemCk = Nothing
    
End Property

Property Get CK_I01() As c_itemCheck3D
'Les champs suivants sont présents dans les propriétés du Product :
'       material, thickness, length, width, mass, observation, recognition, Dtemplate
    Dim ResTemp As String
    Set cItemCk = New c_itemCheck3D
        cItemCk.Cote = "I-01"
        cItemCk.Check = True
        cItemCk.statut = "OK"

    Dim coll_PropertEC As c_Properties
    Set coll_PropertEC = New c_Properties

    Set coll_PropertEC = GrilleASs_EC.xProperties
    ResTemp = AllPropExist(coll_PropertEC, GrilleASs_EC.Numero)
    If ResTemp = "OK" Then
        cItemCk.statut = "OK"
    Else
        cItemCk.Comment = cItemCk.Comment & ResTemp & Chr(10)
        cItemCk.statut = "KO"
    End If

    'Renvoi l'objet crée
    Set CK_I01 = cItemCk
    'Liberation des classes
    Set cItemCk = Nothing
    Set coll_PropertEC = Nothing
    
End Property

Property Get CK_I02() As c_itemCheck3D
'Le champ "SOURCE" est renseigné "FABRIQUE" ou "ACHETE"
    Set cItemCk = New c_itemCheck3D
        cItemCk.Cote = "I-02"
        cItemCk.Check = True
        cItemCk.statut = "OK"

    If GrilleASs_EC.prm_Source = 1 Then
        cItemCk.Comment = "Product : " & GrilleASs_EC.Numero & " - Champs source : Inconnu "
        cItemCk.statut = "KO"
    Else
        cItemCk.Comment = "Champs source : " & Choose(GrilleASs_EC.prm_Source, "Inconnu", "FABRIQUE", "ACHETE")
    End If

    'Renvoi l'objet crée
    Set CK_I02 = cItemCk
    'Liberation des classes
    Set cItemCk = Nothing

End Property

Property Get CK_I03() As c_itemCheck3D
'Le champ "DESCRIPTION" doit contenir l'intitulé de la grille défini dans le DSCGP. Les autres products ont une brève description
On Error GoTo Erreur
    Set cItemCk = New c_itemCheck3D
        cItemCk.Cote = "I-03"
        cItemCk.Check = True
        cItemCk.statut = "OK"

    If GrilleASs_EC.Prm_DescriptionRef <> DSCGP_EC.DesignGrille Then
        cItemCk.Comment = "Product : " & GrilleASs_EC.Numero & " designation incorrecte" & Chr(10) _
                            & "         DescriptionRef sur Part = " & GrilleASs_EC.Prm_DescriptionRef & Chr(10) _
                            & "         valeur DSCGP = " & DSCGP_EC.DesignGrille & Chr(10)
        cItemCk.statut = "KO"
    End If
    GoTo Fin
Erreur:
    cItemCk.statut = "KO"
    If Err.Number > vbObjectError + 512 Then
        cItemCk.Comment = Err.Description
    Else
        cItemCk.Comment = "Erreur dans la récupération du nom de l'assemblage"
    End If
Fin:
    'Renvoi l'objet crée
    Set CK_I03 = cItemCk
    'Liberation des classes
    Set cItemCk = Nothing

End Property

Property Get CK_I04() As c_itemCheck3D
'Le champ "MASS" est rempli. Pour la grille assemblée, il est renseigné sans tenir compte des éléments non solidaires de la grille
    Set cItemCk = New c_itemCheck3D
        cItemCk.Cote = "I-04"
        cItemCk.Check = True
        cItemCk.statut = "OK"
    Dim vMass As String
    
    On Error Resume Next
    vMass = GrilleASs_EC.xProperties.Item("MASS").Valeur
    If Err.Number <> 0 Then
        Err.Clear
        cItemCk.Comment = "Product : " & GrilleASs_EC.Numero & " - attribut MASS absent"
        cItemCk.statut = "KO"
        On Error GoTo 0
    Else
        If vMass = "" Then
            cItemCk.Comment = "Product : " & GrilleASs_EC.Numero & " - MASS : non renseigné"
            cItemCk.statut = "KO"
        End If
    End If
    
    'Renvoi l'objet crée
    Set CK_I04 = cItemCk
    'Liberation des classes
    Set cItemCk = Nothing

End Property

Property Get CK_J01() As c_itemCheck3D
'Les coordonnées Xe,Ye et Ze renseignées dans le "fastener" sont identiques à celles mesurées pour le point d'attaque de ce "fastener" par rapport au repère 3D CATIA
    Set cItemCk = New c_itemCheck3D
        cItemCk.Cote = "J-0x"
        cItemCk.Check = True
        cItemCk.statut = "measured"

    Dim part_ec As c_ckPart
        Set part_ec = New c_ckPart
    'Dim tMes As c_mesure
    '    Set tMes = New c_mesure
    For i = 1 To cGriNues.Count
        Set part_ec = cGriNues.Item(i)
        cItemCk.Comment = cItemCk.Comment & part_ec.nom & "Voir Onglet 'Mesures'" & Chr(10)
    Next i

    'Renvoi l'objet crée
    Set CK_J01 = cItemCk
    'Liberation des classes
    'Set tMes = Nothing
    Set cItemCk = Nothing
    Set part_ec = Nothing
End Property

Property Get CK_N01() As c_itemCheck3D
    'test si le part est en conception hybride
    'Tente de créer un point dans le corp de pièce (l'efface ensuite)
    'et détecte l'erreur
    Set cItemCk = New c_itemCheck3D
        cItemCk.Cote = "N-01"
        cItemCk.Check = True
        cItemCk.statut = "OK"

    Dim part_ec As c_ckPart
        Set part_ec = New c_ckPart
    Dim tMes As c_mesure
        Set tMes = New c_mesure
    Dim mBody As Body
    
    For i = 1 To cGriNues.Count
        Set part_ec = cGriNues.Item(i)
        Set mBody = part_ec.ctPart.MainBody
        If mBody.HybridShapes Is Nothing Then
            cItemCk.Comment = cItemCk.Comment & part_ec.nom & " n'est pas en mode hybride" & Chr(10)
        Else
            cItemCk.statut = "KO"
            cItemCk.Comment = cItemCk.Comment & part_ec.nom & " est en mode hybride !" & Chr(10)
        End If
    Next i

    'Renvoi l'objet crée
    Set CK_N01 = cItemCk
    'Liberation des classes
    Set part_ec = Nothing
    Set cItemCk = Nothing

End Property

Property Get CK_N02() As c_itemCheck3D
'Les sets géométriques suivants sont présents dans la CatPart : références externes isolées, surf0, pinules, feet, pointsA, std, surf100, pointsB, gravure, travail, points de construction
'(tous en No SHOW sauf "gravure")
    Set cItemCk = New c_itemCheck3D
        cItemCk.Cote = "N-02"
        cItemCk.Check = True
        cItemCk.statut = "OK"
    Dim ResTemp As String
    Dim part_ec As c_ckPart
    Set part_ec = New c_ckPart
    For i = 1 To cGriNues.Count
        Set part_ec = cGriNues.Item(i)
        
        ResTemp = AllSetExist(part_ec.xBodies, part_ec.nom)
        If ResTemp = "OK" Then
            cItemCk.statut = "OK"
        Else
            cItemCk.Comment = cItemCk.Comment & ResTemp & Chr(10)
            cItemCk.statut = "KO"
        End If
        
        ResTemp = ""
     Next i
    'Renvoi l'objet crée
    Set CK_N02 = cItemCk
    'Liberation des classes
    Set cItemCk = Nothing
    Set part_ec = Nothing
    
End Property

Property Get CK_N04() As c_itemCheck3D
'Les points de pinnules "P" sant dans le set "pinules"
    Set cItemCk = New c_itemCheck3D
        cItemCk.Cote = "N-04"
        cItemCk.Check = True
        cItemCk.statut = "OK"

    Dim mBody As HybridBody
    Dim mShapes As HybridShapes
    Dim mShape As HybridShape
    
    Dim part_ec As c_ckPart
    Set part_ec = New c_ckPart
    Dim tComment As String
    For i = 1 To cGriNues.Count
        Set part_ec = cGriNues.Item(i)
        If part_ec.xBodies Is Nothing Then
            Exit For
        End If
        For j = 1 To part_ec.xBodies.Count
            Set mBody = part_ec.xBodies.Item(j).HBody
            If mBody.HybridShapes Is Nothing Then
            Else
                Set mShapes = mBody.HybridShapes
                If mShapes.Count > 0 Then
                    For Each mShape In mShapes
                        If UCase(Left(mShape.Name, 1)) = "P" And Len(mShape.Name) < 4 Then
                            If mBody.Name <> nHBPin Then
                                tComment = "     - le point " & mShape.Name & " est dans le set : " & mBody.Name & Chr(10)
                                cItemCk.statut = "KO"
                            End If
                        End If
                    Next
                End If
            End If
            If cItemCk.statut = "KO" Then
                cItemCk.Comment = "Erreur de set géométrique dans le part " & part_ec.nom & Chr(10) & tComment
            End If
        Next
    Next i

    'Renvoi l'objet crée
    Set CK_N04 = cItemCk
    'Liberation des classes
    Set cItemCk = Nothing
    Set part_ec = Nothing

End Property

Property Get CK_N06() As c_itemCheck3D
'Les points de contrôle "F" des patins sont dans le set "feet"
    Set cItemCk = New c_itemCheck3D
        cItemCk.Cote = "N-06"
        cItemCk.Check = True
        cItemCk.statut = "OK"

    Dim mBody As HybridBody
    Dim mShapes As HybridShapes
    Dim mShape As HybridShape
    
    Dim part_ec As c_ckPart
    Set part_ec = New c_ckPart
    Dim tComment As String
    For i = 1 To cGriNues.Count
        Set part_ec = cGriNues.Item(i)
        If part_ec.xBodies Is Nothing Then
            Exit For
        End If
        For j = 1 To part_ec.xBodies.Count
            Set mBody = part_ec.xBodies.Item(j).HBody
            If mBody.HybridShapes Is Nothing Then
            Else
                Set mShapes = mBody.HybridShapes
                If mShapes.Count > 0 Then
                    For Each mShape In mShapes
                        If UCase(Left(mShape.Name, 1)) = "F" And Len(mShape.Name) < 4 Then
                            If mBody.Name <> nHBFeet Then
                                tComment = "     - le point " & mShape.Name & " est dans le set : " & mBody.Name & Chr(10)
                                cItemCk.statut = "KO"
                            End If
                        End If
                    Next
                End If
            End If
            If cItemCk.statut = "KO" Then
                cItemCk.Comment = "Erreur de set géométrique dans le part " & part_ec.nom & Chr(10) & tComment
            End If
        Next
    Next i

    'Renvoi l'objet crée
    Set CK_N06 = cItemCk
    'Liberation des classes
    Set cItemCk = Nothing
    Set part_ec = Nothing

End Property

Property Get CK_O06() As c_itemCheck3D
'Si la grille possède des pinnules : quantité en accord avec le DSCGP
    Set cItemCk = New c_itemCheck3D
        cItemCk.Cote = "O-06"
        cItemCk.Check = True
        cItemCk.statut = "OK"

    Dim part_ec As c_ckPart
    Set part_ec = New c_ckPart
    Dim NomPinule As String
    Dim NbPinules As Integer
For j = 1 To cGriNues.Count
    Set part_ec = cGriNues.Item(1)
    If part_ec.xBodies.Count > 0 Then
        If Not (part_ec.xBodies.Item(nHBPin) Is Nothing) Then
            For i = 1 To part_ec.xBodies.Item(nHBPin).HBody.HybridShapes.Count
                If UCase(Left(part_ec.xBodies.Item(nHBPin).HBody.HybridShapes.Item(i).Name, 1)) = "P" Then
                    NbPinules = NbPinules + 1
                End If
            Next i
        End If
    End If
    If NbPinules <> DSCGP_EC.Nb_Pinules Then
        cItemCk.Comment = "Nb de pinules incorrect." & Chr(10)
        cItemCk.Comment = cItemCk.Comment & "    - valeur DSCGP = " & DSCGP_EC.Nb_Pinules & Chr(10)
        cItemCk.Comment = cItemCk.Comment & "    - nombre de pinules trouvées dans le set " & nHBPin & " = " & NbPinules
        cItemCk.statut = "KO"
    Else
        cItemCk.Comment = "Nb de pinules correct." & Chr(10)
        cItemCk.Comment = cItemCk.Comment & "    - valeur DSCGP = " & DSCGP_EC.Nb_Pinules & Chr(10)
        cItemCk.Comment = cItemCk.Comment & "    - nombre de pinules trouvées dans le set " & nHBPin & " = " & NbPinules
        cItemCk.statut = "OK"
    End If
Next
    'Renvoi l'objet crée
    Set CK_O06 = cItemCk
    'Liberation des classes
    Set cItemCk = Nothing
    Set part_ec = Nothing
    
End Property

'Property Get CK_G0x() As c_itemCheck3D
''*********************
'    Set cItemCk = New c_itemCheck3D
'        cItemCk.Cote = "G-0x"
'        cItemCk.Check = True
'        cItemCk.Statut = "OK"
'
''###########
''#
''# a documenter
''#
''##########
'
'    'Renvoi l'objet crée
'    Set CK_G0x = cItemCk
'    'Liberation des classes
'    Set cItemCk = Nothing
'
'End Property

Private Function EstGriNue(str As String, itg As Integer) As Boolean
'Test si c'est une grille.
'Les grille sont numéroyé de 2 à 8999 sur les 4 derniers digit ou sur les 12 à 16eme selon les system de numérotation
'Les grille A350 standard (Broughton, Nantes, Saint-Nazaire) ont un "P" en 17eme pos
    Select Case itg
    Case 1, 11
        If Interv(str, 12, 5, 2, 8999) And Mid(str, 17, 1) = "P" Then
            EstGriNue = True
        Else
            EstGriNue = False
        End If
    Case 2, 3, 12, 13
        If Interv(str, Len(str) - 4, 4, 2, 8999) Then
            EstGriNue = True
        Else
            EstGriNue = False
        End If
    Case Else
        EstGriNue = False
    End Select
End Function

Private Function getDocNom(mVar As String) As Document
'renvoi le document "nomdoc" de la collection des documents chargés
    Set coll_docs = CATIA.Documents
    Dim DocTemp As Document
    On Error Resume Next
    Set DocTemp = coll_docs.Item(mVar)
    If Err.Number <> 0 Then
        Err.Clear
        Set getDocNom = Nothing
    Else
        Set getDocNom = DocTemp
    End If
    On Error GoTo 0

End Function

'#############################################################
'# Fonctions de controle des set géométrique
'#############################################################
Private Function AllSetExist(col_set As c_Hbodies, str As String) As String
'Vérifie que tous les set géométrique
'sont présents
'str = Nom du part
Init_ColSetAirbus

Dim Set_EC As Variant
Dim TempCheck As String
    TempCheck = "OK"
Dim TempResult As String
    TempResult = str & Chr(10)
Dim TempHB As c_Hbody
Set TempHB = New c_Hbody
    For Each Set_EC In Col_SetAirbus
        On Error Resume Next
        Set TempHB = col_set.Item(Set_EC)
        'teste la présence du set géométrique
        If Err.Number <> 0 Then
            Err.Clear
            TempCheck = "KO"
            TempResult = TempResult & " - Set géométrique manquants : " & Set_EC & Chr(10)
        Else
            On Error GoTo 0
            'Test si les set géométriques sont masqués
            If TempHB.ShowStatus = catVisPropertyShowAttr Then
                If TempHB.nom <> nHBGrav Then 'sauf le set gravure qui peux resté affiché
                    TempCheck = "KO"
                    TempResult = TempResult & " - Set géométrique en Show : " & Set_EC & Chr(10)
                End If
            End If
        End If
    Next
    If TempCheck = "OK" Then
        AllSetExist = "OK"
    Else
        AllSetExist = TempResult
    End If
Set TempHB = Nothing
Set Col_SetAirbus = Nothing
End Function

'#############################################################
'# Fonctions de controle des Propriétés
'#############################################################
Private Function AllPropExist(col_prop As c_Properties, str As String) As String
'Vérifie que toutes les propriétés
'material, thickness, length, width, mass, observation, recognition et Dtemplate sont présentes
'Str =  Nom du Part
init_ColPropAirbus
Dim Prop_EC As Variant

Dim TempCheck As String
    TempCheck = "OK"
Dim TempResult As String
    TempResult = str & " - Propriétes manquantes : " & Chr(10)
Dim TempProp As c_Property
Set TempProp = New c_Property
    For Each Prop_EC In Col_PropAirbus
        On Error Resume Next
        Set TempProp = col_prop.Item(Prop_EC)
        If Err.Number <> 0 Then
            Err.Clear
            TempCheck = "KO"
            TempResult = TempResult & "   . " & Prop_EC & Chr(10)
        End If
        On Error GoTo 0
    Next
    If TempCheck = "OK" Then
        AllPropExist = "OK"
    Else
        AllPropExist = TempResult
    End If
Set TempProp = Nothing
Set Col_PropAirbus = Nothing
End Function


'#############################################################
'# Fonctions de controle numérotation des parts/product
'#############################################################

Private Function NumGriOK(strNom As String, intgr As Integer, strNumlot As String) As Boolean
'Vérifie que le nom passé en argument est conforme à la numérotation d'une part de grille nue
'strNom = partnumber du part
'strNumlot = N° du Lot dans le DSCGP

Dim RepGr As Integer
NumGriOK = True
    Select Case intgr
        Case 1, 11
            'Nombre de digit
            If Not Nbdigit(strNom, 19) Then
                NumGriOK = False
            End If
            'Corrélation avec le numéro de lot
            If Not EQNumlot(strNom, strNumlot) Then
                NumGriOK = False
            End If
            'Tiret en 11eme position
            If Not CarSpe(strNom, "-", 11) Then
                NumGriOK = False
            End If
            '"P" en 17eme position
            If Not CarSpe(strNom, "P", 17) Then
                NumGriOK = False
            End If
            '12 à 16eme carractère compris entre 2 et 8999
            If Not Interv(strNom, 12, 16, 2, 8999) Then
                NumGriOK = False
            End If

        Case 2, 12
            'Nombre de digit
            If Not Nbdigit(strNom, 18) Then
                NumGriOK = False
            End If
            'Corrélation avec le numéro de lot
            If Not EQNumlot(strNom, strNumlot) Then
                NumGriOK = False
            End If
            
        Case 3, 13
            'Nombre de digit
            If Not Nbdigit(strNom, 18) Then
                NumGriOK = False
            End If
            'Corrélation avec le numéro de lot
            'Exepté le 5eme Carractère
            If Not EQNumlot(Left(strNom, 4), Left(strNumlot, 4)) Then
                NumGriOK = False
            End If
            If Not EQNumlot(Mid(strNom, 6, Len(strNom) - 6), Mid(strNumlot, 6, Len(strNumlot) - 6)) Then
                NumGriOK = False
            End If
    End Select



End Function

Private Function Nbdigit(ByVal str As String, ByVal Nb As Integer) As Boolean
'Renvois true ou false si la longueur de str est égale a Nb
    If Len(str) = Nb Then
        Nbdigit = True
    Else
        Nbdigit = False
    End If
End Function

Private Function EQNumlot(ByVal strGri As String, strLot As String) As Boolean
'Renvois true ou false si le debut de la chaine strGri est égale à la chaine strLot
    If Left(strGri, Len(strLot)) = strLot Then
        EQNumlot = True
    Else
        EQNumlot = False
    End If
End Function

Private Function CarSpe(ByVal str As String, car As String, pos As Integer) As Boolean
'Renvoi true si le caractère "Car" est présent à la position "pos" de la chaine "str"
    If Mid(str, pos, 1) = car Then
        CarSpe = True
    Else
        CarSpe = False
    End If
End Function

Private Function ConvMasse(str As String) As Double
'converti la chaine "xxKg" en valeur xx
'Renvoi 0 si la masse est illisible
Dim tval As Double
    On Error Resume Next
    str = Replace(str, ",", ".", 1, , vbTextCompare)
    tval = val(str)
    If Err.Number <> 0 Then
        Err.Clear
        ConvMasse = 0
    Else
        ConvMasse = tval
    End If
    On Error GoTo 0
End Function

Private Function LstFasteners(partEc As part) As c_Fasteners
'Renvoi la collection des fasterners
Dim tmpFasts As c_Fasteners
Set tmpFasts = New c_Fasteners
Dim tmpFast As c_Fastener
Set tmpFast = New c_Fastener
Dim tmp_Param As Parameter
Dim HBRefextIsol As HybridBody
Set HBRefextIsol = partEc.HybridBodies.Item(nHBRefExtIsol)
Dim tFastEC As HybridShape


        For i = 1 To HBRefextIsol.HybridShapes.Count
            Set tFastEC = HBRefextIsol.HybridShapes.Item(i)
            'test si c'est un "Fastener"
            If TypeName(tFastEC) = "HybridShapeInstance" Then
                'Collecte les attibuts
                tmpFast.nom = tFastEC.Name
                Set tmp_Param = tFastEC.GetParameter("Comments")
                tmpFast.Comments = tmp_Param.Value
                Set tmp_Param = tFastEC.GetParameter("Fastener_Diameter")
                tmpFast.FastDiam = tmp_Param.Value
                Set tmp_Param = tFastEC.GetParameter("Xe")
                tmpFast.Xe = tmp_Param.Value
                Set tmp_Param = tFastEC.GetParameter("Ye")
                tmpFast.Ye = tmp_Param.Value
                Set tmp_Param = tFastEC.GetParameter("Ze")
                tmpFast.Ze = tmp_Param.Value
                Set tmp_Param = tFastEC.GetParameter("Xdir")
                tmpFast.Xdir = tmp_Param.Value
                Set tmp_Param = tFastEC.GetParameter("Ydir")
                tmpFast.Ydir = tmp_Param.Value
                Set tmp_Param = tFastEC.GetParameter("Zdir")
                tmpFast.Zdir = tmp_Param.Value
                'Collecte des coordonnnées
                tmpFast.X = CoordFast(partEc, tFastEC, "X")
                tmpFast.Y = CoordFast(partEc, tFastEC, "Y")
                tmpFast.Z = CoordFast(partEc, tFastEC, "Z")
                
               tmpFasts.Add tmpFast.nom, tmpFast.Comments, tmpFast.FastDiam, tmpFast.Xe, tmpFast.Ye, tmpFast.Ze, tmpFast.Xdir, tmpFast.Ydir, tmpFast.Zdir, tmpFast.X, tmpFast.Y, tmpFast.Z
            End If
        Next

    Set LstFasteners = tmpFasts
    'liberation des objets
    Set HBRefextIsol = Nothing
    Set tmpFast = Nothing
    Set tmpFasts = Nothing
End Function

Private Sub LstMesures(tpart As part, cFasteners As c_Fasteners)
'Construit la liste des mesures des Fasteneres, des Pts et des trous
    Dim HB_EC As HybridBody
    Dim HS_Ec As HybridShape
    Dim cFast As c_Fastener
    Dim cMes As c_mesure
    Dim nSearch As String
    Dim lNoPt As Integer
    Dim coPta As Coord
    
    Set cFast = New c_Fastener
    Set cMes = New c_mesure
    
    'Collecte des mesures sur les fasteners
    Set HB_EC = tpart.HybridBodies.Item(nHBRefExtIsol)
    For i = 1 To cFasteners.Count
        Set cFast = cFasteners.Item(i)
        col_Mes.Add cFast.nom, cFast.X, cFast.Y, cFast.Z, CDbl(cFast.Xe), CDbl(cFast.Ye), CDbl(cFast.Ze)
    Next i
    
    'Collecte des mesures des PtsA
    Set HB_EC = tpart.HybridBodies.Item(nHBPtA)
    For i = 1 To HB_EC.HybridShapes.Count
        Set HS_Ec = HB_EC.HybridShapes.Item(i)
        With coPta
            .X = CoordPt(tpart, HS_Ec, "X")
            .Y = CoordPt(tpart, HS_Ec, "Y")
            .Z = CoordPt(tpart, HS_Ec, "Z")
        End With
        
        'Extraction du nom du Fastener dans le nom du PointA
        lNoPt = Len("A" & i) 'A1 = 2 A27 = 3
        If Len(HS_Ec.Name) >= lNoPt Then 'il y a un complémnt apres le N° du point ex A27-blablabla
            lNoPt = lNoPt + 1 'ajout du tiret
            nSearch = Right(HS_Ec.Name, Len(HS_Ec.Name) - lNoPt)
            'recherhe du fastener dans la collection des mesures
            On Error Resume Next
            Set cMes = col_Mes.Item(nSearch)
            If Err.Number <> 0 Then ' la recherche a echouée
                Err.Clear
                Set cMes = SearcFastCoord(coPta, HS_Ec.Name) 'Recherche du fastener par les coordonnée du point A
                On Error GoTo 0
'##Recherche par le "comment"
'## a developper
            Else
                'Ajout des coordonnées du point a la mesure
                With col_Mes.Item(nSearch)
                    .NomPtA = HS_Ec.Name
                    .PtAX = coPta.X
                    .PtAY = coPta.Y
                    .PtAZ = coPta.Z
                End With
            End If
        Else 'pas des complement apres le N0 de point ex : A27
            'Recherche par les coordonnées du point A
            Set cMes = SearcFastCoord(coPta, HS_Ec.Name)
        End If
        If cMes.NomPtA = "" Then
            'Ajout d'un item mesure
            col_Mes.Add "", 0, 0, 0, 0, 0, 0, HS_Ec.Name, coPta.X, coPta.Y, coPta.Z
        Else
            'Ajout des coordonnées du point a la mesure
                With col_Mes.Item(cMes.nom)
                    .NomPtA = HS_Ec.Name
                    .PtAX = coPta.X
                    .PtAY = coPta.Y
                    .PtAZ = coPta.Z
                End With
        End If
    Set cMes = Nothing
    Next i
    'Libération des classes
    
    Set cFast = Nothing
    
End Sub

Private Function SearcFastCoord(cpt As Coord, str As String) As c_mesure
'Recherche un fastener dont les coordonnées correspondent a cpt
Dim i As Long
Set SearcFastCoord = New c_mesure
For i = 1 To col_Mes.Count
    If Round(Abs(col_Mes.Item(i).Xe - cpt.X), 3) < 0.2 Then
        If Round(Abs(col_Mes.Item(i).Ye - cpt.Y), 3) < 0.2 Then
            If Round(Abs(col_Mes.Item(i).Ze - cpt.Z), 3) < 0.2 Then
                Set SearcFastCoord = col_Mes.Item(i)
                SearcFastCoord.NomPtA = str
                SearcFastCoord.PtAX = cpt.X
                SearcFastCoord.PtAY = cpt.Y
                SearcFastCoord.PtAZ = cpt.Z
                Exit For
            End If
        End If
    End If
Next i

End Function
Private Function CoordFast(tpart As part, tHSEC As HybridShape, Axe As String) As Double
'Renvoi la coordonnées X, Y ou Z mesurée du fastener passé en argument
Dim spa_workbench As SPAWorkbench
Dim mMes 'as Measurable
Dim mRef As Reference
Dim Pt(8)

    Set spa_workbench = CATIA.ActiveDocument.GetWorkbench("SPAWorkbench")
    Set mRef = tpart.CreateReferenceFromObject(tHSEC)
    Set mMes = spa_workbench.GetMeasurable(mRef)
    mMes.GetPointsOnCurve Pt
    Select Case Axe
            Case "X"
                CoordFast = Pt(0)
            Case "Y"
                CoordFast = Pt(1)
            Case "Z"
                CoordFast = Pt(2)
    End Select
End Function

Private Function CoordPt(tpart As part, tHSEC As HybridShape, Axe As String) As Double
'Renvoi la coordonnées X, Y ou Z mesurée du point passé en argument
Dim spa_workbench As SPAWorkbench
Dim mMes 'as Measurable
Dim mRef As Reference
Dim Pt(8)

    Set spa_workbench = CATIA.ActiveDocument.GetWorkbench("SPAWorkbench")
    Set mRef = tpart.CreateReferenceFromObject(tHSEC)
    Set mMes = spa_workbench.GetMeasurable(mRef)
    mMes.GetPoint Pt
    Select Case Axe
            Case "X"
                CoordPt = Pt(0)
            Case "Y"
                CoordPt = Pt(1)
            Case "Z"
                CoordPt = Pt(2)
    End Select
End Function

Private Sub Class_Terminate()
'Liberation des sous classes
Set DSCGP_EC = Nothing
Set GrilleASs_EC = Nothing
End Sub
