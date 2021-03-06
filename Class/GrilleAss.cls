VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GrilleAss"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*********************************************************************
'* Classe : GrilleASS
'*
'* Fonctions :  Teste la pr�sence des �l�ments normalis� des grilles assembl�es
'*
'*
'* Version : 1
'* Cr�ation :  CFR
'* Modification : 16/07/15
'*
'**********************************************************************
Option Explicit
Private mvarProdDoc As ProductDocument
Private mvarProduct As Product
Private mvarProducts As Products
Private mvarPNumber As String
Private mvarContraintes As Constraints
Private mvarProductParametres As Parameters
Private mvarNomGrille As String
Private mvarNomGrilleSym As String
Private mvarNumLot As String
Private mvarSysNum As Integer
Private mvarParam As Parameter
Private mColProperties As c_Properties  'Collection des attibuts Grilles (x......)

Private i As Long

'===========================
' Initialisation de l'objet
'===========================

Private Sub Class_initialize()

    'On Error Resume Next

    Set mvarProdDoc = CATIA.ActiveDocument
    Set mvarProduct = mvarProdDoc.Product
    Set mvarProducts = mvarProduct.Products
    mvarPNumber = mvarProduct.PartNumber
    Set mvarContraintes = mvarProduct.Connections("CATIAConstraints")
    Set mvarProductParametres = mvarProduct.ReferenceProduct.UserRefProperties
    Set mColProperties = New c_Properties
    'Construit la collection des propri�t�s
    For Each mvarParam In mvarProductParametres
        mColProperties.Add DecoupeSlash(mvarParam.Name), mvarParam.Value
    Next
End Sub

Property Let ProductDocGrille(ByVal ProdDoc As ProductDocument)
' R�initialisation de l'objet

    Set mvarProdDoc = ProdDoc
    Set mvarProduct = mvarProdDoc.Product
    Set mvarProducts = mvarProduct.Products
    mvarPNumber = mvarProduct.PartNumber
    Set mvarContraintes = mvarProduct.Connections("CATIAConstraints")
    Set mvarProductParametres = mvarProduct.ReferenceProduct.UserRefProperties
    Set mColProperties = New c_Properties
    'Construit la collection des propri�t�s
    For Each mvarParam In mvarProductParametres
        mColProperties.Add DecoupeSlash(mvarParam.Name), mvarParam.Value
    Next
      
End Property

Public Property Get ProductDocGrille() As ProductDocument
'Renvoi le document de la grille assembl�e
    Set ProductDocGrille = mvarProdDoc
End Property

Property Get Produits() As Products
'renvois la collection des products de la grille assembl�e
    Set Produits = mvarProducts
End Property

Property Get Produit() As Product
'Renvoi le produit de la grille assembl�e
    Set Produit = mvarProduct
End Property

Property Get Col_Const() As c_Contraintes
'Renvoi la collection des contraites de la grille assembl�e
    Set Col_Const = List_Contraintes(mvarContraintes)

End Property

Private Function List_Contraintes(mConst As Constraints) As c_Contraintes
'collecte les contraintes d'assemblage du lot ou de l'env
    Dim col_members As c_Contraintes
    Set col_members = New c_Contraintes
    If mConst.Count > 0 Then
        For i = 1 To mConst.Count
            col_members.Add mConst.Item(i).Name, mConst.Item(i).Status, mConst.Item(i).ReferenceType
        Next
    End If
    Set List_Contraintes = col_members
    Set col_members = Nothing
End Function

Property Get Contraintes() As Constraints
'Renvoi la collection des contraintes de la grille assembl�e
    Set Contraintes = mvarContraintes
End Property

Property Get DocGrilleNue() As PartDocument
'Renvoi le document de la grille nue
    DocGrilleNue = mvarProducts.Item(GrilleNueNom).Parent
End Property

Property Get GrilleNueNom() As String
'Renvoi le nom de la grille nue
    For i = 1 To mvarProducts.Count
        If Left(mvarProducts.Item(i).Name, 10) = Left(mvarProduct.Name, 10) Then
            GrilleNueNom = mvarProducts.Item(i).PartNumber
            Exit For
        End If
    Next
End Property

Property Get Exist_PartGrilleNue() As Boolean
'D�tecte si une part grille nue est pr�sente dans le product Grille assembl�e
' On v�rifie qu'il y a une pi�ce dans l'assemblage
    Exist_PartGrilleNue = False
    If mvarProducts.Count = 0 Then
        Exist_PartGrilleNue = False
    Else
    ' Recherche de la grille nue
        For i = 1 To mvarProducts.Count
            If Left(mvarProducts.Item(i).Name, 10) = Left(mvarProduct.Name, 10) Then
                Exist_PartGrilleNue = True
                Exit For
            End If
        Next
    End If
End Property

Property Let SysNumGrille(intg As Integer)
'Collecte le system de num�rotation des grille tel que d�fini dans le DSCGP
   mvarSysNum = intg
End Property

Property Let NomGrillenueDscgp(str As String)
'collecte le nom de la grille nue tel que d�fini dans le DSCGP
    mvarNomGrille = str
End Property

Property Let NomGrillenueSymDscgp(str As String)
'collecte le nom de la grille sym nue tel que d�fini dans le DSCGP
    mvarNomGrilleSym = str
End Property

Property Let NomLotDscgp(str As String)
'Collecte le N� du lot tel que d�fini dans le DSCGP
    mvarNumLot = str
End Property

Property Get ExistGrilleDSCGP() As Boolean
'D�tecte si une part grille nue conforme au nom d�fini dans le DSCGP
'est pr�sente dans le product Grille assembl�e
    ExistGrilleDSCGP = False
    If mvarProducts.Count = 0 Then
        ExistGrilleDSCGP = False
    Else
    ' Recherche de la grille nue
        For i = 1 To mvarProducts.Count
            If mvarProducts.Item(i).Name = mvarNomGrille Then
                ExistGrilleDSCGP = True
                Exit For
            End If
        Next
    End If
End Property

Property Get ExistGrillesymDSCGP() As Boolean
'D�tecte si une part grille nue sym conforme au nom d�fini dans le DSCGP
'est pr�sente dans le product Grille assembl�e
    ExistGrillesymDSCGP = False
    If mvarProducts.Count = 0 Then
        ExistGrillesymDSCGP = False
    Else
    ' Recherche de la grille nue
        For i = 1 To mvarProducts.Count
            If mvarProducts.Item(i).Name = mvarNomGrilleSym Then
                ExistGrillesymDSCGP = True
                Exit For
            End If
        Next
    End If
End Property

Property Get Numero() As String
'Renvoi le PartNumber de la Grille Assembl�e
    Numero = mvarPNumber
End Property

'#############################################################
'# Les Param�tres
'#############################################################
Public Function TestParamExist(ByVal vData As String) As Boolean
'test si le param�tre pass� en argument existe dans le part.
'si oui renvoi true sinon False,
On Error Resume Next
    Dim PrmCherche As StrParam
    Set PrmCherche = mvarProductParametres.Item(vData)
    If (Err.Number <> 0) Then
        Err.Clear
        TestParamExist = False
    Else
        TestParamExist = True
    End If
End Function

Public Function LectureParam(ByVal vData As String) As String
'Renvoi le contenu du param�tre pass� en argument dans le part de la grille.
'Si le param�tre n'existe pas renvois une chaine vide
    If TestParamExist(vData) Then
        LectureParam = mvarProductParametres.Item(vData).Value
    Else
        LectureParam = ""
    End If
End Function

Public Property Let Prm_DescriptionRef(ByVal vData As String)
'Documente le contenu du param�tre "DescriptionRef"
'Cas particulier param�tre de base sur les Parts
    mvarProduct.DescriptionRef = vData
End Property

Public Property Get Prm_DescriptionRef() As String
'Renvoi le contenu du param�tre "DescriptionRef"
'Cas particulier param�tre de base sur les Parts
    Prm_DescriptionRef = mvarProduct.DescriptionRef
End Property

Property Get prm_Source()
'Renvoi le contenu du param�tre "Source"
    prm_Source = mvarProduct.Source
End Property

Public Property Let xProperties(ByVal vCol As c_Properties)
    Set mColProperties = vCol
End Property

Public Property Get xProperties() As c_Properties
    Set xProperties = mColProperties
End Property


