VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GrilleLot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*********************************************************************
'* Classe : GrilleLot
'*
'* Fonctions :  Teste la présence des éléments normalisé dans le product Lot
'*
'* Version : 1
'* Création :  CFR
'* Modification : 17/05/16
'*
'**********************************************************************
Option Explicit
Private GL_ProdDoc As ProductDocument
Private GL_Product As Product
Private GL_ProductEnv As Product
Private GL_Products As Products
Private GL_Numero As String
Private GL_ConstLot As Constraints
Private GL_ConstEnv As Constraints
Private GL_ProductParametres As Parameters

Private i As Long

'===========================
' Initialisation de l'objet
'===========================

Private Sub Class_Initialize()

    'On Error Resume Next

    Set GL_ProdDoc = CATIA.ActiveDocument
    Set GL_Product = GL_ProdDoc.Product
    Set GL_Products = GL_Product.Products
    GL_Numero = GL_Product.PartNumber
    Set GL_ConstLot = GL_Product.Connections("CATIAConstraints")
    Set GL_ProductParametres = GL_Product.ReferenceProduct.UserRefProperties
    
End Sub

Property Let GL_ProdDocLot(ProdDocLot As ProductDocument)
' Réinitialisation de l'objet

    Set GL_ProdDoc = ProdDocLot
    Set GL_Product = GL_ProdDoc.Product
    Set GL_Products = GL_Product.Products
    GL_Numero = GL_Product.PartNumber
    Set GL_ConstLot = GL_Product.Connections("CATIAConstraints")
    Set GL_ProductParametres = GL_Product.ReferenceProduct.UserRefProperties
    
End Property

Property Get NoLot()
'Renvoi le Numero de Lot
    NoLot = GL_Numero
End Property

'Property Get Exist_NoeudEnv() As Boolean
''Détecte si un composant nommé "Env" est présent
'    Exist_NoeudEnv = False
'    If GL_Products.Count = 0 Then
'        Exist_NoeudEnv = False
'    Else
'        For i = 1 To GL_Products.Count
'            If GL_Products.Item(i).PartNumber = "env" Then
'                Exist_NoeudEnv = True
'                Exit For
'            End If
'        Next
'    End If
'End Property

Property Get Exist_NoeudEnv() As Boolean
'Détecte si un composant nommé "Env" est présent
    Exist_NoeudEnv = False
'    Dim tProd As Product
    On Error GoTo Err_Exist_NoeudEnv
'    Set tProd = GL_Products.GetItem("env")
'    Exist_NoeudEnv = True
'    GoTo Quit_Exist_NoeudEnv
    
    If GL_Products.Count = 0 Then
        Exist_NoeudEnv = False
    Else
        For i = 1 To GL_Products.Count
            If GL_Products.Item(i).PartNumber = "env" Then
                Exist_NoeudEnv = True
                Exit For
            End If
        Next
    End If
    GoTo Quit_Exist_NoeudEnv
    
Err_Exist_NoeudEnv:
    On Error GoTo 0
Quit_Exist_NoeudEnv:
End Property

Property Get NoeudEnv() As Product
'Renvoi le product "NoeudEnv"
    For i = 1 To GL_Products.Count
        If GL_Products.Item(i).PartNumber = "env" Then
            Set NoeudEnv = GL_Products.Item(i)
            Exit For
        End If
    Next
    i = 0
End Property

Property Get Col_ConstLot() As c_Contraintes
'Renvoi la collection des contraites d'assemblage du lot
    Set Col_ConstLot = List_Contraintes(GL_ConstLot)

End Property

Property Get Col_ConstEnv() As c_Contraintes
'renvoi la collection de contraintes d'assemblage de l'environement
    Set GL_ProductEnv = GL_Products.Item("env.1")
    Set GL_ConstEnv = GL_ProductEnv.Connections("CATIAConstraints")
    Set Col_ConstEnv = List_Contraintes(GL_ConstEnv)
    
End Property

Private Function List_Contraintes(mConst As Constraints) As c_Contraintes
'collecte les contraintes d'assemblage du lot ou de l'env
    Dim col_members As c_Contraintes
    Set col_members = New c_Contraintes
    Dim toto
    If mConst.Count > 0 Then
        On Error GoTo Erreur
        For i = 1 To mConst.Count
            col_members.Add mConst.Item(i).Name, mConst.Item(i).Status, mConst.Item(i).ReferenceType
        Next
    End If
    Set List_Contraintes = col_members
    Set col_members = Nothing
GoTo Fin

Erreur:
    Err.Raise vbObjectError + 540, "GrilleLot", "Erreur dans la collecte des contraintes de positionnement de l'environnement"
Fin:

End Function

'Private Function List_Contraintes(mConst As Constraints) As c_Contraintes
''collecte les contraintes d'assemblage du lot ou de l'env
'    Dim col_members As c_Contraintes
'    Set col_members = New c_Contraintes
'    If mConst.Count > 0 Then
'        For i = 1 To mConst.Count
'            col_members.Add mConst.Item(i).Name, mConst.Item(i).Status, mConst.Item(i).ReferenceType
'        Next
'    End If
'    Set List_Contraintes = col_members
'    Set col_members = Nothing
'End Function

'Private Function GetProdChargé(oProds As Products) As Products
''test le le product passé en argument est chargé en session
''on vérifier que le partNumber est accessible
'Dim tProd As Product
'Dim tProds As Products
'Dim i As Integer
'    For i = 1 To oProds.Count
'        On Error Resume Next
'        Set tProd = tProds.AddNewProduct(oProds.Item(i).PartNumber)
'        If Err.Number <> 0 Then
'            Err.Clear
'        End If
'    Next
'    On Error GoTo 0
'Set GetProdChargé = tProds
'End Function
