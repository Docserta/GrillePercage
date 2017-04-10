Attribute VB_Name = "A1_AjoutGrille"
Option Explicit

'*********************************************************************
'* Macro : A1_AjoutGrille
'*
'* Fonctions :  Ajoute un product de grille à un product général existant
'*              Ajoute des attributs provenant d'un fichier excel
'*              Crée les set géométriques et les contrainte de fixation
'*
'* Version : 4
'* Création :  CFR
'* Modification : 15/04/14
'*
'**********************************************************************
Sub catmain()

'Log de l'utilisation de la macro
LogUtilMacro nPath, nFicLog, nMacro, "A1_AjoutGrille", VMacro

'Ouvre la boite de dlg "Frm_AjoutGrille"
    Load Frm_AjoutGrille
    
Dim Doc_AssGen As ProductDocument
Dim ProdAssGen As Product
Dim AssGenProds As Products
Dim AssGenConstrs As Constraints
Dim MesParametres As Parameters

Dim NameRef As String
Dim i As Integer
  
'Recupération du Product général
    Set coll_docs = CATIA.Documents
    Set Doc_AssGen = CATIA.ActiveDocument
    Set ProdAssGen = Doc_AssGen.Product
    Set AssGenProds = ProdAssGen.Products
    Set AssGenConstrs = ProdAssGen.Connections("CATIAConstraints")

'Récupère les paramètres enregistrés dans le fichier d'assemblage
    Set MesParametres = ProdAssGen.UserRefProperties
    Frm_AjoutGrille.TBX_NomAss = TestParamExist(MesParametres, "Param_Assembl")
    Frm_AjoutGrille.TBX_RepSave = TestParamExist(MesParametres, "Param_RepSauv")
    Frm_AjoutGrille.Show
    
'Sort du programme si click sur bouton Annuler dans Frm_AjoutGrille
    If Not (Frm_AjoutGrille.ChB_OkAnnule) Then
        Unload Frm_AjoutGrille
        Exit Sub
    End If

AjoutGrille ProdAssGen, Frm_AjoutGrille.TBX_NomGriAss, ValDscgp.design, Frm_AjoutGrille.TBX_NomGriNue, Frm_AjoutGrille.TBX_NomAss, Frm_AjoutGrille.TBX_NomU01
 
'Fixe le product Grille Assemblée
    For i = 1 To Doc_AssGen.Product.Products.Count
        If InStr(1, Doc_AssGen.Product.Products.Item(i).Name, Frm_AjoutGrille.TBX_NomGriAss, vbTextCompare) <> 0 Then
            FixePart2 Frm_AjoutGrille.TBX_NomAss, Doc_AssGen.Product.Products.Item(i).Name
        End If
    Next
    
Unload Frm_AjoutGrille
End Sub

