VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_itemChecks3D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'c_CollChecks3D
'Collection des checks 3D collectées dans le fichier excel

Option Explicit
Private mcol As Collection

Public Function Add(Cote As String, Optional Check As Boolean = False, Optional statut As String = "", Optional Comment As String = "") As c_itemCheck3D
    'Crée un nouvel objet
    Dim objNewMember As c_itemCheck3D
    Set objNewMember = New c_itemCheck3D

    'défini les propriétés transférées à la méthode
    objNewMember.Cote = Cote
    objNewMember.Check = Check
    objNewMember.statut = statut
    objNewMember.Comment = Comment
        
    'Ajoute un objet à la collection
    mcol.Add objNewMember, Cote
        
    'Renvoi l'objet crée
    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Function WriteItem(Cote As String, Optional Check As Boolean = False, Optional statut As String = "", Optional Comment As String = "") As c_itemCheck3D
   'Met a jour l'objet recherché
    Dim objNewMember As c_itemCheck3D
    
    Set objNewMember = mcol.Item(Cote)
    'Met à jour les propriétés transférées à la méthode
    objNewMember.Check = Check
    objNewMember.statut = statut
    objNewMember.Comment = Comment
    
    'Renvoi l'objet crée
    Set WriteItem = objNewMember
    Set objNewMember = Nothing
  
End Function

Public Property Get Item(vntIndexKey As Variant) As c_itemCheck3D
    'utilisé pour se référer à un élément de collection
    'vntIndexKey contient l'index ou la clé de la collection,
    'donc il est déclaré comme Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mcol(vntIndexKey)
End Property

Public Property Get Count() As Long
    'permet d'obtenir le nombre d'éléments
    'Syntax: Debug.Print x.Count
    Count = mcol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    'est utilisé pour supprimer un élément de la collection
    'vntIndexKey contient l'index ou la clé de la collection,
    'donc il est déclaré comme Variant
    'Syntax: x.Remove(xyz)

    mcol.Remove vntIndexKey
End Sub

Private Sub Class_initialize()
    'crée la collection lorsque la classe est créée
    Set mcol = New Collection
End Sub

Private Sub Class_Terminate()
    'supprime la collection lorsque la classe est fermée
    Set mcol = Nothing
End Sub


