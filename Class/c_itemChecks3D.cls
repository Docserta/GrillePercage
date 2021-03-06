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
'Collection des checks 3D collect�es dans le fichier excel

Option Explicit
Private mcol As Collection

Public Function Add(Cote As String, Optional Check As Boolean = False, Optional statut As String = "", Optional Comment As String = "") As c_itemCheck3D
    'Cr�e un nouvel objet
    Dim objNewMember As c_itemCheck3D
    Set objNewMember = New c_itemCheck3D

    'd�fini les propri�t�s transf�r�es � la m�thode
    objNewMember.Cote = Cote
    objNewMember.Check = Check
    objNewMember.statut = statut
    objNewMember.Comment = Comment
        
    'Ajoute un objet � la collection
    mcol.Add objNewMember, Cote
        
    'Renvoi l'objet cr�e
    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Function WriteItem(Cote As String, Optional Check As Boolean = False, Optional statut As String = "", Optional Comment As String = "") As c_itemCheck3D
   'Met a jour l'objet recherch�
    Dim objNewMember As c_itemCheck3D
    
    Set objNewMember = mcol.Item(Cote)
    'Met � jour les propri�t�s transf�r�es � la m�thode
    objNewMember.Check = Check
    objNewMember.statut = statut
    objNewMember.Comment = Comment
    
    'Renvoi l'objet cr�e
    Set WriteItem = objNewMember
    Set objNewMember = Nothing
  
End Function

Public Property Get Item(vntIndexKey As Variant) As c_itemCheck3D
    'utilis� pour se r�f�rer � un �l�ment de collection
    'vntIndexKey contient l'index ou la cl� de la collection,
    'donc il est d�clar� comme Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mcol(vntIndexKey)
End Property

Public Property Get Count() As Long
    'permet d'obtenir le nombre d'�l�ments
    'Syntax: Debug.Print x.Count
    Count = mcol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    'est utilis� pour supprimer un �l�ment de la collection
    'vntIndexKey contient l'index ou la cl� de la collection,
    'donc il est d�clar� comme Variant
    'Syntax: x.Remove(xyz)

    mcol.Remove vntIndexKey
End Sub

Private Sub Class_initialize()
    'cr�e la collection lorsque la classe est cr��e
    Set mcol = New Collection
End Sub

Private Sub Class_Terminate()
    'supprime la collection lorsque la classe est ferm�e
    Set mcol = Nothing
End Sub


