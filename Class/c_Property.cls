VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_Property"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'c_Property

Private mvarNom As String
Private mvarValeur As String
Public Property Let nom(ByVal vData As String)
    mvarNom = vData
End Property

Public Property Get nom() As String
    nom = mvarNom
End Property

Public Property Let Valeur(ByVal vData As String)
    mvarValeur = vData
End Property

Public Property Get Valeur() As String
    Valeur = mvarValeur
End Property
