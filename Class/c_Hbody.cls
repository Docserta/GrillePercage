VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_Hbody"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'c_Hbody
'Objet HybridBody

Private mvarNom As String               'Nom de l'hybridbody
Private mvarShow As CatVisPropertyShow  'statut Show/NoShow de l'objet
Private mvarHBody As HybridBody         'HybridBody

Public Property Let nom(ByVal vData As String)
    mvarNom = vData
End Property

Public Property Get nom() As String
    nom = mvarNom
End Property

Public Property Let ShowStatus(ByVal vProp As CatVisPropertyShow)
    mvarShow = vProp
End Property

Public Property Get ShowStatus() As CatVisPropertyShow
    ShowStatus = mvarShow
End Property

Public Property Let HBody(ByVal vHbod As HybridBody)
    Set mvarHBody = vHbod
End Property

Public Property Get HBody() As HybridBody
    Set HBody = mvarHBody
End Property

