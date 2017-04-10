VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NumMachines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*********************************************************************
'* Classe : NumMachines
'*
'* Fonctions :  'Récupération des diamètres de perçage et autres info en fonction des Numéro de machine
'*              'Info stockées dans un fichier excel à alimenter au fur et a mesure
'*
'*
'* Version : 4
'* Création :  CFR
'* Modification : 12/05/15
'*                Ajout profondeur de taraudage de la vis arretoir
'*
'**********************************************************************
Private MachinesCC, MachinesVT, Agrafes, ListMachine

Property Let OpenBibliMachine(NomFicBibLMAchine As String)
ReDim MachinesCC(2, 0)
ReDim MachinesVT(10, 0)
ReDim Agrafes(1, 0)

If NomFicBibLMAchine <> "" Then
    Dim objexcel
    Dim objWorkBook
    Dim objWorkSheetCC, obkWorkSheetVT, objWorkSheetAgrafes
    Set objexcel = CreateObject("EXCEL.APPLICATION")
    Set objWorkBook = objexcel.Workbooks.Open(CStr(NomFicBibLMAchine), True, True)
    
    'Récupération de la liste des grilles CC
    Set objWorkSheetCC = objWorkBook.Sheets("CC")
       Dim LigEC As Integer
       LigEC = 2
    While objWorkSheetCC.cells(LigEC, 1) <> ""
        ReDim Preserve MachinesCC(2, LigEC - 2)
        MachinesCC(0, LigEC - 2) = objWorkSheetCC.cells(LigEC, 1) ' N° machine
        MachinesCC(1, LigEC - 2) = objWorkSheetCC.cells(LigEC, 2) ' Diam perçage avion
        MachinesCC(2, LigEC - 2) = objWorkSheetCC.cells(LigEC, 3) ' Diam perçage grille
        LigEC = LigEC + 1
    Wend
    
    'Récupération de la liste des grilles VT
    Set obkWorkSheetVT = objWorkBook.Sheets("VT")
       LigEC = 2
    While obkWorkSheetVT.cells(LigEC, 1) <> ""
        ReDim Preserve MachinesVT(10, LigEC - 2)
        MachinesVT(0, LigEC - 2) = obkWorkSheetVT.cells(LigEC, 1) ' N° machine
        MachinesVT(1, LigEC - 2) = obkWorkSheetVT.cells(LigEC, 2) ' Diam perçage avion
        MachinesVT(2, LigEC - 2) = obkWorkSheetVT.cells(LigEC, 3) ' Diam perçage grille
        MachinesVT(3, LigEC - 2) = obkWorkSheetVT.cells(LigEC, 4) ' Diam Position Vis Arretoir
        MachinesVT(4, LigEC - 2) = obkWorkSheetVT.cells(LigEC, 5) ' Diam perçage vis Arretoir
        MachinesVT(5, LigEC - 2) = obkWorkSheetVT.cells(LigEC, 6) ' Profondeur Vis Arretoir
        MachinesVT(6, LigEC - 2) = obkWorkSheetVT.cells(LigEC, 7) ' Profondeur Taraudage Vis Arretoir
        MachinesVT(7, LigEC - 2) = obkWorkSheetVT.cells(LigEC, 8) ' Diam Lamage
        MachinesVT(8, LigEC - 2) = obkWorkSheetVT.cells(LigEC, 9) ' Nbr de vis arretoir
        MachinesVT(9, LigEC - 2) = obkWorkSheetVT.cells(LigEC, 10) ' Ref de la bague
        MachinesVT(10, LigEC - 2) = obkWorkSheetVT.cells(LigEC, 11) ' Ref de la vis arretoir
        LigEC = LigEC + 1
    Wend
       
    'Recuperation de la liste des agrafes
    Set objWorkSheetAgrafes = objWorkBook.Sheets("Agrafes")
        LigEC = 2
    While objWorkSheetAgrafes.cells(LigEC, 1) <> ""
        ReDim Preserve Agrafes(1, LigEC - 1)
        Agrafes(0, LigEC - 1) = objWorkSheetAgrafes.cells(LigEC, 1) ' Ref de l'agrafe
        Agrafes(1, LigEC - 1) = objWorkSheetAgrafes.cells(LigEC, 2) ' Diam perçage grille
        LigEC = LigEC + 1
    Wend
    
    objWorkBook.Close False
Else
    MsgBox "Pas de fichier excel sélectionné !"
    Exit Property
End If

End Property

Property Get ListeMachinesCC()
    ListeMachinesCC = MachinesCC
End Property

Property Get ListeMachinesVT()
    ListeMachinesVT = MachinesVT
End Property

Property Get ListAgrafes()
    ListAgrafes = Agrafes
End Property

Property Get DiamPercageAvionCC(NumMachine As String) As String
'Renvoi le Diamètre de perçage (Grille CC) de la machine passée en argument
Dim i As Long
DiamPercageAvionCC = 0
For i = 0 To UBound(MachinesCC, 2)
    If MachinesCC(0, i) = NumMachine Then
        DiamPercageAvionCC = MachinesCC(1, i)
    End If
Next
End Property

Property Get DiamPercageAvionVT(NumMachine As String) As String
'Renvoi le Diamètre de perçage (Grille VT) de la machine passée en argument
Dim i As Long
DiamPercageAvionVT = 0
For i = 0 To UBound(MachinesVT, 2)
    If MachinesVT(0, i) = NumMachine Then
        DiamPercageAvionVT = MachinesVT(1, i)
    End If
Next
End Property

Property Get DiamPercageGrilleCC(NumMachine As String)
'Renvoi le Diamètre de perçage (Grille CC) de la machine passée en argument
Dim i As Long
DiamPercageGrilleCC = 0
For i = 0 To UBound(MachinesCC, 2)
    If MachinesCC(0, i) = NumMachine Then
        DiamPercageGrilleCC = MachinesCC(2, i)
    End If
Next
End Property

Property Get DiamPercageGrilleVT(NumMachine As String)
'Renvoi le Diamètre de perçage (Grille VT) de la machine passée en argument
Dim i As Long
DiamPercageGrilleVT = 0
For i = 0 To UBound(MachinesVT, 2)
    If MachinesVT(0, i) = NumMachine Then
        DiamPercageGrilleVT = MachinesVT(2, i)
    End If
Next
End Property

Property Get PosArretVT(NumMachine As String)
'Renvoi le Rayon de position de la vis arretoir de la machine passée en argument
Dim i As Long
PosArretVT = 0
For i = 0 To UBound(MachinesVT, 2)
    If MachinesVT(0, i) = NumMachine Then
        PosArretVT = MachinesVT(3, i)
    End If
Next
End Property

Property Get DiamArretVT(NumMachine As String)
'Renvoi le diametre de taraudage de la vis arretoir de la machine passée en argument
DiamArretVT = ""
Dim i As Long
For i = 0 To UBound(MachinesVT, 2)
    If MachinesVT(0, i) = NumMachine Then
        DiamArretVT = MachinesVT(4, i)
    End If
Next
End Property

Property Get ProfArretVT(NumMachine As String)
'Renvoi la profondeur de perçage de la vis arretoir de la machine passée en argument
ProfArretVT = 0
Dim i As Long
For i = 0 To UBound(MachinesVT, 2)
    If MachinesVT(0, i) = NumMachine Then
        ProfArretVT = MachinesVT(5, i)
    End If
Next
End Property

Property Get ProfTarauArretVT(NumMachine As String)
'Renvoi la profondeur du taraudage de la vis arretoir de la machine passée en argument
ProfTarauArretVT = 0
Dim i As Long
For i = 0 To UBound(MachinesVT, 2)
    If MachinesVT(0, i) = NumMachine Then
        ProfTarauArretVT = MachinesVT(6, i)
    End If
Next
End Property

Property Get DiamLamageVT(NumMachine As String)
'Renvoi le diamètre du lamage autour du trou de nez avion de la machine passée en argument
DiamLamageVT = 0
Dim i As Long
For i = 0 To UBound(MachinesVT, 2)
    If MachinesVT(0, i) = NumMachine Then
        DiamLamageVT = MachinesVT(7, i)
    End If
Next
End Property

Property Get NBVisArretoirVT(NumMachine As String)
'Renvoi le nombre de vis arretoir de la machine passée en argument
NBVisArretoirVT = ""
Dim i As Long
For i = 0 To UBound(MachinesVT, 2)
    If MachinesVT(0, i) = NumMachine Then
        NBVisArretoirVT = MachinesVT(8, i)
    End If
Next
End Property

Property Get RefBagueVT(NumMachine As String)
'Renvoi la reference de la bague de la machine passée en argument
RefBagueVT = ""
Dim i As Long
For i = 0 To UBound(MachinesVT, 2)
    If MachinesVT(0, i) = NumMachine Then
        RefBagueVT = MachinesVT(9, i)
    End If
Next
End Property

Property Get RefVisArretoirVT(NumMachine As String)
'Renvoi la reference de la vis arretoir de la machine passée en argument
RefVisArretoirVT = ""
Dim i As Long
For i = 0 To UBound(MachinesVT, 2)
    If MachinesVT(0, i) = NumMachine Then
        RefVisArretoirVT = MachinesVT(10, i)
    End If
Next
End Property
