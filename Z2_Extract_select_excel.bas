Attribute VB_Name = "Z2_Extract_select_excel"
Option Explicit

Sub catmain()

'Log de l'utilisation de la macro
LogUtilMacro nPath, nFicLog, nMacro, "Z2_Extract_select_excel", VMacro

Dim info_filters As String, AdminLevel As String
'Dim RootProduct As Product
Dim ActiveDoc As Document
Dim part As part
Dim HybridBodies_tmp As HybridBodies
Dim MySelection As Selection
Dim ShapeA As HybridShape
Dim ShapeB As HybridShape
Dim Name_Input As String
Dim xls As Variant
Dim cPt As Long
Dim oref As Reference
Dim TheSPAWorkbench As SPAWorkbench
Dim Coord(2)
Dim TheMeasurable 'As Measurable

    Set ActiveDoc = CATIA.ActiveDocument
    Set part = ActiveDoc.part
    
    Set MySelection = ActiveDoc.Selection
    Set HybridBodies_tmp = part.HybridBodies
    
    cPt = 1
    
    Set xls = CreateObject("Excel.Application")
        xls.WindowState = 1
        xls.Visible = True
        xls.Workbooks.Add
        xls.worksheets(1).Name = "Feuille1"
    
        xls.worksheets(1).range("A1") = "points selectionnes"
        xls.worksheets(1).range("B1") = "X"
        xls.worksheets(1).range("C1") = "Y"
        xls.worksheets(1).range("D1") = "Z"
        xls.worksheets(1).range("A1:D1").Interior.Color = RGB(255, 0, 0)

    Set TheSPAWorkbench = ActiveDoc.GetWorkbench("SPAWorkbench")
    
    While (cPt <= MySelection.Count)
    
        Set ShapeA = MySelection.Item(cPt)
        xls.worksheets(1).range("A" & cPt + 1) = ShapeA.Value.Name
        Set oref = part.CreateReferenceFromObject(ShapeA.Value)
        Set TheMeasurable = TheSPAWorkbench.GetMeasurable(oref)
        TheMeasurable.GetPoint Coord
    
        xls.worksheets(1).range("A" & cPt + 1) = ShapeA.Value.Name
        xls.worksheets(1).range("B" & cPt + 1) = Coord(0)
        xls.worksheets(1).range("C" & cPt + 1) = Coord(1)
        xls.worksheets(1).Columns("C").AutoFit
        xls.worksheets(1).range("D" & cPt + 1) = Coord(2)
        xls.worksheets(1).Columns("D").AutoFit
    
        cPt = cPt + 1
    Wend

End Sub


