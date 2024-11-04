Public Sub SAPConsolidation()
    Dim wsKNKK As Worksheet: Set wsKNKK = ThisWorkbook.Sheets("KNKK")
    Dim wsAllEU As Worksheet: Set wsAllEU = ThisWorkbook.Sheets("all eu")
    Dim wbConsolidation As Workbook: Set wbConsolidation = ThisWorkbook

    Dim TextFile As Workbook
    Dim OpenFiles() As Variant
    Dim i As Integer
    
    'First of all, delete all old data from the Excel file
    wsAllEU.Range("A4").CurrentRegion.Delete
    wsKNKK.Range("A1").CurrentRegion.Delete
    
    'Select and open the reports extracted from SAP. The loop will open and copy the info in each file and consolidate it in sheet 1
    OpenFiles = Application.GetOpenFilename(Title:="Select SAP Aging Reports to import", MultiSelect:=True)
    
    Application.ScreenUpdating = False
    
    For i = 1 To Application.CountA(OpenFiles)
        Set TextFile = Workbooks.Open(OpenFiles(i))
        
        If wsAllEU.Range("A4").Value = "" Then
                If TextFile.Name = "HU" Then
                    Call HUFormat
                End If
            TextFile.Sheets(1).Range("A4").CurrentRegion.Copy
            wbConsolidation.Activate
            wsAllEU.Activate
            wsAllEU.Range("A1").Select
            ActiveSheet.Paste
            Application.CutCopyMode = False
            TextFile.Saved = True
            TextFile.Close
            
        Else:
                If TextFile.Name = "HU" Then
                    Call HUFormat
                End If
            TextFile.Sheets(1).Rows("5:5").Select
            Range(selection, selection.End(xlDown)).Select
            selection.Copy
            wbConsolidation.Activate
            wsAllEU.Activate
            wsAllEU.Range("A4").Select
            selection.End(xlDown).Select
            ActiveCell.Offset(1, 0).Select
            ActiveSheet.Paste
            Application.CutCopyMode = False
            TextFile.Saved = True
            TextFile.Close
        End If
    Next i


End Sub

' SAP extracts the file in a different format for amounts, so doing the necessary using this function
Private Sub HUFormat()
    Columns("S:AK").Select
    selection.Replace What:=",", Replacement:=""
    selection.Replace What:=".", Replacement:=""
End Sub


'The following Sub will integrate another type of report (KNKK) extracted from SAP (Credit Risk Report)
'Important to extract it as excel from SAP in order to preserve columns and rows
Public Sub KNKKIntegration()
    Dim wsKNKK As Worksheet: Set wsKNKK = ThisWorkbook.Sheets("KNKK")
    Dim wsAllEU As Worksheet: Set wsAllEU = ThisWorkbook.Sheets("all eu")
    Dim wbConsolidation As Workbook: Set wbConsolidation = ThisWorkbook

    Dim KNKKFile As Workbook
    Dim OpenFile As Variant

    'Below code opens KNKK File and adds it in SAP Consolidation file in Sheet 2
    OpenFile = Application.GetOpenFilename(Title:="Select SAP KNKK Report to import")
    Set KNKKFile = Workbooks.Open(OpenFile)
    
    Application.ScreenUpdating = False
    
    KNKKFile.Sheets(1).Range("A1").CurrentRegion.Copy
    wbConsolidation.Activate
    wsKNKK.Activate
    wsKNKK.Range("A1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    KNKKFile.Saved = True
    KNKKFile.Close

    'Going further to the formula in Sheet 1, to make the necessary for the crosscheck between the aging reports and the KNKK Credit Risk report
    'Column headings used by the formula will be updated and the data in the KNKK sheet will be converted from "Text" to "General"
    'Index and Match excel formula is more suitable here, because it can check the data no matter the position of the column headings
    wsAllEU.Activate
    wsAllEU.Columns("S:U").Select
    selection.ClearContents
    Range("S4").Select
    ActiveCell.Value = "Credit limit"
    Range("T4").Select
    ActiveCell.Value = "Risk category"
    Range("U4").Select
    ActiveCell.Value = "Rating"
    wsKNKK.Activate
    wsKNKK.Range("A2").CurrentRegion.Select
    With selection
        .NumberFormat = "General"
        .Value = .Value
    End With
    
    Dim KNKKSelection As Range
    Set KNKKSelection = wsKNKK.Range("A2").CurrentRegion
    
    wsAllEU.Activate
    wsAllEU.Range("S5").Formula = "=INDEX(" & wsKNKK.Name & "!" & KNKKSelection.Address & ",MATCH($J5,KNKK!$A:$A,0),MATCH(S$4,KNKK!$1:$1,0))"
    
    Dim lastRow As Long
    
    lastRow = Cells(Rows.Count, "J").End(xlUp).Row
        
    Range("S5").Select
    selection.AutoFill Destination:=Range("S5:U5"), Type:=xlFillDefault
    
    Range(Cells(5, "S"), Cells(lastRow, "S")).FillDown
    Range(Cells(5, "T"), Cells(lastRow, "T")).FillDown
    Range(Cells(5, "U"), Cells(lastRow, "U")).FillDown
    
    Range("S:U").Copy
    Range("S:U").PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    
End Sub




