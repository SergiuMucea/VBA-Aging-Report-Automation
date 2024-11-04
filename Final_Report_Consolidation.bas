Sub MainMacro()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim ws As Worksheet
    Dim wsAllEU As Worksheet: Set wsAllEU = wb.Sheets("ALL EU")
    Dim AllEUSelection As Range: Set AllEUSelection = wsAllEU.Range("A2").CurrentRegion
   
    Application.ScreenUpdating = False
    
    'Converts the data to values, removing formulas and other formatting
    AllEUSelection.Copy
    wsAllEU.Range("A1").PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    
    'Deletes first two rows
    wsAllEU.Rows("1:2").Delete
    
    'Deletes first column
    wsAllEU.Columns("A").Delete
   
    'Removes rows with #N/A
    AllEUSelection.AutoFilter Field:=2, Criteria1:=0
    AllEUSelection.Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete
    wsAllEU.ShowAllData

    'Sorts Total Overdue USD column R in ALL EU sheet
    AllEUSelection.Sort key1:=wsAllEU.Range("Q2"), Order1:=xlDescending, Header:=xlYes
        
    'Wrap text in review columns
    Range("AH:AK").WrapText = True
    
    
    'The following procedure goes through each separate country sheet and consolidates the data in ALL EU sheet per country
    For Each ws In Worksheets
         
        If Len(ws.Name) <= 3 Then
        
            'Deletes existing data in each country sheet
            Worksheets(ws.Name).Select
            Range("A1").CurrentRegion.Delete
            
            'Filters data according to Sheet name in AllEU sheet and copy-pastes it in each separate country sheet
            AllEUSelection.AutoFilter Field:=2, Criteria1:=ws.Name
            
            AllEUSelection.Copy
            Worksheets(ws.Name).Activate
            Worksheets(ws.Name).Range("A1").Select
            ActiveSheet.Paste
            Application.CutCopyMode = False
            
            'Set Filter on each sheet
            Range("A2").Select
            Selection.AutoFilter
                        
            'Separate function for Hungary so that the formatting of the amounts is solved.
            If ws.Name = "HU" Then
                Worksheets(ws.Name).Activate
                Range("M1").EntireColumn.Insert
                
                Range("L1").Copy
                Range("M1").Select
                ActiveSheet.Paste
                Application.CutCopyMode = False
                
                Range("M2").Formula = "=L2/100"
                Dim lastRowHU As Long
                lastRowHU = Cells(Rows.Count, "L").End(xlUp).Row
                Range(Cells(2, "M"), Cells(lastRowHU, "M")).FillDown
                
                Columns("M").Copy
                Columns("L").PasteSpecial xlPasteValues
                
                Columns("M").Delete
            End If
    
        End If
        
    Next ws
    
    'removes filter in All EU sheet
    wsAllEU.ShowAllData

End Sub
