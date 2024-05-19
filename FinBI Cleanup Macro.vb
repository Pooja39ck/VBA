
Sub FinBI_cleanup()

    ' Format entire sheet to General
    Sheets("Raw data").Activate
    ActiveSheet.Cells.NumberFormat = "General"

    'Remove blanks in I Columns

    Rows("1:1").Select
    Range("B1").Activate
    Range(Selection, Selection.End(xlDown)).Select
    Selection.AutoFilter
    Range("I:I").Cells.SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    
    'Save entire data into FinBI sheet
    
    Range("B1").Select
    Range(Selection, Selection.End(xlDown).EntireRow).Select
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveSheet.Name = "FinBI"
     
    
    'Delete unwanted columns
    
    Sheets("Raw Data").Activate
    Range("A:A,E:L,P:R,T:T,X:X,Z:Z,AC:AE").EntireColumn.Delete
    
    'Save -9s & -7s to new sheet
    
    Sheets("Raw Data").Activate
    ActiveSheet.Range("A1", Selection.End(xlDown).End(xlToRight)).AutoFilter Field:=6, Criteria1:= _
        "=-9999", Operator:=xlOr, Criteria2:="=-7777"
    Range("F1").Select
    Range(Selection, Selection.End(xlDown).EntireRow).Select
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveSheet.Name = "-9s & -7s"
    Sheets("Raw Data").Activate
    Application.DisplayAlerts = False
    On Error Resume Next
    Range("A2:N2", Range("f" & Rows.Count).End(xlUp)).SpecialCells(xlCellTypeVisible).EntireRow.Delete
    
   
    ActiveSheet.ShowAllData
    
    
    '-6666 replace with web order id
    
    
    ActiveSheet.Range("F1", Selection.End(xlDown).End(xlToRight)).AutoFilter Field:=6, Criteria1:="=-6666"
    Range("F1", Selection.End(xlDown)).Select
    On Error Resume Next
    ActiveSheet.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Cells(1, 6).Select
    ActiveCell.FormulaR1C1 = "=RC[-1]"
    ActiveCell.Copy
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveSheet.ShowAllData
    Columns("F:F").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues
    On Error GoTo 0
    Application.CutCopyMode = False

    ' Product ID contains -s-  & ends with -s Paste to switching sheet & delete in main sheet
    
    ActiveSheet.Range("J1", Selection.End(xlDown).End(xlToRight)).AutoFilter Field:=10, Criteria1:="=*-s-*", Operator:=xlOr, Criteria2:="=*-s"
    Range("J1").Select
    Range(Selection, Selection.End(xlDown).EntireRow).Select
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveSheet.Name = "Switching"
    Sheets("Raw Data").Activate
    Application.DisplayAlerts = False
    On Error Resume Next
    Range("A2:N2", Range("J" & Rows.Count).End(xlUp)).SpecialCells(xlCellTypeVisible).EntireRow.Delete
    ActiveSheet.ShowAllData
        
    ' Format Qty to number
    Columns(4).NumberFormat = "0"
    
    '(-ve)negative qty paste to -ve sheet & delete data in main sheet
    
    ActiveSheet.Range("D1", Selection.End(xlDown).End(xlToRight)).AutoFilter Field:=4, Criteria1:="<0"
    Range("D1").Select
    Range(Selection, Selection.End(xlDown).EntireRow).Select
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveSheet.Name = "Negative"
    Sheets("Raw Data").Activate
    Application.DisplayAlerts = False
    On Error Resume Next
    Range("A2:N2", Range("D" & Rows.Count).End(xlUp)).SpecialCells(xlCellTypeVisible).EntireRow.Delete
    ActiveSheet.ShowAllData
    

    '0qty paste to 0's sheet & delete data in main sheet
    
    ActiveSheet.Range("D1", Selection.End(xlDown).End(xlToRight)).AutoFilter Field:=4, Criteria1:="0"
    Range("D1").Select
    Range(Selection, Selection.End(xlDown).EntireRow).Select
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveSheet.Name = "0's"
    Sheets("Raw Data").Activate
    Application.DisplayAlerts = False
    On Error Resume Next
    Range("A2:N2", Range("D" & Rows.Count).End(xlUp)).SpecialCells(xlCellTypeVisible).EntireRow.Delete
    ActiveSheet.ShowAllData
    
    'SEA = Highest license tag to highest qty & remove duplicates by Sales order
    
    Range("a1").Select
    Range(Selection, Selection.End(xlDown).End(xlToRight)).Select
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveSheet.Name = "SEA"
    Application.DisplayAlerts = Fals
    
    ' See
    
    Sheets("Raw Data").Activate
    Range("K1").EntireColumn.Insert
    Range("k1").Value = "License"
    Range("j1", Selection.End(xlDown).End(xlToRight)).AutoFilter Field:=10, Criteria1:="=*See*"
    ActiveSheet.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Cells(1, 11).Select
    ActiveCell.Value = "SEE"
    ActiveCell.Copy
    
    Dim LastRow As Long
    Dim LastCellColC As Range

    LastRow = Cells(Rows.Count, "j").End(xlUp).Row
    Set LastCellColC = Range("j" & LastRow).Offset(0, 1)
    LastCellColC.Select
    Range(Selection, Selection.End(xlUp)).Select
    Selection.FillDown
    ActiveSheet.ShowAllData
    
    ' EXT
    
    Range("j1", Selection.End(xlDown).End(xlToRight)).AutoFilter Field:=10, Criteria1:="=*EXT*"
    ActiveSheet.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Cells(1, 11).Select
    ActiveCell.Value = "EXT"
    ActiveCell.Copy
    
    LastRow = Cells(Rows.Count, "j").End(xlUp).Row
    Set LastCellColC = Range("j" & LastRow).Offset(0, 1)
    LastCellColC.Select
    Range(Selection, Selection.End(xlUp)).Select
    Selection.FillDown
    ActiveSheet.ShowAllData
    
    ' ACT
    
    Range("j1", Selection.End(xlDown).End(xlToRight)).AutoFilter Field:=10, Criteria1:="=*Act*", Operator:=xlOr, Criteria2:="=*IOT*"
    ActiveSheet.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Cells(1, 11).Select
    ActiveCell.Value = "ACT"
    ActiveCell.Copy
    
    LastRow = Cells(Rows.Count, "j").End(xlUp).Row
    Set LastCellColC = Range("j" & LastRow).Offset(0, 1)
    LastCellColC.Select
    Range(Selection, Selection.End(xlUp)).Select
    Selection.FillDown
    ActiveSheet.ShowAllData
    
    ' Vlookup Highest qty to Cisco Bookings Quantity column & Delete Added Sheet
    
    Range("A1", Range("A1").End(xlDown).End(xlToRight)).Sort Key1:=Range("k1"), Order1:=xlAscending, Header:=xlYes
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "Qtylookup"
    Sheets("Raw Data").Range("F:F").EntireColumn.Copy Sheets("Qtylookup").Range("A:A")
    Sheets("Raw Data").Range("D:D").EntireColumn.Copy Sheets("Qtylookup").Range("B:B")
    Range("A1", Range("A1").End(xlDown).End(xlToRight)).Sort Key1:=Range("B1"), Order1:=xlDescending, Header:=xlYes
    Sheets("Raw Data").Activate
    Range("D2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[2],Qtylookup!C1:C2,2,0)"
    ActiveCell.Copy
    Range("D2").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Columns("D:D").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    ' Delete Qtylookup Sheet
    
    Sheets("Qtylookup").Activate
    Application.DisplayAlerts = False
    ActiveSheet.Delete
    Application.DisplayAlerts = True
        
    ' Delete Added license Column
    
    Sheets("Raw data").Activate
    Range("K:K").EntireColumn.Delete
    
    'Remove Duplicates by SO
    
    Range("A1", Range("A1").End(xlDown).End(xlToRight)).RemoveDuplicates Columns:=6, Header:=xlYes
    
    Sheets("Raw Data").Range("A1").Select
    
    MsgBox ("Yeah, Time to update Vlookup Sheet! then Click on RUN[Ready]")
    
    
    
End Sub


    
    Sub FinBI_Cleanup2()
    

    'After updating Vlookup sheet run these steps
    
    'Remove Dual SO on Hubspot
    
    Sheets("Raw Data").Activate
    Range("G1").EntireColumn.Insert
    Range("G1").Value = "Deal Name on Hubspot"
    
    Range("G2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1],Vlookup!C7:C8,2,0)"
    ActiveCell.Copy
    
    Range("f" & Cells(Rows.Count, "f").End(xlUp).Row).Offset(0, 1).Select
    Range(Selection, Selection.End(xlUp)).Select
    Selection.FillDown

    Columns("G:G").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    'Filter only not equal to NA
    
    Range("G1", Selection.End(xlDown).End(xlToRight)).AutoFilter Field:=7, Criteria1:="<>#N/A"
    
    ' Paste found Dupl to new sheet & Remove the same in Raw data sheet
    
    Range("g1").Select
    Range(Selection, Selection.End(xlDown).EntireRow).Select
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveSheet.Name = "dupl SO on hubspot"
    Sheets("Raw Data").Activate
    Application.DisplayAlerts = False
    On Error Resume Next
    Range("A2:o2", Range("g" & Rows.Count).End(xlUp)).SpecialCells(xlCellTypeVisible).EntireRow.Delete
    ActiveSheet.ShowAllData
    
    Range("g:g").EntireColumn.Delete
    
    ' Goto Negative Sheet Sort qty by ascending & Bring SO before qty
    
    Sheets("Negative").Activate
    Range("D1").EntireColumn.Insert
    ActiveSheet.Range("g:g").EntireColumn.Copy Range("D:D")
    Range("g:g").EntireColumn.Delete
    Range("A1", Range("A1").End(xlDown).End(xlToRight)).Sort Key1:=Range("e1"), Order1:=xlAscending, Header:=xlYes
    
    ' Goto Raw Sheet add 3 qty columns next to qty
    
    Sheets("Raw Data").Activate
    Range("E1:G1").EntireColumn.Insert
    Range("E1").Select
    ActiveCell.Value = "SO Based -ve Qty"
    Range("F1").Select
    ActiveCell.Value = "ERP Based -ve Qty"
    Range("G1").Select
    ActiveCell.Value = "Update -ve Qty"
    
    ' Lookup So based -ve qty from negative sheet
    
    Range("E2").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[4],Negative!C4:C5,2,0),"""")"
    ActiveCell.Copy
    
    Range("D" & Cells(Rows.Count, "D").End(xlUp).Row).Offset(0, 1).Select
    Range(Selection, Selection.End(xlUp)).Select
    Selection.FillDown
    
    ' Lookup ERP based -ve qty from negative sheet
    
    Range("F2").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-3],Negative!C3:C5,3,0),"""")"
    ActiveCell.Copy
    
    Range("E" & Cells(Rows.Count, "E").End(xlUp).Row).Offset(0, 1).Select
    Range(Selection, Selection.End(xlUp)).Select
    Selection.FillDown
    
    ' Cisco qty + ERP based -ve qty
    
    Range("G2").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(RC[-3]+RC[-1],"""")"
    ActiveCell.Copy
    
    Range("F" & Cells(Rows.Count, "F").End(xlUp).Row).Offset(0, 1).Select
    Range(Selection, Selection.End(xlUp)).Select
    Selection.FillDown
    
    ' Value paste all formula's
    
    Range("e:g").EntireColumn.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    'Filter -Ve & 0 from Update qty & Paste it to neg & 0's sheet
    
    Sheets("Raw Data").Activate
    ActiveSheet.Range("A1", Selection.End(xlDown).End(xlToRight)).AutoFilter Field:=7, Criteria1:="<=0"
    Range("g1").Select
    Range(Selection, Selection.End(xlDown).EntireRow).Select
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveSheet.Name = "neg & 0's"
    Sheets("Raw Data").Activate
    Application.DisplayAlerts = False
    On Error Resume Next
    Range("A2:q2", Range("g" & Rows.Count).End(xlUp)).SpecialCells(xlCellTypeVisible).EntireRow.Delete
    ActiveSheet.ShowAllData
    
    'Update Qty blanks filled with +cisco booking qty
    
    ActiveSheet.Range("A1", Selection.End(xlDown).End(xlToRight)).AutoFilter Field:=7, Criteria1:="="
    ActiveSheet.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Cells(1, 7).Select
    ActiveCell.FormulaR1C1 = "=RC[-3]"
    ActiveCell.Copy
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveSheet.ShowAllData
    Columns("G:G").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues
    On Error GoTo 0
    
    
    'Add  6 Columns next to SO & Update column name
    
    Sheets("Raw Data").Activate
    Range("J1:O1").EntireColumn.Insert
    Range("J1").Select
    ActiveCell.Value = "WebID Based SA"
    Range("K1").Select
    ActiveCell.Value = "SO Based SA"
    Range("L1").Select
    ActiveCell.Value = "Company ID"
    Range("M1").Select
    ActiveCell.Value = "Customer Journey Stage"
    Range("N1").Select
    ActiveCell.Value = "Customer Source"
    Range("O1").Select
    ActiveCell.Value = "Skip"
    
    ' Look up to WO ID & Get Smart a/c in lookup sheet
    
    Range("J2").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-1],Vlookup!C2:C3,2,0),"""")"
    ActiveCell.Copy
    
    Range("I" & Cells(Rows.Count, "I").End(xlUp).Row).Offset(0, 1).Select
    Range(Selection, Selection.End(xlUp)).Select
    Selection.FillDown
    
    ' Look up to SO ID & Get  Smart a/c in lookup sheet
    
    Range("K2").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-2],Vlookup!C1:C3,3,0),"""")"
    ActiveCell.Copy
    
    Range("j" & Cells(Rows.Count, "j").End(xlUp).Row).Offset(0, 1).Select
    Range(Selection, Selection.End(xlUp)).Select
    Selection.FillDown
    
    Range("j:K").EntireColumn.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    

    ' Filter Non Blanks in WebID Based SA column & SO Based SA
    
    Sheets("Raw Data").Activate
    With ActiveSheet.Range("A1", Selection.End(xlDown).End(xlToRight))
        .AutoFilter Field:=11, Criteria1:="="
        .AutoFilter Field:=10, Criteria1:="<>"
    End With
    
    ActiveSheet.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Cells(1, 11).Select
    ActiveCell.FormulaR1C1 = "=RC[-1]"
    ActiveCell.Copy
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveSheet.ShowAllData
    Columns("K:K").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues
    On Error GoTo 0
    Application.CutCopyMode = False
    
    'Lookup to CCW Company dump export
    
    'Bring Company Name
    
    Range("L2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1],Vlookup!C11:C15,2,0)"
    ActiveCell.Copy
    
    Range("k" & Cells(Rows.Count, "k").End(xlUp).Row).Offset(0, 1).Select
    Range(Selection, Selection.End(xlUp)).Select
    Selection.FillDown
    
    'Bring Journey Stage
    
    Range("M2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1],Vlookup!C12:C15,2,0)"
    ActiveCell.Copy
    
    Range("L" & Cells(Rows.Count, "L").End(xlUp).Row).Offset(0, 1).Select
    Range(Selection, Selection.End(xlUp)).Select
    Selection.FillDown
    
    'Bring Provision source
    
    Range("N2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-2],Vlookup!C12:C15,3,0)"
    ActiveCell.Copy
    
    Range("M" & Cells(Rows.Count, "M").End(xlUp).Row).Offset(0, 1).Select
    Range(Selection, Selection.End(xlUp)).Select
    Selection.FillDown
    
    'Bring Reachout?
    
    Range("O2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-3],Vlookup!C12:C15,4,0)"
    ActiveCell.Copy
    
    Range("N" & Cells(Rows.Count, "N").End(xlUp).Row).Offset(0, 1).Select
    Range(Selection, Selection.End(xlUp)).Select
    Selection.FillDown
    
    Range("J1").Select
    Range("J1").EntireColumn.Insert
    Range("J1").Select
    ActiveCell.Value = "Comment"
    
    Range("E1:E1, F1:F1, G1:G1, J1:J1, K1:K1, L1:L1, M1:M1, N1:N1, O1:O1, P1:P1").Interior.ColorIndex = 6


    MsgBox ("Yay..!, Data is ready to save")
    

    
    End Sub
    
    
    Sub Savereport()
    
    ActiveWorkbook.Sheets.Copy
    Application.DisplayAlerts = False
    ActiveWorkbook.Sheets("Macro").Delete
    Application.DisplayAlerts = False
    save_name = Application.GetSaveAsFilename(fileFilter:="Excel File (*.xlsx), *.xlsx")
    ActiveWorkbook.SaveAs Filename:=save_name, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    ActiveWorkbook.Save
    ActiveWindow.Close
    
    
    End Sub

    Sub ResetMacro()
    
    Dim xWs As Worksheet
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    For Each xWs In Application.ActiveWorkbook.Worksheets
        If xWs.Name <> "Macro" And xWs.Name <> "Vlookup" And xWs.Name <> "Raw Data" Then
            xWs.Delete
        End If
    Next
    Worksheets("Raw Data").Cells.Clear
    ActiveWorkbook.Save
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
End Sub
    
    Sub TEST()

Worksheets("Raw Data").Cells.Clear

End Sub
    