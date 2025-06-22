' 1. Clear all filters in a worksheet
Sub ClearAllFilters()
    On Error Resume Next
    If ActiveSheet.AutoFilterMode Then ActiveSheet.AutoFilterMode = False
    On Error GoTo 0
End Sub

' 2. Convert text to proper case (capitalize first letter of each word)
Sub ConvertToProperCase()
    Dim rng As Range
    For Each rng In Selection
        If Not IsEmpty(rng) Then rng.Value = WorksheetFunction.Proper(rng.Value)
    Next rng
End Sub

' 3. Remove duplicates based on selected columns
Sub RemoveDuplicatesSelectedColumns()
    Dim colCount As Integer
    colCount = Selection.Columns.Count
    ActiveSheet.Range(Selection.Address).RemoveDuplicates Columns:=Evaluate("=ROW(1:" & colCount & ")"), Header:=xlYes
End Sub

' 4. Highlight cells with formulas
Sub HighlightFormulas()
    Dim cell As Range
    For Each cell In ActiveSheet.UsedRange
        If cell.HasFormula Then
            cell.Interior.Color = RGB(255, 255, 0) ' Yellow
        End If
    Next cell
End Sub

' 5. Freeze top row and first column
Sub FreezePaneTopRowFirstColumn()
    With ActiveWindow
        .SplitColumn = 1
        .SplitRow = 1
        .FreezePanes = True
    End With
End Sub

' 6. Filter organic fruits only
Sub FilterOrganicFruits()
    Dim ws As Worksheet
    Set ws = Worksheets("Fruits") ' Change to your sheet name
    ws.AutoFilterMode = False
    ws.Range("A1:H6").AutoFilter Field:=8, Criteria1:="TRUE"
End Sub

' 7. Apply conditional formatting to highlight expired fruits
Sub HighlightExpiredFruits()
    Dim ws As Worksheet
    Set ws = Worksheets("Fruits")
    
    With ws.Range("G2:G6") ' Adjust range as needed
        .FormatConditions.Delete
        .FormatConditions.Add Type:=xlExpression, Formula1:="=G2<TODAY()"
        .FormatConditions(1).Interior.Color = RGB(255, 0, 0) ' Red
    End With
End Sub

' 8. Calculate total inventory value
Sub CalculateTotalInventoryValue()
    Dim ws As Worksheet
    Dim total As Double
    Set ws = Worksheets("Fruits")
    
    For i = 2 To ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        If ws.Cells(i, 6).Value = True Then ' InStock = TRUE
            total = total + ws.Cells(i, 5).Value ' PricePerKg
        End If
    Next i
    
    MsgBox "Total inventory value: $" & Format(total, "0.00"), vbInformation, "Inventory Value"
End Sub

' 9. Import fruits from a text file
Sub ImportFruitsFromTextFile()
    Dim filePath As String
    filePath = Application.GetOpenFilename("Text Files (*.txt), *.txt")
    
    If filePath <> "False" Then
        With ActiveSheet.QueryTables.Add(Connection:="TEXT;" & filePath, _
            Destination:=Range("A1"))
            .TextFileParseType = xlDelimited
            .TextFileCommaDelimiter = True
            .Refresh
        End With
    End If
End Sub

' 10. Generate summary statistics for fruits
Sub FruitSummaryStatistics()
    Dim ws As Worksheet
    Set ws = Worksheets("Fruits")
    
    ws.Range("J1").Value = "Category"
    ws.Range("K1").Value = "Count"
    ws.Range("L1").Value = "Avg Price"
    
    ' Get unique categories
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    ws.Range("C2:C" & lastRow).AdvancedFilter Action:=xlFilterCopy, _
        CopyToRange:=ws.Range("J2"), Unique:=True
        
    ' Calculate count and average for each category
    Dim catRow As Long
    catRow = 2
    Do While ws.Cells(catRow, "J").Value <> ""
        Dim cat As String
        cat = ws.Cells(catRow, "J").Value
        
        ' Count
        ws.Cells(catRow, "K").Value = Application.WorksheetFunction.CountIf(ws.Range("C2:C" & lastRow), cat)
        
        ' Average price
        ws.Cells(catRow, "L").Value = Application.WorksheetFunction.AverageIf(ws.Range("C2:C" & lastRow), cat, ws.Range("E2:E" & lastRow))
        
        catRow = catRow + 1
    Loop
End Sub

' 11. Filter active employees only
Sub FilterActiveEmployees()
    Dim ws As Worksheet
    Set ws = Worksheets("Employees") ' Change to your sheet name
    ws.AutoFilterMode = False
    ws.Range("A1:I6").AutoFilter Field:=7, Criteria1:="FALSE"
End Sub

' 12. Calculate employee tenure in years
Sub CalculateEmployeeTenure()
    Dim ws As Worksheet
    Set ws = Worksheets("Employees")
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Add tenure column if it doesn't exist
    If ws.Range("J1").Value <> "Tenure (Years)" Then
        ws.Range("J1").Value = "Tenure (Years)"
    End If
    
    ' Calculate tenure for each employee
    For i = 2 To lastRow
        ws.Cells(i, "J").Value = Year(Date) - Year(ws.Cells(i, "F").Value) + _
            (DateSerial(Year(Date), Month(ws.Cells(i, "F").Value), Day(ws.Cells(i, "F").Value)) > Date)
    Next i
End Sub

' 13. Generate department salary report
Sub DepartmentSalaryReport()
    Dim ws As Worksheet
    Set ws = Worksheets("Employees")
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Create summary table
    ws.Range("L1").Value = "Department"
    ws.Range("M1").Value = "Avg Salary"
    ws.Range("N1").Value = "Max Salary"
    ws.Range("O1").Value = "Min Salary"
    ws.Range("P1").Value = "Employee Count"
    
    ' Get unique departments
    ws.Range("E2:E" & lastRow).AdvancedFilter Action:=xlFilterCopy, _
        CopyToRange:=ws.Range("L2"), Unique:=True
        
    ' Calculate statistics for each department
    Dim deptRow As Long
    deptRow = 2
    Do While ws.Cells(deptRow, "L").Value <> ""
        Dim dept As String
        dept = ws.Cells(deptRow, "L").Value
        
        ' Average salary
        ws.Cells(deptRow, "M").Value = Application.WorksheetFunction.AverageIf(ws.Range("E2:E" & lastRow), dept, ws.Range("H2:H" & lastRow))
        
        ' Max salary
        ws.Cells(deptRow, "N").Value = Application.WorksheetFunction.MaxIfs(ws.Range("H2:H" & lastRow), ws.Range("E2:E" & lastRow), dept)
        
        ' Min salary
        ws.Cells(deptRow, "O").Value = Application.WorksheetFunction.MinIfs(ws.Range("H2:H" & lastRow), ws.Range("E2:E" & lastRow), dept)
        
        ' Employee count
        ws.Cells(deptRow, "P").Value = Application.WorksheetFunction.CountIf(ws.Range("E2:E" & lastRow), dept)
        
        deptRow = deptRow + 1
    Loop
End Sub

' 14. Send birthday reminders
Sub BirthdayReminders()
    Dim ws As Worksheet
    Set ws = Worksheets("Employees")
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Dim bdayList As String
    bdayList = "Upcoming birthdays:" & vbCrLf & vbCrLf
    
    For i = 2 To lastRow
        If Month(ws.Cells(i, "F").Value) = Month(Date) And Day(ws.Cells(i, "F").Value) - Day(Date) <= 7 And Day(ws.Cells(i, "F").Value) >= Day(Date) Then
            bdayList = bdayList & ws.Cells(i, "B").Value & " " & ws.Cells(i, "C").Value & ": " & _
                MonthName(Month(ws.Cells(i, "F").Value)) & " " & Day(ws.Cells(i, "F").Value) & vbCrLf
        End If
    Next i
    
    If Len(bdayList) > 25 Then
        MsgBox bdayList, vbInformation, "Birthday Reminders"
    Else
        MsgBox "No birthdays in the next 7 days", vbInformation, "Birthday Reminders"
    End If
End Sub

' 15. Create employee directory (name, role, department)
Sub CreateEmployeeDirectory()
    Dim ws As Worksheet
    Set ws = Worksheets("Employees")
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Create new worksheet for directory
    Dim dirSheet As Worksheet
    On Error Resume Next
    Set dirSheet = Worksheets("Employee Directory")
    On Error GoTo 0
    
    If dirSheet Is Nothing Then
        Set dirSheet = Worksheets.Add(After:=Worksheets(Worksheets.Count))
        dirSheet.Name = "Employee Directory"
    Else
        dirSheet.Cells.Clear
    End If
    
    ' Set up directory headers
    dirSheet.Range("A1").Value = "Name"
    dirSheet.Range("B1").Value = "Role"
    dirSheet.Range("C1").Value = "Department"
    dirSheet.Range("D1").Value = "Email"
    
    ' Populate directory
    For i = 2 To lastRow
        dirSheet.Cells(i, 1).Value = ws.Cells(i, "B").Value & " " & ws.Cells(i, "C").Value
        dirSheet.Cells(i, 2).Value = ws.Cells(i, "D").Value
        dirSheet.Cells(i, 3).Value = ws.Cells(i, "E").Value
        dirSheet.Cells(i, 4).Value = LCase(ws.Cells(i, "B").Value & "." & ws.Cells(i, "C").Value & "@company.com")
    Next i
    
    ' Format directory
    dirSheet.Range("A1:D1").Font.Bold = True
    dirSheet.Columns("A:D").AutoFit
    dirSheet.Range("A1:D" & lastRow).Borders.LineStyle = xlContinuous
End Sub

' 16. Calculate salary increase (3% for all active employees)
Sub CalculateSalaryIncrease()
    Dim ws As Worksheet
    Set ws = Worksheets("Employees")
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Add new salary column if it doesn't exist
    If ws.Range("K1").Value <> "New Salary" Then
        ws.Range("K1").Value = "New Salary"
    End If
    
    ' Calculate new salary
    For i = 2 To lastRow
        If ws.Cells(i, "G").Value = False Then ' Not terminated
            ws.Cells(i, "K").Value = ws.Cells(i, "H").Value * 1.03
        Else
            ws.Cells(i, "K").Value = ws.Cells(i, "H").Value
        End If
    Next i
End Sub

' 17. Export terminated employees to CSV
Sub ExportTerminatedEmployees()
    Dim ws As Worksheet
    Set ws = Worksheets("Employees")
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Filter for terminated employees
    ws.AutoFilterMode = False
    ws.Range("A1:I" & lastRow).AutoFilter Field:=7, Criteria1:="TRUE"
    
    ' Copy to new workbook
    Dim newBook As Workbook
    Set newBook = Workbooks.Add
    ws.UsedRange.SpecialCells(xlCellTypeVisible).Copy newBook.Sheets(1).Range("A1")
    
    ' Save as CSV
    Dim savePath As String
    savePath = Application.GetSaveAsFilename(InitialFileName:="Terminated_Employees_" & Format(Date, "yyyymmdd"), _
        FileFilter:="CSV Files (*.csv), *.csv")
    
    If savePath <> "False" Then
        newBook.SaveAs Filename:=savePath, FileFormat:=xlCSV
        newBook.Close False
        MsgBox "Terminated employees
