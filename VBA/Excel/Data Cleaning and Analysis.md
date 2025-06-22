# Excel VBA Codes for Data Cleaning and Analysis

Here you can find VBA macros to help with your data cleaning, filtering, and analysis tasks across four reference tables. Here's my approach:

## 1. Reference Tables Creation

First, let's create four reference tables to use with our VBA macros:

### Table 1: Fruits Inventory
```
FruitID | FruitName | Category | OriginCountry | PricePerKg | InStock | ExpiryDate | Organic
--------|-----------|----------|---------------|------------|---------|------------|--------
1       | Apple     | Pome     | USA           | 2.99       | TRUE    | 2023-12-15 | TRUE
2       | Banana    | Tropical | Ecuador       | 1.49       | TRUE    | 2023-11-20 | FALSE
3       | Orange    | Citrus   | Spain         | 3.29       | FALSE   | 2023-11-10 | TRUE
4       | Mango     | Tropical | Mexico        | 4.99       | TRUE    | 2023-11-25 | FALSE
5       | Blueberry | Berry    | Canada        | 8.99       | TRUE    | 2023-11-05 | TRUE
```

### Table 2: Employee Data
```
BadgeID | FirstName | LastName | Role           | Department | HireDate   | Terminated | Salary  | ManagerID
--------|-----------|----------|----------------|------------|------------|------------|---------|----------
1001    | John      | Smith    | Data Analyst   | Analytics  | 2018-05-15 | FALSE      | 75000   | 2001
1002    | Sarah     | Johnson  | BI Developer   | IT         | 2020-02-10 | FALSE      | 82000   | 2002
1003    | Michael   | Williams | DBA            | IT         | 2019-11-05 | TRUE       | 95000   | 2002
1004    | Emily     | Brown    | Data Scientist | Analytics  | 2021-08-22 | FALSE      | 88000   | 2001
1005    | David     | Jones    | Analyst        | Finance    | 2017-03-30 | FALSE      | 68000   | 2003
```

### Table 3: Computer Assets
```
AssetID | SerialNumber | Hostname    | Model          | OSVersion  | PurchaseDate | WarrantyEnd | AssignedTo | Status
--------|--------------|-------------|----------------|------------|--------------|-------------|------------|--------
C001    | SN12345678   | WS-JSMITH   | Dell XPS 15    | Win10 21H2 | 2021-03-10   | 2024-03-10  | 1001       | Active
C002    | SN23456789   | WS-SJOHNSON | MacBook Pro M1 | macOS 13   | 2022-01-15   | 2025-01-15  | 1002       | Active
C003    | SN34567890   | WS-MWILLIAM | Lenovo T490    | Win11 22H2 | 2020-08-20   | 2023-08-20  | 1003       | Retired
C004    | SN45678901   | WS-EBROWN   | HP EliteBook   | Win10 22H2 | 2022-05-05   | 2025-05-05  | 1004       | Active
C005    | SN56789012   | WS-DJONES   | Surface Pro 8  | Win11 22H2 | 2021-11-12   | 2024-11-12  | 1005       | Active
```

### Table 4: Financial Data
```
TransactionID | Date       | AccountID | Description         | Category    | Amount  | Currency | Reconciled
--------------|------------|-----------|---------------------|-------------|---------|----------|-----------
T1001         | 2023-10-01 | A101      | Office Supplies     | Expenses    | 245.50  | USD      | TRUE
T1002         | 2023-10-02 | A102      | Software License    | IT          | 1200.00 | USD      | FALSE
T1003         | 2023-10-03 | A103      | Client Payment      | Revenue     | 5500.00 | USD      | TRUE
T1004         | 2023-10-04 | A101      | Travel Expenses     | Travel      | 875.30  | USD      | FALSE
T1005         | 2023-10-05 | A104      | Equipment Purchase  | Capital     | 3200.00 | USD      | TRUE
```

## 2. VBA Codes

Here are 30 practical VBA macros for working with these tables:

### General Utilities (5 macros)

```vba
' 1. Clear all filters in a worksheet
Sub ClearAllFilters()
    On Error Resume Next
    If ActiveSheet.AutoFilterMode Then ActiveSheet.AutoFilterMode = False
    On Error GoTo 0
End Sub
```  
```vba   
' 2. Convert text to proper case (capitalize first letter of each word)
Sub ConvertToProperCase()
    Dim rng As Range
    For Each rng In Selection
        If Not IsEmpty(rng) Then rng.Value = WorksheetFunction.Proper(rng.Value)
    Next rng
End Sub
```
```vba  
' 3. Remove duplicates based on selected columns
Sub RemoveDuplicatesSelectedColumns()
    Dim colCount As Integer
    colCount = Selection.Columns.Count
    ActiveSheet.Range(Selection.Address).RemoveDuplicates Columns:=Evaluate("=ROW(1:" & colCount & ")"), Header:=xlYes
End Sub
```
```vba
' 4. Highlight cells with formulas
Sub HighlightFormulas()
    Dim cell As Range
    For Each cell In ActiveSheet.UsedRange
        If cell.HasFormula Then
            cell.Interior.Color = RGB(255, 255, 0) ' Yellow
        End If
    Next cell
End Sub
```
```vba   
' 5. Freeze top row and first column
Sub FreezePaneTopRowFirstColumn()
    With ActiveWindow
        .SplitColumn = 1
        .SplitRow = 1
        .FreezePanes = True
    End With
End Sub
```

### Fruit Table Macros (5 macros)

```vba
' 6. Filter organic fruits only
Sub FilterOrganicFruits()
    Dim ws As Worksheet
    Set ws = Worksheets("Fruits") ' Change to your sheet name
    ws.AutoFilterMode = False
    ws.Range("A1:H6").AutoFilter Field:=8, Criteria1:="TRUE"
End Sub
```
```vba
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
```   
```vba   
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
```
```vba   
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
```
```vba   
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
```

### Employee Data Macros (6 macros)

```vba
' 11. Filter active employees only
Sub FilterActiveEmployees()
    Dim ws As Worksheet
    Set ws = Worksheets("Employees") ' Change to your sheet name
    ws.AutoFilterMode = False
    ws.Range("A1:I6").AutoFilter Field:=7, Criteria1:="FALSE"
End Sub
```
```vba
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
```
```vba   
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
```
```vba
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
```
```vba
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
```   
```vba   
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
```   
