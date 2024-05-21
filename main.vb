Sub SummarizeAndCheckOrders()
    Dim wsSource As Worksheet, wsMaster As Worksheet, wsDiscrepancies As Worksheet
    Dim strFilePath As String
    Dim fd As FileDialog
    Dim lastRow As Long, i As Long, j As Long
    Dim dict As Object
    Dim rng As Range
    Dim key As Variant
    Dim startColumn As String
    Dim startColIndex As Integer
    Dim dateA1 As Date, dateJ1 As Date
    Dim hasDateA1 As Boolean, hasDateJ1 As Boolean

    ' Turn off screen updating, events, and automatic calculations
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    ' Ensure the necessary sheets exist or create them
    On Error Resume Next
    Set wsMaster = ThisWorkbook.Sheets("Master")
    If wsMaster Is Nothing Then
        Set wsMaster = ThisWorkbook.Sheets.add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsMaster.Name = "Master"
    End If
    Set wsDiscrepancies = ThisWorkbook.Sheets("Discrepancies")
    If wsDiscrepancies Is Nothing Then
        Set wsDiscrepancies = ThisWorkbook.Sheets.add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsDiscrepancies.Name = "Discrepancies"
    End If
    On Error GoTo 0
    
     ' Call the log function
    LogOldestComparison
    
    ' Check if dates exist in A1 and J1
    hasDateA1 = False
    hasDateJ1 = False
    If IsDate(wsMaster.Range("A1").value) Then
        dateA1 = CDate(wsMaster.Range("A1").value)
        hasDateA1 = True
    End If
    If IsDate(wsMaster.Range("J1").value) Then
        dateJ1 = CDate(wsMaster.Range("J1").value)
        hasDateJ1 = True
    End If
    
    ' Determine the start column based on existing data and dates
    If hasDateA1 And hasDateJ1 Then
        If dateA1 <= dateJ1 Then
            startColumn = "A"
            startColIndex = 1
        Else
            startColumn = "J"
            startColIndex = 10
        End If
    ElseIf hasDateA1 Then
        startColumn = "J"
        startColIndex = 10
    Else
        startColumn = "A"
        startColIndex = 1
    End If
    
    ' Open file dialog to select the source workbook
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Select the Excel file with the orders"
        If .Show = True Then
            strFilePath = .SelectedItems(1)
        Else
            MsgBox "No file selected. Exiting macro."
            Exit Sub
        End If
    End With
    
    ' Open the selected workbook
    Dim wbSource As Workbook
    Set wbSource = Workbooks.Open(strFilePath)
    Set wsSource = wbSource.Worksheets(1) ' Assumes the data is on the first sheet
    
    ' Create dictionary to store customer orders
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Read through the data and fill the dictionary
    With wsSource
        lastRow = .Cells(.Rows.Count, "A").End(xlUp).row ' Find the last row with data in column A
        For i = 2 To lastRow ' Assuming row 1 has headers
            ' Convert the PromShip date to the day of the week
            Dim promShipDay As String
            promShipDay = Format(.Cells(i, "N").value, "dddd")
            
            ' Only include weekdays (exclude Saturday and Sunday)
            If promShipDay <> "Saturday" And promShipDay <> "Sunday" Then
                ' Construct a unique key for each customer and day of the week
                key = .Cells(i, "A").value & "|" & promShipDay
                If Not dict.exists(key) Then
                    dict(key) = 1
                Else
                    dict(key) = dict(key) + 1
                End If
            End If
        Next i
    End With
    
    ' Close the source workbook without saving
    wbSource.Close SaveChanges:=False
    
    ' Write the data from the dictionary to the master sheet
    With wsMaster
        ' Clear previous data in the target columns
        If startColumn = "A" Then
            .Range("A:G").Clear
        Else
            .Range("J:P").Clear
        End If
        
        ' Write date and time above the headers
        .Cells(1, startColIndex).value = Now
        
        ' Write headers if this is the first run in this column range
        If .Cells(2, startColIndex).value = "" Then
            .Cells(2, startColIndex).value = "CustID"
            .Cells(2, startColIndex + 1).value = "Monday"
            .Cells(2, startColIndex + 2).value = "Tuesday"
            .Cells(2, startColIndex + 3).value = "Wednesday"
            .Cells(2, startColIndex + 4).value = "Thursday"
            .Cells(2, startColIndex + 5).value = "Friday"
            .Cells(2, startColIndex + 6).value = "Number of Days Shipped"
        End If
        
        ' Initialize row number for writing
        i = 3
        
        ' Process dictionary and write to the master sheet
        For Each key In dict.Keys
            ' Split the key to get customer ID and day of the week
            Dim parts() As String
            parts = split(key, "|")
            
            ' Find or create the row for the customer
            Set rng = .Columns(startColIndex).Find(What:=parts(0), LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlPrevious)
            If rng Is Nothing Then
                ' If not found, write the new customer ID
                .Cells(i, startColIndex).value = parts(0)
                Set rng = .Cells(i, startColIndex)
                i = i + 1
            End If
            
            ' Write the order count to the appropriate day column
            Dim columnIndex As Integer
            columnIndex = MatchDayToColumnIndex(parts(1)) + (startColIndex - 1)
            If columnIndex <> 0 Then
                .Cells(rng.row, columnIndex).value = dict(key)
            End If
        Next key
        
        ' Calculate number of days shipped
        For Each row In .Range(.Cells(3, startColIndex), .Cells(i - 1, startColIndex))
            Dim numDaysShipped As Long
            numDaysShipped = Application.WorksheetFunction.CountIf(row.Offset(0, 1).Resize(1, 5), ">0")
            row.Offset(0, 6).value = numDaysShipped
        Next row
        
        ' Format the master sheet and apply filter to the top row
        .Rows(2).Font.Bold = True
        .Rows(2).HorizontalAlignment = xlCenter
        .Columns("A:P").AutoFit
    End With
    
    ' Compare the two sets of columns and log discrepancies
    With wsDiscrepancies
        ' Clear previous discrepancies
        .Cells.Clear
        .Cells(1, 1).value = "CustID"
        .Cells(1, 2).value = "Discrepancy"
        .Cells(1, 3).value = "Count of Discrepancies"
        j = 2
        
        ' Loop through each customer in the master sheet
        For Each row In wsMaster.Range(wsMaster.Cells(3, 1), wsMaster.Cells(i - 1, 1))
            Dim custID As String
            custID = row.value
            
            ' Find the corresponding rows in both column sets
            Dim rowA As Range, rowJ As Range
            Set rowA = wsMaster.Columns(1).Find(What:=custID, LookIn:=xlValues, LookAt:=xlWhole)
            Set rowJ = wsMaster.Columns(10).Find(What:=custID, LookIn:=xlValues, LookAt:=xlWhole)
            
            ' Check for discrepancies
            If Not rowA Is Nothing And Not rowJ Is Nothing Then
                Dim discrepancy As String
                Dim discrepancyCount As Integer
                discrepancy = ""
                discrepancyCount = 0
                
                For k = 1 To 5
                    If (rowA.Offset(0, k).value > 0 And rowJ.Offset(0, k).value = 0) Or _
                       (rowA.Offset(0, k).value = 0 And rowJ.Offset(0, k).value > 0) Then
                        discrepancy = discrepancy & " " & Choose(k, "Monday", "Tuesday", "Wednesday", "Thursday", "Friday")
                        discrepancyCount = discrepancyCount + 1
                    End If
                Next k
                
                If discrepancy <> "" Then
                    .Cells(j, 1).value = custID
                    .Cells(j, 2).value = discrepancy
                    .Cells(j, 3).value = discrepancyCount
                    j = j + 1
                End If
            End If
        Next row
    End With
    
        ' Turn screen updating, events, and automatic calculations back on
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic

    
    ' Display message to user that the summary and discrepancy check have been completed
    MsgBox "Order summary has been created in the master sheet and discrepancies logged in the Discrepancies sheet."
End Sub

' Helper function to match day names to column indexes
Function MatchDayToColumnIndex(dayName As String) As Integer
    Select Case dayName
        Case "Monday": MatchDayToColumnIndex = 2
        Case "Tuesday": MatchDayToColumnIndex = 3
        Case "Wednesday": MatchDayToColumnIndex = 4
        Case "Thursday": MatchDayToColumnIndex = 5
        Case "Friday": MatchDayToColumnIndex = 6
        Case Else: MatchDayToColumnIndex = 0
            MsgBox "Invalid day name: " & dayName, vbExclamation
    End Select
End Function

Sub LogOldestComparison()
    Dim wsComparison As Worksheet, wsLogs As Worksheet
    Dim dateA7 As Date, dateG7 As Date
    Dim lastLogRow As Long
    Dim sourceRange As Range
    Dim dateFormat As String
    Dim formattedDateA7 As String, formattedDateG7 As String

    ' Define the desired date format
    dateFormat = "DD/MM/YYYY HH:MM"

    ' Ensure the Comparison sheet exists
    On Error Resume Next
    Set wsComparison = ThisWorkbook.Sheets("Comparison")
    If wsComparison Is Nothing Then
        MsgBox "Comparison sheet does not exist. Exiting log function."
        Exit Sub
    End If
    On Error GoTo 0

    ' Ensure the Logs sheet exists or create it
    On Error Resume Next
    Set wsLogs = ThisWorkbook.Sheets("Logs")
    If wsLogs Is Nothing Then
        Set wsLogs = ThisWorkbook.Sheets.add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsLogs.Name = "Logs"
    End If
    On Error GoTo 0

    ' Check dates in A7 and G7
    If Not IsDate(wsComparison.Range("A7").value) Or Not IsDate(wsComparison.Range("G7").value) Then
        MsgBox "One or both dates in A7 and G7 are invalid. Exiting log function."
        Exit Sub
    End If

    dateA7 = CDate(wsComparison.Range("A7").value)
    dateG7 = CDate(wsComparison.Range("G7").value)

    ' Convert the dates to the desired string format
    formattedDateA7 = Format(dateA7, dateFormat)
    formattedDateG7 = Format(dateG7, dateFormat)

    ' Determine the oldest date and set the corresponding range to copy
    If dateA7 <= dateG7 Then
        Set sourceRange = wsComparison.Range("A7:B13")
    Else
        Set sourceRange = wsComparison.Range("G7:H13")
    End If

    ' Paste the copied table as values to the Logs sheet
    With wsLogs
        lastLogRow = .Cells(.Rows.Count, 1).End(xlUp).row
        If lastLogRow > 1 Then lastLogRow = lastLogRow + 2 ' Ensure there's a gap between logs
        
        ' Paste the date in the desired format
        .Cells(lastLogRow, 1).value = formattedDateA7
        .Cells(lastLogRow + 1, 1).Resize(sourceRange.Rows.Count - 1, sourceRange.Columns.Count).value = sourceRange.Offset(1, 0).value

        Application.CutCopyMode = False
    End With
End Sub



