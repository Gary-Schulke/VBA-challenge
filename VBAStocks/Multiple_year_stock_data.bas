Attribute VB_Name = "Module1"
'The button event handler for calculating all sheets in the workbook.
Sub CalculateAllSheets_Click()
    Call CalculateAllSheets
End Sub

'The button evnet handler for calculating the active sheet.
Sub CalculateActiveSheet_Click()
    Call CalculateActiveSheet
End Sub

'Make a string for using the Range() function.
'Takes the sheet name and current row and creates the string needed to use Range(...)
'Returns the current row and then next row as a range.
Function mkRng(sht As String, rwcnt As Long) As Range
    'sht - The sheet name.
    'rwcnt - The current row of interest.
    Dim rw As String
    Dim nrw As String
    rw = CStr(rwcnt)
    nrw = CStr(rwcnt + 1)
    Set mkRng = Range(sht + "!A" + CStr(rw) + ":G" + CStr(nrw))
End Function

'Writes the results from each stock tag and writes it to the designated area on each sheet.
'Formats yearlyChange, red for negative values and green for positive values
Sub writeSheetResults(sht As String, resultRow As Integer, ticker As String, yearlyChange As Double, percentchange As Double, volume As Double)
    'sht - The sheet name.
    'resultRow - The row the results should be written to.
    'ticker - The ticker name being analyzed. (From column A)
    'yearlyChange - Year closing price - Year opening price.
    'percentchange - year opening price / year closing price
    'volume - The sum of daily volumes for the year.
    Dim rw As String
    Dim sRng As String
    Dim rng As Range
    rw = CStr(resultRow)
    sRng = sht + "!I" + CStr(rw) + ": L" + CStr(rw)
    Set rng = Range(sRng)
    rng.Cells(1, 1).Value = ticker
    rng.Cells(1, 2).Value = yearlyChange
    rng.Cells(1, 2).FormatConditions.Delete
    With rng.Cells(1, 2).FormatConditions.Add(xlCellValue, xlLess, "=0")
            .Interior.Color = -16776961
            .StopIfTrue = False
    End With
    With rng.Cells(1, 2).FormatConditions.Add(xlCellValue, xlGreater, "=0")
            .Interior.Color = -16744448
            .StopIfTrue = False
    End With
    rng.Cells(1, 3).NumberFormat = "0.00%"
    rng.Cells(1, 3).Value = percentchange
    rng.Cells(1, 4).Value = volume
      
    'Update best and worst performers.
    sRng = (sht + "!O1:Q4")
    Set rng = Range(sRng)
    If (percentchange > rng.Cells(2, 3).Value) Then
        rng.Cells(2, 2).Value = ticker
        rng.Cells(2, 3).Value = percentchange
    End If
    If (percentchange < rng.Cells(3, 3).Value) Then
        rng.Cells(3, 2).Value = ticker
        rng.Cells(3, 3).Value = percentchange
    End If
     If (volume > rng.Cells(4, 3).Value) Then
        rng.Cells(4, 2).Value = ticker
        rng.Cells(4, 3).Value = volume
    End If

End Sub

'Clears previous results.  The table may not show it when the script is running.
'Creates the column headers for the results table.
Sub writeSheetHeaders(sht As String)
    Dim sRng As String
    Dim rng As Range
    Dim lastRowData As Long
    Dim clearRange As String
    
    'Clear previous results
    lastRowData = Range("I1").End(xlDown).Row
    clearRange = (sht + "!I1:L" + CStr(lastRowData))
    Range(clearRange).Clear
            
    'Table with each ticker
    sRng = (sht + "!I1:L1")
    Set rng = Range(sRng)

    rng.Cells(1, 1).Value = "Ticker"
    rng.Cells(1, 2).Value = "Yearly Change"
    rng.Cells(1, 3).Value = "Percent Change"
    rng.Cells(1, 4).Value = "Total Stock Volume"
    rng.Columns.AutoFit
    
    'Best and Worst Performers
    'Clear previous results, label and format.
    sRng = (sht + "!O1:Q4")
    Set rng = Range(sRng)
    rng.Clear
    rng.Cells(1, 2).Value = "Ticker"
    rng.Cells(1, 3).Value = "Value     "
    rng.Cells(2, 1).Value = "Greatest % Increase"
    rng.Cells(2, 3).NumberFormat = "0.00%"
    rng.Cells(3, 1).Value = "Greatest % Decrease"
    rng.Cells(3, 3).NumberFormat = "0.00%"
    rng.Cells(4, 1).Value = "Greatest Total Volume"
    rng.Columns.AutoFit

End Sub
'Loops through all sheets in the workbook
Sub CalculateAllSheets()
    For Each wks In ThisWorkbook.Sheets
        Call CaluclateSheet(wks)
    Next wks
End Sub
'Calculates only the active sheet.
Sub CalculateActiveSheet()
        Call CaluclateSheet(ActiveSheet)
End Sub
'Performs the analysis for on a worksheet.
Sub CaluclateSheet(wks As Variant)
    'wks - the worksheet being analyzed.
Dim ticker As String
Dim OpenJan1 As Double
Dim CloseDec31 As Double
Dim lastRow As Long
Dim sumVolume As Double
Dim sheetName As String
Dim currentRange As Range
Dim percentchange As Double
Dim resultRow As Integer
Dim rowCount As Long
Dim nextTicker As String
'Get workbook names
    lastRow = wks.Cells.SpecialCells(xlCellTypeLastCell).Row
    sheetName = wks.Name
   ' Debug.Print (wks.Name + Str(lastRow))
rowCount = 2
'Create the results headers.
Call writeSheetHeaders(sheetName)

Set currentRange = mkRng(sheetName, rowCount)

ticker = currentRange.Cells(1, 1).Value
sumVolume = 0 'currentRange.Cells(1, 7).Value
OpenJan1 = 0

nextTicker = currentRange.Cells(2, 1).Value
resultRow = 2
'Loop until the ticker cell value is empty
While (nextTicker <> "")
    'Loop while the ticker name is unchanged for the next row
    While (ticker = nextTicker)
        If OpenJan1 = 0 Then
            OpenJan1 = currentRange.Cells(1, 3).Value
        End If

        sumVolume = sumVolume + currentRange.Cells(1, 7).Value
        rowCount = rowCount + 1
        Set currentRange = mkRng(sheetName, rowCount)
        nextTicker = currentRange.Cells(2, 1).Value
    Wend
    'The loop is complete.  Get data from the last line.
    sumVolume = sumVolume + currentRange.Cells(1, 7).Value
    CloseDec31 = currentRange.Cells(1, 6).Value
    'Calculate the yearly percent change.
    If CloseDec31 = 0 And OpenJan1 = 0 Then
        percentchange = 0
    ElseIf CloseDec31 = 0 Then
        percentchange = 1
    ElseIf OpenJan1 <= CloseDec31 Then
                percentchange = 1 - (OpenJan1 / CloseDec31)
    Else
        percentchange = ((CloseDec31 - OpenJan1) / OpenJan1)
    End If
    'Calculate the yearly change
    Dim yearlyChange As Double
    yearlyChange = CloseDec31 - OpenJan1
    'Debug.Print (ticker + " " + Str(yearlyChange) + Str(percentchange) + Str(sumVolume))
    'Write the results to the results table.
    Call writeSheetResults(sheetName, resultRow, ticker, yearlyChange, percentchange, sumVolume)
    'Reset variables for the next loop.
    resultRow = resultRow + 1
    OpenJan1 = 0
    nextTicker = currentRange.Cells(2, 1).Value
    rowCount = rowCount + 1
    ticker = nextTicker
    sumVolume = 0
    'Advance the range.
    Set currentRange = mkRng(sheetName, rowCount)
Wend
    resultRow = 2
End Sub
