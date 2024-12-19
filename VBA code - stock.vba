Sub year_stock()
    ' Create main data table
    Dim wsCombined As Worksheet
    Set wsCombined = Sheets.Add
    wsCombined.Name = "Q1-Q4"

    ' Copy the headers from the first sheet ("Q1")
    Sheets("Q1").Range("A1:G1").Copy wsCombined.Range("A1:G1")
    
    ' Combine all data together from all sheets
    Dim i As Integer
    Dim LastRow As Long
    Dim LastRowCombined As Long

    ' Iterate through sheets to combine data
    Dim ShtCount As Long
    ShtCount = ActiveWorkbook.Sheets.Count

    For i = 2 To ShtCount
    If Worksheets(i).Name <> "Q1-Q4" Then
     With Worksheets(i)
     LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
    .Range("A2:G" & LastRow).Copy
     End With
     With wsCombined
     LastRowCombined = .Cells(.Rows.Count, "A").End(xlUp).Row + 1
     .Cells(LastRowCombined, 1).PasteSpecial
     End With
    End If
    Next i

    ' Create new columns
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Quarterly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"

    ' variables for calculations
    Dim t As Long, j As Long
    Dim QChange As Double
    Dim perChange As Double
    Dim LastRowA As Long

    ' Set ws to the combined data sheet
    Set ws = wsCombined
    LastRowA = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Initialize counters
    t = 2
    j = 2

    For f = 2 To LastRowA
    
    ' Check for a new ticker
    If ws.Cells(f + 1, 1).Value <> ws.Cells(f, 1).Value Then
    ' Ticker
    ws.Cells(t, 9).Value = ws.Cells(f, 1).Value
    ' Quarterly change
    QChange = ws.Cells(f, 6).Value - ws.Cells(j, 3).Value
    ws.Cells(t, 10).Value = QChange
    
    ' Color formating
    If QChange < 0 Then
    ws.Cells(t, 10).Interior.Color = RGB(255, 0, 0)
    Else
    ws.Cells(t, 10).Interior.Color = RGB(0, 255, 0)
    End If

    ' Percent change calculation
    If ws.Cells(j, 3).Value <> 0 Then
    perChange = (ws.Cells(f, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value
    Else
    perChange = 0
    End If
    
    'Percent formating
    ws.Cells(t, 11).Value = perChange
    ws.Cells(t, 11).NumberFormat = "0%"

    ' Total stock volume calculation
    ws.Cells(t, 12).Value = Application.Sum(ws.Range(ws.Cells(j, 7), ws.Cells(f, 7)))
    
    t = t + 1
    End If
    j = f + 1
    Next f

'Summary data
Dim LastRowT As Long
LastRowT = ws.Cells(Rows.Count, 9).End(xlUp).Row

'Setting variables
Dim GreatVol As Double
Dim GreatIncr As Double
Dim GreatDecr As Double


'Create new columns
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"

'Initialize data for summary
GreatVol = ws.Cells(2, 12).Value
GreatIncr = ws.Cells(2, 11).Value
GreatDecr = ws.Cells(2, 11).Value

'Loop for summary
'Greatest volume
For i = 2 To LastRowT
    If ws.Cells(i, 12).Value > GreatVol Then
    GreatVol = ws.Cells(i, 12).Value
    ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
    Else
    GreatVol = GreatVol
    End If
'Greatest increase
    If ws.Cells(i, 11).Value > GreatIncr Then
    GreatIncr = ws.Cells(i, 11).Value
    ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
    Else
    GreatIncr = GreatIncr
    End If
'Greatest decrease
    If ws.Cells(i, 11).Value < GreatDecr Then
    GreatDecr = ws.Cells(i, 11).Value
    ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
    Else
    GreatDecr = GreatDecr
    End If
                
'Write summary results
    ws.Cells(2, 17).Value = Format(GreatIncr, "Percent")
    ws.Cells(3, 17).Value = Format(GreatDecr, "Percent")
    ws.Cells(4, 17).Value = Format(GreatVol, "Scientific")
    Next i
End Sub
