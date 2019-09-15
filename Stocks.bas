Attribute VB_Name = "Module1"
Sub Stocks():
'Declare variables
Dim StartYear As Double, EndYear As Double, total As Double, count As Integer, change As Double, PerChange As Double
Dim IncTick As String, DecTick As String, VolTick As String, IncVal As Double, DecVal As Double, VolVal As Double
    
'Loop through each worksheet
For Each ws In Worksheets
    'Activate the sheet being worked on
    ws.Activate
    
    'Set up headers for tables
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    'Challenge Chart
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    ActiveSheet.Columns("I:O").AutoFit
    
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    
    'Set Greatest Increase, Decrease, and volume
    IncVal = 0
    DecVal = 0
    VolVal = 0
    
    'Set total variable to 0
    total = 0
    'set count variable to 2, this tracks where to place values, starting second row
    count = 2
    'Set StartYear value to start
    StartYear = Cells(2, 3).Value
    'Loop through data in sheet
    For i = 2 To ws.Cells(Rows.count, 1).End(xlUp).Row
        'Add to total
        total = total + Cells(i, 7).Value
        'Find when the Ticker changes
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            'Get EndYear value
            EndYear = Cells(i, 6).Value
            
            'Calculate change and add to sheet
            change = Round(EndYear - StartYear, 8)
            Cells(count, 9).Value = Cells(i, 1).Value
            Cells(count, 10).Value = change
            Cells(count, 10).NumberFormat = "0.00000000"
            
            'Set conditional formating
            If change > 0 Then
                Cells(count, 10).Interior.Color = vbGreen
            ElseIf change < 0 Then
                Cells(count, 10).Interior.Color = vbRed
            End If
            
            'Make sure StartYear is not zero (divide by zero issue)
            If StartYear <> 0 Then
                'Calculate % change
                PerChange = change / StartYear
            Else
                PerChange = 0
            End If
            
            'Set % change
            Cells(count, 11).Value = PerChange
            Cells(count, 11).NumberFormat = "0.00%"
            Cells(count, 12).Value = total
            
            'Check if the % Change is greatest increase or decrease
            If PerChange > IncVal Then
                IncTick = Cells(i, 1).Value
                IncVal = PerChange
            ElseIf PerChange < DecVal Then
                DecTick = Cells(i, 1).Value
                DecVal = PerChange
            End If
            
            'Check if greastest total volume
            If total > VolVal Then
                VolTick = Cells(i, 1).Value
                VolVal = total
            End If
            
            'Reset StartYear value
            StartYear = Cells(i + 1, 3).Value
            'Increment count
            count = count + 1
            'Reset total value
            total = 0
        End If
            
    
    Next i
    
    'Set Challenge values
    Range("P2").Value = IncTick
    Range("Q2").Value = IncVal
    Range("Q2").NumberFormat = "0.00%"
    Range("P3").Value = DecTick
    Range("Q3").Value = DecVal
    Range("Q3").NumberFormat = "0.00%"
    Range("P4").Value = VolTick
    Range("Q4").Value = VolVal
    Range("Q4").NumberFormat = "General"

Next ws
    
End Sub
