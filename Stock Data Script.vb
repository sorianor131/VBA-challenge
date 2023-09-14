Sub stock_data():

    'Loop through all sheets within this workbook and add header values to given range
For Each ws In Worksheets

    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"


    'Declare variables for ticker name, total volume, open/close price, yearly/percent change and summary row
    Dim Ticker As String
    Dim Volume As Double
        Volume = 0
    Dim OpenPrice As Double
        OpenPrice = ws.Cells(2, 3).Value
    Dim ClosePrice As Double
    Dim YearlyChange As Double
        YearlyChange = 0
    Dim PercentChange As Double
    Dim SummaryRow As Integer
        SummaryRow = 2
            
    'Determine last row in column A
    LastRow = ws.Range("A1").End(xlDown).Row

    'Loop through the first column to determine if next value is not equal to previous value
    For i = 2 To LastRow
    
    'If not equal to next value then set ticker name to that value as well as close price, yearly/percent change and volume for said ticker
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        Ticker = ws.Cells(i, 1).Value
        ClosePrice = ws.Cells(i, 6).Value
        'Yearly change calculation
        YearlyChange = ClosePrice - OpenPrice
        'Percent change calculation
        PercentChange = YearlyChange / OpenPrice
        Volume = Volume + ws.Cells(i, 7).Value
        
        'Print variable values
        ws.Range("I" & SummaryRow).Value = Ticker
        ws.Range("J" & SummaryRow).Value = YearlyChange
        ws.Range("K" & SummaryRow).Value = PercentChange
        ws.Range("L" & SummaryRow).Value = Volume
        
        'Add one row to the summary table
        SummaryRow = SummaryRow + 1
        'Reset variables
        OpenPrice = ws.Cells(i + 1, 3).Value
        YearlyChange = 0
        Volume = 0
        
        Else
        
        'Add to the volume total
        Volume = Volume + ws.Cells(i, 7).Value
        
        End If
        
    Next i
        
    'Declare variables for greatest increase/decrease and total volume
    Dim Increase As Double
        Increase = 0
    Dim Decrease As Double
        Decrease = 0
    Dim VolumeMax As Double
        VolumeMax = 0
    
    'Determine last row in column I
    LastRow = ws.Range("I1").End(xlDown).Row
    
    'Loop through all percent changes to identify greatest increase/decrease
    For i = 2 To LastRow
    
        If Increase < ws.Cells(i, 11).Value Then
        Ticker = ws.Cells(i, 9).Value
        Increase = ws.Cells(i, 11).Value
        ws.Range("P2").Value = Ticker
        ws.Range("Q2").Value = Increase
        'Format value to percent
        ws.Range("Q2").NumberFormat = "0.00%"
        
        ElseIf Decrease > ws.Cells(i, 11).Value Then
        Ticker = ws.Cells(i, 9).Value
        Decrease = ws.Cells(i, 11).Value
        ws.Range("P3").Value = Ticker
        ws.Range("Q3").Value = Decrease
        'Format value to percent
        ws.Range("Q3").NumberFormat = "0.00%"
        
        ElseIf VolumeMax < ws.Cells(i, 12).Value Then
        Ticker = ws.Cells(i, 9).Value
        VolumeMax = ws.Cells(i, 12).Value
        ws.Range("P4").Value = Ticker
        ws.Range("Q4").Value = VolumeMax
        
        End If
        
    Next i
    
    'Determine last row in column I
    LastRow = ws.Range("I1").End(xlDown).Row

    'Loop through yearly change to apply conditional formatting
    For i = 2 To LastRow

        If ws.Cells(i, 10).Value < YearlyChange Then
        ws.Cells(i, 10).Interior.ColorIndex = 3 'Red

        ElseIf ws.Cells(i, 10).Value >= YearlyChange Then
        ws.Cells(i, 10).Interior.ColorIndex = 4 'Green

        End If
        
    Next i
    
    'Format value to percent
    ws.Range("K2:K" & LastRow).NumberFormat = "0.00%"
    
    'Auto fit worksheet columns
    ws.UsedRange.EntireColumn.AutoFit

    Next ws
    
End Sub