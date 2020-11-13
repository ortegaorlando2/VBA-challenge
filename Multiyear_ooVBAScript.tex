Sub cycles():

Dim ticker As String
Dim l As Long
Dim year As Integer
Dim priceopen As Double
Dim priceclose As Double
Dim PercentChange As Double
Dim volume As Double
Dim change As Double
Dim MaxIncreaseTicker As String
Dim MaxDecreaseTicker As String
Dim GreatestVolumeTicker As String
Dim MaxIncrease As Double
Dim MaxDecrease As Double
Dim GreatestVolume As Double

MaxIncrease = 0
MaxDecrease = 0
GreatestVolume = 0

'loop to change the worksheet
For year = 1 To 3 '(1,2
Worksheets(year).Select
Range("A1").Select

'Creates table headers
Range("A1:G1").Interior.ColorIndex = 42
Range("J1:M1").Interior.ColorIndex = 42
Range("Q1:R1").Interior.ColorIndex = 42
Range("P2:P4").Interior.ColorIndex = 38
Cells(1, 10).Value = "Ticker"
Cells(1, 11).Value = "Yearly_Change"
Cells(1, 12).Value = "Percent_Change"
Cells(1, 13).Value = "Total_Stock_Volume"
Cells(1, 17).Value = "Ticker"
Cells(1, 18).Value = "Value"
Cells(2, 16).Value = "Greatest % Increase"
Cells(3, 16).Value = "Greatest % Decrease"
Cells(4, 16).Value = "Greatest Total Volume"


'Assigns first ticker for the table
ticker = Cells(2, 1).Value
Cells(2, 10).Value = ticker


'counting number of lines(rows) in worksheet using Stack Overflow posting by docjay
l = Range("A1", Range("A1").End(xlDown)).Rows.Count

'Initializes counter for summary table rows
k = 2

'Initializes stock price and volume
priceopen = Cells(2, 3).Value
volume = 0
change = 0

'loop to advance row by row
For i = 2 To l + 1
'Adds stock volume
volume = volume + Cells(i, 7)

'Determines openprice and closing price
    If (ticker = Cells(i, 1).Value) Then
        priceopen = priceopen
    Else
        priceclose = Cells(i - 1, 6).Value
        'Calculates yearly change and percentage change,
        change = priceclose - priceopen
        If (ticker <> "") Then
            If (priceopen <> 0) Then 'checking for divide by cero cases
                PercentChange = change / priceopen '
            Else
               Cells(i, 12).Value = 0
            End If
        Else
            'End of file write out results
            MsgBox ("end of data for year " & ActiveSheet.Name)
            Cells(1, 16).Value = ActiveSheet.Name
            Cells(2, 17).Value = MaxIncreaseTicker
            Cells(3, 17).Value = MaxDecreaseTicker
            Cells(4, 17).Value = GreatestVolumeTicker
            Cells(2, 18).Value = MaxIncrease
            Cells(3, 18).Value = MaxDecrease
            Cells(4, 18).Value = GreatestVolume
            Exit Sub
        End If
        
        'Reinitializes openprice
        priceopen = Cells(i, 3).Value
        Cells(i, 3).Interior.ColorIndex = 6

    End If

'Writes results to summary table, add colors, and checks for greatest and lowest changes
    If (ticker <> Cells(i, 1).Value) Then
        k = k + 1 '(summary table row control
        Cells(k - 1, 10).Value = ticker
        Cells(k - 1, 11).Value = change
        Cells(k - 1, 12).Value = FormatPercent(PercentChange, 2)
        Cells(k - 1, 13).Value = volume
        
        If (change < 0) Then
            Cells(k - 1, 11).Interior.ColorIndex = 3
        Else
            Cells(k - 1, 11).Interior.ColorIndex = 4
        End If
        
        'Checks for max and min
        If (PercentChange > MaxIncrease) Then
            MaxIncrease = PercentChange
            MaxIncreaseTicker = Cells(k - 1, 10).Value
        End If
        If (PercentChange < MaxDecrease) Then
            MaxDecrease = PercentChange
            MaxDecreaseTicker = Cells(k - 1, 10).Value
        End If
        If (volume > GreatestVolume) Then
            GreatestVolume = volume
            GreatestVolumeTicker = Cells(k - 1, 10).Value
        End If

        'Resets volume for next ticker
        volume = 0
    End If
    
'Resets stock ticker
ticker = Cells(i, 1)
Next i

'Writes out results to max and min %change and max volume per year
Cells(2, 17).Value = MaxIncreaseTicker
Cells(3, 17).Value = MaxDecreaseTicker
Cells(4, 17).Value = GreatestVolumeTicker
Cells(2, 18).Value = FormatPercent(MaxIncrease, 2)
Cells(3, 18).Value = FormatPercent(MaxDecrease, 2)
Cells(4, 18).Value = GreatestVolume

'Resets maxs and mins for next year
MaxIncrease = 0
MaxDecrease = 0
GreatestVolume = 0

Next year

End Sub



