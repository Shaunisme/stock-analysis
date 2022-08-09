Sub runEverySheet()          'Please run this procedure at all
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call stockStat
    Next
    Application.ScreenUpdating = True
End Sub

Sub stockStat()
    Dim i As Integer         'Row counter
    Dim j As Integer         'Ticker counter
    Dim volume As Double     'Total stock volume
    Dim ticker As String
    Dim openingPrice As Double
    
'Bonus begin
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestTotalVolumn As Double
    Dim gTicker(3) As String
'end
    
    Cells(1, 9).Value = "Ticker"
    Columns(9).ColumnWidth = 6
    Cells(1, 10).Value = "Yearly Change"
    Columns(10).ColumnWidth = 14
    Cells(1, 11).Value = "Percent Change"
    Columns(11).ColumnWidth = 15
    Cells(1, 12).Value = "Total Stock Volume"
    Columns(12).ColumnWidth = 18
    
    i = 2
    j = 2
    volume = 0
    ticker = Cells(i, 1).Value
    openingPrice = Cells(i, 3).Value
    
'Bonus begin
    greatestIncrease = -100000
    greatestDecrease = 100000
    greatestTotalVolumn = 0
'end

    Do While True
        If IsEmpty(Cells(i, 1).Value) Or Cells(i, 1).Value <> ticker Then   'The sheet end OR next ticker
    
            Cells(j, 9).Value = ticker
            Cells(j, 10).Value = Cells(i - 1, 6).Value - openingPrice

            Cells(j, 11).Value = Cells(j, 10) / openingPrice
            Cells(j, 12).Value = volume
            
            If Cells(j, 10).Value > 0 Then                              'Color
                Cells(j, 10).Interior.Color = vbGreen
            ElseIf Cells(j, 10).Value < 0 Then
                Cells(j, 10).Interior.Color = vbRed
            End If
            Cells(j, 11).NumberFormatLocal = "0.00%"                    'Percent
                        
'Bonus begin
            If Cells(j, 11).Value > greatestIncrease Then
                greatestIncrease = Cells(j, 11).Value
                gTicker(0) = ticker
            End If
            If Cells(j, 11).Value < greatestDecrease Then
                greatestDecrease = Cells(j, 11).Value
                gTicker(1) = ticker
            End If
            If Cells(j, 12).Value > greatestTotalVolumn Then
                greatestTotalVolumn = Cells(j, 12).Value
                gTicker(2) = ticker
            End If
'End
            
            If IsEmpty(Cells(i, 1).Value) Then                           'The sheet end will exit
                Exit Do
            End If
            
            j = j + 1
            volume = 0
            ticker = Cells(i, 1).Value
            openingPrice = Cells(i, 3).Value
        End If
    
        volume = volume + Cells(i, 7).Value
        i = i + 1
    Loop
    
'Bonus begin
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Columns(15).ColumnWidth = 22
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volumn"
    
    Columns(16).ColumnWidth = 6
    Cells(2, 16).Value = gTicker(0)
    Cells(3, 16).Value = gTicker(1)
    Cells(4, 16).Value = gTicker(2)
    
    Columns(17).ColumnWidth = 16
    Cells(2, 17).Value = greatestIncrease
    Cells(3, 17).Value = greatestDecrease
    Cells(4, 17).Value = greatestTotalVolumn
    
    Cells(3, 17).NumberFormatLocal = "0.00%"
    Cells(2, 17).NumberFormatLocal = "0.00%"
'End

End Sub


