Sub WallStreet():

    Dim Ticker As String
    Dim OpenCloseDelta As Double
    Dim TotalVolume As Double
    Dim OpPrice As Double
    Dim ClPrice As Double
    Dim PrcntCh As Variant
    Dim GreatestInc As Variant
    Dim GreatestDec As Variant
    Dim GreatestTot As Variant
    Dim Summary_Table_Row As Integer
    Dim WS As Worksheet
    

    'Loop through all sheets
    Set WS = ActiveSheet
    For Each WS In ThisWorkbook.Worksheets
        WS.Activate

        'Initial value of summaray table row
        Summary_Table_Row = 2

        'Initial value of total volume
        TotalVolume = 0

        'Initial value of open price
        OpPrice = 0

        'Initial value of close price
        ClPrice = 0

        'Add headers summary tables
        Range("I1").Value = "Ticker"
        Range("P1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        Range("Q1").Value = "Value"
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"

        lastrow = WS.Cells(Rows.Count, "A").End(xlUp).Row
        'MsgBox(lastrow)

        'Loop thorugh ticker colum
        For i = 2 To (lastrow)

            'Set inital open price
            If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then

                'Set open price
                OpPrice = Format(Cells(i, 3).Value, "00.00")
                'MsgBox(OpPrice)

                'Start total volume value
                TotalVolume = TotalVolume + Cells(i, 7).Value

            'Check if still the same ticker ticker, if not
            ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

                'Capture the Ticker
                Ticker = Cells(i, 1).Value

                'Insert Ticker into summary table
                Range("I" & Summary_Table_Row).Value = Ticker

                'Capture end year close price
                ClPrice = Format(Cells(i, 6).Value, "00.00")
                'MsgBox(ClPrice)

                'Calculate the year difference in price
                OpenCloseDelta = ClPrice - OpPrice
                'MsgBox(OpenCloseDelta)

                'Set cell color to match stock difference
                If OpenCloseDelta >= 0 Then

                    'Set cell to green if positive
                    Range("J" & Summary_Table_Row).Interior.ColorIndex = 4

                Else
                
                    'Set cell to red if negative
                    Range("J" & Summary_Table_Row).Interior.ColorIndex = 3

                End If

                'Insert difference into summary table
                Range("J" & Summary_Table_Row).Value = OpenCloseDelta

                'Determine Percentage and account for divisble by 0
                If OpPrice = 0 Then
                    PrcntCh = 0
                    
                    Else
                    PrcntCh = FormatPercent(OpenCloseDelta / OpPrice, 2)
                    
                End If

                'Insert percentage value into summary Table
                Range("K" & Summary_Table_Row).Value = PrcntCh

                'Add last ticker volume to total volume
                TotalVolume = TotalVolume + Cells(i, 7).Value

                'Insert total volume to summary table
                Range("L" & Summary_Table_Row).Value = TotalVolume

                'Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1

                'Reset total volume to zero before next ticker or end
                TotalVolume = 0

                'Reset open price of ticker
                OpPrice = 0

                'Reset close price of ticker
                ClPrice = 0
                
                'Reset percent change of ticker
                PrcntCh = 0

            Else

                'Add last ticker volume to total volume
                TotalVolume = TotalVolume + Cells(i, 7).Value


            
            End If

        Next i

        'Set initial value of greatest percentage increase
        GreatestInc = 0

        'Set initial value of greatest percentage decrease
        GreatestDec = 0

        'Set initial value of greatest total volume
        GreatestTot = 0

        'Find last row of summary table
        lastrow_sumtbl = WS.Cells(Rows.Count, "I").End(xlUp).Row
        'MsgBox(lastrow_Sumbtbl)

        For j = 2 to (lastrow_sumtbl)

            'Search greatest increase by looking for largest positive percentage
            If Cells(j, 11).Value > 0 And Cells(j, 11).Value > Cells(2, 17).Value Then

                'Capture percentage value
                GreatestInc = FormatPercent(Cells(j, 11).Value)
                'MsgBox(GreatestInc)
                
                'Store Ticker of largest positive percentage
                Ticker = Cells(j, 9).Value

                'Insert ticker into second summary table
                Range("P2").Value = Ticker

                'Insert percentage value into second summary table
                Range("Q2").Value = GreatestInc

            End If

            'Search for greatest decrease by looking for largest negative percentage
            If Cells(j, 11).Value < 0 And Cells(j, 11).Value < Cells(3, 17).Value Then

                'Capture percentage value
                GreatestDec = FormatPercent(Cells(j, 11).Value)
                
                'Store Ticker of largest negative percentage
                Ticker = Cells(j, 9).Value

                'Insert ticker into second summary table
                Range("P3").Value = Ticker

                'Insert percentage value into second summary table
                Range("Q3").Value = GreatestDec

            End If

            'Search for greatest total volume
            If Cells(j, 12).Value > Cells(4, 17).Value Then

                'Capture percentage value
                GreatestTot = Cells(j, 12).Value
                
                'Store Ticker of largest positive percentage
                Ticker = Cells(j, 9).Value

                'Insert ticker into second summary table
                Range("P4").Value = Ticker

                'Insert percentage value into second summary table
                Range("Q4").Value = GreatestTot

            End If

        Next j

        'Auto adjust summary table colums to content size
        Columns(10).AutoFit

        Columns(11).AutoFit

        Columns(12).AutoFit

        Columns(15).AutoFit

        Columns(16).AutoFit

        Columns(17).AutoFit
    
    Next

End Sub





