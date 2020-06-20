Sub WallStreet():

    Dim Ticker As String
    Dim OpenCloseDelta As Double
    Dim TotalVolume As Double
    Dim OpPrice As Double
    Dim ClPrice As Double
    Dim PrcntCh As Variant
    Dim Summary_Table_Row As Integer
    
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
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"

    lastrow = Cells(Rows.Count, "A").End(xlUp).Row
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

    'Auto adjust summary table colums to content size
    Columns(10).AutoFit

    Columns(11).AutoFit

    Columns(12).AutoFit

End Sub





