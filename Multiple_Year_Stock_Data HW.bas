Attribute VB_Name = "Module1"
Sub Multiple_Year_Stock_Data()
    
    'Declare "ws" as Worksheet
    Dim ws As Worksheet
    
    'Loop through each worksheet
    For Each ws In Worksheets
    
    'Label column headers of the tables
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    
    'Declare variables
    Dim Ticker As String
    Dim LastRowA As Long
    Dim LastRowK As Long
    Dim TotalStockVolume As Double
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim YearlyChange As Double
    Dim PreviousAmount As Long
    Dim PercentChange As Double
    Dim GreatestIncrease As Double
    Dim GreatestDecrease As Double
    Dim LastRowValue As Long
    Dim GreatestTotalVolume As Long
    
    TotalStockVolume = 0
    SummaryTableRow = 2
    PreviousAmount = 2
    GreatestIncrease = 0
    GreatestDecrease = 0
    GreatestTotalVolume = 0
    
    'Determine value of the last row in column A
    LastRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
    For i = 2 To LastRowA

        TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
    
        'Check if the next row has the same ticker name as the previous row
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            Ticker = ws.Cells(i, 1).Value
            ws.Range("I" & SummaryTableRow).Value = Ticker
            ws.Range("L" & SummaryTableRow).Value = TotalStockVolume
        
            TotalStockVolume = 0

            OpenPrice = ws.Range("C" & PreviousAmount)
     
            ClosePrice = ws.Range("F" & i)
          
            YearlyChange = ClosePrice - OpenPrice
            ws.Range("J" & SummaryTableRow).Value = YearlyChange
                
            'Change format of Column J with "$"
            ws.Range("J" & SummaryTableRow).NumberFormat = "$0.00"

            'Determine Percent Change, if Yearly Open Price is 0, then Percent Change is 0
            If OpenPrice = 0 Then
                PercentChange = 0
                    
                'Otherwise, set Percent Change to Yearly Change divided by Yearly Open Price
                Else
                YearlyOpen = ws.Range("C" & PreviousAmount)
                PercentChange = YearlyChange / OpenPrice
                        
            End If
                
            ws.Range("K" & SummaryTableRow).Value = PercentChange
                
            ws.Range("K" & SummaryTableRow).NumberFormat = "0.00%"

            'If Yearly change is Positive, highlight cell in Green
            If ws.Range("J" & SummaryTableRow).Value >= 0 Then
            ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 4
                    
                Else
                'If Yearly change is Negative, highlight cell in Red
                ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 3
                
            End If
            
            SummaryTableRow = SummaryTableRow + 1
            PreviousAmount = i + 1
                
        End If
                
        Next i

        'Determine value of the last row in column K
        LastRowK = ws.Cells(Rows.Count, 11).End(xlUp).Row
        
        For i = 2 To LastRowK
            
            'Determine Greatest % Increase
            If ws.Range("K" & i).Value > ws.Range("Q2").Value Then
                ws.Range("Q2").Value = ws.Range("K" & i).Value
                ws.Range("P2").Value = ws.Range("I" & i).Value
                
            End If

            'Determine Greatest % Decrease
            If ws.Range("K" & i).Value < ws.Range("Q3").Value Then
                ws.Range("Q3").Value = ws.Range("K" & i).Value
                ws.Range("P3").Value = ws.Range("I" & i).Value
                    
            End If

            'Determine Greatest Total Volume
            If ws.Range("L" & i).Value > ws.Range("Q4").Value Then
                ws.Range("Q4").Value = ws.Range("L" & i).Value
                ws.Range("P4").Value = ws.Range("I" & i).Value
                    
            End If

            Next i
            
        'Change format of Q2 and Q3 to %
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").NumberFormat = "0.00%"

    Next ws

End Sub
