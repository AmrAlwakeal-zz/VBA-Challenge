sub WallStreet():
    Dim ws As Worksheet
    'Source : https://excel-vba.programmingpedia.net/en/tutorial/1144/loop-through-all-sheets-in-active-workbook

    For Each WS In ActiveWorkbook.Worksheets
    WS.Activate
    'Loop Through WorkSheets
        '============================================================================================================
        For Each ws in Worksheets

        'Part I Define Variables and their Values 
        '============================================================================================================
        'A - Declare Variables
        '----------------------------------------------------------------
            Dim Ticker As String
            Dim OpenYear As Double
            Dim CloseYear As Double
            Dim YearlyChange As Double
            Dim StockTotal As Double
            Dim YearlyChange As Double
            Dim PercentChange As Double    
            Dim SummaryTableRow  As Integer
            Dim GreatTotal as Double
            Dim PercentIncrease as Long
            Dim PercentDecrease as Long
            Dim lastRow as Long
            lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

            Dim OpenDate As Boolean
            OpenDate = True
            'B- Variabnles initial values
            '----------------------------------------------------------------
            SummaryTableRow = 2
            GreatTotal = 0   
            PercentIncrease = 0                       
            PercentDecrease = 0
            'C-Result Tables Headers & 
            '----------------------------------------------------------------
            ws.Range("I1").Value = "Ticker"
            ws.Range("J1").Value = "Total Stock Volume"
            ws.Range("K1").Value = "Yearly Change"
            ws.Range("L1").Value = "Percent Change"
            ws.Range("O2").Value = "Greatest % Increase"
            ws.Range("O3").Value = "Greatest % Decrease"
            ws.Range("O4").Value = "Greatest Total Volume"
            ws.Range("P1").Value = "Ticker"
            ws.Range("Q1").Value = "Value"    
        
            '============================================================================================================
            'Part II : For loop
            '============================================================================================================
            ' 1- Ticker, StockTotal, OpenYear, CloseYear
            '------------------------------------------------------------------------------------------------
            For i = 2 to lastRow 
         
                If  (ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value) Then 
                    Ticker = ws.Cells(i, 1).Value
                    StockTotal = StockTotal + ws.Cells(i, 7).Value
                    CloseYear = ws.Cells(i, 6).Value
                    ws.Cells(SummaryTableRow, 9).Value = Ticker
                    ws.Cells(SummaryTableRow, 10).Value = StockTotal
                    CloseYear = Cells(i, 6).Value
                    YearlyChange = CloseYear - OpenYear
                    ws.Cells(SummaryTableRow, 11).Value = YearlyChange
                        
                    ' Get the Percent Change 
                    PercentChange = (CloseYear / OpenYear) * 100
                    ws.Cells(SummaryTableRow, 12).Value = PercentChange
                    ws.Cells((SummaryTableRow, 12).NumberFormat = "0.00%"
                    ' Add colors 
                    'Source: https://stackoverflow.com/questions/50588153/value-error-setting-interior-colorindex-property-in-excel-2013

                    If (ws.Cells(SummaryTableRow, 12).Value > 0) Then
                        ws.Cells(SummaryTableRow, 12).Interior.ColorIndex = 4
                    Else
                        ws.Cells(SummaryTableRow, 12).Interior.ColorIndex = 3
                    End if
                    StockTotal = 0
                    SummaryTableRow = SummaryTableRow + 1
                    StockTotal = StockTotal + Cells(i, 7).Value
                    OpenDate = True    

                Else
                    ' Add to the total stock volume
                    StockTotal = StockTotal + Cells(i, 7).Value
                    ' Get the price the stock opened the year
                    If Open_Year_Date And Cells(i, 3).Value <> 0 Then
                        OpenYear = Cells(i, 3).Value
                        OpenDate = False
                
                    End If
                End If
            Next i
            '------------------------------------------------------------------------------------------------
            '2- Greatest Total & add to Summary 
            '------------------------------------------------------------------------------------------------
            
            For i = 2 to SummaryTableRow
                If (ws.Cells(i, 11).Value > GreatTotal) Then
                    GreatTotal = ws.Cells(i, 11).Value
                    ws.Cells(2, 16).Value = ws.Cells(i, 10).Value
                End if
            Next i
            '------------------------------------------------------------------------------------------------
            '3- Greastest Percent Decrease & Increase & add them to Summary
            '------------------------------------------------------------------------------------------------
            For i = 2 to SummaryTableRow
                If (ws.Cells(i, 11).Value > PercentIncrease) Then 
                    PercentIncrease = ws.Cells(i, 11).Value                
                    ws.Cells(3, 16) = ws.Cells(i, 10).Value
                Elseif (ws.Cells(i, 13).Value < PercentDecrease) Then
                    PercentDecrease = ws.Cells(i, 11).Value
                    ws.Cells(4, 16).Value = ws.Cells(i, 10).Value
                End If
            Next i
            '------------------------------------------------------------------------------------------------
            'Source: https://stackoverflow.com/questions/44409090/percent-style-formatting-in-excel-vba
            'set cell format to percent
            ws.Cells(3, 17).Style = "percent"
            ws.Cells(4, 17).Style = "percent"

           'auto fit table columns
            ws.Columns("J:Q").AutoFit
    Next ws
End sub
        