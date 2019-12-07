VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Sub StockChange()
    

        Dim ws As Worksheet
        Dim Check As String
        Dim CheckTicker As String
        Dim CheckCellData As String
        Dim WorkSheetName As String
        
        Dim DailyOpen As Single
        Dim OpeningPrice As Single
        Dim ClosingPrice As Single
        Dim YearlyChange As Double
        Dim PercentChange As Variant
        
        Dim StockVolume As Variant
        Dim TotalVolume As Variant
        Dim MaxStockVolume As String
        Dim MaxStockVolumeLoc As String
        Dim MinPercent As String
        Dim MinPercentLoc As String
        Dim MaxPercent As String
        Dim MaxPercentLoc As String
        Dim LastRow As Long
        Dim LastRowL As Long
        Dim LastRowLetter As String
        Dim Counter As Long
        Dim StockCount As Long
        Dim wColNum As Single
        Dim wRowNum As Single
        
        Dim TickerValue As String
        Dim i As Long
        
    For Each ws In Worksheets
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
    

        WorkSheetName = ws.Name                                                 'Get worksheet name
        Debug.Print (WorkSheetName)

        
    
    
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row         'Get last row

        
        Counter = 0                                                                             'Counter for tracking number of rows for each ticker
        i = 0
        TotalVolume = 0
        StockCount = 1


    
        For i = 2 To LastRow                                                               
        'Iterate through rows
            Counter = Counter + 1
            CheckTicker = ws.Cells(i, 1).Value                                     
            'Get ticker value for stock
            DailyOpen = ws.Cells(i, 3).Value                                       
            'Get daily opening value for stock
            StockVolume = ws.Cells(i, 7).Value
            TotalVolume = TotalVolume + StockVolume

                
        If Counter = 1 Then
            OpeningPrice = DailyOpen

        End If
            
            
        If CheckTicker <> ws.Cells(i + 1, 1).Value Then
            StockCount = StockCount + 1
            ClosingPrice = ws.Cells(i, 6).Value
            Counter = 0
            
            YearlyChange = (ClosingPrice - OpeningPrice)
            'Evaluate for 0 value in data grrrr...
                If YearlyChange <> 0 And OpeningPrice <> 0 Then
                    PercentChange = (YearlyChange / OpeningPrice)
                Else
                    PercentChange = 0
                End If
            
            
            ws.Cells(StockCount, 9).Value = CheckTicker
            
            If YearlyChange > 0 Then
                ws.Cells(StockCount, 10).Interior.Color = vbGreen
            Else
                ws.Cells(StockCount, 10).Interior.Color = vbRed
            End If
                
            ws.Cells(StockCount, 10).NumberFormat = "0.00"
            ws.Cells(StockCount, 10).Value = YearlyChange
            ws.Cells(StockCount, 11).NumberFormat = "0.00%"
            ws.Cells(StockCount, 11).Value = PercentChange
            ws.Cells(StockCount, 12).Value = TotalVolume
            


        End If
        
        
            
        Next i
        
        'Fill cells with data determiniation values for maximum volume, highest and lowest percent change
       
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"

        MaxStockVolume = ws.Evaluate("CELL(""address"",OFFSET(L1,MATCH(MAX(L:L),L:L,0)-1,0))")
        MaxStockVolumeLoc = Replace(MaxStockVolume, "$", "")
        wColNum = Range(MaxStockVolumeLoc & 1).Column
        wRowNum = Range(MaxStockVolumeLoc).Row
        ws.Cells(4, 17).Value = ws.Range(MaxStockVolumeLoc).Value
        TickerValue = ws.Cells(wRowNum, wColNum - 3)
        ws.Cells(4, 16).Value = TickerValue
     
       
        MinPercent = ws.Evaluate("CELL(""address"",OFFSET(K1,MATCH(MIN(K:K),K:K,0)-1,0))")
        ws.Cells(3, 17).NumberFormat = "0.00%"
        ws.Cells(3, 17).Value = ws.Range(MinPercent).Value
        MinPercentLoc = Replace(MinPercent, "$", "")
        wColNum = Range(MinPercentLoc & 1).Column
        wRowNum = Range(MinPercentLoc).Row
        TickerValue = ws.Cells(wRowNum, wColNum - 2)
        ws.Cells(3, 16).Value = TickerValue
       

        MaxPercent = ws.Evaluate("CELL(""address"",OFFSET(K1,MATCH(MAX(K:K),K:K,0)-1,0))")
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(2, 17).Value = ws.Range(MaxPercent).Value
        MaxPercentLoc = Replace(MaxPercent, "$", "")
        wColNum = Range(MaxPercentLoc & 1).Column
        wRowNum = Range(MaxPercentLoc).Row
        TickerValue = ws.Cells(wRowNum, wColNum - 2)
        ws.Cells(2, 16).Value = TickerValue
        

        
    Next ws



End Sub



