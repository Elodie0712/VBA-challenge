VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Sub multipleyears()
    Dim ws As Worksheet
    Dim Lastrow As Long
    Dim tickerName As String
    Dim stockVolume As Double
    Dim tickerRow As Integer
    Dim openPrice As Double
    Dim closePrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    
    'label variables for the last part
    
    Dim GI As Double
    Dim GD As Double
    Dim GV As Double
    Dim MIT As String
    Dim MDT As String
    Dim MVT As String
    

    For Each ws In Worksheets ' Apply to all worksheets
    
        ' Set the variables
        stockVolume = 0
        openPrice = ws.Cells(2, 3).Value ' 1st open price

        ' Name the new rows
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

        ' Number of rows
        Lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

        ' Ticker count
        tickerRow = 2 'start at row 2
        
        For i = 2 To Lastrow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                tickerName = ws.Cells(i, 1).Value
                closePrice = ws.Cells(i, 6).Value

                'yearly change and percent change
                
                yearlyChange = closePrice - openPrice
                percentChange = yearlyChange / openPrice * 100

                ' Update values
                
                ws.Cells(tickerRow, 9).Value = tickerName
                ws.Cells(tickerRow, 10).Value = yearlyChange
                ws.Cells(tickerRow, 11).Value = percentChange & "%"
                ws.Cells(tickerRow, 12).Value = stockVolume

                ' Move to the next row and reset stock volume for the next ticker
                
                tickerRow = tickerRow + 1
                stockVolume = 0
                
                ' Update open price for the next ticker
                openPrice = ws.Cells(i + 1, 3).Value
            Else
                ' Add stock volume for the each diff ticker
                
                stockVolume = stockVolume + ws.Cells(i, 7).Value
            End If
            
        Next i
        
        Lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To Lastrow
        
        If ws.Cells(i, 10).Value > 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 43
        Else
        
        ws.Cells(i, 10).Interior.ColorIndex = 3
        End If
        Next i
       ' Find Greatest % increase, decrease, and total volume
        Lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

       GI = ws.Cells(2, 11).Value
        GD = ws.Cells(2, 11).Value
        GV = ws.Cells(2, 12).Value

        For i = 2 To Lastrow
            If ws.Cells(i, 11).Value > GI Then
                GI = ws.Cells(i, 11).Value
                MIT = ws.Cells(i, 9).Value
            End If

            If ws.Cells(i, 11).Value < GD Then
               GD = ws.Cells(i, 11).Value
                MDT = ws.Cells(i, 9).Value
            End If

            If ws.Cells(i, 12).Value > GV Then
                GV = ws.Cells(i, 12).Value
                MVT = ws.Cells(i, 9).Value
            End If
        Next i

        ' Display the results
        ws.Cells(2, 16).Value = MIT
        ws.Cells(2, 17).Value = GI * 100 & "%"

        ws.Cells(3, 16).Value = MDT
        ws.Cells(3, 17).Value = GD * 100 & "%"

        ws.Cells(4, 16).Value = MVT
        ws.Cells(4, 17).Value = GV
        
       ws.Cells.EntireColumn.AutoFit

    Next ws
End Sub

