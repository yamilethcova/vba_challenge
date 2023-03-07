# vba_challenge
Sub stockmarket():

       
       Dim Ticker As String
        Ticker = " "
        Dim Ticker_Volume As Double
        Ticker_Volume = 0
        
        Dim inputRow As Long
        Dim LastRow As Long
        Dim OutputRow As Long
        Dim open_price As Double
        open_price = 0
        Dim close_price As Double
        close_price = 0
        Dim price_change As Double
        price_change = 0
        Dim price_change_percent As Double
        price_change_percent = 0
        Dim Vol As Long
        Vol = 0
        Dim Max_Ticker_Name As String
        Max_Ticker_Name = " "
        Dim Min_Ticker_Name As String
        Min_Ticker = " "
        Dim Min_Percent As Double
        Min_Percent = 0
        Dim Max_Volume_Ticker_Name As String
        Max_Volume_Ticker_Name = " "
        Dim Max_Volume As Double
        Max_Volume = 0

       
        
 For Each ws In Worksheets
      
      OutputRow = 2
      ws.Cells(1, 9).Value = "Ticker"
      ws.Cells(1, 10).Value = "Yearly Change"
      ws.Cells(1, 11).Value = "Percent Change"
      ws.Cells(1, 12).Value = "Stock Volume"
      ws.Cells(1, 16).Value = "Ticker"
      ws.Cells(1, 17).Value = "Value"
      
      ws.Cells(2, 15).Value = "Greatest % Increase"
      ws.Cells(3, 15).Value = "Greatest % Decrease"
      ws.Cells(4, 15).Value = "Greatest Total Volume"
      
  
  
      LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  
      MsgBox LastRow
      
      Dim name As String
      Dim SumaryRoad As Integer
      sumaryRow = 2
      open_price = ws.Cells(2, 3).Value
        
        Yearly_Price_Change_percent = 0
        Total_Ticker_Volume = 0

For inputRow = 2 To LastRow

        If (ws.Cells(inputRow + 1, 1).Value <> ws.Cells(inputRow, 1).Value) Then
            name = ws.Cells(inputRow, 1).Value
            ws.Cells(sumaryRow, 9).Value = name
            ws.Cells(inputRow, 7).Value = Vol
            ws.Cells(sumaryRow, 12).Value = Vol
            close_price = ws.Cells(inputRow, 6).Value
            price_change = close_price - open_price
            price_percent_change = (close_price - open_price) / open_price
            
            ws.Cells(sumaryRow, 10).Value = price_change
        If price_change > 0 Then
                ws.Cells(sumaryRow, 10).Interior.ColorIndex = 4
            ElseIf price_change < 0 Then
                ws.Cells(sumaryRow, 10).Interior.ColorIndex = 3
            End If
            
            open_price = ws.Cells(inputRow + 1, 3).Value
            If open_price <> 0 Then
                price_change_percent = (close_price - open_price) * 100
                ws.Cells(sumaryRow, 11).Value = price_change_percent
                sumaryRow = sumaryRow + 1
            End If
        End If


Next inputRow
    
        
'    ws.Range("Q2") = WorksheetFunction.Max(ws.Range("K2" & LastRow))
'    ws.Range("P2") = WorksheetFunction.Max(ws.Range("J2" & LastRow))
'
'    ws.Range("Q3") = WorksheetFunction.Min(ws.Range("K3" & LastRow))
'    ws.Range("P3") = WorksheetFunction.Min(ws.Range("J3" & LastRow))
''
'    Total_Ticker_Volume = Total_Ticker_Volume + Cells(inputRow, 7).Value

    
    MsgBox ("done")

Next ws

End Sub

