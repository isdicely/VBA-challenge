Sub Stocks():
Dim ws As Worksheet
Dim I As Integer
Dim ticker As String
Dim num_rows As Long
Dim ticker_index As Long
Dim x As Long
Dim opening As Double
Dim closing As Double
Dim total_volume As Double
Dim Gr_Pct_Increase As Double
Dim Gr_Pct_Decrease As Double
Dim Gr_Ttl_Volume As Double
Dim Gr_Pct_Inc_Ticker As String
Dim Gr_Pct_Dec_Ticker As String
Dim Gr_Ttl_Vol_Ticker As String

ticker = ""
ticker_index = 0
total_volume = 0
Gr_Pct_Increase = 0
Gr_Pct_Decrease = 0
Gr_Ttl_Volume = 0
Gr_Pct_Inc_Ticker = ""
Gr_Pct_Dec_Ticker = ""
Gr_Ttl_Vol_Ticker = ""


For Each ws In Worksheets
    
    num_rows = ws.Cells(Rows.Count, 1).End(xlUp).Row
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly_Change"
    ws.Cells(1, 11).Value = "Percent_Change"
    ws.Cells(1, 12).Value = "Total_Stock_Volume"
    
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value:"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        
    For x = 2 To num_rows
           
       If ws.Cells(x, 1).Value <> ticker Then
        
            If ticker <> "" Then
                ws.Cells(1 + ticker_index, 10).Value = closing - opening
                ws.Cells(1 + ticker_index, 11).NumberFormat = "0.00%"
                                                
                ws.Cells(1 + ticker_index, 12).Value = total_volume
                
                If opening <> 0 Then
                    ws.Cells(1 + ticker_index, 11).Value = ((closing - opening) / opening)
                    
                    If ws.Cells(1 + ticker_index, 11).Value >= 0 Then
                        ws.Cells(1 + ticker_index, 11).Interior.ColorIndex = 4
                    ElseIf ws.Cells(1 + ticker_index, 11).Value < 0 Then
                        ws.Cells(1 + ticker_index, 11).Interior.ColorIndex = 3
                    End If
            
                End If
            End If
            
            ticker = ws.Cells(x, 1).Value
            ws.Cells(2 + ticker_index, 9).Value = ticker
            ticker_index = ticker_index + 1
            opening = ws.Cells(x, 3).Value
            total_volume = 0
            
        End If
        
        closing = ws.Cells(x, 6).Value
        total_volume = total_volume + ws.Cells(x, 7).Value
        
    Next x
        
    ws.Cells(1 + ticker_index, 10).Value = closing - opening
    ws.Cells(1 + ticker_index, 11).NumberFormat = "0.00%"
        
    If opening <> 0 Then
        ws.Cells(1 + ticker_index, 11).Value = ((closing - opening) / opening)
            
            If ws.Cells(1 + ticker_index, 11).Value >= 0 Then
                ws.Cells(1 + ticker_index, 11).Interior.ColorIndex = 4
            ElseIf ws.Cells(1 + ticker_index, 11).Value < 0 Then
                ws.Cells(1 + ticker_index, 11).Interior.ColorIndex = 3
            End If
    
    End If
    
    ws.Cells(1 + ticker_index, 12).Value = total_volume

    ticker_index = 0
    ticker = ""
    Gr_Pct_Increase = 0
    Gr_Pct_Decrease = 0
    Gr_Ttl_Volume = 0
    Gr_Pct_Inc_Ticker = ""
    Gr_Pct_Dec_Ticker = ""
    Gr_Ttl_Vol_Ticker = ""

    num_stocks = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    For I = 2 To num_stocks
        If ws.Cells(I, 11).Value > Gr_Pct_Increase Then
            Gr_Pct_Increase = ws.Cells(I, 11).Value
            Gr_Pct_Inc_Ticker = ws.Cells(I, 9).Value
        End If
        
        If ws.Cells(I, 11).Value < Gr_Pct_Decrease Then
            Gr_Pct_Decrease = ws.Cells(I, 11).Value
            Gr_Pct_Dec_Ticker = ws.Cells(I, 9).Value
        End If
                
        If ws.Cells(I, 12).Value > Gr_Ttl_Volume Then
            Gr_Ttl_Volume = ws.Cells(I, 12).Value
            Gr_Ttl_Vol_Ticker = ws.Cells(I, 9)
        End If
    
    Next I

    ws.Cells(2, 17).Value = Gr_Pct_Increase
    ws.Cells(2, 17).NumberFormat = "0.00%"
    ws.Cells(2, 16).Value = Gr_Pct_Inc_Ticker
    ws.Cells(3, 17).Value = Gr_Pct_Decrease
    ws.Cells(3, 17).NumberFormat = "0.00%"
    ws.Cells(3, 16).Value = Gr_Pct_Dec_Ticker
    ws.Cells(4, 17).Value = Gr_Ttl_Volume
    ws.Cells(4, 16).Value = Gr_Ttl_Vol_Ticker
    
    ws.Columns("A:Q").AutoFit
    
Next ws

End Sub
