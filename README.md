# VBA
First VBA Project

Sub stock_data()

Dim ticker_name As String

Dim yearly_change As Double
yearly_change = 0

Dim percent_change As Double
percent_change = 0

Dim ticker_row As Integer
ticker_row = 2

Dim total_stock As Double

Range("I1") = "Ticker"
Range("J1") = "Yearly Change"
Range("K1") = "Percent Change"
Range("L1") = "Total Stock Volume"

lastrow = Cells(Rows.Count, 1).End(xlUp).Row


    For i = 2 To lastrow


        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            ticker_name = Cells(i, 1).Value

            yearly_change = yearly_change + (Cells(i, 6).Value - Cells(i, 3).Value)

            If Cells(i, 3).Value = 0 Then
            
                percent_change = 0
                
            Else
            
                percent_change = percent_change + (Cells(i, 6).Value - Cells(i, 3).Value / Cells(i, 3).Value)

            End If

            total_stock = total_stock + Cells(i, 7).Value
            
        
            Range("I" & ticker_row).Value = ticker_name

            Range("J" & ticker_row).Value = yearly_change
            
            Range("J" & ticker_row).NumberFormat = "0.00"

            Range("K" & ticker_row).Value = (percent_change / 100000)
        
            Range("K" & ticker_row).NumberFormat = "0.00%"
        
            Range("L" & ticker_row).Value = total_stock
            
        
            ticker_row = ticker_row + 1
        
            yearly_change = 0
        
            percent_change = 0
        
            total_stock = 0
        
        Else

            yearly_change = yearly_change + (Cells(i, 6).Value - Cells(i, 3).Value)
        
                 If Cells(i, 3).Value = 0 Then
            
                percent_change = 0
                
                Else
            
                percent_change = percent_change + (Cells(i, 6).Value - Cells(i, 3).Value / Cells(i, 3).Value)

                End If
            
        
            total_stock = total_stock + Cells(i, 7).Value
        
        End If
        
        
            If Range("J" & ticker_row).Value >= 0 Then
        
                Range("J" & ticker_row).Interior.ColorIndex = 10
            
             Else
        
                Range("J" & ticker_row).Interior.ColorIndex = 9
        
            End If
        
            
    
    Next i
        
    
End Sub


