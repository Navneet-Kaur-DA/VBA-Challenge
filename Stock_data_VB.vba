Sub Stock_Data()
    ' ws for all te worksheets iin file
    For Each ws In Worksheets
    
        Dim j As Integer
        Dim Ticker_name As String ' Print Tickers' name
        Dim Total_volume As Double ' calculate total Volume
        Dim open_value As Variant  'variant to save open value
        Dim Close_value As Variant  ' variant to save close value
        Dim cond As Boolean '  boolean to keep open value unchanged in loop
        Dim yearly_Change As Variant   'Variant to save
        Dim Great_inc As Double   ' to Calculate greatest increment percentage
        Dim Great_dec As Integer   ' to calculate greatest decerment percentage
        Dim Great_total As Double   ' to calculate greatest total volume
        Dim grect_inc_ticker As String  ' to print greatest ticker name
        Dim per_change As Variant   ' percenatge change vaiable
        
        
        
        
        Total_volume = 0
        j = 2
        
        'calculate last row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        cond = False
       
       ' set some header names
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Cells(4, 15).Value = "Greatest % Increase"
        ws.Cells(5, 15).Value = "Greatest % Decrease"
        ws.Cells(6, 15).Value = "Greatest Total Volume"
        ws.Cells(3, 16).Value = "Ticker"
        ws.Cells(3, 17).Value = "Value"
        
        Great_inc = 0
        Great_dec = 0
        Great_total = 0
        
        
        
        'loop starts to go through all the rows
        For i = 2 To LastRow
        
        
            'checking condition value and setting open_value variable so that it cant be changed til the same ticker name is apperaed in  rows
            If cond = False Then
                open_value = ws.Cells(i, 3).Value
                cond = True
            End If
            
            
            'setting condition: if the first cell value in first row is not same as first cell value in next row
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                ' tickers name
                ws.Cells(j, 9).Value = ws.Cells(i, 1).Value
                
                ' calculating close value
                Close_value = ws.Cells(i, 6).Value
                
                'calculating yeraly change
                yearly_Change = Close_value - open_value
                
                ws.Cells(j, 10).Value = yearly_Change
                
                    'changing color if the yearly_change column based on the value
                    If yearly_Change > 0 Then
                        ws.Cells(j, 10).Interior.ColorIndex = 4
                    Else
                        ws.Cells(j, 10).Interior.ColorIndex = 3
                    End If
                    
                ' calculting percentage change
                per_change = (yearly_Change / open_value) * 100
                
                
                ws.Cells(j, 11).Value = per_change
                
                ' changing the color of the percentage column based on the value
                If per_change > 0 Then
                        ws.Cells(j, 11).Interior.ColorIndex = 7
                    Else
                        ws.Cells(j, 11).Interior.ColorIndex = 6
                    End If
                
                ' calculating total stock value
                Total_volume = Total_volume + ws.Cells(i, 7).Value
                ws.Cells(j, 12).Value = Total_volume
                
                ' calculating greatest percentage increment
                If (per_change > Great_inc) Then
                    Great_inc = per_change
                    ws.Cells(4, 16).Value = ws.Cells(j, 9).Value
                    ws.Cells(4, 17).Value = ws.Cells(j, 11).Value
                End If
                
                ' calculating greatest percentage decrement
                If (per_change < Great_dec) Then
                    Great_dec = per_change
                    ws.Cells(5, 16).Value = ws.Cells(j, 9).Value
                    ws.Cells(5, 17).Value = ws.Cells(j, 11).Value
                End If
                
                ' calculating greatest stock volume
                 If (Total_volume > Great_total) Then
                    Great_total = Total_volume
                    ws.Cells(6, 16).Value = ws.Cells(j, 9).Value
                    ws.Cells(6, 17).Value = ws.Cells(j, 12).Value
                End If
                
                j = j + 1
                Total_volume = 0
                cond = False
                
            Else
            ' else part: if the first cell value of row is same as first cell value of next row than caluclating the volume total
                Total_volume = Total_volume + ws.Cells(i, 7).Value
               
                
            End If
        Next i
        
   Next ws
End Sub
