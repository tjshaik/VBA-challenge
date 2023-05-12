Attribute VB_Name = "Module1"

Sub HW()
For Each ws In Worksheets
Dim ticker As String
Dim closing_price As Double
Dim opening_price As Double
Dim yearly_change As Double
Dim total_volume As Double
Dim percent_change As Double
Dim last_row As Long
Dim i As Long
Dim Greatest_increase As Double
Dim Greatest_decrease As Double
Dim Greatest_total As Double


'headers on the same data sheet
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest total volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"

RowCounter = 2


'opening price location
opening_price = ws.Cells(2, 3).Value

'closing price location
'closing_price = ws.Cells(2, 6).Value



last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row

'going through each row
' row selection for
For i = 2 To last_row

'organizing by ticker
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    'closing value
    closing_price = ws.Cells(i, 6).Value
    
    'yearly change value
    yearly_change = (closing_price - opening_price)
    ws.Range("J" & RowCounter).Value = yearly_change
    
    
    ticker = ws.Cells(i, 1).Value
    ws.Range("I" & RowCounter).Value = ticker
    'getting last ticker volume

    total_volume = total_volume + ws.Cells(i, 7).Value
    ws.Range("L" & RowCounter).Value = total_volume
    
 
    'percent change
    percent_change = yearly_change / opening_price
    
    ws.Range("K" & RowCounter).Value = percent_change
    
    ' format percent change into percent
    ws.Range("K" & RowCounter).NumberFormat = "0.00%"
    
    'setting next row counter
    RowCounter = RowCounter + 1
    
    'reseting total value
    total_volume = 0
    ticker = ws.Cells(i + 1, 1).Value
    
    
    opening_price = ws.Cells(i + 1, 3).Value
   
    
Else


' getting total ticker volume
total_volume = total_volume + ws.Cells(i, 7).Value

    
    End If
    
Next i

  ' Conditional formating for yearly change
        
        For i = 2 To last_row
            If ws.Cells(i, 10).Value > 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 10
            Else
            
            ws.Cells(i, 10).Interior.ColorIndex = 3
            End If
        Next i
        
    ' Adding fuctionality to the script
    Greatest_increase = 0
    For j = 2 To last_row
    If ws.Cells(j, 11).Value > Greatest_increase Then
    Greatest_increase = ws.Cells(j, 11).Value
    End If
    Next j
    
    ws.Range("Q2").Value = Greatest_increase
    
    increase_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & last_row)), ws.Range("K2:K" & last_row), 0)
    ws.Range("P2").Value = ws.Cells(increase_number + 1, 9)

' greatest decrease
Greatest_decease = 1000
    For j = 2 To last_row
    If ws.Cells(j, 11).Value < Greatest_decrease Then
    Greatest_decrease = ws.Cells(j, 11).Value
    End If
    Next j
    
    ws.Range("Q3").Value = Greatest_decrease
    
    increase_number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & last_row)), ws.Range("K2:K" & last_row), 0)
    ws.Range("P3").Value = ws.Cells(increase_number + 1, 9)
    
    ' total vol greateest
    Greatest_volume = 0
    For j = 2 To last_row
    If ws.Cells(j, 12).Value > Greatest_volume Then
    Greatest_volume = ws.Cells(j, 12).Value
    End If
    Next j
    
    ws.Range("Q4").Value = Greatest_volume
    
    increase_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & last_row)), ws.Range("L2:L" & last_row), 0)
    ws.Range("P4").Value = ws.Cells(increase_number + 1, 9)

    
    

        
       
Next ws

End Sub

            
        

