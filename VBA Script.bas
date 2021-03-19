Attribute VB_Name = "Module1"
Sub Stock_Market()

Dim ticker_symbol As String


Dim Total_Stock_Volume As Double
Total_Stock_Volume = 0

Dim Summary_Table As Integer
Summary_Table = 2

Dim last_row As Long
 
 last_row = Cells(Rows.Count, 1).End(xlUp).Row
Dim open_date As Double
Dim close_date As Double
Dim yearly_change As Double
Dim percent_change As Double

open_date = 0

For i = 2 To last_row

Dim stock_volume As Long
   stock_volume = Cells(i, 7).Value

If open_date = 0 Then
    open_date = Cells(i, 3).Value
End If


    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
        ticker_symbol = Cells(i, 1).Value
        
        Total_Stock_Volume = Total_Stock_Volume + stock_volume
        
       Cells(Summary_Table, 9).Value = ticker_symbol
        
        Cells(Summary_Table, 10).Value = Total_Stock_Volume
        
        
        close_date = Cells(i, 6).Value
        
        yearly_change = close_date - open_date
        
        Cells(Summary_Table, 11).Value = yearly_change
        
    
        
        If yearly_change > 0 Then
        
            Cells(Summary_Table, 11).Interior.ColorIndex = 4
            
        ElseIf yearly_change < 0 Then
            Cells(Summary_Table, 11).Interior.ColorIndex = 3
            
        Else
            Cells(Summary_Table, 11).Interior.ColorIndex = 6
            
        End If
        
        
     If open_date = 0 Then
        percent_change = 0
        
    Else
        percent_change = (yearly_change / open_date)
    End If
    
    Cells(Summary_Table, 12).Value = Format(percent_change, "Percent")
    
    
    Summary_Table = Summary_Table + 1
        
        Total_Stock_Volume = 0
        
        open_date = 0
        
    Else
    
    
 Total_Stock_Volume = Total_Stock_Volume + stock_volume
 
 
        
    End If
    
Next i

        

End Sub
