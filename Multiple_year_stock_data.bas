Attribute VB_Name = "Module1"
Sub Homework()
Dim Maxrow As Double

Maxrow = Cells(Rows.Count, 1).End(xlUp).Row
   
Dim WS As Worksheet
    Set starting_ws = Sheets(1)

       
   'set summary column headers

        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        
        'Create Variable to hold Value
        Dim Open_Price As Double
        Dim Close_Price As Double
        Dim Yearly_Change As Double
        Dim Ticker_Name As String
        Dim Percent_Change As Double
        Dim Volume As Double
       
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        
        Open_Price = Cells(2, 3).Value
        
         
         ' Loop through all ticker
        
        For i = 2 To Maxrow
        
     
        
         ' Check if we are still within the same ticker symbol, if it is not...
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
                Ticker_Total = Ticker_Total + Cells(i, 7).Value
                ' Find all the values
                
                Ticker_Name = Cells(i, 1).Value
                
                Close_Price = Cells(i, 6).Value
               
                Yearly_Change = Close_Price - Open_Price
                
                ' Add Percent Change
                If Open_Price = 0 Then
                    Percent_Change = 0
                Else
                    Percent_Change = Yearly_Change / Open_Price
                 End If
                 
                 
      ' Print in the Summary Table
    
             Range("I" & Summary_Table_Row).Value = Ticker_Name
             Range("J" & Summary_Table_Row).Value = Yearly_Change
             Range("L" & Summary_Table_Row).Value = Ticker_Total
             Range("K" & Summary_Table_Row).Value = Percent_Change
      
                        
      ' Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
      
       ' Setting the colors
        
          'If value in Yearly Change column is positive
          
            If Cells(i, 10).Value > 0 Then
                Cells(i, 10).Interior.ColorIndex = 4
            ElseIf Cells(i, 10).Value = 0 Then
                Cells(i, 10).Interior.ColorIndex = 2
             ElseIf Cells(i, 10).Value < 0 Then
                Cells(i, 10).Interior.ColorIndex = 3
            End If
        

      ' Reset the Ticker Total
       Ticker_Total = 0
        Summary_Table_Row = Summary_Table_Row + 1
        Open_Price = Cells(i + 1, 3).vlaue
                
        End If
    Next i

        

Summary_Table_Row = Summary_Table - 1
        
        'For Bonus'
        
        ' Set Greatest % Increase, % Decrease, and Total Volume
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
        ' Look through each rows to find the greatest value and its associate ticker
        For Z = 2 To YCLastRow
            If Cells(2, 11).Value = Application.WorksheetFunction.Max(WS.Range("K2:K" & YCLastRow)) Then
                Cells(2, 16).Value = Cells(2, 9).Value
                Cells(2, 17).Value = Cells(2, 11).Value
                
            ElseIf Cells(2, 11).Value = Application.WorksheetFunction.Min(WS.Range("K2:K" & YCLastRow)) Then
                Cells(3, 16).Value = Cells(2, 9).Value
                Cells(3, 17).Value = Cells(2, 11).Value
                
            ElseIf Cells(2, 12).Value = Application.WorksheetFunction.Max(WS.Range("L2:L" & YCLastRow)) Then
                Cells(4, 16).Value = Cells(2, 9).Value
                Cells(4, 17).Value = Cells(2, 12).Value
            End If
        Next Z
        
      

End Sub
