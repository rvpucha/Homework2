Attribute VB_Name = "Module1"
Sub Stock_Change()

    'Declare variables
  
    Dim ticker As Long
    Dim row As Integer
    Dim lrow As Long
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim max_pchange As Double
    Dim max_nchange As Double
    
    'Display year
    X = ActiveSheet.Name
    MsgBox " YEAR  " & X
           
    'New column titles added
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(2, 15).Value = "Greatest % increase"
    Cells(3, 15).Value = "Greatest % decrease"
    Cells(4, 15).Value = "Greatest total volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    
    'Intialize
    column11 = 11
    column12 = 12
    ticker = 2
    row = 2
    max_pchange = 0
    max_nchange = 0
    max_vol = 0
    
    'Find number of rows in a sheet
    lrow = Cells(Rows.Count, 1).End(xlUp).row
    MsgBox ("Number of rows in active sheet  " & lrow)
    
    'Loop through rows
    
start:     For i = ticker To lrow
            'skip rows with zero opening stock
            If Cells(ticker, 3).Value = 0 Then
            ticker = i + 1
            GoTo start
                  
            'Search for ticker change
            ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
               
            'Calculate yearly change
            yearly_change = Cells(i, 6).Value - Cells(ticker, 3).Value
        
            'sum stock volume for each ticker in an year
            tot_stockvol = WorksheetFunction.Sum(Range(Cells(ticker, 7), Cells(i, 7)))
                
            'Check for cases with yearly change being Zero
            If yearly_change = 0 Then
            Cells(row, 11).Value = Format(yearly_change, "percent")
            
            Else
            'Calculate percent yearly change (non-zero)
            percent_change = yearly_change / Cells(ticker, 3).Value
            Cells(row, 11).Value = Format(percent_change, "percent")
            End If
            
            'Write data in new columns 'ticker name', 'yearly change' and 'tot.stock.vol'
            Cells(row, 9).Value = Cells(i, 1).Value
            Cells(row, 10).Value = yearly_change
            Cells(row, 12).Value = tot_stockvol
            
            'Conditional formatting :positive change in green and negative change in red
            If Cells(row, 10).Value < 0 Then
            Cells(row, 10).Interior.ColorIndex = 3
            Else
            Cells(row, 10).Interior.ColorIndex = 4
            End If
          
            'Reset the new ticker for starting loop and new row number to write data
            ticker = i + 1
            row = row + 1
              
            End If
    Next i
    
    'bonus ------------- bonus -------------- bonus
           
        'Find number of rows in Jth column
        lastrow = Cells(Rows.Count, "J").End(xlUp).row
        MsgBox ("Number of Tickers " & lastrow)
    
        'Loop through rows
        For i = 2 To lastrow
 
        If Cells(i, column11).Value > max_pchange Then
       'Find Greatest % Increase in stock
        max_pchange = Cells(i, column11).Value
        Cells(2, 17).Value = Format(max_pchange, "percent")
        Cells(2, 16).Value = Cells(i, 9).Value
        
        'Find Greatest % Decrease in stock
        ElseIf Cells(i, column11).Value < max_nchange Then
        max_nchange = Cells(i, column11).Value
        Cells(3, 17).Value = Format(max_nchange, "percent")
        Cells(3, 16).Value = Cells(i, 9).Value
        End If
    Next i
    
    For i = 2 To lastrow
        If Cells(i, column12).Value > max_vol Then
           
        'Find Greatest total volume of stock
        max_vol = Cells(i, column12).Value
        Cells(4, 17).Value = max_vol
        Cells(4, 16).Value = Cells(i, 9).Value
    
        End If
    Next i
    'show fancy arrows for greatest values
    Range("N2:N4") = ChrW(62) & ChrW(62) & ChrW(62) & ChrW(62) & ChrW(62) & ChrW(62)
    Range("N2:N4").Font.Size = 12
    Range("N2:N4").Font.Color = vbMagenta

   'Autofit column widths for decent look
    Worksheets(X).Columns("A:Q").AutoFit
End Sub
