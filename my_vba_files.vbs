Sub wsloop()

'define variables
Dim mainWS As Worksheet
Dim ws As Workbook
Dim headers() As Variant


Dim ticker As String
Dim total_stock_volumn As Double
Dim vol As Integer
Dim year_open As Double
Dim year_close As Double
Dim yearly_change As Double
Dim percent_change As Double


'set row count for the workbook
Dim lastrow As Long



'set location for variable
Dim summary_table_row As Long


'loop through each worksheet
For Each mainWS In Worksheets

'set column names
mainWS.Cells(1, 9).Value = "ticker"
mainWS.Cells(1, 10).Value = "yearly_change"
mainWS.Cells(1, 11).Value = "percent_change"
mainWS.Cells(1, 12).Value = "total_stock_volumn"

summary_table_row = 2


'loop through all sheets to find last cell that is not empty
lastrow = mainWS.Cells(Rows.Count, 1).End(xlUp).Row


    'loop throught the worksheet
    For i = 2 To lastrow

    If mainWS.Cells(i - 1, 1).Value <> mainWS.Cells(i, 1).Value Then


'set intial value of open stock vlaue for the first ticker of mainws
year_open = mainWS.Cells(i, 3).Value
total_stock_volumn = 0


        End If
 
        total_stock_volumn = total_stock_volumn + mainWS.Cells(i, 7).Value
 

    'same ticker name
        If mainWS.Cells(i + 1, 1).Value <> mainWS.Cells(i, 1).Value Then
        
        
            'set starting point
            ticker = mainWS.Cells(i, 1).Value
            
            'calculate
            year_close = mainWS.Cells(i, 6).Value
            yearly_change = year_close - year_open
            percent_change = (yearly_change / year_open) * 100
  
    
            
            'insert values into summary
            mainWS.Range("I" & summary_table_row).Value = ticker
            mainWS.Range("J" & summary_table_row).Value = yearly_change
            mainWS.Range("k" & summary_table_row).Value = percent_change
            mainWS.Range("L" & summary_table_row).Value = total_stock_volumn
            
            

         'color fill red for negative and green for positive change
    
            If (yearly_change > 0) Then
            mainWS.Range("J" & summary_table_row).Interior.ColorIndex = 4
            

            ElseIf (yearly_change <= 0) Then
             mainWS.Range("J" & summary_table_row).Interior.ColorIndex = 3
             
             

            End If

            summary_table_row = summary_table_row + 1
    
    End If

  Next i

  Next mainWS
  

End Sub
