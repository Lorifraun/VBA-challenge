Attribute VB_Name = "Module1"
Sub VBAMultiYear()

'start with going through each worksheet
For Each ws In Worksheets

'set varible for holding the ticker symbol name
Dim Ticker As String

'set varible for holding the Volume
Dim Volume As Long
Volume = 0

'set varible for holding the Total Volume
Dim Total_Volume As Double
Total_Volume = 0

'set variable for holding the stock open
Dim Year_Open As Double
Year_Open = ws.Cells(2, 3).Value
Year_Open = 0

'set varible for holding the stock close
Dim Year_Close As Double
Year_Close = 0

'set varible for holding the Yearly_Change
Dim Yearly_Change As Double
Yearly_Change = 0

'set varible for holding the Percent_Change
Dim Percent_Change As Single
Percent_Change = 0

'Dim ws As Worksheet
Dim LastRow As Long
Dim i As Long


ws.Columns("J").NumberFormat = "0.00"
ws.Columns("K").NumberFormat = "0.00%"


LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
Lastcolumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column

 ' Keep track of the location for each ticker in the summary table
   Dim Summary_Table_Row As Integer
   Summary_Table_Row = 2
   
      
  'Add the column header
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
        
    'Loop through all Ticker Symbols
    For i = 2 To LastRow
    
        'Checks to see if ticker in the row below is not the same as the one above
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

        'set for Ticker name
        Ticker = ws.Cells(i, 1).Value
                       
        'set and add to the Total Volume
        Total_Volume = Total_Volume + ws.Cells(i, 7).Value
                      
               
        'Print the value assigned to the ticker varible in the Summary Table
        ws.Range("I" & Summary_Table_Row).Value = Ticker
        
       'Print the volume total in the summary table
        ws.Range("L" & Summary_Table_Row).Value = Total_Volume
        
        
        'set Year_Close
        Year_Close = ws.Cells(i, 6).Value
        
        Yearly_Change = (Year_Close - Year_Open)
        
        'Print the year change
        ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
        
        
                       
        'calculate percent change
        If Year_Open <> 0 Then
            Percent_Change = (Yearly_Change / Year_Close)
            ws.Range("K" & Summary_Table_Row).Value = Percent_Change
        Else
            ws.Range("K" & Summary_Table_Row).Value = 0
        End If
        
        'conditional formatting
            If Yearly_Change >= 0 Then
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
            Else
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
            End If
                
        
                  
        'move to next row to select open value
        Year_Open = ws.Cells(i + 1.3).Value
                
        'Print next Ticker to the below row in the Summary Table
        Summary_Table_Row = Summary_Table_Row + 1
        
        'Reset Total Volume
        Total_Volume = 0
                           
                                  
        'If false
        Else

        'Add to the Volume Total
        Total_Volume = Total_Volume + ws.Cells(i, 7).Value

                   
        End If
             
        'move to the next row

    Next i

    
    
    Next ws
End Sub

