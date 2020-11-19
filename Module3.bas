Attribute VB_Name = "Module3"
Sub Multiple_stock()
    Dim ws As Worksheet
    Dim ticker As String
    Dim ticker_volume As Double
    Dim i As Double
    Dim summary_table_row As Integer
    Dim start As Double
    Dim Year_open As Double
    Dim year_Close As Double
    Dim yearly_change As Double
    Dim percent_change As Double
         
    For Each ws In Worksheets
    start = 2
    summary_table_row = 2
    ticker_volume = 0
   
    RowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row
             
    For i = 2 To RowCount
                 
    'check if we are are still within the same ticker,
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
       

        ticker = ws.Cells(i, 1).Value
        ticker_volume = ticker_volume + ws.Cells(i, 7).Value
       
        Year_open = ws.Cells(start, 3).Value
        year_Close = ws.Cells(i, 6).Value
                             
        yearly_change = year_Close - Year_open
        If yearly_change <> 0 And Year_open <> 0 Then
   
        percent_change = (year_Close - Year_open) / Year_open * 100
        
        Else
        percent_change = 0
        End If
        
        
       
        ws.Range("I" & summary_table_row).Value = ticker
        ws.Range("J" & summary_table_row).Value = yearly_change
        ws.Range("K" & summary_table_row).Value = percent_change
        ws.Range("L" & summary_table_row).Value = ticker_volume
        
        'Conditional Formatting of Yearly change
          If yearly_change > 0 Then
          ws.Range("J" & summary_table_row).Interior.ColorIndex = 4
          Else
          ws.Range("J" & summary_table_row).Interior.ColorIndex = 3
          End If
        
       
        summary_table_row = summary_table_row + 1
        ticker_volume = 0
        start = i + 1
           
         Else
          ticker_volume = ticker_volume + ws.Cells(i, 7).Value
       
         End If
         
         Next i
         
         Next ws
                               
           
End Sub
