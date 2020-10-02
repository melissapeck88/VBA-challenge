Sub alphabetical_testing()

  ' Set initial variables
  Dim ticker As String
  Dim ticker_summary As String
  Dim yearly_open As Double
  Dim yearly_change As Double
  Dim percent_change As Integer
  Dim total_stock_volume As Double
  Dim ws As Worksheet
  
  yearly_change = 0
  percent_change = 0
  total_stock_volume = 0
  
  'Set additional summary values
  'Dim greatest_percent_increase_ticker As String
  'Dim greatest_percent_increase As Double
  'Dim greatest_percent_decrease As Double
  'Dim greatest_total_volume As Double
  
  For Each ws In Worksheets
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    Dim start As Long
    start = 2
   
    
    'add headers
    ws.Cells(1, 9).Value = "ticker_summary"
    ws.Cells(1, 12).Value = "yearly_change"
    ws.Cells(1, 11).Value = "percent_change"
    ws.Cells(1, 10).Value = "total_stock_volume"

  ' location in the summary table and last row
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  
 ' lastrowsummary = ws.cells(Rows.Count, 9).End(xlUp).Row
  
    
  ' Loop through all lines
  For i = 2 To lastRow
    'lastRow = ws.cells(Rows.Count, 1).End(xlUp).Row
    ' Check if we are still within the same ticker
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      ' Set the ticker_summary
      ticker_summary = ws.Cells(i, 1).Value

      ' Add to the total_stock_volume
      total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
     yearly_open = ws.Cells(start, 3).Value
      'Calculate the yearly_change
      yearly_change = ws.Cells(i, 6).Value - yearly_open
      
      
      'Calculate the percent_change
       If yearly_open = 0 Then
       percent_change = 0
                  
        Else
        percent_change = yearly_change / yearly_open * 100
        End If

      ' Print the ticker in the Summary Table
      ws.Range("I" & Summary_Table_Row).Value = ticker_summary

      ' Print the total_stock_volume to the Summary Table
      ws.Range("J" & Summary_Table_Row).Value = total_stock_volume
      
      'Print the yearly_change
      ws.Range("L" & Summary_Table_Row).Value = yearly_change
      
        'Conditional formatting for yearly_change
        If yearly_change > 0 Then
            ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
        
        ElseIf yearly_change < 0 Then
            ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
        
        Else
            ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 0
        
        End If
      
      'Print the percent_change
    ws.Range("K" & Summary_Table_Row).Value = percent_change

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the total_stock_volume
      total_stock_volume = 0
      
      start = i + 1
      

    ' If the cell immediately following a row is the same ticker indicator
    
    Else

      ' Add to the total_stock_volume
      total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value

    End If

  Next i
Next ws

End Sub


