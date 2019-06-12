'This function tallies stock volume for each stock and then creates a table within the same worksheet for each stock
Sub Test()

'Set a variable to hold the ticker
Dim ticker As String

'Set a varaible to hold the total stock volume of each ticker
Dim total_volume As Double
total_volume = 0

'Set a variable to keep track of the row on the summary table
Dim summary_table_row As Integer
summary_table_row = 2

'Set a variable for starting the loop at a certain row
Dim r As Long

'Loop through every worksheet
For Each ws In Worksheets
    
    'Name the summary table columns
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Total Stock Volume"
    
    'Last row of the stock data
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Loop through all the stock data
    For r = 2 To LastRow
        
        'If the following cell is different then...
        If ws.Cells(r, 1).Value <> ws.Cells(r + 1, 1).Value Then
        
            'Set the ticker name
            ticker = ws.Cells(r, 1).Value
            
            'Add to the total stock volume
            total_volume = total_volume + ws.Cells(r, 7).Value
            
            'Print the ticker name to the summary table
            ws.Range("I" & summary_table_row).Value = ticker
            
            'Print the total volume to the summary table
            ws.Range("J" & summary_table_row).Value = total_volume
            
            'Reset the total stock volume to 0
            total_volume = 0
            
            'Move to the following row
            summary_table_row = summary_table_row + 1
            
        Else
        
            'Add to the total stock volume
            total_volume = total_volume + ws.Cells(r, 7).Value
            
        End If
            
    Next r
    
    'Reset the summary table for each worksheet
    summary_table_row = 2

Next ws
        

End Sub


