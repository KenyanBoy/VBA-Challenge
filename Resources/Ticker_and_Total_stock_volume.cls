VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Ticker_and_Total_Stock_Volume()

    For Each ws In Worksheets
    
        
        'Determine the Last Row and summary table row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        Summary_Table_Row = 2
        
        'Create titles for columns
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

        'Create variables for above columns
        Total_Stock_Volume = 0
        
            'Loop through all rows
            For i = 2 To LastRow
            
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                    'Set Ticker Name
                    Ticker_Name = ws.Cells(i, 1).Value
                
                    'Add to Totals
                    Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
                
                    'Add Values to summary table
                    ws.Cells(Summary_Table_Row, 9).Value = Ticker_Name
                    ws.Cells(Summary_Table_Row, 12).Value = Total_Stock_Volume
                   
                    'Add one to summary table for next row
                    Summary_Table_Row = Summary_Table_Row + 1
                
                    'Reset Totals
                    Total_Stock_Volume = 0
                
                Else
                    
                    'Add to Total Stock Volume
                    Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
                    
                End If
                
            Next i
        
        'Fixing column widths
        ws.Columns("I:L").AutoFit
        
    Next ws
    

End Sub
