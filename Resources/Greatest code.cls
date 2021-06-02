VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Greatest()

    For Each ws In Worksheets
    
                
        'Determine the Last Row and summary table row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Create Greatest Table
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        
            'Loop through all rows
            For i = 2 To LastRow
            
                'Find if Yearly Change is larger than Maximum
                If ws.Cells(i, 11).Value > Maximum Then
                    
                    'Define Maximum
                    Maximum = ws.Cells(i, 11)
                    
                    'Fill out Greatest Table
                    ws.Cells(2, 15).Value = ws.Cells(i, 9)
                    ws.Cells(2, 16).Value = Maximum
                    
                
                
                'Find if Yearly Change is smaller than Minimum
                ElseIf ws.Cells(i, 11).Value < Minimum Then
                    
                    'Define Minimum
                    Minimum = ws.Cells(i, 11)
                    
                    'Fill out Greatest Table
                    ws.Cells(3, 15).Value = ws.Cells(i, 9)
                    ws.Cells(3, 16).Value = Minimum
                    
                End If
                
                ws.Range("P2:P3").NumberFormat = "0.00%"
                
                'Determine if Stock Volume is larger than Stock_Maximum
                If ws.Cells(i, 12).Value > Stock_Maximum Then
                    
                    'Define Stock_Maximum
                    Stock_Maximum = ws.Cells(i, 12)
                    
                    'Fill out Greatest Table
                    ws.Cells(4, 15).Value = ws.Cells(i, 9)
                    ws.Cells(4, 16).Value = Stock_Maximum
                    
                End If
                
                
            Next i
        
        'Fixing column widths
        ws.Columns("N:P").AutoFit
        
        'Reset totals
        Maximum = 0
        Minimum = 0
        Stock_Maximum = 0
        
    Next ws

End Sub
