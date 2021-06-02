VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Yearly_and_Percent_Change()

Dim OpenPrice As Double
Dim ClosePrice As Double
Dim YearlyChange As Double
Dim PercentChange As Double

    For Each ws In Worksheets
    
        
        'Determine the Last Row and summary table row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        Summary_Table_Row = 2
        
        
        'Set Yearly Open Price
        OpenPrice = ws.Cells(2, 3).Value
        
            'Loop through all rows
            For i = 2 To LastRow
                
                'Check when Ticker changes value
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                    'Set Yearly Close Price
                    ClosePrice = ws.Cells(i, 6).Value
                
                    'Set Value of Yearly Change
                    YearlyChange = ClosePrice - OpenPrice
                    ws.Cells(Summary_Table_Row, 10).Value = YearlyChange
                
                'Determine if YearlyChange is red or green
                If YearlyChange > 0 Then
                
                    'Positive change = Green
                    ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
                    
                ElseIf YearlyChange <= 0 Then
                
                    'Negative change = Red
                    ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
                    
                End If
                
                'Avoiding error
                If OpenPrice <> 0 Then
                
                    PercentChange = (YearlyChange / OpenPrice)
                    
                    ws.Cells(Summary_Table_Row, 11).Value = PercentChange
                
                End If
                
                'Move down one row in summary table
                Summary_Table_Row = Summary_Table_Row + 1
                
                'Set new opening price
                OpenPrice = ws.Cells(i + 1, 3).Value
                
            End If
           
        Next i
        
        'Changing format of Percent Change to percentage
        ws.Range("K2:K" & LastRow).NumberFormat = "0.00%"
        
    Next ws

End Sub
