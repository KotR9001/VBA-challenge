Attribute VB_Name = "Module1"
Sub StockTest()
    
'Create Variable to Loop Through All Worksheets
'Code from Wells_Fargo Activity
For Each ws In Worksheets
    
    'Declare Variable for Ticker
    Dim ticker As String
    
    'Declare Variable for Opening Value
    Dim opening_value As Variant
    
    'Declare Variable for Closing Value
    Dim closing_value As Variant
    
    'Declare Variable for Yearly Change from Opening Price at Beginning of Year to Closing Price at End of Year
    'Dim yearly_change As Variant
    
    'Declare Variable for Percentage Change from Opening Price at Beginning of Year to Closing Price at End of Year
    'Dim percent_change As Single
    
    'Declare Total Stock Volume
    Dim total_volume As Single
    
    'Declare Row Count
    Dim totalrows As Single
    
    'Declare Summary Table Row Position
    Dim summary_table_row As Integer
    
    'Initialize Values
    yearly_change = 0
    percent_change = 0
    total_volume = 0
    summary_table_row = 2
    
    'Find the Total Number of Rows
    'Code format from combined_wells_fargo Activity
    totalrows = ws.Cells(Rows.Count, 1).End(xlUp).Row
    'MsgBox (totalrows)
    
    'Create Summary Table Column Headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    'Set the Opening Value for the First Ticker Value
    opening_value = ws.Cells(2, 3).Value
    'MsgBox (opening_value)
    
    'Run Loop for Obtaining Ticker, Yearly Change, Percentage Change, & Total Stock Volume Values
    
    'Iterate through Rows (Days of Transactions)
    For i = 2 To totalrows
        
        'Compare Next Row with Current Row
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            'Set the Ticker Value
            ticker = ws.Cells(i, 1).Value
            'MsgBox (ticker)
            
            'Assign the Ticker Value to Column I in the Summary Table
            ws.Range("I" & summary_table_row).Value = ticker
            
            'Calculate the Total Stock Volume
            total_volume = total_volume + ws.Cells(i, 7).Value
            'MsgBox (total_volume)
            
            'Assign the Total Volume to Column L in the Summary Table
            ws.Range("L" & summary_table_row).Value = total_volume
            
            'Set the Closing Value
            closing_value = ws.Cells(i, 6).Value
            
            'Calculate the Yearly Change From Opening Price at Beginning of Year to Closing Price at End of Year
            yearly_change = closing_value - opening_value
            
            'Assign the Yearly Change Value to Column J in the Summary Table
            'Set Yearly Change to Two Decimal Places
            ws.Range("J" & summary_table_row).Value = Round(yearly_change, 2)
            'Apply Conditional Formatting to Column J
            'From https://www.automateexcel.com/excel-formatting/color-reference-for-color-index/
            If yearly_change > 0 Then
                ws.Cells(summary_table_row, 10).Interior.Color = vbGreen
            ElseIf yearly_change < 0 Then
                ws.Cells(summary_table_row, 10).Interior.Color = vbRed
            End If
            
            'Calculate the Percent Change From Opening Price at Beginning of Year to Closing Price at End of Year
            If opening_value <> 0 Then
                'Return Percent Change Value
                percent_change = (closing_value - opening_value) / opening_value
            ElseIf opening_value = 0 Then
                'Return Value of 0
                precent_change = 0
            End If
            
            'Assign the Percent Change to Column K of the Summary Table
            ws.Range("K" & summary_table_row).Value = percent_change
            'Apply Percentage Format to Column K
            'From https://excelvbatutor.com/vba_lesson9.htm
            ws.Range("K" & summary_table_row) = Format(percent_change, "Percent")
            
            'Increment the Summary Table Row#
            summary_table_row = summary_table_row + 1
            
            'Check to See If There is a Ticker Value in (i+1)
            If ws.Cells(i + 1, 1).Value <> " " Then
                'Set the Opening Value for the Next Ticker Value
                opening_value = ws.Cells(i + 1, 3).Value
            ElseIf Cells(i + 1, 1).Value = " " Then
                'End the Insertion of Ticker Value
                opening_value = " "
            End If
            'Reset the Total Volume to Zero
            total_volume = 0
        
        Else
        
            'Add to the Total Volume
            total_volume = total_volume + ws.Cells(i, 7).Value
        
        End If
    Next i
    
    'Resize Summary Table Columns
    'From https://docs.microsoft.com/en-us/office/vba/api/excel.range.autofit
    ws.Columns("I:I").AutoFit
    ws.Columns("J:J").AutoFit
    ws.Columns("K:K").AutoFit
    ws.Columns("L:L").AutoFit
    
    'Call the Greats Subroutine
    'Call Greats
    
'Go to the Next Worksheet
Next ws

End Sub
Sub Greats()

'Create Variable to Loop Through All Worksheets
'Code from Wells_Fargo Activity
'For Each ws In Worksheets

    'Declare Variables
    Dim ticker_increase As String
    Dim ticker_decrease As String
    Dim ticker_total_volume As String
    Dim value_increase As Variant
    Dim value_decrease As Variant
    Dim value_total_volume As Variant
    Dim totalrows As Integer
    
    'Initialize the Values
    'From https://www.wallstreetmojo.com/vba-max/
    value_total_volume = 0
    
    'Create Row & Column Headers for Greatest Values Table
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    
    'Find the Total Number of Rows
    'Code format from combined_wells_fargo Activity
    totalrows = Cells(Rows.Count, 9).End(xlUp).Row
    
    'Find the Values
    
    'Loop for Greatest % Increase Values
    For i = 2 To totalrows
        'See If % Change from Next Row is > Current Row
        If Cells(i + 1, 11).Value > Application.WorksheetFunction.Max(Range("K2:K" & i)) Then
            'Replace Greatest % Increase Ticker and Value
            ticker_increase = Cells(i + 1, 9).Value
            value_increase = Format(Cells(i + 1, 11).Value, "Percent")
        End If
    Next i
    
    'Loop for Greatest % Decrease Values
    For i = 2 To totalrows
        'See If % Change from Next Row is < Current Row
        If Cells(i + 1, 11).Value < Application.WorksheetFunction.Min(Range("K2:K" & i)) Then
            'Replace Greatest % Decrease Ticker and Value
            ticker_decrease = Cells(i + 1, 9).Value
            value_decrease = Format(Cells(i + 1, 11).Value, "Percent")
        End If
    Next i
    
    'Loop for Greatest Total Volume Values
    For i = 2 To totalrows
        'See If Greatest Total Volume from Next Row is > Current Row
        If Cells(i + 1, 12).Value > value_total_volume Then
            'Replace Greatest Total Volume Ticker and Value
            ticker_total_volume = Cells(i + 1, 9).Value
            value_total_volume = Cells(i + 1, 12).Value
        End If
    Next i
    
    'Assign the Values to the Greatest Values Summary Table
    Cells(2, 16).Value = ticker_increase
    Cells(2, 17).Value = value_increase
    Cells(3, 16).Value = ticker_decrease
    Cells(3, 17).Value = value_decrease
    Cells(4, 16).Value = ticker_total_volume
    Cells(4, 17).Value = value_total_volume
    
    'Resize the Columns for the Greatest Values Summary Table
    Columns("O:O").AutoFit
    Columns("P:P").AutoFit
    Columns("Q:Q").AutoFit

'Go to the Next Worksheet
'Next ws

End Sub
Sub Delete()

'Cycle Through Every Worksheet
For Each ws In Worksheets

    'Delete Created Columns
    'From https://software-solutions-online.com/how-to-delete-columns-in-excel-using-vba/#Example_6_Delete_multiple_columns_in_a_table
    ws.Range("I:L", "O:Q").EntireColumn.Delete
    
    'Remove Conditional Formatting
    'From https://www.excel-easy.com/vba/examples/background-colors.html
    ws.Columns("J:J").Interior.ColorIndex = 0
    
'Go to the Next Worksheet
Next ws

End Sub


