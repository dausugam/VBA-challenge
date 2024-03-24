Attribute VB_Name = "Module1"
Option Explicit

Sub Sheet_Lopp()

    Dim ws_count As Integer
    Dim i As Integer
    
    ws_count = ActiveWorkbook.Worksheets.Count
    
    For i = 1 To ws_count
        ActiveWorkbook.Sheets(i).Activate
        Ticker_Summary
    Next i

End Sub

Sub Ticker_Summary()

    'Define Variables
        Dim row As Single
        Dim summary_row As Integer
        Dim ticker As String
        Dim open_value As Single
        Dim close_value As Single
        Dim total_value As Double
    
    'Print Ticker Summary Titles
        Cells(1, 9) = "Ticker"
        Cells(1, 10) = "Yearly Change"
        Cells(1, 11) = "Percent Change"
        Cells(1, 12) = "Total Stock Volume"
        
    'Initialize Variables
        row = 2
        summary_row = 2
        ticker = Cells(row, 1)
        open_value = Cells(row, 3)
        Cells(summary_row, 9) = ticker

    'Loop through all rows
        Do Until Cells(row, 1) = Empty
            'Indentify the ticker on row and and compare with current ticker
                If ticker = Cells(row, 1) Then
                    'If it is the same ticker, increase total value and update close value
                        close_value = Cells(row, 6)
                        total_value = total_value + Cells(row, 7)
                Else
                    'if it is a different ticker, create summary of current ticker
                        Cells(summary_row, 10) = close_value - open_value
                        Cells(summary_row, 11) = Cells(summary_row, 10) / open_value
                        Cells(summary_row, 12) = total_value
                    'Reset information for new ticker
                        summary_row = summary_row + 1
                        ticker = Cells(row, 1)
                        open_value = Cells(row, 3)
                        total_value = Cells(row, 7)
                        Cells(summary_row, 9) = ticker
                End If
            row = row + 1
        Loop
        
    'Summarize the information for the last ticker
        Cells(summary_row, 10) = close_value - open_value
        Cells(summary_row, 11) = Cells(summary_row, 10) / open_value
        Cells(summary_row, 12) = total_value
        
    'Summary Table Formatting
        For row = 2 To summary_row
            'Format Yearly Change Column
                Cells(row, 10).Select
                    Selection.NumberFormat = "0.00"
                    If Selection.Value > 0 Then
                        Selection.Interior.ColorIndex = 4
                    Else
                        Selection.Interior.ColorIndex = 3
                    End If
            'Format Percent Change
                Cells(row, 11).Select
                    Selection.NumberFormat = "0.00%"
                    If Selection.Value > 0 Then
                        Selection.Interior.ColorIndex = 4
                    Else
                        Selection.Interior.ColorIndex = 3
                    End If
            'Format Total Column
                Cells(row, 12).Select
                    Selection.NumberFormat = "#,##0"
        Next row
        
    'Return the greatest increase, decrease and volume
        'Print Titles and Labels
            Cells(2, 15) = "Greatest % Increase"
            Cells(3, 15) = "Greatest % Decrease"
            Cells(4, 15) = "Greatest Total Volume"
            Cells(1, 16) = "Ticker"
            Cells(1, 17) = "Value"
            
        'Initialize values
            Cells(2, 16) = Cells(2, 9)
            Cells(3, 16) = Cells(2, 9)
            Cells(4, 16) = Cells(2, 9)
            Cells(2, 17) = Cells(2, 11)
            Cells(3, 17) = Cells(2, 11)
            Cells(4, 17) = Cells(2, 12)
            
        'Loop through summary table and update values
            For row = 3 To summary_row
                'Update Increase
                    If Cells(row, 11) > Cells(2, 17) Then
                        Cells(2, 16) = Cells(row, 9)
                        Cells(2, 17) = Cells(row, 11)
                    End If
                'Update Decrease
                    If Cells(row, 11) < Cells(3, 17) Then
                        Cells(3, 16) = Cells(row, 9)
                        Cells(3, 17) = Cells(row, 11)
                    End If
                'Update Total Volume
                    If Cells(row, 12) > Cells(4, 17) Then
                        Cells(4, 16) = Cells(row, 9)
                        Cells(4, 17) = Cells(row, 12)
                    End If
            Next row
        'Format Increase, Decrease and Total Volume Value
            Cells(2, 17).NumberFormat = "0.00%"
            Cells(3, 17).NumberFormat = "0.00%"
            Cells(4, 17).NumberFormat = "#,##0"
            
    Cells(1, 1).Select
            
End Sub
