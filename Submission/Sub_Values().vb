Sub Values()

    ' stocks are pre-sorted

    ' LOOP by rows

    ' Big idea: create a leaderboard

    ' IF next stock is different, that means we have finished our group

    ' ELSE, then keep summing the volume

    ' variables
    Dim WS As Worksheet
    
    Dim ticker As String
    Dim next_ticker As String
    Dim vol As LongLong
    Dim vol_total As LongLong
    Dim row As Long
    Dim leaderboard_row As Long
    Dim last_row As Long
    Dim opened As Double
    Dim closed As Double
        
    ' new variables
    Dim qchange As Double
    Dim pct_qchange As Double
    
    ' begin the loop per worksheet
    ' (my own psuedo code, borrowing an opening line from Prof. Booth)
    For Each WS In ThisWorkbook.Worksheets

        ' Set Column Headers
        WS.Cells(1, 9).Value = "Ticker"
        WS.Cells(1, 10).Value = "Quarterly Change"
        WS.Cells(1, 11).Value = "Percent Change"
        WS.Cells(1, 12).Value = "Total Volume"
        

        ' Reset per stock
        ' (This peice of Code is borrowed from Prof. Booth's solution with variables changed to match those used for this xlsm)
        vol_total = 0
        opened = WS.Cells(2, 3).Value
        leaderboard_row = 2
        last_row = WS.Cells(Rows.Count, 1).End(xlUp).row
    
        
            ' extract values from workbook
            ' (this code a mix of original creation a copied from Prof. solution, ended up being more or less identical, unsure of other methods [such as range] to exttract values since we are using specific cells)
        For row = 2 To last_row
            ticker = WS.Cells(row, 1).Value
            opened = WS.Cells(row, 3).Value
            closed = WS.Cells(row, 6).Value
            vol = WS.Cells(row, 7).Value
            next_ticker = WS.Cells((row + 1), 1).Value
            
            ' if statement
            ' (my own psudeo) This will set up the process to add a new stock to the leaderboard.
            If ticker <> next_ticker Then
                                   
                ' add total
                vol_total = vol_total + vol
                
                ' Change logic
                closed = WS.Cells(row, 6).Value
                qchange = (closed - opened)
                pct_qchange = (qchange / opened)
                            
        
                ' write to leaderboard
                WS.Cells(leaderboard_row, 9).Value = ticker
                WS.Cells(leaderboard_row, 10).Value = qchange
                WS.Cells(leaderboard_row, 11).Value = FormatPercent(pct_qchange) 'used Prof.'s formatting tip
                WS.Cells(leaderboard_row, 12).Value = vol_total
            
                ' Conditional Formatting
                ' (This is all copied from Prof. example with variables adjusted)
                If (qchange > 0) Then
                    ws.Cells(leaderboard_row, 10).Interior.ColorIndex = 4
                ElseIf (qchange < 0) Then
                    ws.Cells(leaderboard_row, 10).Interior.ColorIndex = 3
                Else
                    ' Do Nothing (default White)
                End If
    
                ' reset total
                vol_total = 0
                leaderboard_row = leaderboard_row + 1
                
                'reset new open price
                opened = ws.cells(row + 1, 3).Value

            Else

                ' add total
                vol_total = vol_total + vol
        
            '(My own Psuedo) This will close the loop that created the first leaderboard
            End If
        Next row
            
        ' Second Loop for Second Leaderboard

        ' (My own Psuedo) Will need new headers in Columns P and Q
        WS.Cells(1, 16).Value = "Ticker"
        WS.Cells(1, 17).Value = "Value"
        ' (My own Psuedo) And new labels in O2, O3, O4
        ws.Cells(2,15).Value = "Greatest % increase"
        ws.Cells(3,15).Value = "Greatest % decrease"
        ws.Cells(4,15).Value = "Greatest Total Volume"

        ' (My own Psuedo) And new variables
            Dim max_pct_change As Double
            Dim min_pct_change As Double
            Dim max_pct_change_ticker As String
            Dim min_pct_change_ticker as String
            Dim greatest_volume as LongLong
            Dim greatest_volume_ticker As String  

        ' init to first row of the first leaderboard for comparison
        max_pct_change = ws.cells(2,11).value
        min_pct_change = ws.cells(2,11).value
        greatest_volume = ws.cells(2,11).value
        max_pct_change_ticker = ws.cells(2, 1)
        min_pct_change_ticker = ws.cells(2, 1)
        greatest_volume_ticker = ws.cells(2, 1)

            ' (My own Psuedo) Now open the second loop
            dim i as Integer

            for i = 2 to leaderboard_row 'borrowed from Prof.'s code

            ' Compare current row to the inits (first row)
                ' We have a new Max Percent Change!
                If (ws.cells(i, 11) > max_pct_change) Then
                    max_pct_change = ws.cells(i, 11)
                    max_pct_change_ticker = ws.cells(i, 9).value

                ' We have a new Min Percent Change!
                    ElseIf (ws.cells(i, 11) < min_pct_change) Then
                        min_pct_change = ws.cells(i, 11)
                        min_pct_change_ticker = ws.cells(i, 9).value
            
                End If

                ' We have a new Max Volume!
                If (ws.cells(i, 12)) > greatest_volume Then
                    greatest_volume = ws.cells(i, 12)
                    greatest_volume_ticker = ws.cells(i, 9)

                    End If

            Next i        
        
        ' Write out to Excel Workbook
        ws.cells(2, 16).value = max_pct_change_ticker
        ws.cells(3, 16).value = min_pct_change_ticker
        ws.cells(4, 16).value = greatest_volume_ticker
        ws.cells(2, 17).value = max_pct_change
        ws.cells(3, 17).value = min_pct_change
        ws.cells(4, 17).value = greatest_volume



   
    Next WS
    
End Sub