Sub stock_function():
        'Loop through sheets
    For Each ws In Worksheets
        ws.Activate
            'create column names
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "YearlyChange"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"

            'Declaration of variables
        Dim current_ticker As String, next_ticker As String
        Dim open_price As Double, close_price As Double, volume As Double, perc_change As Double, gr_perc_inc As Double, gr_perc_dec As Double, gr_tot_vol As Double
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        Dim last_row As Long
        last_row = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        volume = 0
        open_price = ws.Cells(2, 3).Value
                ' Using this to debug. Range("O1").Value = close_price
            'Loop through data
        For i = 2 To last_row
                'If ticker is about to change then summarize and paste data.
            If i = last_row Or Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                current_ticker = Cells(i, 1).Value
                volume = volume + Cells(i, 7).Value
                close_price = Cells(i, 6).Value
                    ' Using this to debug.Range("O1").Value = close_price
                Range("I" & Summary_Table_Row).Value = current_ticker
                Range("L" & Summary_Table_Row).Value = volume
                Range("J" & Summary_Table_Row).Value = close_price - open_price
                perc_change = (close_price - open_price) / open_price
                Range("K" & Summary_Table_Row).Value = perc_change
                Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                
                    'Conditional Formatting for Color of YR change and Per Change columns
                If ws.Range("J" & Summary_Table_Row).Value < 0 Then
                    ws.Range("J" & Summary_Table_Row).Interior.Color = vbRed

                ElseIf ws.Range("J" & Summary_Table_Row).Value > 0 Then
                    ws.Range("J" & Summary_Table_Row).Interior.Color = vbGreen
                Else
                    ws.Range("J" & Summary_Table_Row).Interior.Color = vbYellow
                End If
                
                If ws.Range("K" & Summary_Table_Row).Value < 0 Then
                    ws.Range("K" & Summary_Table_Row).Interior.Color = vbRed

                ElseIf ws.Range("K" & Summary_Table_Row).Value > 0 Then
                    ws.Range("K" & Summary_Table_Row).Interior.Color = vbGreen
                Else
                    ws.Range("K" & Summary_Table_Row).Interior.Color = vbYellow
                End If
                    ' End of Main If block, time to reset variables.
                Summary_Table_Row = Summary_Table_Row + 1
                volume = 0
                open_price = ws.Cells(i + 1, 3).Value
                ' Block where ticker will not change, need to update Volume.
            Else
                volume = volume + Cells(i, 7).Value
            End If
        Next i
            'Declaration and gathering of additional information Greatest % increase", "Greatest % decrease", and "Greatest total volume
            'Gather the max min and max volume.
        gr_perc_inc = Application.WorksheetFunction.Max(ws.Columns("K"))
        gr_perc_dec = Application.WorksheetFunction.Min(ws.Columns("K"))
        gr_tot_vol = Application.WorksheetFunction.Max(ws.Columns("L"))
            'enter the information found above.
        ws.Range("Q2").Value = gr_perc_inc
        ws.Range("Q3").Value = gr_perc_dec
        ws.Range("Q4").Value = gr_tot_vol
            'find ticker associated with values found for Greatest % increase", "Greatest % decrease", and "Greatest total volume
        Dim ticker_inc As String
        Dim ticker_dec As String
        Dim ticker_vol As String
        ticker_inc = Application.WorksheetFunction.XLookup(gr_perc_inc, ws.Columns("K"), ws.Columns("I"))
        ticker_dec = Application.WorksheetFunction.XLookup(gr_perc_dec, ws.Columns("K"), ws.Columns("I"))
        ticker_vol = Application.WorksheetFunction.XLookup(gr_tot_vol, ws.Columns("L"), ws.Columns("I"))
            'paste ticker values.
        ws.Range("P2").Value = ticker_inc
        ws.Range("P3").Value = ticker_dec
        ws.Range("P4").Value = ticker_vol
            'format columns for dollar sign and percent signs. Autofit the sheet.
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").NumberFormat = "0.00%"
        ws.Columns("L").NumberFormat = "General"
        ws.Columns("J").NumberFormat = "$#,##0.00"
        ws.Columns.AutoFit
        ws.Rows.AutoFit
            'Move on to next sheet.
        Next ws
    End Sub
    

    
        
   



