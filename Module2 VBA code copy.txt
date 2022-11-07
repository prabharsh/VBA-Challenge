' '-------------------------------
' Bootcamp: UTOR-VIRT-DATA-PT-10-2022-U-LOLC-MTTH
' Module 2 Challenge - VBA
' Objective : VBA scripting to analyze generated stock market data
' Student Name : Prabha Shankar
' EMAIL : prabhars@icloud.com
' '-------------------------------





Sub StockAnalysis()


' '-------------------------------
' Declare 1D variables
' '-------------------------------
 
Dim ws As Worksheet
Dim ticker_name As String
Dim trade_dt As String
Dim trade_date As Date
Dim open_value As Double
Dim high_value As Double
Dim low_value As Double
Dim close_value As Double
Dim trade_volume As Double
Dim row_num As Double
Dim output_rownum As Double
Dim last_ticker As String

  'Declare 1D variables FOR BONUS
    Dim open_value_yr As Double
    Dim close_value_yr As Double
    Dim trade_volume_yr As Double
    
  'Declare 2D Array for BONUS PART
    Dim greatest_movers(3, 2) As Variant

' '---------------------------------------------------------------
' Sort the sheet first by Column A then by Column B
' '---------------------------------------------------------------

    For Each ws In Worksheets
        With ws.Sort
             .SortFields.Add Key:=Range("A1"), Order:=xlAscending
             .SortFields.Add Key:=Range("B1"), Order:=xlAscending
             .SetRange Columns("A:G")
             .Header = xlYes
             .Apply
        End With
' '--------------------------------------------------------------------------------
' Populating headers / Assign variables value for plus1 (Column A and Column I)
' '--------------------------------------------------------------------------------
        
        ws.Range("I1") = "Ticker"
        ws.Range("L1") = "Total Stock Volume"
        ws.Range("J1") = "Yearly Change"
        ws.Range("K1") = "Percent Change"
        output_rownum = 2 'Column I ticker next line switch variable
        row_num = 2 'Column A ticker next line switch variable
        
' '--------------------------------------------------------------
' Read and assign values to varients from C,D,E,F,G Columns
' '--------------------------------------------------------------
         
       Do 'Worksheet calculations start
        
            ticker_name = ws.Range("A" & row_num)
            open_value = CDbl(ws.Range("C" & row_num).Value) 'Constant open value
            high_value = CDbl(ws.Range("D" & row_num).Value)
            low_value = CDbl(ws.Range("E" & row_num).Value)
            close_value = CDbl(ws.Range("F" & row_num).Value)
            trade_volume = CDbl(ws.Range("G" & row_num).Value)
        
        
        'STEP ONE - first time run - one time
            If last_ticker = "" Then
                open_value_yr = open_value 'Constant open value for Ticker greatest calc - BONUS PART
                last_ticker = ticker_name
                close_value_yr = close_value
                greatest_movers(0, 0) = ticker_name ''Greatest % increase Ticker - BONUS PART
                greatest_movers(1, 0) = ticker_name ''Greatest % decrease Ticker- BONUS PART
                greatest_movers(2, 0) = ticker_name ''Greatest total Volume  Ticker- BONUS PART
                'Bonus - if first value is not zero
                If open_value_yr <> 0 Then
                    If close_value_yr > open_value_yr Then
                        greatest_movers(0, 1) = (close_value_yr - open_value_yr) / open_value_yr
                    ElseIf close_value_yr < open_value_yr Then
                        greatest_movers(1, 1) = (close_value_yr - open_value_yr) / open_value_yr
                    End If
                 End If
                 greatest_movers(2, 1) = trade_volume ''Greatest total Volume Value - BONUS PART
            End If
        
         'STEP TWO - Entering results when moving to a new ticker - after first time run condition is met - RESET start point
            If last_ticker <> ticker_name Then
                ws.Range("I" & output_rownum) = last_ticker 'assigning TICKER on Column I
                ws.Range("J" & output_rownum) = close_value_yr - open_value_yr 'assigning YEARLY CHANGE on Column J
                If close_value_yr > open_value_yr Then
                    ws.Range("J" & output_rownum).Interior.Color = vbGreen 'color code green when positive value on Column J
                ElseIf close_value_yr < open_value_yr Then
                        ws.Range("J" & output_rownum).Interior.Color = vbRed 'color code red when negative value on Column J
                End If
            
               If open_value_yr = 0 Then 'in case of empty value Percentage Change/Column K
                    ws.Range("K" & output_rownum) = "N/A"
               Else ''Loop conditional RESET
                    ws.Range("K" & output_rownum) = (close_value_yr - open_value_yr) / open_value_yr 'Assign value for Percentage Change on Column K
                    ws.Range("K" & output_rownum).NumberFormat = "0.00%" 'format as Percentage value column K
                    If ((close_value_yr > open_value_yr) And (((close_value_yr - open_value_yr) / open_value_yr) > greatest_movers(0, 1))) Then ''BONUS PART
                        greatest_movers(0, 0) = last_ticker 'loop last ticker of greatest - BONUS PART
                        greatest_movers(0, 1) = (close_value_yr - open_value_yr) / open_value_yr 'loop last ticker value of greatest% increase  - BONUS PART
                    ElseIf ((close_value_yr < open_value_yr) And (((close_value_yr - open_value_yr) / open_value_yr) < greatest_movers(1, 1))) Then 'loop last ticker value of greatest% decrease -  - BONUS PART
                        greatest_movers(1, 0) = last_ticker
                        greatest_movers(1, 1) = (close_value_yr - open_value_yr) / open_value_yr
                    End If

               End If
                ws.Range("L" & output_rownum) = trade_volume_yr ' Assign Total Volume on Column L
                If trade_volume_yr > greatest_movers(2, 1) Then 'BONUS Greatest Total Volumne
                    greatest_movers(2, 0) = last_ticker
                    greatest_movers(2, 1) = trade_volume_yr
                End If
                output_rownum = output_rownum + 1 'Column I ticker next line switch variable
            
                open_value_yr = open_value
                last_ticker = ticker_name
                trade_volume_yr = trade_volume
                close_value_yr = close_value
        
            Else
                trade_volume_yr = trade_volume_yr + trade_volume ' Total Volume Calculations on Column L
                close_value_yr = close_value 'Switch Close value for Ticker
            End If
            row_num = row_num + 1 'Column A ticker next line switch
            
' '--------------------------------------------------------------------------------
' Populating Greatest value BONUS PART
' '--------------------------------------------------------------------------------
     
     Loop While ticker_name <> ""
     
     
        ws.Range("O2") = "Greatest % Increase"
        ws.Range("O3") = "Greatest % Decrease"
        ws.Range("O4") = "Greatest Total Volume"
        ws.Range("P1") = "Ticker"
        ws.Range("Q1") = "Value"
        ws.Range("P2") = greatest_movers(0, 0)
        ws.Range("P3") = greatest_movers(1, 0)
        ws.Range("P4") = greatest_movers(2, 0)
        ws.Range("Q2") = greatest_movers(0, 1)
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3") = greatest_movers(1, 1)
        ws.Range("Q3").NumberFormat = "0.00%"
        ws.Range("Q4") = greatest_movers(2, 1)
        ws.Cells.EntireColumn.AutoFit 'Format- Autofit columns
        
    Next ws 'Worksheet Calculation loop repeat until no worksheets
    
    
    
End Sub


