Attribute VB_Name = "Module1"
Sub Yearly_Stock_Info():

    Dim Current As Worksheet
    
    Dim data_row, total_row As Integer
    Dim last_row As Long
    Dim last_total_row As Long
    
    Dim year_open As Double
    Dim total_volume As Double
    
    Dim greatest_ticker_increase As String
    Dim greatest_increase As Double
    Dim greatest_ticker_decrease As String
    Dim greatest_decrease As Double
    Dim greatest_ticker_volume As String
    Dim greatest_volume As Double

    'Loop through all of the worksheets
    For Each Current In Worksheets
    
        MsgBox ("Current Worksheet = " + Current.Name)
    
        'Set last stock data row
        last_row = Current.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Initiate total and greatest values
        total_row = 1
        greatest_increase = 0
        greatest_decrease = 0
        greatest_volume = 0
        
        'Create new total columns with headers
        Current.Range("I1").Value = "Ticker"
        Current.Range("J1").Value = "Yearly Change"
        Current.Range("K1").Value = "Percent Change"
        Current.Range("L1").Value = "Total Stock Volume"
        
        'Create new greatest columns with lablels
        Current.Range("O1").Value = "Ticker"
        Current.Range("P1").Value = "Value"
        Current.Range("N2").Value = "Greatest % Increase"
        Current.Range("N3").Value = "Greatest % Decrease"
        Current.Range("N4").Value = "Greatest Total Volume"
             
        'Loop through each row
        For data_row = 2 To last_row
            
            'Save Opening price for Ticker at the beginning of the year
            If Current.Cells(data_row - 1, 1).Value <> Current.Cells(data_row, 1).Value Then
                year_open = Current.Range("C" & data_row).Value
            End If
            
            'Check for new Ticker
            If Current.Cells(data_row + 1, 1).Value <> Current.Cells(data_row, 1).Value Then
                   
                'Increment the row for totals
                total_row = total_row + 1
                
                'Gather total information
                year_close = Current.Range("F" & data_row).Value
                total_volume = total_volume + Current.Range("G" & data_row).Value

                yearly_change = year_close - year_open
                'Check for zero
                If year_open = 0 Then
                    percent_change = 0
                Else
                    percent_change = yearly_change / year_open
                End If
                
                'Insert total information
                Current.Range("I" & total_row).Value = Current.Range("A" & data_row).Value
                Current.Range("J" & total_row).Value = yearly_change
                Current.Range("K" & total_row).Value = percent_change
                Current.Range("K" & total_row).NumberFormat = "0.00%"
                Current.Range("L" & total_row).Value = total_volume
                
                'Format yearly_change
                If yearly_change >= 0 Then
                    Current.Range("J" & total_row).Interior.Color = RGB(0, 255, 0)
                Else
                    Current.Range("J" & total_row).Interior.Color = RGB(255, 0, 0)
                End If
                    
                'Reset total stock volume
                total_volume = 0
         
            Else
                                  
                total_volume = total_volume + Current.Range("G" & data_row).Value
            
            End If
            
        Next data_row
                  
        'Check for Greatest total values
        For data_row = 2 To total_row
        
            'Hold greatest values
            If Current.Cells(data_row, 11).Value > greatest_increase Then
                greatest_ticker_increase = Current.Cells(data_row, 9).Value
                greatest_increase = Current.Cells(data_row, 11).Value
            End If
            
            If Current.Cells(data_row, 11).Value < greatest_decrease Then
                greatest_ticker_decrease = Current.Cells(data_row, 9).Value
                greatest_decrease = Current.Cells(data_row, 11).Value
            End If
            
            If Current.Cells(data_row, 12).Value > greatest_volume Then
                greatest_ticker_volume = Current.Cells(data_row, 9).Value
                greatest_volume = Current.Cells(data_row, 12).Value
            End If
                    
        Next data_row
        
        'Insert greatest information
        Current.Range("O2").Value = greatest_ticker_increase
        Current.Range("P2").Value = greatest_increase
        Current.Range("P2").NumberFormat = "0.00%"
        Current.Range("O3").Value = greatest_ticker_decrease
        Current.Range("P3").Value = greatest_decrease
        Current.Range("P3").NumberFormat = "0.00%"
        Current.Range("O4").Value = greatest_ticker_volume
        Current.Range("P4").Value = greatest_volume

    Next

End Sub

