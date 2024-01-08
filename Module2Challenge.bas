Attribute VB_Name = "Module2"
Sub Challenge02():

' LOOP THROUGH ALL SHEETS

For Each ws In Worksheets

    ' Set up Summary Table
 
    'Create variables for summary table columns
    Dim Ticker As String 'Put in Column 9
    Dim YearlyChange As Double 'Put in Column 10
    Dim PercentChange As Double 'Put in Column 11
    Dim TotalStockVolume As LongLong 'Put in Column 12
    Dim LastClosePrice As Double
    Dim FirstOpenPrice As Double
    Dim LastRow As Long
    
  
    'Add header labels to columns in summary table
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "YearlyChange"
    ws.Cells(1, 11).Value = "PercentChange"
    ws.Cells(1, 12).Value = "TotalStockVolume"
    'ws.Cells(1, 13).Value = "FirstOpenPrice"
    'ws.Cells(1, 14).Value = "LastClosePrice"

    'Set up Greatest table headers
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
         
    'Set TotalStock Volume, YearlyChange, and PercentChange initially to zero
    TotalStockVolume = 0
    YearlyChange = 0
    PercentChange = 0

    'Set SummaryTableRow as an integer to start at 2
     Dim SummaryTableRow As Integer
     SummaryTableRow = 2

    'Declare r (for row) as Long, for the daily stock results table
     Dim r As Long

    'Determine LASTROW - code provided by Hunter Hollis in class
     LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
                       
     'Set Greatest % Increase, Decrease, and Volume types
     Dim GreatIncrease As Double
     Dim GreatDecrease As Double
     Dim GreatVolume As LongLong
     
     'Had to define next three variables to get Greatest Table to populate
     Dim GreatIncreaseCell As Range
     Dim GreatDecreaseCell As Range
     Dim GreatVolumeCell As Range
     
     
       'Loop through the rows, looking for a change in Ticker
        For r = 2 To LastRow
        
        'FIRST CASE SCENARIO - Identify FirstOpenPrice
 
            'If adjacent cells don't have same Ticker value
            If ws.Cells(r - 1, 1).Value <> ws.Cells(r, 1) Then
             
            'Set the Ticker value
            Ticker = ws.Cells(r, 1).Value
             
            'Write Ticker value to Summary Table
            ws.Range("I" & SummaryTableRow).Value = Ticker
             
             'Define and Set FirstOpenPrice
              'Dim FirstOpenPrice As Double
              FirstOpenPrice = ws.Cells(r, 3).Value
              
              'Write FirstOpenPrice to Summary Table
              'ws.Range("M" & SummaryTableRow).Value = FirstOpenPrice
              
             'And add up the total stock volume (column 7)
              TotalStockVolume = TotalStockVolume + ws.Cells(r, 7).Value
             
             'Calculate Yearly Change and Percent Change values for each Ticker
          
             'Define Yearly Changes
              YearlyChange = LastClosePrice - FirstOpenPrice
     
              'Define PercentChange
               PercentChange = (LastClosePrice - FirstOpenPrice) / FirstOpenPrice
    
               'Write YearlyChange and PercentChange values in Summary Table
               ws.Range("J" & SummaryTableRow).Value = YearlyChange
               ws.Range("K" & SummaryTableRow).Value = PercentChange

        '------Color coding cells and Formatting Percent Change ------------------------
 
             'Color code Yearly and % Change cells: green if positive, red if negative
                If PercentChange >= 0 Then
                 ws.Range("K" & SummaryTableRow).Interior.ColorIndex = 4
                 Else: ws.Range("K" & SummaryTableRow).Interior.ColorIndex = 3
                 End If
   
                 If YearlyChange >= 0 Then
                  ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 4
                Else: ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 3
                End If
                
                'Formatting for % Change values
                ws.Range("K" & SummaryTableRow).Value = FormatPercent(ws.Range("K" & SummaryTableRow), 2)

      '--------------------------------------------------------------------------------------
     
     'SECOND CASE SCENARIO - Identify LastClosePrice
     
            ElseIf ws.Cells(r + 1, 1).Value <> ws.Cells(r, 1).Value Then
         
            'Set the Ticker value
             Ticker = ws.Cells(r, 1).Value
                    
            'Write the Ticker to the summary table
            ws.Range("I" & SummaryTableRow).Value = Ticker
             
             'Find LastClosePrice
             LastClosePrice = ws.Cells(r, 6).Value
              
            'Write LastClosePrice to Summary Table
             'ws.Range("N" & SummaryTableRow).Value = LastClosePrice
              
             'Add up the total stock volume (column 7)
             TotalStockVolume = TotalStockVolume + ws.Cells(r, 7).Value
             
             'Write the TotalStockVolume to the Summary Table
             ws.Range("L" & SummaryTableRow).Value = TotalStockVolume
                                                  
            'Calculate Yearly Change and Percent Change values for each Ticker
   
            'Define Yearly Changes
            YearlyChange = LastClosePrice - FirstOpenPrice
     
            'Define PercentChange
             PercentChange = (LastClosePrice - FirstOpenPrice) / FirstOpenPrice
    
            'Write YearlyChange and PercentChange values to Summary Table
             ws.Range("J" & SummaryTableRow).Value = YearlyChange
             ws.Range("K" & SummaryTableRow).Value = PercentChange
    
      '------Color coding cells and Formatting Percent Change ------------------------
 
           'Color code cells green if positive, red if negative for YearlyChange
             If PercentChange >= 0 Then
             ws.Range("K" & SummaryTableRow).Interior.ColorIndex = 4
             Else: ws.Range("K" & SummaryTableRow).Interior.ColorIndex = 3
            End If
   
             If YearlyChange >= 0 Then
             ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 4
             Else: ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 3
            End If
            
            'Formatting for % Change values
            ws.Range("K" & SummaryTableRow).Value = FormatPercent(ws.Range("K" & SummaryTableRow), 2)
        
        '----------------------------------------------------------------------
      
            'And add one to the SummaryTableRow value
            SummaryTableRow = SummaryTableRow + 1

            'And reset TotalStockVolume to zero for new value of r
            TotalStockVolume = 0
                                                      
     
 '--------THIRD CASE SCENARIO -If adjacent row values are the same-------------------------
        
         'If cell immediately following row has same ticker,
         'then add row's volume value to TickerTotalVolume
         'ElseIf ws.Cells(r + 1, 1).Value = ws.Cells(r, 1).Value Then
         Else
         TotalStockVolume = TotalStockVolume + ws.Cells(r, 7).Value
        
End If
 
Next r
             
 Next ws
           
           
    '-----Finding Greatest % Increase, Greatest % Decrease, and Greatest Total Volume-----------------------
    'https://stackoverflow.com/questions/51977446/vba-find-highest-value-in-a-column-c-and-return-its-value-and-the-adjacent-ce
    'https://stackoverflow.com/questions/52191966/finding-min-and-max-in-a-range-in-a-column-vba
    
    'Set Greatest % Increase, Decrease, and Volume types
      
     Dim MaxInc As Double
     Dim MaxVol As LongLong
     Dim MaxNeg As Double
     
     Dim i As Range
     Dim v As Range
     
     Set i = Range("K" & SummaryTableRow)
     Set v = Range("L" & SummaryTableRow)
     
     MaxInc = Application.WorksheetFunction.Max(i)
     MaxNeg = Application.WorksheetFunction.Min(i)
     MaxVol = Application.WorksheetFunction.Max(v)
     
     
     ''''''This approach worked  in the test file but I can't get the variable values to print to the specified cells (or ranges)
      'MaxInc = Application.WorksheetFunction.Max(ws.Range("K" & SummaryTableRow))
      'MaxVol = Application.WorksheetFunction.Max(ws.Range("L" & SummaryTableRow))
      'MaxNeg = Application.WorksheetFunction.Min(ws.Range("K" & SummaryTableRow))
       
     'Set MaxIncCell = ws.Range("K" & SummaryTableRow).Find(MaxInc, Lookat:=xlWhole)
     'Set MaxVolCell = ws.Range("L" & SummaryTableRow).Find(MaxVol, Lookat:=xlWhole)
     'Set MaxNegCell = ws.Range("K" & SummaryTableRow).Find(MaxNeg, Lookat:=xlWhole)
     
     
     '''''''Getting object errors for ws.Cells= and ws.Range = statements
     'ws.Cells(2, 17).Value = MaxInc
     'ws.Range("Q2").Value = MaxInc
     'ws.Range("P2").Value = GreatIncreaseCell.Offset(, -2) 'This code doesn't work
    
     'ws.Cells(4, 17).Value = MaxVol
     'ws.Range("Q4").Value = MaxVol
     'ws.Range("P4").Value = GreatVolumeCell.Offset(, -3)   'This code doesn't work
     
     'ws.Cells(3, 17).Value = MaxNeg
     'ws.Range("Q3").Value = MaxNeg
    
    '-----------------------------------------------------------------
    
  
 End Sub
