# VBA-Challenge

Code contributed by teacher, classmates, and in Stack Overflow incude:

'Determine LASTROW - code provided by Hunter Hollis in class
     LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

How to identify FirstOpenPrice provided by Tianyue Li
'If adjacent cells don't have same Ticker value
      If ws.Cells(r - 1, 1).Value <> ws.Cells(r, 1) Then
      Ticker = ws.Cells(r, 1).Value
      ws.Range("I" & SummaryTableRow).Value = Ticker
      FirstOpenPrice = ws.Cells(r, 3).Value

How to format numbers as percentages and color code them provided by Melissa Krachmer
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

Methods for how to identify the Greatest % Increase, Greatest % Decrease, and Greatest
Total Volume, plus their corresponding Ticker values, suggested in several Stack Overflow posts found at 
 https://stackoverflow.com/questions/51977446/vba-find-highest-value-in-a-column-c-and-return-its-value-and-the-adjacent-ce
https://stackoverflow.com/questions/52191966/finding-min-and-max-in-a-range-in-a-column-vba

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
     
      MaxInc = Application.WorksheetFunction.Max(ws.Range("K" & SummaryTableRow))
      MaxVol = Application.WorksheetFunction.Max(ws.Range("L" & SummaryTableRow))
      MaxNeg = Application.WorksheetFunction.Min(ws.Range("K" & SummaryTableRow))
       
     Set MaxIncCell = ws.Range("K" & SummaryTableRow).Find(MaxInc, Lookat:=xlWhole)
     Set MaxVolCell = ws.Range("L" & SummaryTableRow).Find(MaxVol, Lookat:=xlWhole)
     Set MaxNegCell = ws.Range("K" & SummaryTableRow).Find(MaxNeg, Lookat:=xlWhole)
