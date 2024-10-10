Attribute VB_Name = "Module2"
Sub ticker()

'Applying the code to each worksheet

Dim Current As Worksheet

For Each Current In Worksheets

Current.Activate
 
 'Defining variables and putting in the starting value if relevant
 
 Dim ticker As String
 Dim Summary As Integer
 Summary = 2
 Dim Volume As Double
 Volume = 0
 Dim Open_Value As Double
 Open_Value = Cells(2, 3).Value
 Dim Close_Value As Double
 Dim quarterly_change As Double
 Dim percent_change As Double
 lastrow = Cells(Rows.Count, 1).End(xlUp).Row
 
 'put in the headers for the summary table
 
 Range("I1").Value = "Ticker"
 Range("j1").Value = "Quarterly Change"
 Range("k1").Value = "Percentage Change"
 Range("l1").Value = "Total Stock Volume"
 Range("P1").Value = "Ticker"
 Range("Q1").Value = "Value"
 Range("O2").Value = "Greatest % Increase"
 Range("O3").Value = "Greatest % Decrease"
 Range("O4").Value = "Greatest Total Volume"
 
 
'starting For loop
 For i = 2 To lastrow
 
    'comparing the tickers to get to the last row for each ticker
    
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
   'Finding values
   
    ticker = Cells(i, 1).Value
    
    Volume = Volume + Cells(i, 7).Value
    
    Close_Value = Cells(i, 6).Value
    
    quarterly_change = (Close_Value - Open_Value)
    
    percent_change = ((Close_Value - Open_Value) / Open_Value)
    
    'Conditional formatting the percentages
    
    Range("K" & Summary).NumberFormat = "0.00%"
    
    'Put values in a table
    
    Range("I" & Summary).Value = ticker
    
    Range("L" & Summary).Value = Volume
    
    Range("J" & Summary).Value = quarterly_change
    
    Range("K" & Summary).Value = percent_change
    
        'Conditional formatting for the percentages
        
        If quarterly_change > 0 Then
    
        Range("J" & Summary).Interior.ColorIndex = 4
    
        ElseIf quarterly_change < 0 Then
        
        Range("J" & Summary).Interior.ColorIndex = 3
        
        End If
    

    'reset values
    
    Summary = Summary + 1
    
    Volume = 0
    
    Open_Value = Cells(i + 1, 3).Value
    
    
Else
    
    'if the tickers are the same then, add the volume together
    
    Volume = Volume + Cells(i, 7).Value
    

End If

Next i
 
 MsgBox Current.Name
  
 
        ' Take the max and min and place them in a separate part in the worksheet
        Range("Q2") = "%" & WorksheetFunction.Max(Range("K2:K" & lastrow)) * 100
        Range("Q3") = "%" & WorksheetFunction.Min(Range("K2:K" & lastrow)) * 100
        Range("Q4") = WorksheetFunction.Max(Range("L2:L" & lastrow))

        ' Returns one less because header row not a factor
        increase_number = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & lastrow)), Range("K2:K" & lastrow), 0)
        decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & lastrow)), Range("K2:K" & lastrow), 0)
        volume_number = WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" & lastrow)), Range("L2:L" & lastrow), 0)

        ' Final ticker symbol for total, greatest % of increase and decrease, and average
        Range("P2") = Cells(increase_number + 1, 9)
        Range("P3") = Cells(decrease_number + 1, 9)
        Range("P4") = Cells(volume_number + 1, 9)
 
 
 Next
 
 
 
 
 End Sub
