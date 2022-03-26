Sub stock_analysis():
    ' Set Dimensions
    Dim ws As Worksheet
    Dim RowCounta, RowCountb, Start As Long
    Dim Ticker As Integer
    Dim TotalVolume, Change, PercentChange As Double

    ' Dimensions for Extention Table
    Dim MaxIncrease, MaxDecrease As Double
    Dim MaxIncreaseTicker, MaxDecreaseTicker, MaxVolumeTicker As String
    Dim MaxVolume As LongLong
    
    ' Creating a loop for all worksheets
    For Each ws In Worksheets
    

'Sort Data by Ticker, then by Date
    Columns.Sort Key1:=Columns("A"), Order1:=xlAscending, key2:=Columns("B"), Order2:=xlAscending, Header:=xlYes
        
' Part 1 - Creating Tables and Adding Data into First Table
' -----------------------------------

    ' Set Title Row for Table
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    
      
    ' Set initial values
    Ticker = 0
    TotalVolume = 0
    Change = 0
    Start = 2
    
    ' Get the Total row count and create the loop to the last row with data
    RowCounta = ws.Cells(Rows.Count, "A").End(xlUp).Row
    For i = 2 To RowCounta
        
        ' If ticker changes then print results
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
        ' Stores results in variables
        TotalVolume = TotalVolume + ws.Cells(i, 7).Value
        ' Handle a ticker with zero total volume
            If TotalVolume = 0 Then
            
            ' Print the results (adding the ticker and default values to columns i to l)
            ws.Range("I" & 2 + Ticker).Value = Cells(i, 1).Value
            ws.Range("J" & 2 + Ticker).Value = 0
            ws.Range("K" & 2 + Ticker).Value = "%" & 0
            ws.Range("L" & 2 + Ticker).Value = 0
            
            Else
            ' Find the first non-zero total volume value and add to temp table
                If ws.Cells(Start, 3) = 0 Then
                    For find_value = Start To i
                        If ws.Cells(find_value, 3).Value <> 0 Then
                            Start = find_value
                            Exit For
                        End If
                    Next find_value
                End If
               
                ' Calculate the change for the current ticker value and add to second column of temp table
                Change = (ws.Cells(i, 6) - ws.Cells(Start, 3))
                PercentChange = Round((Change / ws.Cells(Start, 3) * 100), 2)
                
                ' Start of the next stock ticker
                Start = i + 1
                
                ' Print the results to the summary table
                ws.Range("I" & 2 + Ticker).Value = ws.Cells(i, 1).Value
                ws.Range("J" & 2 + Ticker).Value = Round(Change, 2)
                ws.Range("K" & 2 + Ticker).Value = "%" & PercentChange
                ws.Range("L" & 2 + Ticker).Value = TotalVolume
                
                ' Formatting of Percentage Change Data - Positives green and negatives red
                Select Case Change
                    Case Is > 0
                        ws.Range("J" & 2 + Ticker).Interior.ColorIndex = 4
                    Case Is < 0
                        ws.Range("J" & 2 + Ticker).Interior.ColorIndex = 3
                    Case Else
                        ws.Range("J" & 2 + Ticker).Interior.ColorIndex = 0
                End Select
            
            End If
            
            ' Reset the variables for each new stock Ticker
            TotalVolume = 0
            Change = 0
            Ticker = Ticker + 1
            
            ' If ticker is still the same add results
            Else
                TotalVolume = TotalVolume + ws.Cells(i, 7).Value

       End If
    
    Next i
    
' Part 2 - Adding to the second table
' -----------------------------------

' Setting Default Values of variables for Bonus Table
MaxIncrease = 0
MaxDecrease = 0
MaxVolume = 0
RowCountb = 0

' a - Get the Total row count and create the loop to the last row with data
    RowCountb = ws.Cells(Rows.Count, "K").End(xlUp).Row
    For j = 2 To RowCountb

    ' Identify the Greatest % Increase - Identify the greatest change row by row, if larger than previous - update variable and overwrite existing value
        If ws.Cells(j + 1, 11).Value > MaxIncrease Then
            MaxIncrease = ws.Cells(j + 1, 11).Value
            ws.Range("Q2").Value = "%" & MaxIncrease * 100
           
            MaxIncreaseTicker = ws.Cells(j + 1, 9).Value
            ws.Range("P2").Value = MaxIncreaseTicker
       End If
        
    Next j

' b - Get the Total row count and create the loop to the last row with data
    RowCountb = ws.Cells(Rows.Count, "K").End(xlUp).Row
    For k = 2 To RowCountb
    
    ' Identify the Greatest % Decrease - Identify the greatest change row by row, if larger than previous - update variable and overwrite existing value
        
        If ws.Cells(k + 1, 11).Value < MaxDecrease Then
            MaxDecrease = ws.Cells(k + 1, 11).Value
            ws.Range("Q3").Value = "%" & MaxDecrease * 100
        
            MaxDecreaseTicker = ws.Cells(k + 1, 9).Value
            ws.Range("P3").Value = MaxDecreaseTicker
        End If
       
    Next k

' c - Get the Total row count and create the loop to the last row with data
    RowCountb = ws.Cells(Rows.Count, "K").End(xlUp).Row
    For m = 2 To RowCountb
    
    ' Identify the Greatest Volume - Identify the greatest change row by row, if larger than previous - update variable and overwrite existing value
    
        If ws.Cells(m + 1, 12).Value > MaxVolume Then
            MaxVolume = ws.Cells(m + 1, 12).Value
            ws.Range("Q4").Value = MaxVolume
          
            MaxVolumeTicker = ws.Cells(m + 1, 9).Value
            ws.Range("P4").Value = MaxVolumeTicker
            
        End If
               
    Next m
    
Next ws
    
End Sub
