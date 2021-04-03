Sub TickerAnalysis1()

  ' Set an initial variable for holding the ticker symbol
  Dim Ticker As String

  ' Set an initial variable for holding the total volume per year
  Dim Volume As Single

  ' Keep track of the location for each Ticker in the summary table
  Dim Summary_Table_Row As Single
  
  ' Set an initial variable to mark the row representing 1st Jan of a Ticker
  Dim RowFirstJan As Single
  
  ' Set an initial variable to mark the row representing 31st Dec of a Ticker
  Dim RowEndDecember As Single

  ' Set an inital variable for the % price change over the year
  Dim PriceChangePercent As Double
  
   ' Set an inital variable for the % price change over the year
  Dim PriceChange As Single
  
  ' Set an initial object variable as the current worksheet
  Dim ws As Worksheet
  
  'Set initial variables as the greatest percent increase
  Dim GreatestPriceIncreasePercent As Double
  Dim GreatestPriceIncreaseTicker As String
  
  'Set initial variables as the greatest percent decrease
  Dim GreatestPriceDecreasePercent As Double
  Dim GreatestPriceDecreaseTicker As String
  
  'Set an initial variable as the greatest volume
  Dim GreatestVolume As Single
  Dim GreatestVolumeTicker As String
  
  'Start with the first worksheet in this workbook
  Set ws = ThisWorkbook.Worksheets(1)
  
  'Create a loop through each worksheet
  For Each ws In Worksheets
  
  'Give initial values to variables (reset each Worksheet)
  Summary_Table_Row = 2
  RowFirstJan = 2
  RowEndDecember = 2
  PriceChange = 0
  PriceChangePercent = 0
  GreatestVolume = 0
  GreatestPriceDecreasePercent = 0
  GreatestPriceIncreasePercent = 0
  
      ' Loop through all Tickers and dates
      For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
      
        ' Check if still within the same Ticker, if not we're at the end of that Ticker series
        ' Subsequent actions update the summary table
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
          ' Update variable for last row (end of December) for current Ticker
          RowEndDecember = i
          
          ' Set the Ticker
          Ticker = ws.Cells(i, 1).Value
    
          ' Add to the Volume
          Volume = Volume + ws.Cells(i, 7)
          
          'Update the Greatest Volume
          If Volume > GreatestVolume Then
            GreatestVolume = Volume
            GreatestVolumeTicker = Ticker
            Else: GreatestVolume = GreatestVolume
          End If
          
          ' Calculate the Price Change
          PriceChange = ws.Cells(RowEndDecember, 6).Value - ws.Cells(RowFirstJan, 3).Value
          
          'Calculate Price Change Percent over the year unless divisor is 0
          If ws.Cells(RowFirstJan, 3).Value = 0 Then
            PriceChangePercent = 0
            Else: PriceChangePercent = (ws.Cells(RowEndDecember, 6).Value / ws.Cells(RowFirstJan, 3).Value) - 1
          End If
          
          'Update the GreatestPriceIncreasePercent
          If PriceChangePercent > GreatestPriceIncreasePercent Then
            GreatestPriceIncreasePercent = PriceChangePercent
            GreatestPriceIncreaseTicker = Ticker
            Else: GreatestPriceIncreasePercent = GreatestPriceIncreasePercent
          End If
          
          'Update the GreatestPriceDecreasePercent
          If PriceChangePercent < GreatestPriceDecreasePercent Then
            GreatestPriceDecreasePercent = PriceChangePercent
            GreatestPriceDecreaseTicker = Ticker
            Else: GreatestPriceDecreasePercent = GreatestPriceDecreasePercent
          End If
          
          'Print Results in the Summary Table
    
          'Print the Ticker in the Summary Table
          ws.Range("J" & Summary_Table_Row).Value = Ticker
          
          ' Print the Price Change in the Summary Table and color format based on value
          ws.Range("K" & Summary_Table_Row).Value = Format(PriceChange, "#.00")
          If PriceChange > 0 Then
            ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
            Else: ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
          End If
          
          'Print the Price Change Percent as a Percentage
          ws.Range("L" & Summary_Table_Row).Value = FormatPercent(PriceChangePercent, 2)
          
          ' Print the Volume total to the Summary Table
          ws.Range("M" & Summary_Table_Row).Value = Volume
          
          ' The following lines of code were used during development
          ' ws.Range("M" & Summary_Table_Row).Value = ws.Cells(RowEndDecember, 6).Value
          ' ws.Range("N" & Summary_Table_Row).Value = RowEndDecember
          ' ws.Range("O" & Summary_Table_Row).Value = ws.Cells(RowFirstJan, 3).Value
          ' ws.Range("P" & Summary_Table_Row).Value = RowFirstJan
    
          ' Add one to the summary table row for next Ticker
          Summary_Table_Row = Summary_Table_Row + 1
          
          ' Reset the Volume for next Ticker
          Volume = 0
                  
          'Reset RowFirstJan row marker for next Ticker
          RowFirstJan = i + 1
          
    
        ' If the cell immediately following a row is the same Ticker (just update the volume total)
        Else
          Volume = Volume + ws.Cells(i, 7).Value
             
        End If
    
      'Move to the next row of the Worksheet
      Next i
      
      'Add Titles to Summary Table
      ws.Range("J1").Value = "Ticker"
      ws.Range("K1").Value = "Yearly Change"
      ws.Range("L1").Value = "Percent Change"
      ws.Range("M1").Value = "Total Stock Volume"
      
      'Print to the Greatest values table
      ws.Range("Q1").Value = "Ticker"
      ws.Range("R1").Value = "Value"
      ws.Range("P2").Value = "Greatest % Increase"
      ws.Range("P3").Value = "Greatest % Decrease"
      ws.Range("P4").Value = "Greatest Total Volume"
      
      ws.Range("Q2").Value = GreatestPriceIncreaseTicker
      ws.Range("Q3").Value = GreatestPriceDecreaseTicker
      ws.Range("Q4").Value = GreatestVolumeTicker
      
      ws.Range("R2").Value = FormatPercent(GreatestPriceIncreasePercent, 2)
      ws.Range("R3").Value = FormatPercent(GreatestPriceDecreasePercent, 2)
      ws.Range("R4").Value = GreatestVolume
      
'Move to the next Worksheet
Next ws

End Sub


