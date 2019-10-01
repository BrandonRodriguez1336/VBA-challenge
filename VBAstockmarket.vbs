Attribute VB_Name = "Module1"
Sub VBAstockmarket()

' Loop through all sheets
Dim ws As Worksheet

For Each ws In Worksheets

   

  ' create variables and headers
  
  Dim openT As Double
  Dim closeT As Double
  Dim lastrow As Long
  
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  
  openT = ws.Cells(2, 3).Value
  
  ' Set an initial variable for holding the total Stock Volume
  Dim TotalStockVol As Double
  TotalStockVol = 0
    
    
  ' Keep track of the location for Stock in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

  ' Loop through all Stock transactions
  For i = 2 To lastrow
   
    ' Check if we are still within the same Stock Ticker, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
      ' Set the Close stock value
       closeT = ws.Cells(i, 6)
      ' Add to the Total Stock Volume
      TotalStockVol = TotalStockVol + Cells(i, 7).Value

      ' Print the Ticker in the Summary Table
      ws.Range("I" & Summary_Table_Row).Value = Cells(i, 1).Value
      ws.Range("J" & Summary_Table_Row).Value = (closeT - openT)
      'format the yearly change to allow 8 decimal places
      ws.Range("J" & Summary_Table_Row).NumberFormat = "0.00000000"
      'to handle stocks begining with zero '0'
      If openT = 0 Then
      ws.Range("K" & Summary_Table_Row).Value = (closeT - openT)
      Else: ws.Range("K" & Summary_Table_Row).Value = ((closeT - openT) / openT)
      End If
      
      
      ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
      'Update the interior color in column K if positive or negative
      If ws.Range("J" & Summary_Table_Row).Value > 0 Then
      ' set the cell color to green
      ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
      ' set the cell color to red
      Else: ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
   
      
      End If
      
      
      ' Print the Stock Ticker Volume to the Summary Table
      ws.Range("L" & Summary_Table_Row).Value = TotalStockVol

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Total Stock Volume
      TotalStockVol = 0
       openT = ws.Cells(i + 1, 3).Value
    ' If the cell immediately following a row is the same Ticker...
    Else

      ' Add to the Brand Total
      TotalStockVol = TotalStockVol + Cells(i, 7).Value

    End If
    
  Next i
  
'Begin to calculate the greatest increase, greatest decrease, greatest total volume
ws.Cells(2, 14).Value = "Greatest % Increase"
ws.Cells(3, 14).Value = "Greatest % Decrease"
ws.Cells(4, 14).Value = "Greatest Total Volume"
ws.Cells(1, 15).Value = "Ticker"
ws.Cells(1, 16).Value = "Value"

Dim GreatestPcntIncrease As Double
Dim GreatestPcntDecrease As Double
Dim GreatestTotalVolume As Currency
Dim ticker(2) As String

GreatestTotalVolume = 0
'set last row to the ticker table the previous loop built
lastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row

  'set the percent change value to the greatest increase, greatest decrease, greatest volume to perform comparisons
   GreatestPcntIncrease = ws.Cells(2, 11).Value
   GreatestPcntDecrease = ws.Cells(2, 11).Value
   GreatestTotalVolume = ws.Cells(2, 12).Value
   ticker(0) = ws.Cells(2, 9).Value
   ticker(1) = ws.Cells(2, 9).Value
   ticker(2) = ws.Cells(2, 9).Value
   
' Loop through all Stock transactions
  For i = 2 To lastrow

   
    ' Check if the current percent change is greater than the value in the next cell
    If ws.Cells(i + 1, 11).Value > GreatestPcntIncrease Then
        GreatestPcntIncrease = ws.Cells(i + 1, 11).Value
        ticker(0) = ws.Cells(i + 1, 9).Value
    End If
    ' check if the current percent change is less than the value in the next cell
     If ws.Cells(i + 1, 11).Value < GreatestPcntDecrease Then
        GreatestPcntDecrease = ws.Cells(i + 1, 11).Value
        ticker(1) = ws.Cells(i + 1, 9).Value
    End If
    ' check if the current total stock volume is the greatest value
    If ws.Cells(i + 1, 12).Value > GreatestTotalVolume Then
        GreatestTotalVolume = ws.Cells(i + 1, 12).Value
        ticker(2) = ws.Cells(i + 1, 9).Value
    End If

    
    
  Next i
  
  ws.Cells(2, 15).Value = ticker(0)
  ws.Cells(2, 16).Value = GreatestPcntIncrease
  ws.Cells(2, 16).NumberFormat = "0.00%"
  
  ws.Cells(3, 15).Value = ticker(1)
  ws.Cells(3, 16).Value = GreatestPcntDecrease
  ws.Cells(3, 16).NumberFormat = "0.00%"
   
  ws.Cells(4, 15).Value = ticker(2)
  ws.Cells(4, 16).Value = GreatestTotalVolume
  ws.Cells(4, 16).NumberFormat = "0"
  
ws.Columns("I:P").AutoFit




Next ws

End Sub

