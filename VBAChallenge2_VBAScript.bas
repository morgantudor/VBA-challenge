Attribute VB_Name = "Module1"
Sub Stocks()
  For Each ws In Worksheets
    
    'Set variables
    Dim ticker As String
    Dim stockvolume As Double
     stockvolume = 0
    Dim stockopen As Double
    stockopen = ws.Cells(2, 3).Value
    Dim stockclose As Double
    Dim yearlychange As Double
    Dim percentchange As Double
    Dim summarytablerow As Integer
     summarytablerow = 2
    Dim highestincrease As Double
      highestincrease = 0
    Dim highestdecrease As Double
      highestdecrease = 0
    Dim highestvolume As Double
      highestvolume = 0
    
    'last row function
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'create new headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    'create small summary table
     ws.Cells(1, 15).Value = "Ticker"
     ws.Cells(1, 16).Value = "Value"
     ws.Cells(2, 14).Value = "Greatest % Increase"
     ws.Cells(3, 14).Value = "Greatest % Decrease"
     ws.Cells(4, 14).Value = "Greatest Total Volume"
    
    'loop functions
    For i = 2 To lastrow
     
     'ticker
     If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
      ticker = ws.Cells(i, 1).Value
      ws.Range("I" & summarytablerow).Value = ticker
     
      'volume
      stockvolume = stockvolume + ws.Cells(i, 7).Value
      ws.Range("L" & summarytablerow).Value = stockvolume
     
      'yearly change
      stockclose = ws.Cells(i, 6).Value
      yearlychange = stockclose - stockopen
      ws.Range("J" & summarytablerow).Value = yearlychange
     
     'percent change
      percentchange = (yearlychange / stockopen)
      ws.Range("K" & summarytablerow).Value = percentchange
      ws.Cells(summarytablerow, 11).NumberFormat = "0.00%"
    
      
     'fill greatest increase
      If ws.Cells(summarytablerow, 11).Value > highestincrease Then
       highestincrease = ws.Cells(summarytablerow, 11).Value
       ws.Cells(2, 16).Value = highestincrease
       ws.Cells(2, 15).Value = ticker
       ws.Cells(2, 16).NumberFormat = "0.00%"
      End If
    
     'fill greatest decrease
      If ws.Cells(summarytablerow, 11).Value < highestdecrease Then
        highestdecrease = ws.Cells(summarytablerow, 11).Value
        ws.Cells(3, 16).Value = highestdecrease
        ws.Cells(3, 15).Value = ticker
        ws.Cells(3, 16).NumberFormat = "0.00%"
       End If
     
     'fill highest volume
      If ws.Cells(summarytablerow, 12).Value > highestvolume Then
        highestvolume = ws.Cells(summarytablerow, 12).Value
        ws.Cells(4, 16).Value = highestvolume
        ws.Cells(4, 15).Value = ticker
      End If
      
      'add another row
      summarytablerow = summarytablerow + 1
      stockvolume = 0
      stockopen = ws.Cells(i + 1, 3).Value
     
     Else
      stockvolume = stockvolume + ws.Cells(i, 7).Value
      
     End If
     
     'Color coding yearly change
     If ws.Cells(i, 10).Value > 0 Then
      ws.Cells(i, 10).Interior.ColorIndex = 4
     
     ElseIf ws.Cells(i, 10).Value = 0 Then
      ws.Cells(i, 10).Interior.ColorIndex = 0
      
     Else
      ws.Cells(i, 10).Interior.ColorIndex = 3
     End If

    Next i
 Next ws
End Sub


