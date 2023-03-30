Attribute VB_Name = "Module1"
Sub VBChalleng()

For Each ws In Worksheets
  Dim WorksheetName As String
 'Current Raw
 Dim i As Long
 'start raw of ticker
 Dim j As Long
 'fill ticker raw
 Dim tickcount As Long
 'last Raw column A
 Dim lastRawA As Long
 'last row column I
 Dim LastRowI As Long
 'percent change calculation
 Dim PerChange As Double
 ' greatest increase calculation
 Dim GreatIncr As Double
 'greatest decrease calculation
 Dim GreatDecr As Double
 ' greatest total volume
 Dim GreatVol As Double
 
 
 'Go the WorksheetName
        WorksheetName = ws.Name
        
        'headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
 ' Ticker Counter to first row
    tickcount = 2
 'Set start row to 2
    j = 2

     LastRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row
 'MsgBox ("Last row in column A is " & LastRowA)
        
      'Loop through all rows
            For i = 2 To LastRowA
        'Check if ticker name changed
        
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        'Write ticker in column I (Coumn9)
        
            ws.Cells(tickcount, 9).Value = ws.Cells(i, 1).Value
                
                
        'Calculate and write Yearly Change in column J (Column10)
        
            ws.Cells(tickcount, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
                
 
   'Conditional formating for colors
          If ws.Cells(tickcount, 10).Value < 0 Then
   'Set cell background color to red
          ws.Cells(tickcount, 10).Interior.ColorIndex = 3
 
          Else
                
    'Set cell background color to green
        ws.Cells(tickcount, 10).Interior.ColorIndex = 4
                
         End If

  'Calculate and write percent change in column K (#11)
                    If ws.Cells(j, 3).Value <> 0 Then
                    PerChange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                    'Percent formating
                    ws.Cells(tickcount, 11).Value = Format(PerChange, "Percent")
                    
                    Else
                    
                    ws.Cells(tickcount, 11).Value = Format(0, "Percent")
                    
                    End If
 'Calculate and write total volume in column L (#12)
                ws.Cells(tickcount, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
                
                'Increase TickCount by 1
                tickcount = tickcount + 1
                
                'Set new start row of the ticker block
                j = i + 1
                
                End If
            
            Next i
            
            
            
 'Find last non-blank cell in column I
        LastRowI = ws.Cells(Rows.Count, 9).End(xlUp).Row
        'MsgBox ("Last row in column I is " & LastRowI)
        
 'Prepare for summary
    GreatVol = ws.Cells(2, 12).Value
    GreatIncr = ws.Cells(2, 11).Value
    GreatDecr = ws.Cells(2, 11).Value
        
 'Loop for Great values
    For i = 2 To LastRowI
            
 'For greatest total volume-check
    If ws.Cells(i, 12).Value > GreatVol Then
    GreatVol = ws.Cells(i, 12).Value
    ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                
    Else
                
    GreatVol = GreatVol
                
    End If
                
'check if next value is larger--if yes take over a new value
    If ws.Cells(i, 11).Value > GreatIncr Then
    GreatIncr = ws.Cells(i, 11).Value
    ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                
    Else
                
    GreatIncr = GreatIncr
                
    End If
                
    'check if next value is smaller
    If ws.Cells(i, 11).Value < GreatDecr Then
    GreatDecr = ws.Cells(i, 11).Value
    ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                
    Else
                
    GreatDecr = GreatDecr
                
    End If
                
    'Write  results in ws.Cells
     ws.Cells(2, 17).Value = Format(GreatIncr, "Percent")
     ws.Cells(3, 17).Value = Format(GreatDecr, "Percent")
     ws.Cells(4, 17).Value = Format(GreatVol, "Scientific")
            
     Next i
            
    'Djust column width automatically
    Worksheets(WorksheetName).Columns("A:Z").AutoFit
            
    Next ws
        
End Sub
