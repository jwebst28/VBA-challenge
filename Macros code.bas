Attribute VB_Name = "Module1"



Sub MultipleYearStockData():

    For Each ws In Worksheets
    
        Dim WorksheetName As String
        'Current row
        Dim i As Long
        'Start row of ticker block
        Dim j As Long
        'Index counter to fill Ticker row
        Dim TickCount As Long
        'Last row in column A
        Dim LastRowA As Long
        'last row column I
        Dim LastRowI As Long
        'Variable for percent change calculation
        Dim PerChange As Double
        'Variable for greatest increase calculation
        Dim GreatIncr As Double
        'Variable for greatest decrease calculation
        Dim GreatDecr As Double
        'greatest total volume
        Dim GreatVol As Double
        
        'worksheet name
        WorksheetName = ws.Name
        
        'Insert headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        'Set ticker counter to  second row
        TickCount = 2
        
        'Set start row to two
        j = 2
        
        'Find  last non-blank cell within column A
        LastRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row
        'MsgBox ("Last row in column A is " & LastRowA)
        
       'Loop through all rows
        For i = 2 To LastRowA
            
        'Check if ticker name changed
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
       'Write ticker in column I (#9)
       ws.Cells(TickCount, 9).Value = ws.Cells(i, 1).Value
                
       'Calculate and write Yearly Change in column J (#10)
        ws.Cells(TickCount, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
              
        'Conditional formating
         If ws.Cells(TickCount, 10).Value < 0 Then
                
        'Set cell background color to red
          ws.Cells(TickCount, 10).Interior.ColorIndex = 3
                
         Else
                
          'Set cell background color to green
           ws.Cells(TickCount, 10).Interior.ColorIndex = 4
                
          End If
                 
          'Calculate and write percent change in column K (#11)
           If ws.Cells(j, 3).Value <> 0 Then
           PerChange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                    
          'Percent formating
           ws.Cells(TickCount, 11).Value = Format(PerChange, "Percent")
                    
            Else
                    
            ws.Cells(TickCount, 11).Value = Format(0, "Percent")
                    
            End If
                    
           'Calculate and write total volume in column L (#12)
            ws.Cells(TickCount, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
                
            'crease TickCount by 1
            TickCount = TickCount + 1
                
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
        
            'Loop for summary
            For i = 2 To LastRowI
            
           'For greatest total volume--check if next value is larger--if yes take over a new value and populate ws.Cells
             If ws.Cells(i, 12).Value > GreatVol Then
        GreatVol = ws.Cells(i, 12).Value
        ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                
          Else
                
          GreatVol = GreatVol
                
           End If
                
         'For greatest increase--check if next value is larger--if yes take over a new value and populate ws.Cells
          If ws.Cells(i, 11).Value > GreatIncr Then
          GreatIncr = ws.Cells(i, 11).Value
          ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                
           Else
                
          GreatIncr = GreatIncr
                
           End If
                
        'For greatest decrease--check if next value is smaller--if yes take over a new value and populate ws.Cells
        If ws.Cells(i, 11).Value < GreatDecr Then
        GreatDecr = ws.Cells(i, 11).Value
         ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                
         Else
                
        GreatDecr = GreatDecr
                
        End If
                
            'Display all summary results in ws.Cells
            ws.Cells(2, 17).Value = Format(GreatIncr, "Percent")
            ws.Cells(3, 17).Value = Format(GreatDecr, "Percent")
            ws.Cells(4, 17).Value = Format(GreatVol, "Scientific")
            
            Next i

        
    Next ws
        
End Sub

