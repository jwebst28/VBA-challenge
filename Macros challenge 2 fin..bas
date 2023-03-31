Attribute VB_Name = "Module1"
Sub MultipleYearStockData():

    For Each ws In Worksheets
    
       'Column headers
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        
        Dim WorksheetName As String
        'Current row
        Dim i As Long
        'Start row of ticker block
        Dim LastRowI As Long
        'Variable for percent change calculation
        Dim PerChange As Double
        'Variable for greatest inc. calculation
        Dim GreatIncr As Double
         Dim j As Long
        'Index counter to fill ticker row
        Dim TickCount As Long
        'Last row column A
        Dim LastRowA As Long
        'last row column I
        'Variable for greatest dec. calculation
        Dim GreatDecr As Double
        'Variable for greatest total volume
        Dim GreatVol As Double
        
        'Register worksheet name
        WorksheetName = ws.Name

        
        'Set ticker counter to first row
        TickCount = 2
        
        'Set start row to two
        j = 2
        
        'Locate last cell in column A
        LastRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row
        'MsgBox ("Last row in column A is " & LastRowA)
        
       'Loop through all rows
        For i = 2 To LastRowA
            
        'Check to see if ticker label changed
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
               
                           
       'Note yearly change in column J number 10
        ws.Cells(TickCount, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
        
       'Write ticker in column I number 9
       ws.Cells(TickCount, 9).Value = ws.Cells(i, 1).Value
    
              
        'CF
      If ws.Cells(TickCount, 10).Value < 0 Then
                
     'Change color to red
     ws.Cells(TickCount, 10).Interior.ColorIndex = 3
                
      Else
      
      'Change color to green
      ws.Cells(TickCount, 10).Interior.ColorIndex = 4
                
     End If
                 
     'Note percent change in column K (#11)
      If ws.Cells(j, 3).Value <> 0 Then
      PerChange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                    
     'Percent formating
     ws.Cells(TickCount, 11).Value = Format(PerChange, "Percent")
                    
      Else
                    
      ws.Cells(TickCount, 11).Value = Format(0, "Percent")
                    
      End If
                
                    
    'Crease tick count by 1
    TickCount = TickCount + 1
                
    'Set new start row of the ticker block
    j = i + 1
           
           
    'Calculate and write total volume in column L (#12)
    ws.Cells(TickCount, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
                
         
     End If
            
     Next i
            
     'Find last filled cell in column I
    LastRowI = ws.Cells(Rows.Count, 9).End(xlUp).Row
    'MsgBox ("Last row in column I is " & LastRowI)
        
    'Labels for summary
    GreatVol = ws.Cells(2, 12).Value
    GreatIncr = ws.Cells(2, 11).Value
    GreatDecr = ws.Cells(2, 11).Value
        
     'Loop for summary
     For i = 2 To LastRowI
            
    'For total volume_check if next value is larger_if yes, replace new value & fill in ws.Cells
     If ws.Cells(i, 12).Value > GreatVol Then
    GreatVol = ws.Cells(i, 12).Value
   ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                
                
    Else
                
   GreatVol = GreatVol
                
   End If
                
    'For greatest increase, check if next value is larger, if yes replace new value & populate ws.Cells
    If ws.Cells(i, 11).Value > GreatIncr Then
   GreatIncr = ws.Cells(i, 11).Value
   ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                
   Else
                
  GreatIncr = GreatIncr
                
 End If
                
  'For biggest decrease_check if  number is smaller_if yes, substitute over a new value & fill in ws.Cells
  If ws.Cells(i, 11).Value < GreatDecr Then
  GreatDecr = ws.Cells(i, 11).Value
  ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                
 Else
                
       GreatDecr = GreatDecr
 End If
                
  'Note summary results in ws.Cells
 ws.Cells(2, 17).Value = Format(GreatIncr, "Percent")
ws.Cells(3, 17).Value = Format(GreatDecr, "Percent")
ws.Cells(4, 17).Value = Format(GreatVol, "Scientific")
            
Next i

        
Next ws
        
End Sub

