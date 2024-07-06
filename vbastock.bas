Attribute VB_Name = "modFirst"

Sub stock_analysis():

  Dim total As Double
  Dim rowIndex As Long
  Dim change As Double
  Dim columnindex As Integer
  Dim start As Long
  Dim rowCount As Long
  Dim PercentChange As Double
  Dim days As Integer
  Dim dailyChange As Single
  Dim AverageChange As Double
  Dim ws As Worksheet
  
  
   For Each ws In Worksheets
    
    
  
   columnindex = 0
   total = 0
   change = 0
   start = 2
   dailyChange = 0
   
   ws.Range("I1").Value = "Ticker"
   ws.Range("J1").Value = "Yearly Change"
   ws.Range("K1").Value = "Percent Change"
   ws.Range("L1").Value = "Total Stock Volume"
   ws.Range("P1").Value = "Ticker"
   ws.Range("Q1").Value = "Value"
   ws.Range("O2").Value = "Greatest % Increase"
   ws.Range("O3").Value = "Greatest % Decrease"
   ws.Range("O4").Value = "Greatest Total Volume"
   
   rowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row
   
   For rowIndex = 2 To rowCount
   
   If ws.Cells(rowIndex + 1, 1).Value <> ws.Cells(rowIndex, 1).Value Then
   
    total = total + ws.Cells(rowIndex, 7).Value
    
    If total = 0 Then
    
     ws.Range("I" & 2 + columnindex).Value = Cells(rowIndex, 1).Value
     ws.Range("J" & 2 + columnIndx).Value = 0
     ws.Range("K" & 2 + columnindex).Value = "%" & 0
     ws.Range("L" & 2 + columnindex).Value = 0
      
 Else
   If ws.Cells(start, 3) = 0 Then
    For Find_value = start To rowIndex
     If ws.Cells(Find_value, 3).Value <> 0 Then
      start = Find_value
       Exit For
      End If
     Next Find_value
     End If
     
      change = (ws.Cells(rowIndex, 6) - ws.Cells(start, 3))
      PercentChange = change / ws.Cells(start, 3)
      
      start = rowIndex + 1
      
      ws.Range("I" & 2 + columnindex) = ws.Cells(rowIndex, 1).Value
      ws.Range("J" & 2 + columnindex) = change
      ws.Range("J" & 2 + columnindex).NumberFormat = "0.00"
      ws.Range("K" & 2 + columnindex).Value = PercentChange
      ws.Range("K" & 2 + columnindex).NumberFormat = "0.00%"
      ws.Range("L" & 2 + columnindex).Value = total
      
      Select Case change
        Case Is > 0
         ws.Range("J" & 2 + columnindex).Interior.ColorIndex = 4
        Case Is < 0
         ws.Range("J" & 2 + columnindex).Interior.ColorIndex = 3
        Case Else
         ws.Range("J" & 2 + columnindex).Interior.ColorIndex = 0
    End Select
    
   End If
   
    total = 0
    change = 0
    columnindex = columnindex + 1
    days = 0
    dailyChange = 0
    
    Else
    total = total + ws.Cells(rowIndex, 7).Value
   
   End If
   
   
   Next rowIndex
   
   
   ws.Range("Q2") = "%" & WorksheetFunction.Max(ws.Range("k2:K" & rowCount)) * 100
   ws.Range("Q3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & rowCount)) * 100
   ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2:L" & rowCount))
   
   Increase_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0)
   decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0)
   volume_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & rowCount)), ws.Range("L2:L" & rowCount), 0)
   
   ws.Range("P2") = ws.Cells(increaase_number + 1, 9)
   ws.Range("P3") = ws.Cells(decrease_number + 1, 9)
   ws.Range("P4") = ws.Cells(volume_number + 1, 9)
   
   
   
   
   
   
Next ws
  
  
End Sub


