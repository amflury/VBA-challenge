Attribute VB_Name = "Module1"
Sub Stocks_List()

   
    Dim worksheetname As String
    
    
    
    
    
    
    
    
    Dim Ticker As String
    Dim Change As Double
    Dim Volume As Double
    Volume = 0

    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    'Dim LastRow As Integer
    
    'For i = 2 To LastRow
    For i = 2 To 10000
    
  '     LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
   
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
            Ticker = Cells(i, 1).Value
        
            Volume = Volume + Cells(i, 7).Value
            
  '          Change = Cells(i, 3).Start(x1Up).Select - Cells(i, 6).End(x1Up).Select
            
   '         Change = Range("C" & Rows.Count).Start(x1Up).Select - Range("F" & Rows.Count).End(x1Up).Select

            Range("I" & Summary_Table_Row).Value = Ticker
        
            Range("L" & Summary_Table_Row).Value = Volume
            
  '          Range("J" & Summary_Table_Row).Value = Change
        
            Summary_Table_Row = Summary_Table_Row + 1
        
            Volume = 0
        
            Else

            Volume = Volume + Cells(i, 7).Value
    
        End If
    
    Next i
                        
End Sub
