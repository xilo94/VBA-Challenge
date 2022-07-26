Attribute VB_Name = "Module1"
Sub Stock()
    Dim Row_Number As Long
    Dim Total_row As Long
    
    Total_row = Cells(Rows.Count, 1).End(xlUp).Row
    For Row_Number = 2 To Total_row
        Cells(Row_Number, 9).Value = Cells(Row_Number, 1)
        'MsgBox (i)
        Cells(Row_Number, 10).Value = Cells(Row_Number, 6).Value - Cells(Row_Number, 3).Value
        Cells(Row_Number, 11).Value = FormatPercent((Cells(Row_Number, 6).Value - Cells(Row_Number, 3).Value) / 100)
        Range("l2") = Application.WorksheetFunction.Sum(Total_row, (Cells(Row_Number, 7).Value))
   Next Row_Number
  
  
  
  
  
    
End Sub
