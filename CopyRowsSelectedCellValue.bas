Sub CopyRowsSelectedCellValue()
'Updateby Extendoffice
    Dim xCount As Integer
    
    xCount = ActiveCell.Value
    
    While xCount > 0
        If xCount > 1 Then
            
            ActiveCell.EntireRow.Copy
            Range(ActiveCell.Offset(1, 0), ActiveCell.Offset(xCount - 1, 0)).EntireRow.Insert Shift:=xlDown
            Application.CutCopyMode = False
            
        End If
        
        Selection.Offset(xCount, 0).Select
   
    xCount = ActiveCell.Value
    
    Wend

End Sub

