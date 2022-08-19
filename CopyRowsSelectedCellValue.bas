Sub CopyRowsSelectedCellValue()
'Updateby Extendoffice
    Dim xCount As Integer
    
    If ActiveCell.Value = "-" Then
        Selection.Offset(1, 0).Select
    End If
    
    
    
    xCount = ActiveCell.Value
    
    While xCount > 0
        If xCount > 1 Then
            
            ActiveCell.EntireRow.Copy
            Range(ActiveCell.Offset(1, 0), ActiveCell.Offset(xCount - 1, 0)).EntireRow.Insert Shift:=xlDown
            Application.CutCopyMode = False
            
        End If
        
        Selection.Offset(xCount, 0).Select
   
    While ActiveCell.Value = "-"
        Selection.Offset(1, 0).Select
    Wend

    xCount = ActiveCell.Value
    
    Wend

End Sub
