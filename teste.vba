Sub deleteDupes()
  Dim i As Integer
  With Sheets("Sheet1")

'Dim killRng As Range

  Dim cont As Integer

For i = 2 To .Cells(Rows.Count, 3).End(xlUp).Row
    If .Cells(i, 3).Value = .Cells(i - 1, 3).Value Then
       
        cont = cont + 1
    Else
        If (cont >= 4) Then
        'formatar
        .Range(.Cells(i - 2, 1), .Cells(i - 2, 8)).Select
        With Selection
            .HorizontalAlignment = xlCenter
             .VerticalAlignment = xlBottom
             .WrapText = False
              .Orientation = 0
               .AddIndent = False
               .IndentLevel = 0
              .ShrinkToFit = False
              .ReadingOrder = xlContext
              .MergeCells = True
          End With
          Selection.Merge
          .Range(.Cells(i - 2, 1), .Cells(i - 2, 8)).Select
          ActiveCell.Formula = "11+12"
        
        
        'deletar
        
           For j = 3 To cont - 1
            .Rows(i - j).Delete
            Next j
        Else
         cont = 0
        End If
    End If
Next i
End With
 
 
End Sub
