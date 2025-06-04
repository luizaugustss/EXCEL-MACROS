Sub CompararTabelasEDestacarItensIguais()
    Dim wsFirst As Worksheet
    Dim wsOther As Worksheet
    Dim tabela1 As ListObject
    Dim tabelaOutra As ListObject
    Dim celFirst As Range
    Dim celOther As Range
    Dim i As Long, j As Long
    Dim colCount As Long, rowCount As Long
    
    ' Define a primeira planilha com a tabela1
    Set wsFirst = ThisWorkbook.Worksheets(1)
    
    ' Verifica se existe uma tabela  na primeira planilha
    On Error Resume Next
    Set tabela1 = wsFirst.ListObjects(1)
    On Error GoTo 0
    
    If tabela1 Is Nothing Then
        MsgBox "Não foi encontrada uma tabela chamada 'tabela1' na primeira planilha.", vbExclamation
        Exit Sub
    End If
    
    ' Loop através de todas as outras planilhas no workbook
    For Each wsOther In ThisWorkbook.Worksheets
        If wsOther.Name <> wsFirst.Name Then
            ' Loop através de todas as tabelas na planilha atual
            For Each tabelaOutra In wsOther.ListObjects
                ' Loop através de cada célula de dados na tabela1
                For i = 1 To tabela1.DataBodyRange.Rows.Count
                    For j = 1 To tabela1.DataBodyRange.Columns.Count
                        Set celFirst = tabela1.DataBodyRange.Cells(i, j)
                        Set celOther = tabelaOutra.DataBodyRange.Cells(i, j)
                        celFirst.Interior.Color = xlNone
                        If celFirst.Value <> celOther.Value And celFirst.Value <> "" And celFirst.Interior.Color <> RGB(255, 0, 200) Then
                            celFirst.Interior.ColorIndex = 7
                            
                            If celOther.Value <> "" Then
                            celFirst.Interior.ColorIndex = 4 'Rosa
                            End If
                        End If
                 
                    Next j
                Next i
            Next tabelaOutra
        End If
    Next wsOther
    
    MsgBox "Comparação concluída! Os itens iguais foram destacados em amarelo.", vbInformation
End Sub
Sub trocarTabela()
    Dim wsFirst As Worksheet
    Dim wsOther As Worksheet
    Dim tabela1 As ListObject
    Dim tabelaOutra As ListObject
    Dim celFirst As Range
    Dim celOther As Range
    Dim i As Long, j As Long
    Dim colCount As Long, rowCount As Long
    
    ' Define a primeira planilha com a tabela1
    Set wsFirst = ThisWorkbook.Worksheets(1)
    
    'titulo
    If wsFirst.Range("C1").Value = "" Then
    wsFirst.Range("D1").Value = "tabela 1"
    End If
    
    ' Verifica se existe uma tabela chamada "tabela1" na primeira planilha
    On Error Resume Next
    Set tabela1 = wsFirst.ListObjects(1)
    On Error GoTo 0
    
    If tabela1 Is Nothing Then
        MsgBox "Não foi encontrada uma tabela na primeira planilha.", vbExclamation
        Exit Sub
    End If
    If wsFirst.Range("C1").Value = "tabela 1" Then
    MsgBox "wb 3"
      wsFirst.Range("C1").Value = "tabela 2"
    Set wsOther = ThisWorkbook.Worksheets(3)
    Else
    MsgBox "wb 2"
      wsFirst.Range("C1").Value = "tabela 1"
    Set wsOther = ThisWorkbook.Worksheets(2)
    End If
            Set tabelaOutra = wsOther.ListObjects(1)
                ' Loop através de cada célula de dados na tabela1
                For i = 1 To tabela1.DataBodyRange.Rows.Count
                    For j = 1 To tabela1.DataBodyRange.Columns.Count
                        tabela1.DataBodyRange.Cells(i, j) = tabelaOutra.DataBodyRange.Cells(i, j)
                    Next j
                Next i
    
End Sub
Function addTabela(TableName As String, Posicao As Integer)


    Dim wsFirst As Worksheet
    Dim wsOther As Worksheet
    Dim tabela1 As ListObject
    Dim tabelaOutra As ListObject
    Dim celFirst As Range
    Dim celOther As Range
    Dim i As Long, j As Long
    Dim colCount As Long, rowCount As Long
    
    ' Define a primeira planilha com a tabela1
    Set wsFirst = ThisWorkbook.Worksheets(1)
     
    ' Verifica se existe uma tabela na primeira planilha
    On Error Resume Next
    Set tabela1 = wsFirst.ListObjects(1)
    On Error GoTo 0
    
    If tabela1 Is Nothing Then
        MsgBox "Não foi encontrada uma tabela na primeira planilha.", vbExclamation
        Exit Function
    End If
    
    If wsFirst.Range("D1").Value = "tabela 1" Then
      wsFirst.Range("D1").Value = "tabela 2"
    Set wsOther = ThisWorkbook.Worksheets(3)
    Else
      wsFirst.Range("D1").Value = "tabela 1"
     ' Loop através de cada tabela procurando igual a busca
    For Each wsOther In ThisWorkbook.Worksheets
    If wsOther.Name = TableName Then
    End If
            Set tabelaOutra = wsOther.ListObjects(1)
                ' Loop através de cada célula de dados na tabela1
                For i = 1 To tabela1.DataBodyRange.Rows.Count
                    For j = 1 To tabela1.DataBodyRange.Columns.Count
                        tabela1.DataBodyRange.Cells(i, j) = tabelaOutra.DataBodyRange.Cells(i, j)
                    Next j
                Next i
    End If
    Next
End Function
