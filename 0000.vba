Sub ProcurarCodigoEmTodasAbas()
    Dim codigo As String
    Dim ws As Worksheet
    Dim celula As Range
    Dim resultado As String
    Dim primeiraOcorrencia As Boolean
    
    codigo = InputBox("Digite o código a ser procurado:")
    
    If codigo = "" Then
        MsgBox "Nenhum código foi digitado.", vbExclamation
        Exit Sub
    End If
    
    resultado = ""
    primeiraOcorrencia = True
    
    For Each ws In ThisWorkbook.Worksheets
        For Each celula In ws.Range("F1:F" & ws.Cells(ws.Rows.Count, "F").End(xlUp).Row)
            If celula.Value = codigo Then
                resultado = resultado & "Planilha: " & ws.Name & vbCrLf & _
                            "Coluna C: " & celula.Offset(0, -3).Value & vbCrLf & _
                            "Coluna D: " & celula.Offset(0, -2).Value & vbCrLf & _
                            "Coluna E: " & celula.Offset(0, -1).Value & vbCrLf & vbCrLf
                primeiraOcorrencia = False
            End If
        Next celula
    Next ws
    
    If primeiraOcorrencia Then
        MsgBox "Código não encontrado em nenhuma aba.", vbInformation
    Else
        MsgBox resultado, vbInformation, "Resultados encontrados"
    End If
End Sub