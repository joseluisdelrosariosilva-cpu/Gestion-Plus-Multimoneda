Attribute VB_Name = "Calculo_precio_UBase"
Sub Calcular_UBase()
    Dim i As Long
    Dim resultado As Double
    Dim ultima_fila As Integer
    Dim HojaBD As Worksheet
    
    Set HojaBD = ThisWorkbook.Sheets("Gestión de Ingredientes")
        ultima_fila = HojaBD.Range("B" & Rows.Count).End(xlUp).Row
        
    ' Recorrer desde fila 5 hasta la última fila
    For i = 5 To ultima_fila
        ' Verificar que ambas celdas tengan valores
        If HojaBD.Range("G" & i).Value <> "" And HojaBD.Range("D" & i).Value <> "" Then
            ' Verificar que el divisor no sea cero
            If HojaBD.Range("D" & i).Value <> 0 Then
                ' Calcular la división
                resultado = HojaBD.Range("G" & i).Value / HojaBD.Range("D" & i).Value
                
                ' Poner el resultado en columna F
                HojaBD.Range("F" & i).Value = resultado
                
                ' Formato opcional: mostrar como porcentaje o decimal
                'HojaBD.Range("F" & i).NumberFormat = "0.00" ' Para 2 decimales

            Else
                ' Si el divisor es cero, poner error
                HojaBD.Range("F" & i).Value = "#DIV/0!"
            End If
        Else
            ' Si falta algún valor, dejar vacío
            HojaBD.Range("F" & i).Value = ""
        End If
    Next i
    
End Sub
