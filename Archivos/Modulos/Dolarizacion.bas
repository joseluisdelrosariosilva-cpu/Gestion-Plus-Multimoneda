Attribute VB_Name = "Dolarizacion"
Sub Cambiar_Precio_Dolar()
    Dim numeroUsuario As Variant
    Dim HojaBD As Worksheet
    Dim tabla As ListObject
    Dim fila As ListRow
    Dim tipoCambioNuevo As Double
    Dim monedaCosto As String
    Dim monedaVenta As String
    Dim fondo As Double
    Dim Ingreso As Double
    Dim ganancia As Double
    
    Set HojaBD = ThisWorkbook.Sheets("Gestión de Almacén")
    Set tabla = HojaBD.ListObjects("INVENTARIO")
    
    Do
        numeroUsuario = InputBox("Ingrese un número:", "Número", HojaBD.Range("G3").Value)
        
        If numeroUsuario = "" Then
            Exit Sub
        End If
        
        If IsNumeric(numeroUsuario) Then
            tipoCambioNuevo = CDbl(numeroUsuario)
            
            ' Actualizar el tipo de cambio en la celda G3
            HojaBD.Range("G3").Value = tipoCambioNuevo
            
            ' Recorrer todas las filas de la tabla
            For Each fila In tabla.ListRows
                ' Obtener las monedas de costo y venta
                ' Columna I = 9 (Moneda Costo), Columna G = 7 (Moneda Venta)
                monedaCosto = fila.Range(9).Value
                monedaVenta = fila.Range(7).Value
                
                ' Verificar si el producto requiere conversión (monedas diferentes)
                If monedaCosto <> monedaVenta Then
                    
                    ' Obtener valores actuales
                    ' Columna J = 10 (Fondo), Columna L = 12 (Ingreso Esperado)
                    fondo = fila.Range(10).Value
                    Ingreso = fila.Range(12).Value
                    
                    ' Recalcular ganancia según el caso
                    Select Case monedaCosto & "-" & monedaVenta
                        Case "MN-USD"
                            ' Caso: Costo MN, Venta USD
                            ' Columna N = 14 (Ganancia)
                            ganancia = Ingreso - (fondo / tipoCambioNuevo)
                            fila.Range(14).Value = ganancia
                            
                        Case "USD-MN"
                            ' Caso: Costo USD, Venta MN
                            ' Columna N = 14 (Ganancia)
                            ganancia = Ingreso - (fondo * tipoCambioNuevo)
                            fila.Range(14).Value = ganancia
                    End Select
                End If
            Next fila
            
            Exit Do
        End If
    Loop
End Sub
