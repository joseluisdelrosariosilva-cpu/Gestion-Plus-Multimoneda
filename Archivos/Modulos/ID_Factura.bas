Attribute VB_Name = "ID_Factura"

'Esta macro genera el ID de factura del producto basado en la fecha y el numero consecutivo de factura
'Basicamente se busca la ultima fila y se convierte la fecha a numero entero que lo que realiza es un
'cálculo de cuantos dias han pasado desde 1/1/1900 hasta la actualidad
'Luego recorre todas las filas de la columna H en busca del numero consecutivo de la factura
'el cual es la combinacion del prefijo formado por los 5 digitos de la fecha convertida a entero
'y el sufijo de los otros 5 digitos correspondiente al numero consecutivo del dia que coincide con el prefijo
'por lo que este prefijo se reinicia diariamente y comienza a sumar sufijos a partir de 0 en adelante
'por cada venta que se realice en el dia. Esto lo hace si el sufijo es mayor que el numero que tiene.
'fue una adecuacion que tuve que hacer debido al ordenamiento

Sub generar_id_factura()

    Dim listarray(50000) As Long
    Dim numero As Long
    Dim fecha As Date
    Dim identificador_fecha As Long
    Dim maximo As Integer
    Dim theWs As Worksheet
    Dim theTbl As ListObject
    
    Set theWs = ThisWorkbook.Sheets("Historial de ventas")
    Set theTbl = theWs.ListObjects("HISTORIALVENTAS")
    
    fecha = facturacion.fecha_de_orden_TextBox.Value
    identificador_fecha = CLng(fecha)
    i = 1
    
    For Each fila In theTbl.ListRows
        If Left(CStr(fila.Range(9).Value), 5) = identificador_fecha Then
            numero = CInt(Right(fila.Range(9).Value, 5))

            listarray(i) = numero
            i = i + 1
        End If
    Next fila
    
    maximo = WorksheetFunction.Max(listarray) + 1
    generador_id_factura = CStr(identificador_fecha) & CStr(Format(maximo, "00000"))
    facturacion.factura_id_TextBox.Value = generador_id_factura
    
End Sub
