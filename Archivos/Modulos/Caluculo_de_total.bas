Attribute VB_Name = "Caluculo_de_total"
Public Sub calcular_total()
'Macro para calcular el importe total a pagar en MN y USD
'Convierte todos los productos a ambas monedas para mostrar el total conjunto

    Dim total_de_linea As Double
    Dim total_MN As Double
    Dim total_USD As Double
    Dim total_MN_Convertido As Double
    Dim total_USD_Convertido As Double
    Dim tipoCambio As Double
    Dim precio As Double
    Dim Cantidad As Double
    Dim Moneda As String
    Dim i As Integer
    
    ' Obtener tipo de cambio de la celda G3
    tipoCambio = ThisWorkbook.Sheets("Gestión de Almacén").Range("G3").Value
    
    ' Inicializar totales
    total_MN = 0
    total_USD = 0
    
    ' Recorrer todos los productos en el listbox de facturación
    For i = 0 To facturacion.productos_facturacion_ListBox.ListCount - 1
        
        ' Obtener valores del listbox
        ' Índices: 0-código, 1-nombre, 2-precio, 3-moneda, 4-cantidad
        precio = CDbl(facturacion.productos_facturacion_ListBox.List(i, 2))
        Cantidad = CDbl(facturacion.productos_facturacion_ListBox.List(i, 4))
        Moneda = facturacion.productos_facturacion_ListBox.List(i, 3)
        
        ' Calcular total de línea
        total_de_linea = precio * Cantidad
        
        ' Acumular según la moneda
        If Moneda = "MN" Then
            total_MN = total_MN + total_de_linea
        ElseIf Moneda = "USD" Then
            total_USD = total_USD + total_de_linea
        End If
        
    Next i
    
    ' Calcular totales convertidos
    ' Total en MN = MN + (USD * tipoCambio)
    total_MN_Convertido = total_MN + (total_USD * tipoCambio)
    
    ' Total en USD = (MN / tipoCambio) + USD
    total_USD_Convertido = (total_MN / tipoCambio) + total_USD
    
    ' Mostrar totales
    facturacion.total_facturacion_TextBox.Value = Format(total_MN_Convertido, "0.00")
    facturacion.total_facturacion_USD_TextBox.Value = Format(total_USD_Convertido, "0.00")
    
End Sub
