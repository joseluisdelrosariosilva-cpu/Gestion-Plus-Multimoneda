VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} facturacion 
   Caption         =   "Orden de venta"
   ClientHeight    =   9580.001
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   13620
   OleObjectBlob   =   "facturacion.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "facturacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public monto As Double
Public descripcion As String
Public fecha As Date
Public Moneda_Costo As String


Private Sub boton_add_covered_Click()
'Este evento pertenece al frame de facturacion y primeramente verifica que la disponibilidad
    'sea mayor que 0 y tambien que las casillas de cantidad y de codigo no esten vacias
    'luego de eso toma los datos de los textboxes y los introduce en el listbox
    'luego limpia los campos y llama al modulo que calcula el total de la suma

    On Error Resume Next
    
    ' PRIMERO: Validaciones básicas de campos obligatorios
    If Me.codigo_facturacion_TextBox.Value = "" Then
        MsgBox "Añada algún producto desde la lista de búsqueda para continuar", vbExclamation
        Exit Sub
    ElseIf Me.cantidad_facturacion_TextBox.Value = "" Then
        MsgBox "Indique la cantidad de unidades que desea agregar", vbExclamation
        Exit Sub
    ElseIf Me.Disponibilidad_facturacio_TextBox.Value <= 0 Then
        MsgBox "No hay disponibilidad en almacén", vbCritical
        Exit Sub
    ElseIf Me.cantidad_facturacion_TextBox.Value <= 0 Then
        MsgBox "Debe añadir una cantidad superior a 0", vbExclamation
        Exit Sub
    End If
    
    ' SEGUNDO: Validar repetición
    Call comprobar_repeticion
    If repeticion = True Then
        MsgBox "No puede añadir un producto añadido anteriormente. Por favor modifique la cantidad registrada en la factura", vbExclamation
        Exit Sub
    End If
    
    ' TERCERO: Validar cantidad vs disponibilidad
    If CDbl(Me.cantidad_facturacion_TextBox.Value) > CDbl(Me.Disponibilidad_facturacio_TextBox.Value) Then
        MsgBox "La disponibilidad en almacén es: " & Me.Disponibilidad_facturacio_TextBox.Value, vbExclamation
        Me.cantidad_facturacion_TextBox.Value = CInt(Me.Disponibilidad_facturacio_TextBox.Value)
        Me.SpinButton1.Value = CInt(Me.Disponibilidad_facturacio_TextBox.Value)
        Exit Sub
    End If
    
    ' CUARTO: Si pasa todas las validaciones, agregar al ListBox
    Me.productos_facturacion_ListBox.AddItem
    Me.productos_facturacion_ListBox.List(Me.productos_facturacion_ListBox.ListCount - 1, 0) = Me.codigo_facturacion_TextBox.Value
    Me.productos_facturacion_ListBox.List(Me.productos_facturacion_ListBox.ListCount - 1, 1) = Me.nombre_de_producto_facturacionTextBox.Value
    Me.productos_facturacion_ListBox.List(Me.productos_facturacion_ListBox.ListCount - 1, 2) = Me.Precio_facturación_TextBox.Value
    Me.productos_facturacion_ListBox.List(Me.productos_facturacion_ListBox.ListCount - 1, 3) = Me.Moneda.Value
    Me.productos_facturacion_ListBox.List(Me.productos_facturacion_ListBox.ListCount - 1, 4) = Me.cantidad_facturacion_TextBox.Value
    
    ' Limpiar campos
    Me.codigo_facturacion_TextBox.Value = ""
    Me.nombre_de_producto_facturacionTextBox.Value = ""
    Me.Precio_facturación_TextBox.Value = ""
    Me.Disponibilidad_facturacio_TextBox.Value = ""
    Me.cantidad_facturacion_TextBox.Value = ""
    Me.Moneda.Value = ""
    
    Call calcular_total
End Sub

Private Sub boton_add_reposo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.boton_add_covered.Visible = True
    Me.boton_add_covered.ZOrder msoBringToFront
End Sub

Private Sub boton_admin_covered_Click()
    Login.Show
End Sub

Private Sub boton_admin_reposo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.boton_admin_covered.Visible = True
    Me.boton_admin_covered.ZOrder msoBringToFront
End Sub

Private Sub boton_atras_reposo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.boton_atras_covered.Visible = True
    Me.boton_atras_covered.ZOrder msoBringToFront
End Sub

Private Sub boton_calendario_covered_Click()
'Aqui se configura el calendario para trabajar en la facturacio
'esto se hizo para diferenciar el textbox que deber recibir la salida
'pero usando el mismo calendario, y asi evitar poner otro
    seleccion_destino_calendario = 2
    frmCalendario.Show
End Sub

Private Sub boton_calendario_reposo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.boton_calendario_covered.Visible = True
    Me.boton_calendario_covered.ZOrder msoBringToFront
End Sub

Private Sub boton_delete_covered_Click()
    Dim respuesta
    If Me.productos_facturacion_ListBox.ListIndex = -1 Then
        MsgBox "Seleccione el producto que desea quitar de la facturación", vbExclamation
    Else
        respuesta = MsgBox("¿Está seguro que desea quitar el producto seleccionado?", vbYesNo, "Eliminar producto")
            Select Case respuesta
                Case Is = 6
                Me.productos_facturacion_ListBox.RemoveItem (Me.productos_facturacion_ListBox.ListIndex)
                Call calcular_total
                Case Is = 7
                Exit Sub
            End Select
    End If
End Sub

Private Sub boton_delete_reposo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.boton_delete_covered.Visible = True
    Me.boton_delete_covered.ZOrder msoBringToFront
End Sub


Private Sub boton_edit_covered_Click()
    On Error GoTo Salir
    If Me.productos_facturacion_ListBox.ListIndex = -1 Then
        MsgBox "Seleccione el producto al que desea modificar su cantidad", vbExclamation
    Else
        Dim nueva_cantidad As Integer
        nueva_cantidad = InputBox("Inserte la nueva cantidad", "Editar Cantidad")
        Me.productos_facturacion_ListBox.List(Me.productos_facturacion_ListBox.ListIndex, 4) = nueva_cantidad
        Call calcular_total
    End If
Salir:
    Exit Sub
End Sub

Private Sub boton_edit_reposo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.boton_edit_covered.Visible = True
    Me.boton_edit_covered.ZOrder msoBringToFront
End Sub

Private Sub boton_procesar_pago_covered_Click()
    Dim total_facturacion As Double
    Dim restan_por_pagar As Double
    Dim total_facturacion_USD As Double
    Dim restan_por_pagar_USD As Double
    
    If Me.productos_facturacion_ListBox.ListCount = 0 Then
        MsgBox "Añada productos a la factura para continuar", vbExclamation
    Else
    
    total_facturacion = Me.total_facturacion_TextBox.Value
    total_facturacion_USD = Me.total_facturacion_USD_TextBox.Value
    restan_por_pagar = total_facturacion
    restan_por_pagar_USD = total_facturacion_USD
    Me.total_a_pagar_TextBox.Value = Format(total_facturacion, "0.00")
    Me.total_a_pagar_USD_TextBox.Value = Format(total_facturacion_USD, "0.00")
        Me.factura_id_TextBox.Enabled = False
        Me.fecha_de_orden_TextBox.Enabled = False
        Me.boton_calendario_reposo.Visible = False
        Me.nombre_producto_busqueda_TextBox.Enabled = False
        Me.productos_busqueda_ListBox.Enabled = False
        Me.codigo_facturacion_TextBox.Enabled = False
        Me.nombre_de_producto_facturacionTextBox.Enabled = False
        Me.Precio_facturación_TextBox.Enabled = False
        Me.Disponibilidad_facturacio_TextBox.Enabled = False
        Me.cantidad_facturacion_TextBox.Enabled = False
        Me.SpinButton1.Enabled = False
        Me.productos_facturacion_ListBox.Enabled = False
        Me.total_facturacion_TextBox.Enabled = False
        Me.total_a_pagar_TextBox.Enabled = True
        Me.restan_por_pagar_TextBox.Enabled = True
        Me.transferecia_CheckBox.Enabled = True
        Me.efectivo_CheckBox.Enabled = True
        Me.USD_CheckBox.Enabled = True
        Me.MultiPage1.Value = 1
    End If
    Me.restan_por_pagar_TextBox.Value = Format(restan_por_pagar, "0.00")
    Me.restan_por_pagar_USD_TextBox.Value = Format(restan_por_pagar_USD, "0.00")
    ActualizarSaldos
End Sub

Private Sub boton_procesar_pago_reposo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.boton_procesar_pago_covered.Visible = True
    Me.boton_procesar_pago_covered.ZOrder msoBringToFront
End Sub

Private Sub boton_registrar_ventas_reposo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.boton_registrar_ventas_covered.ZOrder msoBringToFront
    Me.Label26.ZOrder msoBringToFront
End Sub

Private Sub cantidad_a_pagar_transeferencia_TextBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If (KeyAscii >= 46 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub

'Esto es una validación para evitar que se escriban letras
Private Sub cantidad_facturacion_TextBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If (KeyAscii >= 46 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub


Private Sub fecha_de_orden_TextBox_Change()
    Call generar_id_factura
End Sub

Private Sub fecha_de_orden_TextBox_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.boton_calendario_reposo.ZOrder msoBringToFront
    Me.boton_calendario_covered.Visible = False
End Sub

Private Sub ingreso_cliente_pago_TextBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If (KeyAscii >= 46 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub ingreso_cliente_pago_USD_TextBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If (KeyAscii >= 46 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub Label26_Click()
    boton_registrar_ventas_covered_Click
End Sub


Private Sub MultiPage1_MouseMove(ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.boton_add_covered.Visible = False
    Me.boton_add_reposo.ZOrder msoBringToFront
    Me.boton_edit_covered.Visible = False
    Me.boton_edit_reposo.ZOrder msoBringToFront
    Me.boton_delete_covered.Visible = False
    Me.boton_delete_reposo.ZOrder msoBringToFront
    Me.boton_procesar_pago_covered.Visible = False
    Me.boton_procesar_pago_reposo.ZOrder msoBringToFront
    Me.boton_atras_covered.Visible = False
    Me.boton_atras_reposo.ZOrder msoBringToFront
    Me.boton_registrar_ventas_reposo.ZOrder msoBringToFront
    Me.Label26.ZOrder msoBringToFront
End Sub

Private Sub nombre_producto_busqueda_TextBox_Change()
    ' Macro para buscar productos en la tabla Inventario
    ' Filtra por código o nombre del producto
    
    Dim HojaBD As Worksheet
    Dim tabla As ListObject
    Dim filaTabla As ListRow
    Dim textoBusqueda As String
    Dim Y As Integer
    
    Set HojaBD = ThisWorkbook.Sheets("Gestión de Almacén")
    Set tabla = HojaBD.ListObjects("INVENTARIO")
    
    ' Limpiar el ListBox
    Me.productos_busqueda_ListBox.Clear
    Me.productos_busqueda_ListBox.ColumnCount = 5
    
    ' Obtener el texto de búsqueda
    textoBusqueda = LCase(Trim(Me.nombre_producto_busqueda_TextBox.Value))
    
    Y = 0
    
    ' Recorrer todas las filas de la tabla
    For Each filaTabla In tabla.ListRows
        ' Obtener código y producto de la fila actual
        Dim codigo As String
        Dim producto As String
        
        ' Columna A = 1 (Código)
        codigo = LCase(filaTabla.Range(1).Value)
        ' Columna B = 2 (Producto)
        producto = LCase(filaTabla.Range(2).Value)
        
        ' Verificar si el texto de búsqueda coincide con código o producto
        If textoBusqueda = "" Or InStr(1, codigo, textoBusqueda) > 0 Or InStr(1, producto, textoBusqueda) > 0 Then
            
            ' Agregar fila al ListBox
            Me.productos_busqueda_ListBox.AddItem
            
            ' Columna 1 (índice 0): Código - Columna A = 1
            Me.productos_busqueda_ListBox.List(Y, 0) = filaTabla.Range(1).Value
            
            ' Columna 2 (índice 1): Producto - Columna B = 2
            Me.productos_busqueda_ListBox.List(Y, 1) = filaTabla.Range(2).Value
            
            ' Columna 3 (índice 2): Stock - Columna E = 5
            Me.productos_busqueda_ListBox.List(Y, 2) = filaTabla.Range(5).Value
            
            ' Columna 4 (índice 3): Precio Venta - Columna F = 6
            Me.productos_busqueda_ListBox.List(Y, 3) = filaTabla.Range(6).Value
            
            ' Columna 5 (índice 4): Moneda Venta - Columna G = 7
            Me.productos_busqueda_ListBox.List(Y, 4) = filaTabla.Range(7).Value
            
            Y = Y + 1
        End If
    Next filaTabla
    
    ' Ajustar el ancho de las columnas
    Me.productos_busqueda_ListBox.ColumnWidths = "70;85;55;55;40"
    
End Sub

Private Sub productos_busqueda_ListBox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    'Esta macro traspasa los datos del listbox de busqueda a los campos(textbox) de facturacion
    On Error Resume Next
    Me.codigo_facturacion_TextBox.Value = Me.productos_busqueda_ListBox.List(Me.productos_busqueda_ListBox.ListIndex, 0)
    Me.nombre_de_producto_facturacionTextBox.Value = Me.productos_busqueda_ListBox.List(Me.productos_busqueda_ListBox.ListIndex, 1)
    Me.Disponibilidad_facturacio_TextBox.Value = Me.productos_busqueda_ListBox.List(Me.productos_busqueda_ListBox.ListIndex, 2)
    Me.Precio_facturación_TextBox.Value = Me.productos_busqueda_ListBox.List(Me.productos_busqueda_ListBox.ListIndex, 3)
    Me.Moneda.Value = Me.productos_busqueda_ListBox.List(Me.productos_busqueda_ListBox.ListIndex, 4)

End Sub

Private Sub productos_facturacion_ListBox_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.boton_add_covered.Visible = False
    Me.boton_add_reposo.ZOrder msoBringToFront
    Me.boton_edit_covered.Visible = False
    Me.boton_edit_reposo.ZOrder msoBringToFront
    Me.boton_delete_covered.Visible = False
    Me.boton_delete_reposo.ZOrder msoBringToFront
End Sub

'spinbutton para controlar el campo de cantidad
Private Sub SpinButton1_Change()
    Me.cantidad_facturacion_TextBox = Me.SpinButton1.Value
End Sub

Private Sub total_facturacion_TextBox_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.boton_procesar_pago_covered.Visible = False
    Me.boton_procesar_pago_reposo.ZOrder msoBringToFront
End Sub

Private Sub UserForm_Activate()
    Me.MultiPage1.Value = 0
    Me.fecha_de_orden_TextBox.Value = Format(Date, "d-mmm-yyyy")
    If Application.Visible = True Then
        'Me.adminacces_boton.Enabled = False
        Me.boton_admin_reposo.Visible = False
    ElseIf Application.Visible = False Then
        'Me.adminacces_boton.Enabled = True
        Me.boton_admin_reposo.Visible = True
    End If
End Sub
Private Sub UserForm_Initialize()

'************************** Macro para rellenar listbox***********************
'Esta recibo los datos de la tabla de inventario, y se realiza un bucle para agregar solo
'las columnas deseadas

    Dim HojaBD As Worksheet
    Dim tabla As ListObject
    Dim filaTabla As ListRow
    Dim Y As Integer
    
    Set HojaBD = ThisWorkbook.Sheets("Gestión de Almacén")
    Set tabla = HojaBD.ListObjects("INVENTARIO")
    
    ' Limpiar el ListBox
    Me.productos_busqueda_ListBox.Clear
    Me.productos_busqueda_ListBox.ColumnCount = 5
    Me.productos_facturacion_ListBox.ColumnCount = 5
    
    Y = 0
    
    ' Recorrer todas las filas de la tabla
    For Each filaTabla In tabla.ListRows
        ' Agregar nueva fila al ListBox
        Me.productos_busqueda_ListBox.AddItem
        
        ' Columna 1 (índice 0): Código - Columna A = 1
        Me.productos_busqueda_ListBox.List(Y, 0) = filaTabla.Range(1).Value
        
        ' Columna 2 (índice 1): Producto - Columna B = 2
        Me.productos_busqueda_ListBox.List(Y, 1) = filaTabla.Range(2).Value
        
        ' Columna 3 (índice 2): Stock - Columna E = 5
        Me.productos_busqueda_ListBox.List(Y, 2) = CInt(filaTabla.Range(5).Value)
        
        ' Columna 4 (índice 3): Precio Venta - Columna F = 6
        Me.productos_busqueda_ListBox.List(Y, 3) = Format(filaTabla.Range(6).Value, "0.00")
        
        ' Columna 5 (índice 4): Moneda Venta - Columna G = 7
        Me.productos_busqueda_ListBox.List(Y, 4) = filaTabla.Range(7).Value
        
        Y = Y + 1
    Next filaTabla
    
'*********************************************************************

'******** Macro para cambiar tamano de columnas en facturacion *******
' Ajustar el ancho de las 5 columnas
    Me.productos_busqueda_ListBox.ColumnWidths = "70;85;55;55;40"
    Me.productos_facturacion_ListBox.ColumnWidths = "70;65;45;40;25"
'**********************************************************************

    Me.total_a_pagar_TextBox.Enabled = False
    Me.restan_por_pagar_TextBox.Enabled = False
    Me.transferecia_CheckBox.Enabled = False
    Me.efectivo_CheckBox.Enabled = False

End Sub

'Macro para agregar el gasto de inversión en la tabla de gastos
Sub Agregar_gasto_inv(mon As String)

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim newrow As ListRow

    Set ws = ThisWorkbook.Sheets("Historial de Gastos")
    Set tbl = ws.ListObjects("Gastos")

    ws.Unprotect password:="theSheetProtectpassisthe011rd"
    
    Set newrow = tbl.ListRows.Add
    With newrow
        .Range(1) = CDate(fecha)
        .Range(2) = "Inversión"
        .Range(3) = descripcion
        .Range(4) = CDbl(monto)
        .Range(5) = mon
    End With
    
    tbl.Sort.Apply

    ws.Protect password:="theSheetProtectpassisthe011rd", userinterfaceonly:=True

End Sub
'funcion para calcular el monto del producto vendido que fue de inversion
Private Function CalcularMonto(idProducto As String, Cantidad As Integer) As Double

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim fila As ListRow
    Dim costo As Double
    
    Set ws = ThisWorkbook.Sheets("Gestión de Almacén")
    Set tbl = ws.ListObjects("INVENTARIO")
    
    For Each fila In tbl.ListRows
        If fila.Range(1).Value = idProducto Then
            costo = fila.Range(8).Value
            Moneda_Costo = fila.Range(9).Value
            CalcularMonto = costo * Cantidad
            Exit Function
        End If
    Next fila
    
End Function

'Macro para agregar la fecha a la tabla de fechas unicas
Private Sub Agregar_Fecha()

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim newrow As ListRow
    
    Set ws = ThisWorkbook.Sheets("Fechas")
    Set tbl = ws.ListObjects("FechasUnicas")
    
    For Each fila In tbl.ListRows
        If CDate(fila.Range(1)) = CDate(Me.fecha_de_orden_TextBox.Value) Then
            Exit Sub
        End If
    Next fila
    
    ws.Unprotect password:="theSheetProtectpassisthe011rd"
    
    Set newrow = tbl.ListRows.Add
    
    newrow.Range(1) = CDate(Me.fecha_de_orden_TextBox.Value)
    
    tbl.Sort.Apply
    
    ws.Protect password:="theSheetProtectpassisthe011rd", userinterfaceonly:=True

End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.boton_admin_reposo.ZOrder msoBringToFront
    Me.boton_admin_covered.Visible = False
    Me.boton_calendario_reposo.ZOrder msoBringToFront
    Me.boton_calendario_covered.Visible = False
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu And Me.boton_admin_reposo.Visible = True Then
    ThisWorkbook.Save
    Application.Quit
    Else
        Exit Sub
    End If
End Sub

Private Sub Restar_Cantidad(codigo As String, Cantidad As Integer)
    Dim ws_almacen As Worksheet
    Dim tbl_almacen As ListObject
    Dim fila As ListRow
    Dim tipoCambio As Double
    Dim monedaCosto As String
    Dim monedaVenta As String
    Dim precioCosto As Double
    Dim precioVenta As Double
    Dim stockActual As Integer
    Dim fondo As Double
    Dim Ingreso As Double
    Dim ganancia As Double
    
    Set ws_almacen = ThisWorkbook.Worksheets("Gestión de Almacén")
    Set tbl_almacen = ws_almacen.ListObjects("INVENTARIO")
    
    ' Obtener tipo de cambio de la celda G3
    tipoCambio = ws_almacen.Range("G3").Value
    If tipoCambio <= 0 Then tipoCambio = 1 ' Evitar división por cero
    
    For Each fila In tbl_almacen.ListRows
        ' Verificar si el código coincide (Columna A = índice 1)
        If codigo = fila.Range(1).Value Then
            
            ' Leer valores actuales
            stockActual = fila.Range(5).Value
            precioVenta = fila.Range(6).Value
            monedaVenta = fila.Range(7).Value
            precioCosto = fila.Range(8).Value
            monedaCosto = fila.Range(9).Value
            
            ' Actualizar stock (Columna E = 5)
            fila.Range(5).Value = stockActual - Cantidad
            
            ' Recalcular fondo (Columna J = 10)
            ' Fondo = Precio Costo * Stock actual
            fondo = precioCosto * fila.Range(5).Value
            fila.Range(10).Value = fondo
            fila.Range(11).Value = monedaCosto ' Moneda Fondo (Columna K = 11)
            
            ' Recalcular ingreso esperado (Columna L = 12)
            ' Ingreso = Precio Venta * Stock actual
            Ingreso = precioVenta * fila.Range(5).Value
            fila.Range(12).Value = Ingreso
            fila.Range(13).Value = monedaVenta ' Moneda Ingreso (Columna M = 13)
            
            ' Recalcular ganancia según las monedas (Columna N = 14)
            Select Case monedaCosto & "-" & monedaVenta
                Case "MN-MN"
                    ' Caso 1: Costo MN, Venta MN
                    ganancia = Ingreso - fondo
                    fila.Range(14).Value = ganancia
                    fila.Range(15).Value = "MN" ' Moneda Ganancia (Columna O = 15)
                    
                Case "MN-USD"
                    ' Caso 2: Costo MN, Venta USD
                    ganancia = Ingreso - (fondo / tipoCambio)
                    fila.Range(14).Value = ganancia
                    fila.Range(15).Value = "USD"
                    
                Case "USD-MN"
                    ' Caso 3: Costo USD, Venta MN
                    ganancia = Ingreso - (fondo * tipoCambio)
                    fila.Range(14).Value = ganancia
                    fila.Range(15).Value = "MN"
                    
                Case "USD-USD"
                    ' Caso 4: Costo USD, Venta USD
                    ganancia = Ingreso - fondo
                    fila.Range(14).Value = ganancia
                    fila.Range(15).Value = "USD"
            End Select
            
            Exit Sub
        End If
    Next fila

End Sub

Private Sub boton_atras_covered_Click()
    Me.MultiPage1.Value = 0
    Me.factura_id_TextBox.Enabled = True
    Me.fecha_de_orden_TextBox.Enabled = True
    Me.nombre_producto_busqueda_TextBox.Enabled = True
    Me.productos_busqueda_ListBox.Enabled = True
    Me.codigo_facturacion_TextBox.Enabled = True
    Me.nombre_de_producto_facturacionTextBox.Enabled = True
    Me.Precio_facturación_TextBox.Enabled = True
    Me.Disponibilidad_facturacio_TextBox.Enabled = True
    Me.cantidad_facturacion_TextBox.Enabled = True
    Me.SpinButton1.Enabled = True
    Me.productos_facturacion_ListBox.Enabled = True
    Me.total_a_pagar_TextBox.Enabled = False
    Me.restan_por_pagar_TextBox.Enabled = False
    Me.total_facturacion_TextBox.Enabled = True
    Me.total_a_pagar_TextBox.Value = 0
    Me.restan_por_pagar_TextBox.Value = 0
    Me.total_a_pagar_USD_TextBox.Value = 0
    Me.restan_por_pagar_USD_TextBox.Value = 0
    Me.ingreso_cliente_pago_TextBox.Value = 0
    Me.efectivo_CheckBox.Value = False
    Me.efectivo_CheckBox.Enabled = False
    Me.transferecia_CheckBox.Value = False
    Me.transferecia_CheckBox.Enabled = False
    Me.USD_CheckBox.Value = False
    Me.USD_CheckBox.Enabled = False
    Me.ingreso_cliente_pago_TextBox.Value = 0
    Me.ingreso_cliente_pago_TextBox.Enabled = True
    Me.cantidad_a_pagar_transeferencia_TextBox.Enabled = True
    Me.boton_calendario_reposo.Visible = True
End Sub

'' Módulo del formulario (ejemplo: FacturacionForm)
'' Los totales a pagar en MN y USD son equivalentes (mismo valor en distintas monedas)
'
'Option Explicit

' -------------------------------------------------------------------
' Eventos de los CheckBox: habilitan/deshabilitan campos y limpian
' -------------------------------------------------------------------
Private Sub efectivo_CheckBox_Click()
    If Me.efectivo_CheckBox.Value = True Then
        Me.ingreso_cliente_pago_TextBox.Enabled = True
        Me.ingreso_cliente_pago_TextBox.Visible = True
    Else
        Me.ingreso_cliente_pago_TextBox.Enabled = False
        Me.ingreso_cliente_pago_TextBox.Visible = False
        Me.ingreso_cliente_pago_TextBox.Value = ""
    End If
    Call ActualizarSaldos
End Sub

Private Sub transferecia_CheckBox_Click()
    If Me.transferecia_CheckBox.Value = True Then
        Me.cantidad_a_pagar_transeferencia_TextBox.Enabled = True
        Me.cantidad_a_pagar_transeferencia_TextBox.Visible = True
    Else
        Me.cantidad_a_pagar_transeferencia_TextBox.Value = ""
        Me.cantidad_a_pagar_transeferencia_TextBox.Enabled = False
        Me.cantidad_a_pagar_transeferencia_TextBox.Visible = False
    End If
    Call ActualizarSaldos
End Sub

Private Sub USD_CheckBox_Click()
    If Me.USD_CheckBox.Value = True Then
        Me.ingreso_cliente_pago_USD_TextBox.Enabled = True
        Me.ingreso_cliente_pago_USD_TextBox.Visible = True
    Else
        Me.ingreso_cliente_pago_USD_TextBox.Value = ""
        Me.ingreso_cliente_pago_USD_TextBox.Enabled = False
        Me.ingreso_cliente_pago_USD_TextBox.Visible = False
    End If
    Call ActualizarSaldos
End Sub

' -------------------------------------------------------------------
' Eventos de cambio en los campos de pago
' -------------------------------------------------------------------
Private Sub ingreso_cliente_pago_TextBox_Change()
    Call ActualizarSaldos
End Sub

Private Sub cantidad_a_pagar_transeferencia_TextBox_Change()
    Call ActualizarSaldos
End Sub

Private Sub ingreso_cliente_pago_USD_TextBox_Change()
    Call ActualizarSaldos
End Sub

' -------------------------------------------------------------------
' Procedimiento principal: actualiza saldos restantes con conversión
' -------------------------------------------------------------------
Private Sub ActualizarSaldos()
    Dim totalMN As Double
    Dim totalUSD As Double
    Dim pagadoMN As Double
    Dim pagadoUSD As Double
    Dim restoMN As Double
    Dim restoUSD As Double
    Dim tipoCambio As Double
    
    ' Obtener tipo de cambio desde la celda G3 de la hoja "Gestión de Almacén"
    tipoCambio = ThisWorkbook.Sheets("Gestión de Almacén").Range("G3").Value
    If tipoCambio <= 0 Then tipoCambio = 1   ' evitar división por cero
    
    ' Obtener totales a pagar (asumimos que ya están cargados)
    totalMN = CDbl(Me.total_a_pagar_TextBox.Value)
    totalUSD = CDbl(Me.total_a_pagar_USD_TextBox.Value)
    
    ' Sumar pagos en MN (efectivo + transferencia)
    pagadoMN = 0
    If Me.efectivo_CheckBox.Value Then
        pagadoMN = pagadoMN + Val(Me.ingreso_cliente_pago_TextBox.Value)
    End If
    If Me.transferecia_CheckBox.Value Then
        pagadoMN = pagadoMN + Val(Me.cantidad_a_pagar_transeferencia_TextBox.Value)
    End If
    
    ' Pagos en USD
    pagadoUSD = 0
    If Me.USD_CheckBox.Value Then
        pagadoUSD = Val(Me.ingreso_cliente_pago_USD_TextBox.Value)
    End If
    
    ' Calcular restantes en cada moneda, aplicando conversión
    restoMN = totalMN - pagadoMN - (pagadoUSD * tipoCambio)
    restoUSD = totalUSD - pagadoUSD - (pagadoMN / tipoCambio)
    
    If restoUSD < 0.00999 And restoUSD > -0.00999 Then restoMN = 0
    
    ' Mostrar en los TextBox (formato con dos decimales)
    Me.restan_por_pagar_TextBox.Value = Format(restoMN, "0.00")
    Me.restan_por_pagar_USD_TextBox.Value = Format(restoUSD, "0.00")
    
    ' Aplicar colores según el valor
    Call AplicarColor(Me.restan_por_pagar_TextBox, restoMN)
    Call AplicarColor(Me.restan_por_pagar_USD_TextBox, restoUSD)
End Sub

' -------------------------------------------------------------------
' Procedimiento auxiliar para cambiar el color de fondo del TextBox
' -------------------------------------------------------------------
Private Sub AplicarColor(txtBox As Object, valor As Double)
    If valor > 0.00999 Then
        txtBox.ForeColor = RGB(255, 0, 0)    ' Rojo: falta pagar
    ElseIf valor > -0.00999 Then
        txtBox.ForeColor = RGB(0, 255, 0)    ' Verde: saldo cero
    Else
        txtBox.ForeColor = RGB(255, 255, 0)  ' Amarillo: pagó de más
    End If
End Sub

Private Sub boton_registrar_ventas_covered_Click()

    ' Validar que el saldo restante sea cero en ambas monedas
    If Format(Me.restan_por_pagar_TextBox.Value, "0.00") <> 0 Or Format(Me.restan_por_pagar_USD_TextBox.Value, "0.00") <> 0 Then
        MsgBox "El saldo pendiente no es cero. No se puede registrar la venta.", vbExclamation, "Saldo Pendiente"
        Exit Sub
    End If
    
    ' Obtener referencias a las tablas
    Dim wsVentas As Worksheet, wsFacturas As Worksheet
    Dim tablaVentas As ListObject, tablaFacturas As ListObject
    
    Set wsVentas = ThisWorkbook.Sheets("Historial de ventas")
    Set wsFacturas = ThisWorkbook.Sheets("Registro de facturas")
    Set tablaVentas = wsVentas.ListObjects("HISTORIALVENTAS")
    Set tablaFacturas = wsFacturas.ListObjects("HISTORIALVENTAS4")
    
    wsVentas.Unprotect password:="theSheetProtectpassisthe011rd"
    wsFacturas.Unprotect password:="theSheetProtectpassisthe011rd"
    
    ' Datos comunes
    Dim facturaID As String
    Dim fechaVenta As Date
    
    facturaID = Trim(Me.factura_id_TextBox.Value)
    fechaVenta = CDate(Me.fecha_de_orden_TextBox.Value)
    
    ' Registrar cada producto en la tabla HISTORIALVENTAS
    Dim i As Integer
    Dim nuevaFilaVenta As ListRow
    Dim codigo As String, nombre As String
    Dim Cantidad As Double, precioUnitario As Double
    Dim Moneda As String, totalVenta As Double
    
    For i = 0 To Me.productos_facturacion_ListBox.ListCount - 1
        ' Leer datos del listbox
        codigo = Me.productos_facturacion_ListBox.List(i, 0)        ' índice 0: código
        nombre = Me.productos_facturacion_ListBox.List(i, 1)        ' índice 1: nombre
        precioUnitario = CDbl(Me.productos_facturacion_ListBox.List(i, 2)) ' índice 2: precio
        Moneda = Me.productos_facturacion_ListBox.List(i, 3)        ' índice 3: moneda
        Cantidad = CDbl(Me.productos_facturacion_ListBox.List(i, 4)) ' índice 4: cantidad
        
        totalVenta = precioUnitario * Cantidad
        
        ' Agregar fila a la tabla HISTORIALVENTAS
        Set nuevaFilaVenta = tablaVentas.ListRows.Add
        
        ' Asignar valores por índice de columna (ajusta según el orden de columnas)
        ' Suponiendo columnas: Código, Nombre, Fecha, Cantidad, Precio, Moneda, Total, MonedaTotal, FacturaID
        ' Los índices comienzan en 1 (primera columna de la tabla)
        nuevaFilaVenta.Range(1).Value = codigo           ' Código
        nuevaFilaVenta.Range(2).Value = nombre           ' Nombre
        nuevaFilaVenta.Range(3).Value = fechaVenta       ' Fecha
        nuevaFilaVenta.Range(4).Value = Cantidad         ' Cantidad
        nuevaFilaVenta.Range(5).Value = precioUnitario   ' Precio
        nuevaFilaVenta.Range(6).Value = Moneda           ' Moneda
        nuevaFilaVenta.Range(7).Value = totalVenta       ' Total
        nuevaFilaVenta.Range(8).Value = Moneda           ' MonedaTotal (igual a moneda de venta)
        nuevaFilaVenta.Range(9).Value = facturaID        ' FacturaID
        fecha = fechaVenta
        descripcion = nombre
        monto = CalcularMonto(codigo, CInt(Cantidad))
        Call Agregar_gasto_inv(Moneda_Costo)
        Call Restar_Cantidad(codigo, CInt(Cantidad))
    Next i
    
    ' Registrar en tabla HISTORIALVENTAS4 (Registro de facturas)
    Dim efectivoMN As Double, transferenciaMN As Double, pagoUSD As Double
    
    ' Obtener valores de los TextBox, con validación si están vacíos o deshabilitados
    efectivoMN = 0
    transferenciaMN = 0
    pagoUSD = 0
    
    If Me.efectivo_CheckBox.Value And Me.ingreso_cliente_pago_TextBox.Enabled Then
        efectivoMN = Val(Me.ingreso_cliente_pago_TextBox.Value)
    End If
    
    If Me.transferecia_CheckBox.Value And Me.cantidad_a_pagar_transeferencia_TextBox.Enabled Then
        transferenciaMN = Val(Me.cantidad_a_pagar_transeferencia_TextBox.Value)
    End If
    
    If Me.USD_CheckBox.Value And Me.ingreso_cliente_pago_USD_TextBox.Enabled Then
        pagoUSD = Val(Me.ingreso_cliente_pago_USD_TextBox.Value)
    End If
    
    ' Agregar nueva fila a la tabla de facturas
    Dim nuevaFilaFactura As ListRow
    Set nuevaFilaFactura = tablaFacturas.ListRows.Add
    
    ' Asignar valores (ajusta los índices según la estructura de la tabla)
    ' Columnas: Factura ID, Fecha, Efectivo MN, Transferencia MN, USD
    nuevaFilaFactura.Range(1).Value = facturaID
    nuevaFilaFactura.Range(2).Value = fechaVenta
    nuevaFilaFactura.Range(3).Value = efectivoMN
    nuevaFilaFactura.Range(4).Value = transferenciaMN
    nuevaFilaFactura.Range(5).Value = pagoUSD
    
    tablaVentas.Sort.Apply
    tablaFacturas.Sort.Apply
    
    wsVentas.Protect password:="theSheetProtectpassisthe011rd", userinterfaceonly:=True
    wsFacturas.Protect password:="theSheetProtectpassisthe011rd", userinterfaceonly:=True
    
    Call reinicio_POS
    ' Opcional: limpiar formulario o cerrar
    ' Unload Me   ' Si se desea cerrar el formulario después de registrar
End Sub
