Attribute VB_Name = "calling"
Public fila_a_modificar As Integer
Public seleccion_destino_calendario As Integer
' --- VARIABLES PÚBLICAS ---
Public modoModificacion As Boolean
Public recetaAModificar As String

'Macro que ejecuta el fomrulario en la configuración registrar la cual
'Ejecuta la generación automática del código y las validaciones de los campos vacíos

Sub llamar_ingresar_form()
    ingresar_form.modificar_registrar_textbox.Value = "R"
    ingresar_form.Show
End Sub

Sub llamar_Elaborar()
    Elaborar.Show
End Sub


'Mediante esta macro se realiza la comprobación de que se seleccionaron
'las líneas adecuadas para modifcar, y luego de verificar con el usuario
'se procede a abrir el formulario con la configuracion modificar, la cual
'carga los valores procedentes de la línea seleccionada en los campos
'del formulario

Sub llamar_modificar_form()
    ' Esta macro obtiene la fila seleccionada en la tabla Inventario
    ' y almacena el índice de la fila (dentro de la tabla) en la variable global fila_a_modificar
    
    Dim HojaBD As Worksheet
    Dim tabla As ListObject
    Dim celdaSeleccionada As Range
    Dim filaSeleccionada As ListRow
    
    Set HojaBD = ThisWorkbook.Sheets("Gestión de Almacén")
    Set tabla = HojaBD.ListObjects("INVENTARIO")
    
    ' Verificar que se haya seleccionado una celda dentro de la tabla
    Set celdaSeleccionada = Selection
    
    If Intersect(celdaSeleccionada, tabla.Range) Is Nothing Then
        MsgBox "Por favor, seleccione una celda dentro de la tabla 'Inventario'.", vbExclamation, "Selección Inválida"
        Exit Sub
    End If
    
    ' Obtener el número de fila dentro de la tabla (no el número de fila de la hoja)
    ' ListRows es una colección que comienza en 1
    fila_a_modificar = celdaSeleccionada.Row - tabla.Range.Row
    
    ' Opcional: Mostrar información de depuración
    ' Debug.Print "Fila de hoja: " & celdaSeleccionada.Row
    ' Debug.Print "Primera fila de tabla: " & tabla.Range.Row
    ' Debug.Print "Índice dentro de tabla: " & fila_a_modificar
    
    ' Verificar que el índice está dentro del rango válido
    If fila_a_modificar < 1 Or fila_a_modificar > tabla.ListRows.Count Then
        MsgBox "Error al determinar la fila seleccionada.", vbCritical, "Error"
        Exit Sub
    End If
    
    ' Activar el formulario
    ingresar_form.modificar_registrar_textbox.Value = "M"
    ingresar_form.Show
End Sub


Sub llamar_formulario_facturacion()
    facturacion.Show
End Sub


Sub reinicio_POS()
    Unload facturacion
    MsgBox "Venta realizada con éxito", vbInformation
    facturacion.Show
    'facturacion.Hide
    'facturacion.Show
    'Call generar_id_factura
End Sub


'Esta macro vuelve a setear la funcion en la celda correspondiente para que su valor sea actualizado
'ya que no lo hace de forma automatica
Sub setear_funcion()

    Dim HojaBD As Worksheet
    Dim tabla As ListObject
    Dim fila As ListRow
    Dim tipoCambio As Double
    Dim monedaCosto As String
    Dim monedaVenta As String
    Dim stock As Double
    Dim precioCosto As Double
    Dim precioVenta As Double
    Dim fondo As Double
    Dim Ingreso As Double
    Dim ganancia As Double
    Dim codigo As String
    Dim cantidadInicial As Double
    
    Set HojaBD = ThisWorkbook.Sheets("Gestión de Almacén")
    Set tabla = HojaBD.ListObjects("INVENTARIO")
    
    ' Obtener tipo de cambio de la celda G3
    tipoCambio = HojaBD.Range("G3").Value
    If tipoCambio <= 0 Then tipoCambio = 1 ' Evitar división por cero
    
    ' Recorrer todas las filas de la tabla
    For Each fila In tabla.ListRows
        ' Obtener código y cantidad inicial (col D = índice 4)
        codigo = fila.Range(1).Value
        cantidadInicial = fila.Range(4).Value
        
        ' Calcular stock: Cantidad Inicial - Total Vendido (VALOR ESTÁTICO)
        stock = cantidadInicial - TotalVendido(codigo) - ObtenerMermasProducto(codigo)
        If stock < 0 Then stock = 0
        fila.Range(5).Value = stock ' YA NO escribe la fórmula
        
        precioVenta = fila.Range(6).Value
        monedaVenta = fila.Range(7).Value
        precioCosto = fila.Range(8).Value
        monedaCosto = fila.Range(9).Value
        
        ' Calcular fondo (Columna J = 10)
        ' Fondo = Precio Costo * Stock
        fondo = precioCosto * stock
        fila.Range(10).Value = fondo
        fila.Range(11).Value = monedaCosto ' Moneda Fondo (Columna K = 11)
        
        ' Calcular ingreso esperado (Columna L = 12)
        ' Ingreso = Precio Venta * Stock
        Ingreso = precioVenta * stock
        fila.Range(12).Value = Ingreso
        fila.Range(13).Value = monedaVenta ' Moneda Ingreso (Columna M = 13)
        
        ' Calcular ganancia según las monedas (Columna N = 14)
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
    Next fila

End Sub

' -------------------------------------------------------------------
' Calcula el total vendido de un producto recorriendo Historial de ventas
' Busca el código en columna B, suma cantidades de columna E
' -------------------------------------------------------------------
Private Function TotalVendido(codigo As String) As Double
    Dim wsVentas As Worksheet
    Dim ultima_fila As Long
    Dim celda As Range
    Dim total As Double
    
    Set wsVentas = ThisWorkbook.Sheets("Historial de ventas")
    ultima_fila = wsVentas.Range("B" & Rows.Count).End(xlUp).Row
    
    total = 0
    For Each celda In wsVentas.Range("B5:B" & ultima_fila)
        If CStr(celda.Value) = CStr(codigo) Then
            total = total + wsVentas.Range("E" & celda.Row).Value
        End If
    Next celda
    
    TotalVendido = total
End Function

Sub ModificarReceta()
    Dim ws As Worksheet
    Dim tblRecetas As ListObject
    Dim celdaSeleccionada As Range
    Dim filaSeleccionada As Long
    Dim nombreReceta As String
    
    ' Obtener la hoja y tabla de recetas
    Set ws = ThisWorkbook.Sheets("Elaboración de Productos")
    Set tblRecetas = ws.ListObjects("Recetas")
    
    ' Verificar si la celda está dentro de la tabla de recetas
    On Error Resume Next
    Set celdaSeleccionada = Intersect(Selection, tblRecetas.Range)
    On Error GoTo 0
    
    If celdaSeleccionada Is Nothing Then
        MsgBox "Por favor, seleccione una celda dentro de la tabla de recetas", vbExclamation
        Exit Sub
    End If
    
    ' Obtener el nombre de la receta de la fila seleccionada
    filaSeleccionada = celdaSeleccionada.Row
    nombreReceta = ws.Cells(filaSeleccionada, tblRecetas.ListColumns("Receta").Range.Column).Value
    
    ' Preguntar confirmación
    Dim respuesta As VbMsgBoxResult
    respuesta = MsgBox("¿Está seguro que desea modificar la receta: '" & nombreReceta & "'?", _
                      vbYesNo + vbQuestion, "Confirmar modificación")
    
    If respuesta = vbYes Then
        ' Abrir formulario en modo modificación
        ShowModoModificacion nombreReceta
    End If
End Sub

' --- MÉTODO PARA MOSTRAR EL FORMULARIO EN MODO MODIFICACIÓN ---
Public Sub ShowModoModificacion(nombreReceta As String)
    modoModificacion = True
    recetaAModificar = nombreReceta
    Ingr_Receta.Show
End Sub

Sub llamar_ingresar_nuevo_ingrediente()
    ingresar_ingr_nuevo_form.modificar_ingresar_nuevo_ingr_textbox.Value = "I"
    ingresar_ingr_nuevo_form.Show
End Sub


Sub llamar_modificar_form_ingrediente()

    Dim respuesta As String
    fila_a_modificar = Application.ActiveCell.Row

    If fila_a_modificar > 4 And Range("B" & fila_a_modificar) <> "" Then
    
        respuesta = MsgBox("¿Está seguro que desea modificar el producto seleccionado?", vbYesNo, "Eliminar producto")
        Select Case respuesta
            Case Is = 6
        
            ingresar_ingr_nuevo_form.modificar_ingresar_nuevo_ingr_textbox = "M"
            ingresar_ingr_nuevo_form.Show
            
            Case Is = 7
            Unload ingresar_ingr_nuevo_form
            Exit Sub
        End Select
    Else
        MsgBox "Seleccione la línea con el producto que desea modificar", vbExclamation
    End If
End Sub

Sub llamar_ingresar_ingr_existente()
    ingresar_ingr_existente_form.Show
End Sub

