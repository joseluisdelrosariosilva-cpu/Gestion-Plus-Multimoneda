Attribute VB_Name = "Validar_campos_vacios"
Public contador

Sub validar_vacios()

contador = 0


'validar nombre de producto
    If ingresar_form.nombre_de_producto_TextBox = "" Then
'    ingresar_form.nombre_de_producto_TextBox.BackStyle = fmBackStyleOpaque
'    ingresar_form.nombre_de_producto_TextBox.BackColor = VBA.vbYellow
    contador = contador + 1
    Else
   ' ingresar_form.nombre_de_producto_TextBox.BackStyle = fmBackStyleTransparent
    End If

'validar fecha de entrada
    If ingresar_form.fecha_de_entrada_TextBox = "" Then
    'ingresar_form.fecha_de_entrada_TextBox.BackStyle = fmBackStyleOpaque
    'ingresar_form.fecha_de_entrada_TextBox.BackColor = VBA.vbYellow
    contador = contador + 1
    Else
    'ingresar_form.fecha_de_entrada_TextBox.BackStyle = fmBackStyleTransparent
    End If

'validar cantidad
    If ingresar_form.cantidad_TextBox = "" Then
    'ingresar_form.cantidad_TextBox.BackStyle = fmBackStyleOpaque
    'ingresar_form.cantidad_TextBox.BackColor = VBA.vbYellow
    contador = contador + 1
    Else
    'ingresar_form.cantidad_TextBox.BackStyle = fmBackStyleTransparent
    End If

'validar precio venta
    If ingresar_form.precio_venta_TextBox = "" Then
    'ingresar_form.precio_venta_TextBox.BackStyle = fmBackStyleOpaque
    'ingresar_form.precio_venta_TextBox.BackColor = VBA.vbYellow
    contador = contador + 1
    Else
    'ingresar_form.precio_venta_TextBox.BackStyle = fmBackStyleTransparent
    End If
    
'validar precio costo
    If ingresar_form.precio_costo_TextBox = "" Then
    'ingresar_form.precio_costo_TextBox.BackStyle = fmBackStyleOpaque
    'ingresar_form.precio_costo_TextBox.BackColor = VBA.vbYellow
    contador = contador + 1
    Else
    'ingresar_form.precio_costo_TextBox.BackStyle = fmBackStyleTransparent
    End If
End Sub

