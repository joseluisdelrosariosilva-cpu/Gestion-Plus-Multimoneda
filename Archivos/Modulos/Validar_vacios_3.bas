Attribute VB_Name = "Validar_vacios_3"
Public contador3

Sub validar_vacios_ingredientes_existentes()

contador3 = 0


'validar nombre de producto
    If ingresar_ingr_existente_form.Nombre_ingr_ComboBox = "" Then
    ingresar_ingr_existente_form.Nombre_ingr_ComboBox.BackStyle = fmBackStyleOpaque
    ingresar_ingr_existente_form.Nombre_ingr_ComboBox.BackColor = VBA.vbYellow
    contador3 = contador3 + 1
    Else
   ingresar_ingr_existente_form.Nombre_ingr_ComboBox.BackStyle = fmBackStyleTransparent
    End If

'validar cantidad
    If ingresar_ingr_existente_form.Cantidad_ingresar_producto_existente_TextBox = "" Then
    ingresar_ingr_existente_form.Cantidad_ingresar_producto_existente_TextBox.BackStyle = fmBackStyleOpaque
    ingresar_ingr_existente_form.Cantidad_ingresar_producto_existente_TextBox.BackColor = VBA.vbYellow
    contador3 = contador3 + 1
    Else
    ingresar_ingr_existente_form.Cantidad_ingresar_producto_existente_TextBox.BackStyle = fmBackStyleTransparent
    End If

'validar unidad
    If ingresar_ingr_existente_form.Unidad_ComboBox = "" Then
    ingresar_ingr_existente_form.Unidad_ComboBox.BackStyle = fmBackStyleOpaque
    ingresar_ingr_existente_form.Unidad_ComboBox.BackColor = VBA.vbYellow
    contador3 = contador3 + 1
    Else
    ingresar_ingr_existente_form.Unidad_ComboBox.BackStyle = fmBackStyleTransparent
    End If

'validar costo
    If ingresar_ingr_existente_form.Costo_TextBox = "" Then
    ingresar_ingr_existente_form.Costo_TextBox.BackStyle = fmBackStyleOpaque
    ingresar_ingr_existente_form.Costo_TextBox.BackColor = VBA.vbYellow
    contador3 = contador3 + 1
    Else
    ingresar_ingr_existente_form.Costo_TextBox.BackStyle = fmBackStyleTransparent
    End If
End Sub
