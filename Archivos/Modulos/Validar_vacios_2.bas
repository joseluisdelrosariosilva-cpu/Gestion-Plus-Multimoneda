Attribute VB_Name = "Validar_vacios_2"
Public contador2

Sub validar_vacios_ingredientes()

contador2 = 0


'validar nombre de producto
    If ingresar_ingr_nuevo_form.Nombre_ingrediente_Nuevo_TextBox = "" Then
    ingresar_ingr_nuevo_form.Nombre_ingrediente_Nuevo_TextBox.BackStyle = fmBackStyleOpaque
    ingresar_ingr_nuevo_form.Nombre_ingrediente_Nuevo_TextBox.BackColor = VBA.vbYellow
    contador2 = contador2 + 1
    Else
   ingresar_ingr_nuevo_form.Nombre_ingrediente_Nuevo_TextBox.BackStyle = fmBackStyleTransparent
    End If

'validar cantidad
    If ingresar_ingr_nuevo_form.Cantidad_ingresar_nuevo_producto_TextBox = "" Then
    ingresar_ingr_nuevo_form.Cantidad_ingresar_nuevo_producto_TextBox.BackStyle = fmBackStyleOpaque
    ingresar_ingr_nuevo_form.Cantidad_ingresar_nuevo_producto_TextBox.BackColor = VBA.vbYellow
    contador2 = contador2 + 1
    Else
    ingresar_ingr_nuevo_form.Cantidad_ingresar_nuevo_producto_TextBox.BackStyle = fmBackStyleTransparent
    End If

'validar cantidad referencia
    If ingresar_ingr_nuevo_form.Cantidad_referencia_TextBox = "" Then
    ingresar_ingr_nuevo_form.Cantidad_referencia_TextBox.BackStyle = fmBackStyleOpaque
    ingresar_ingr_nuevo_form.Cantidad_referencia_TextBox.BackColor = VBA.vbYellow
    contador2 = contador2 + 1
    Else
    ingresar_ingr_nuevo_form.Cantidad_referencia_TextBox.BackStyle = fmBackStyleTransparent
    End If

'validar costo total
    If ingresar_ingr_nuevo_form.Costo_Total_Ingr_Nuevo_TextBox = "" Then
    ingresar_ingr_nuevo_form.Costo_Total_Ingr_Nuevo_TextBox.BackStyle = fmBackStyleOpaque
    ingresar_ingr_nuevo_form.Costo_Total_Ingr_Nuevo_TextBox.BackColor = VBA.vbYellow
    contador2 = contador2 + 1
    Else
    ingresar_ingr_nuevo_form.Costo_Total_Ingr_Nuevo_TextBox.BackStyle = fmBackStyleTransparent
    End If
    
'validar Unidad Base
    If ingresar_ingr_nuevo_form.UBase_nuevo_ComboBox.Value = "" Then
    ingresar_ingr_nuevo_form.UBase_nuevo_ComboBox.BackStyle = fmBackStyleOpaque
    ingresar_ingr_nuevo_form.UBase_nuevo_ComboBox.BackColor = VBA.vbYellow
    contador2 = contador2 + 1
    Else
    ingresar_ingr_nuevo_form.UBase_nuevo_ComboBox.BackStyle = fmBackStyleTransparent
    End If
End Sub
