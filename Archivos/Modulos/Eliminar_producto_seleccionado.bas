Attribute VB_Name = "Eliminar_producto_seleccionado"
'Macro para eliminar producto de una fila seleccionada
'Esta macro recoge en una variable la fila seleccionada y comprueba que
'la fila es mayor a 4 y que no está vacía y luego de eso activa un msgbox
'para confirmar que el ususario desea eliminar

Sub eliminar_producto()

Dim fila_a_eliminar As Integer
Dim respuesta As String
fila_a_eliminar = Application.ActiveCell.Row

    If fila_a_eliminar > 4 And Range("B" & fila_a_eliminar) <> "" Then

        respuesta = MsgBox("¿Está seguro que desea eliminar el producto seleccionado?", vbYesNo, "Eliminar producto")
        Select Case respuesta
            Case Is = 6
            Range("B" & fila_a_eliminar).EntireRow.Delete
            Case Is = 7
            Exit Sub
        End Select
    Else
        MsgBox "Seleccione la línea con el producto que desea eliminar"
    End If
End Sub

Sub EliminarReceta()
    Dim wsRecetas As Worksheet
    Dim wsIngredientes As Worksheet
    Dim tblRecetas As ListObject
    Dim tblIngredientesReceta As ListObject
    Dim celdaSeleccionada As Range
    Dim rngBusqueda As Range
    Dim celdaEncontrada As Range
    Dim nombreReceta As String
    Dim respuesta As VbMsgBoxResult
    Dim filaTabla As Long
    Dim filaIngrediente As ListRow
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    ' Obtener hojas de trabajo
    Set wsRecetas = ThisWorkbook.Sheets("Elaboración de Productos")
    Set wsIngredientes = ThisWorkbook.Sheets("Ingredientes")
    
    ' Obtener tablas
    Set tblRecetas = wsRecetas.ListObjects("Recetas")
    Set tblIngredientesReceta = wsIngredientes.ListObjects("Ingredientes_Receta")
    
    ' Verificar que la celda activa esté dentro de la tabla Recetas
    Set celdaSeleccionada = ActiveCell
    
    ' Verificar que estamos en la hoja correcta
    If Not celdaSeleccionada.Parent.Name = wsRecetas.Name Then
        MsgBox "Por favor, seleccione una celda en la hoja 'Elaboración de Productos'.", vbExclamation
        Exit Sub
    End If
    
    ' Verificar que la celda esté dentro de la tabla
    If Intersect(celdaSeleccionada, tblRecetas.Range) Is Nothing Then
        MsgBox "Por favor, seleccione una celda dentro de la tabla de recetas.", vbExclamation
        Exit Sub
    End If
    
    ' Obtener el nombre de la receta usando la fila seleccionada
    Dim filaSeleccionada As Long
    filaSeleccionada = celdaSeleccionada.Row
    
    ' Verificar que no sea la fila de encabezado
    If filaSeleccionada <= tblRecetas.HeaderRowRange.Row Then
        MsgBox "Por favor, seleccione una celda en una fila de datos (no en el encabezado).", vbExclamation
        Exit Sub
    End If
    
    ' Calcular la fila dentro de la tabla
    filaTabla = filaSeleccionada - tblRecetas.HeaderRowRange.Row
    
    ' Obtener el nombre de la receta
    nombreReceta = tblRecetas.DataBodyRange.Cells(filaTabla, 1).Value
    
    ' Confirmar con el usuario
    respuesta = MsgBox("¿Está seguro que desea eliminar la receta: '" & nombreReceta & "'?" & vbCrLf & _
                       "Esta acción no se puede deshacer.", _
                       vbYesNo + vbExclamation + vbDefaultButton2, "Confirmar eliminación")
    
    If respuesta = vbNo Then
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    wsRecetas.Unprotect password:="theSheetProtectpassisthe011rd"
    wsIngredientes.Unprotect password:="theSheetProtectpassisthe011rd"
    
    ' --- ELIMINAR INGREDIENTES DE LA RECETA ---
    ' Usar un bucle inverso para eliminar ingredientes
    For i = tblIngredientesReceta.ListRows.Count To 1 Step -1
        Set filaIngrediente = tblIngredientesReceta.ListRows(i)
        If filaIngrediente.Range.Cells(1, 1).Value = nombreReceta Then
            filaIngrediente.Delete
        End If
    Next i
    
    ' --- ELIMINAR LA RECETA ---
    ' Eliminar la fila de la tabla Recetas
    tblRecetas.ListRows(filaTabla).Delete
    
    tblRecetas.Sort.Apply
    wsRecetas.Protect password:="theSheetProtectpassisthe011rd", userinterfaceonly:=True
    
    tblIngredientesReceta.Sort.Apply
    wsIngredientes.Protect password:="theSheetProtectpassisthe011rd", userinterfaceonly:=True
    
    Application.ScreenUpdating = True
    
    ' Mensaje de confirmación
    MsgBox "Receta '" & nombreReceta & "' eliminada exitosamente.", vbInformation
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Error al eliminar la receta: " & Err.Description, vbCritical
End Sub

