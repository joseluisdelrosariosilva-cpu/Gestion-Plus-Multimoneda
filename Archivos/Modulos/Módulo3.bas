Attribute VB_Name = "Módulo3"
Public repeticion As Boolean
Public repeticion2 As Boolean
Public repeticion3 As Boolean

Sub comprobar_repeticion()
    Dim i As Integer
    Dim cantidad_de_elementos As Integer
    Dim codigo As String
    
    cantidad_de_elementos = facturacion.productos_facturacion_ListBox.ListCount
    codigo = facturacion.codigo_facturacion_TextBox.Value
    
    repeticion = False
    
    For i = 0 To cantidad_de_elementos - 1
        If codigo = facturacion.productos_facturacion_ListBox.List(i, 0) Then
            repeticion = True
            Exit For
        End If
    Next i
End Sub

Sub protegerLibro()

    ThisWorkbook.Protect password:="theBookProtectpassisthe010nd", structure:=True, Windows:=False
    
End Sub

Sub desprotegerLibro()

    ThisWorkbook.Unprotect password:="theBookProtectpassisthe010nd"
    
End Sub

'Macro para comprobar repeticion en el formualario de ingresar nuevo ingredientes

Sub comprobar_repeticion_2()
    Dim i As Integer
    Dim ultima_fila As Integer
    Dim HojaBD As Worksheet
    Dim nombre_producto As String
    
    Set HojaBD = ThisWorkbook.Sheets("Gestión de Ingredientes")
    ultima_fila = HojaBD.Range("B" & Rows.Count).End(xlUp).Row
    nombre_producto = ingresar_ingr_nuevo_form.Nombre_ingrediente_Nuevo_TextBox.Value
    repeticion2 = False
    
    For i = 5 To ultima_fila
        If nombre_producto = HojaBD.Range("B" & i).Value Then
            repeticion2 = True
            Exit For
        End If
    Next i
End Sub

'Esta macro es para comprobar repeticion en el formulario de modificar
'Se diferencia de la anterior en que esta permite poner el mismo nombre actual
'del elemento que se está modificando

Sub comprobar_repeticion_3()
    Dim i As Integer
    Dim ultima_fila As Integer
    Dim fila_activa As Integer
    Dim nombre_nuevo As String
    Dim nombre_actual As String
    Dim HojaBD As Worksheet
    
    Set HojaBD = ThisWorkbook.Sheets("Gestión de Ingredientes")
    
    ultima_fila = HojaBD.Cells(HojaBD.Rows.Count, "B").End(xlUp).Row
    
    fila_activa = Application.ActiveCell.Row
    nombre_nuevo = ingresar_ingr_nuevo_form.Nombre_ingrediente_Nuevo_TextBox.Value
    nombre_actual = HojaBD.Range("B" & fila_activa).Value
    
    repeticion3 = False
    
    For i = 5 To ultima_fila
        If HojaBD.Range("B" & i).Value = nombre_nuevo And i <> fila_activa Then
            repeticion3 = True
            Exit For
        End If
    Next i
End Sub


Sub eliminar_ingrediente()

    Dim ws As Worksheet
    Dim tbl_sel As ListObject
    Dim selected As Range
    Dim selected_row As ListRow
    Dim response As VbMsgBoxResult
    
    Set ws = ThisWorkbook.Sheets("Gestión de Ingredientes")
    Set tbl_sel = ws.ListObjects("INGREDIENTES")
    Set selected = Selection
    
    If Intersect(selected, tbl_sel.Range) Is Nothing Then
        MsgBox "Por favor seleccione un ingrediente dentro de la tabla", vbExclamation, "Selección no válida"
        Exit Sub
    End If
    
    Set selected_row = tbl_sel.ListRows(selected.Row - tbl_sel.HeaderRowRange.Row)
    response = MsgBox("Desea eliminar el ingrediente: " & selected_row.Range(1).Value & "?", _
                                    vbYesNo + vbQuestion, "Confirmar eliminación")
    
    If response = vbYes Then
        ws.Unprotect password:="theSheetProtectpassisthe011rd"
        selected_row.Delete
        ws.Protect password:="theSheetProtectpassisthe011rd", userinterfaceonly:=True
        MsgBox "Ingrediente eliminado correctamente", vbInformation, "Listo!"
    End If
End Sub

