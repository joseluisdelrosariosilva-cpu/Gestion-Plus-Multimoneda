Attribute VB_Name = "Módulo1"
Public fila_sel As Integer
Public cargo As String
Public pago As Double

Sub call_Nueva_Receta_form()
    Ingr_Receta.Show
End Sub

Sub call_busqueda_form()
    Busqueda.Show
End Sub

Sub call_busqueda_Ventas_form()
    Busqueda_Ventas.Show
End Sub

Sub call_busqueda_Facturacion_form()
    Busqueda_Facturacion.Show
End Sub

Sub call_Gasto_Form()
    Gasto_Form.Show
End Sub

Sub call_Trabajador_Form()
    Trabajador.Show
End Sub

'Macro para modificar un valor en la tabla de trabajadores, primero verifica una seleccion valida
Sub call_Modificar_Form()

    Dim ws As Worksheet
    Dim tbl_sel As ListObject
    Dim selected As Range
    Dim selected_row As ListRow
    Dim response As VbMsgBoxResult
    
    Set ws = ThisWorkbook.Sheets("Gestión de Trabajadores")
    Set tbl_sel = ws.ListObjects("trabajadores")
    Set selected = Selection
    
    If Intersect(selected, tbl_sel.Range) Is Nothing Then
        MsgBox "Por favor seleccione un trabajador dentro de la tabla", vbExclamation, "Selección no válida"
        Exit Sub
    End If
    
    Set selected_row = tbl_sel.ListRows(selected.Row - tbl_sel.HeaderRowRange.Row)
    response = MsgBox("Desea modificar el trabajador: " & selected_row.Range(1).Value & "?", _
                                    vbYesNo + vbQuestion, "Confirmar modificación")
    
    If response = vbYes Then
        fila_sel = selected_row.Index
        Trabajador_Mod.Show
    End If
    

End Sub

'Macro para eliminar un valor en la tabla de gastos, primero verifica una seleccion valida
Sub eliminar_gasto()

    Dim ws As Worksheet
    Dim tbl_sel As ListObject
    Dim selected As Range
    Dim selected_row As ListRow
    Dim response As VbMsgBoxResult
    
    Set ws = ThisWorkbook.Sheets("Historial de Gastos")
    Set tbl_sel = ws.ListObjects("Gastos")
    Set selected = Selection
    
    If Intersect(selected, tbl_sel.Range) Is Nothing Then
        MsgBox "Por favor seleccione un gasto dentro de la tabla", vbExclamation, "Selección no válida"
        Exit Sub
    End If
    
    Set selected_row = tbl_sel.ListRows(selected.Row - tbl_sel.HeaderRowRange.Row)
    response = MsgBox("Desea eliminar el gasto: " & selected_row.Range(3).Value & "?", _
                                    vbYesNo + vbQuestion, "Confirmar eliminación")
    
    If response = vbYes Then
        ws.Unprotect password:="theSheetProtectpassisthe011rd"
        selected_row.Delete
        ws.Protect password:="theSheetProtectpassisthe011rd", userinterfaceonly:=True
        MsgBox "Gasto eliminado correctamente", vbInformation, "Listo!"
    End If


End Sub

'Macro para eliminar un valor en la tabla de trabajadores, primero verifica una seleccion valida
Sub eliminar_trabajador()

    Dim ws As Worksheet
    Dim tbl_sel As ListObject
    Dim selected As Range
    Dim selected_row As ListRow
    Dim response As VbMsgBoxResult
    
    Set ws = ThisWorkbook.Sheets("Gestión de Trabajadores")
    Set tbl_sel = ws.ListObjects("trabajadores")
    Set selected = Selection
    
    If Intersect(selected, tbl_sel.Range) Is Nothing Then
        MsgBox "Por favor seleccione un trabajador dentro de la tabla", vbExclamation, "Selección no válida"
        Exit Sub
    End If
    
    Set selected_row = tbl_sel.ListRows(selected.Row - tbl_sel.HeaderRowRange.Row)
    response = MsgBox("Desea eliminar el trabajador: " & selected_row.Range(1).Value & "?", _
                                    vbYesNo + vbQuestion, "Confirmar eliminación")
    
    If response = vbYes Then
        ws.Unprotect password:="theSheetProtectpassisthe011rd"
        selected_row.Delete
        ws.Protect password:="theSheetProtectpassisthe011rd", userinterfaceonly:=True
        MsgBox "Trabajador eliminado correctamente", vbInformation, "Listo!"
    End If


End Sub

'Macro para pagar un salario de la tabla de trabajadores, primero verifica una seleccion valida
Sub pagar_trabajador()

    Dim tbl_sel As ListObject
    Dim selected As Range
    Dim selected_row As ListRow
    Dim response As VbMsgBoxResult
    
    
    Set tbl_sel = ThisWorkbook.Sheets("Gestión de Trabajadores").ListObjects("trabajadores")
    Set selected = Selection
    
    If Intersect(selected, tbl_sel.Range) Is Nothing Then
        MsgBox "Por favor seleccione un trabajador dentro de la tabla", vbExclamation, "Selección no válida"
        Exit Sub
    End If
    
    Set selected_row = tbl_sel.ListRows(selected.Row - tbl_sel.HeaderRowRange.Row)
    
    If ThisWorkbook.Worksheets("Historial de ventas").Range("J5") * selected_row.Range(3) < selected_row.Range(4) Then
        pago = selected_row.Range(4)
    Else
        pago = ThisWorkbook.Worksheets("Historial de ventas").Range("J5") * selected_row.Range(3)
    End If
    
    response = MsgBox("Desea pagar al trabajador: " & selected_row.Range(1).Value & "?" & vbCrLf & "A pagar: " & pago & " CUP", _
                                    vbYesNo + vbQuestion, "Confirmar eliminación")
    
    If response = vbYes Then
        cargo = selected_row.Range(2)
        Agregar_sal
        MsgBox "Salario pagado e incluido en la tabla de gastos correctamente", vbInformation, "Listo!"
    End If

End Sub
'Macro para agregar el salario pagado en la tabla de gastos
Sub Agregar_sal()

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim newrow As ListRow
    
    Set ws = ThisWorkbook.Sheets("Historial de Gastos")
    Set tbl = ws.ListObjects("Gastos")
    
    ws.Unprotect password:="theSheetProtectpassisthe011rd"
    
    Set newrow = tbl.ListRows.Add
    
    With newrow
        .Range(1) = Date
        .Range(2) = "Fijo"
        .Range(3) = "Salario " & cargo
        .Range(4) = CDbl(pago)
    End With
    
    Agregar_Fecha
    
    tbl.Sort.Apply
    
    ws.Protect password:="theSheetProtectpassisthe011rd", userinterfaceonly:=True

End Sub

'Macro para agregar la fecha a la tabla de fechas unicas
Private Sub Agregar_Fecha()

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim newrow As ListRow
    
    Set ws = ThisWorkbook.Sheets("Fechas")
    Set tbl = ws.ListObjects("FechasUnicas")
    
    For Each fila In tbl.ListRows
        If CDate(fila.Range(1)) = Date Then
            Exit Sub
        End If
    Next fila
    
    ws.Unprotect password:="theSheetProtectpassisthe011rd"
    
    Set newrow = tbl.ListRows.Add
    
    newrow.Range(1) = Date
    
    tbl.Sort.Apply
    
    ws.Protect password:="theSheetProtectpassisthe011rd", userinterfaceonly:=True

End Sub

Sub refresh_Dashboard()

ThisWorkbook.Sheets("Dashboard").Unprotect password:="theSheetProtectpassisthe011rd"
    ThisWorkbook.RefreshAll
    ThisWorkbook.Worksheets("Dashboard").Protect password:="theSheetProtectpassisthe011rd", _
        contents:=True, _
        AllowFiltering:=True, _
        AllowSorting:=True, _
        AllowUsingPivotTables:=True, _
        userinterfaceonly:=True

End Sub

Sub verA1()
    'ThisWorkbook.Sheets("Hidden").Range("A1") = Date
    MsgBox ThisWorkbook.Sheets("Hidden").Range("A1")
End Sub

Sub ocultar()
    desprotegerLibro
    'ThisWorkbook.Sheets("Fechas").Visible = True
    ThisWorkbook.Sheets("Fechas").Visible = xlSheetVeryHidden
    protegerLibro
End Sub
