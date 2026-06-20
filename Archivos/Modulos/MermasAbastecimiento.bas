Attribute VB_Name = "MermasAbastecimiento"
Option Explicit

' ============================================================================
' Obtiene la cantidad total de mermas de un producto recorriendo Gastos
' Busca categoría "Merma" y parsea "NombreProducto x cantidad"
' ============================================================================
Public Function ObtenerMermasProducto(codigo As String) As Double
    Dim wsGastos As Worksheet, wsInv As Worksheet
    Dim tblGastos As ListObject, tblInv As ListObject
    Dim fila As ListRow, filaInv As ListRow
    Dim nombreProducto As String, descripcion As String
    Dim totalMermas As Double, posX As Long, cantidadStr As String

    On Error GoTo ErrorHandler

    ' --- Obtener nombre del producto desde INVENTARIO ---
    Set wsInv = ThisWorkbook.Sheets("Gestión de Almacén")
    Set tblInv = wsInv.ListObjects("INVENTARIO")
    nombreProducto = ""

    For Each filaInv In tblInv.ListRows
        If CStr(filaInv.Range(1).Value) = codigo Then
            nombreProducto = CStr(filaInv.Range(2).Value)
            Exit For
        End If
    Next filaInv

    If nombreProducto = "" Then
        ObtenerMermasProducto = 0
        Exit Function
    End If

    ' --- Buscar en Historial de Gastos ---
    Set wsGastos = ThisWorkbook.Sheets("Historial de Gastos")
    Set tblGastos = wsGastos.ListObjects("Gastos")
    totalMermas = 0

    For Each fila In tblGastos.ListRows
        If CStr(fila.Range(2).Value) = "Merma" Then
            descripcion = CStr(fila.Range(3).Value)

            ' Formato esperado: "NombreProducto x cantidad"
            If InStr(1, descripcion, nombreProducto & " x ") = 1 Then
                posX = InStrRev(descripcion, " x ")
                If posX > 0 Then
                    cantidadStr = Trim(Mid(descripcion, posX + 3))
                    If IsNumeric(cantidadStr) Then
                        totalMermas = totalMermas + CDbl(cantidadStr)
                    End If
                End If
            End If
        End If
    Next fila

    ObtenerMermasProducto = totalMermas
    Exit Function

ErrorHandler:
    ObtenerMermasProducto = 0
End Function

' ============================================================================
' Registrar Merma manual — BOTÓN
' ============================================================================
Public Sub RegistrarMerma()
    Dim ws As Worksheet, tbl As ListObject, rng As Range, filaSel As Range
    Dim filaTabla As Long, Cantidad As Double, monto As Double
    Dim codigo As String, nombre As String, monedaCosto As String
    Dim stock As Double, precioCosto As Double, inputStr As String

    Set ws = ThisWorkbook.Sheets("Gestión de Almacén")
    Set tbl = ws.ListObjects("INVENTARIO")

    ' --- Validar selección dentro de la tabla ---
    If tbl.DataBodyRange Is Nothing Then
        MsgBox "No hay productos en el almacén.", vbExclamation: Exit Sub
    End If
    Set rng = Intersect(Selection, tbl.DataBodyRange)
    If rng Is Nothing Then
        MsgBox "Debe seleccionar una fila de la tabla INVENTARIO.", vbExclamation: Exit Sub
    End If

    Set filaSel = rng.Rows(1)
    filaTabla = filaSel.Row - tbl.HeaderRowRange.Row

    ' --- Leer datos del producto ---
    codigo = CStr(tbl.DataBodyRange.Cells(filaTabla, 1).Value)
    nombre = CStr(tbl.DataBodyRange.Cells(filaTabla, 2).Value)
    stock = CDbl(tbl.DataBodyRange.Cells(filaTabla, 5).Value)
    precioCosto = CDbl(tbl.DataBodyRange.Cells(filaTabla, 8).Value)
    monedaCosto = CStr(tbl.DataBodyRange.Cells(filaTabla, 9).Value)

    If stock <= 0 Then
        MsgBox "El producto no tiene stock disponible.", vbExclamation: Exit Sub
    End If

    ' --- Solicitar cantidad ---
    inputStr = InputBox("Ingrese la cantidad de merma para " & nombre & _
                        " (Stock disponible: " & stock & "):", "Registrar Merma", "1")
    If inputStr = "" Then Exit Sub
    If Not IsNumeric(inputStr) Then
        MsgBox "Cantidad inválida. Ingrese un número.", vbExclamation: Exit Sub
    End If

    Cantidad = CDbl(inputStr)
    If Cantidad <= 0 Then
        MsgBox "La cantidad debe ser mayor a 0.", vbExclamation: Exit Sub
    End If
    If Cantidad > stock Then
        MsgBox "La cantidad no puede superar el stock disponible (" & stock & ").", vbExclamation: Exit Sub
    End If

    ' --- Registrar gasto como merma ---
    monto = precioCosto * Cantidad
    Call AgregarGastoMerma(Date, nombre & " x " & Cantidad, monto, monedaCosto)

    ' --- Recalcular stock (setear_funcion actualizado) ---
    Call setear_funcion

    MsgBox "Merma registrada correctamente: " & nombre & " x " & Cantidad, vbInformation
End Sub

' ============================================================================
' Reabastecer Producto manual — BOTÓN
' ============================================================================
Public Sub ReabastecerProducto()
    Dim ws As Worksheet, tbl As ListObject, rng As Range, filaSel As Range
    Dim filaTabla As Long, Cantidad As Double
    Dim stockActual As Double, stockInicial As Double, nombre As String
    Dim cantidadStr As String

    Set ws = ThisWorkbook.Sheets("Gestión de Almacén")
    Set tbl = ws.ListObjects("INVENTARIO")

    ' --- Validar selección ---
    If tbl.DataBodyRange Is Nothing Then
        MsgBox "No hay productos en el almacén.", vbExclamation: Exit Sub
    End If
    Set rng = Intersect(Selection, tbl.DataBodyRange)
    If rng Is Nothing Then
        MsgBox "Debe seleccionar una fila de la tabla INVENTARIO.", vbExclamation: Exit Sub
    End If

    Set filaSel = rng.Rows(1)
    filaTabla = filaSel.Row - tbl.HeaderRowRange.Row

    nombre = CStr(tbl.DataBodyRange.Cells(filaTabla, 2).Value)
    stockActual = CDbl(tbl.DataBodyRange.Cells(filaTabla, 5).Value)
    stockInicial = CDbl(tbl.DataBodyRange.Cells(filaTabla, 4).Value)

    cantidadStr = InputBox("Ingrese la cantidad a añadir para reabastecer:" & vbCrLf & _
                           "Producto: " & nombre & vbCrLf & _
                           "Stock actual: " & stockActual, _
                           "Reabastecer Producto", "0")
    If cantidadStr = "" Then Exit Sub
    If Not IsNumeric(cantidadStr) Or Val(cantidadStr) <= 0 Then
        MsgBox "Debe ingresar una cantidad válida mayor a cero.", vbExclamation: Exit Sub
    End If

    Cantidad = Val(cantidadStr)

    ' --- Sumar a Cantidad Inicial (col 4) y actualizar Fecha (col 3) ---
    tbl.DataBodyRange.Cells(filaTabla, 4).Value = stockInicial + Cantidad
    tbl.DataBodyRange.Cells(filaTabla, 3).Value = Date

    Call setear_funcion

    MsgBox "Reabastecimiento exitoso." & vbCrLf & _
           "Stock anterior: " & stockActual & vbCrLf & _
           "Cantidad añadida: " & Cantidad & vbCrLf & _
           "Nuevo stock: " & (stockActual + Cantidad), vbInformation
End Sub

' ============================================================================
' Auxiliar: Agregar gasto en tabla Gastos (categoría "Merma")
' ============================================================================
Private Sub AgregarGastoMerma(fecha As Date, descripcion As String, _
                               monto As Double, moneda As String)
    Dim ws As Worksheet, tbl As ListObject, newrow As ListRow

    Set ws = ThisWorkbook.Sheets("Historial de Gastos")
    Set tbl = ws.ListObjects("Gastos")

    ws.Unprotect password:="theSheetProtectpassisthe011rd"
    Set newrow = tbl.ListRows.Add

    With newrow
        .Range(1) = CDate(fecha)
        .Range(2) = "Merma"
        .Range(3) = descripcion
        .Range(4) = CDbl(monto)
        .Range(5) = moneda
    End With

    tbl.Sort.Apply
    ws.Protect password:="theSheetProtectpassisthe011rd", userinterfaceonly:=True
End Sub

