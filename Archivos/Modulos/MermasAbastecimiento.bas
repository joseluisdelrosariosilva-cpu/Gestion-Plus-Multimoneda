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
    merma_form.Show
End Sub

' ============================================================================
' Reabastecer Producto manual — BOTÓN
' ============================================================================
Public Sub ReabastecerProducto()
    ingresarProdExistente.Show
End Sub

' ============================================================================
' Auxiliar: Agregar gasto en tabla Gastos (categoría "Merma")
' ============================================================================
Public Sub AgregarGastoMerma(fecha As Date, descripcion As String, _
                               monto As Double, Moneda As String)
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
        .Range(5) = Moneda
    End With

    tbl.Sort.Apply
    ws.Protect password:="theSheetProtectpassisthe011rd", UserInterfaceOnly:=True
End Sub

