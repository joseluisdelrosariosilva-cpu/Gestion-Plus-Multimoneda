VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ingresarProdExistente 
   Caption         =   "Abastecer Producto Existente"
   ClientHeight    =   3800
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   5480
   OleObjectBlob   =   "ingresarProdExistente.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "ingresarProdExistente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' ============================================================================
' FORMULARIO: ingresarProdExistente
' PROPÓSITO:  Interfaz para reabastecer productos del inventario.
'              Reemplaza el InputBox de ReabastecerProducto() con un
'              selector visual de producto + cantidad.
' ============================================================================

Private Sub boton_ingresar_covered_Click()
    If Me.Nombre_ingr_ComboBox.Value = "" Then
        MsgBox "Seleccione un producto.", vbExclamation
        Exit Sub
    End If

    If Not IsNumeric(Me.Cantidad_ingresar_producto_existente_TextBox.Value) Or _
       Val(Me.Cantidad_ingresar_producto_existente_TextBox.Value) <= 0 Then
        MsgBox "Debe ingresar una cantidad valida mayor a cero.", vbExclamation
        Exit Sub
    End If

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim fila As ListRow
    Dim producto As String
    Dim cantidad As Double
    Dim stockInicial As Double

    Set ws = ThisWorkbook.Sheets("Gestión de Almacén")
    Set tbl = ws.ListObjects("INVENTARIO")

    producto = Me.Nombre_ingr_ComboBox.Value
    cantidad = Val(Me.Cantidad_ingresar_producto_existente_TextBox.Value)

    For Each fila In tbl.ListRows
        If fila.Range.Cells(1, 2).Value = producto Then
            stockInicial = fila.Range.Cells(1, 4).Value

            fila.Range.Cells(1, 4).Value = stockInicial + cantidad
            fila.Range.Cells(1, 3).Value = Date

            Call setear_funcion

            Unload Me
            MsgBox "Reabastecimiento exitoso." & vbCrLf & _
                   "Producto: " & producto & vbCrLf & _
                   "Cantidad agregada: " & cantidad, vbInformation
            Exit Sub
        End If
    Next fila

    MsgBox "Producto no encontrado en el inventario.", vbExclamation
End Sub

Private Sub boton_ingresar_reposo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.boton_ingresar_covered.ZOrder msoBringToFront
    Me.Label27.ZOrder msoBringToFront
End Sub

Private Sub Cantidad_ingresar_producto_existente_TextBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If (KeyAscii >= 46 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub Label27_Click()
    boton_ingresar_covered_Click
End Sub

Private Sub UserForm_Activate()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim fila As ListRow

    Set ws = ThisWorkbook.Worksheets("Gestión de Almacén")
    Set tbl = ws.ListObjects("INVENTARIO")

    Me.Nombre_ingr_ComboBox.Clear

    ' Cargar TODOS los nombres de producto del inventario
    If Not tbl.DataBodyRange Is Nothing Then
        For Each fila In tbl.ListRows
            If CStr(fila.Range.Cells(1, 2).Value) <> "" Then
                Me.Nombre_ingr_ComboBox.AddItem CStr(fila.Range.Cells(1, 2).Value)
            End If
        Next fila
    End If
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.boton_ingresar_reposo.ZOrder msoBringToFront
    Me.Label27.ZOrder msoBringToFront
End Sub

