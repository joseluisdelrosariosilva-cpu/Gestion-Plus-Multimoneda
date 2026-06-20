Attribute VB_Name = "Generador_codigo"
Public Codigo_Producto As String

Sub Genera_Codigo()
    Dim wsAlmacen As Worksheet
    Dim tblInventario As ListObject
    Dim ultimaFila As Long
    Dim i As Long
    Dim codigoActual As String
    Dim numeroActual As Long
    Dim maxNumero As Long
    
    Set wsAlmacen = ThisWorkbook.Sheets("Gestión de Almacén")
    Set tblInventario = wsAlmacen.ListObjects("INVENTARIO")
    
    ' Inicializar el máximo
    maxNumero = 0
    
    ' Recorrer todas las filas con datos en la tabla
    If tblInventario.ListRows.Count > 0 Then
        For i = 1 To tblInventario.ListRows.Count
            With tblInventario.ListRows(i)
                codigoActual = .Range.Cells(1, 1).Value
                
                If codigoActual <> "" Then
                    ' Extraer la parte numérica después de "Pr_"
                    If Len(codigoActual) > 3 Then
                        On Error Resume Next
                        numeroActual = CLng(Mid(codigoActual, 4))
                        On Error GoTo 0
                        
                        If numeroActual > maxNumero Then
                            maxNumero = numeroActual
                        End If
                    End If
                End If
            End With
        Next i
    End If
    
    ' Generar nuevo código
    maxNumero = maxNumero + 1
    Codigo_Producto = "Pr_" & Format(maxNumero, "00000")
End Sub
