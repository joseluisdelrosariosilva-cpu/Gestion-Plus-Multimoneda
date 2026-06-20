Attribute VB_Name = "Conversor"
Public Function ConvertirUnidad(ingrediente As String, Cantidad As Double, UnidadOrigen As String) As Double
    Dim tblConversiones As ListObject
    Dim fila As ListRow
    Dim UnidadBase As String
    
    ' Obtener la tabla de conversiones (cambia el nombre según tu tabla)
    Set tblConversiones = ThisWorkbook.Sheets("Gestión de Ingredientes").ListObjects("INGREDIENTES")

    ' Buscar el ingrediente en la tabla
    For Each fila In tblConversiones.ListRows
        If fila.Range.Cells(1, 1).Value = ingrediente Then
            UnidadBase = fila.Range.Cells(1, 4).Value ' Columna 4 = Unidad Base
            Exit For
        End If
    Next fila
    
    ' Si la unidad origen es "u", no convertir
    If LCase(UnidadOrigen) = "u" Then
        ConvertirUnidad = Cantidad
        Exit Function
    End If
    
    ' Convertir a unidad base
    ConvertirUnidad = ConvertirABase(Cantidad, UnidadOrigen, UnidadBase)
End Function

Private Function ConvertirABase(Cantidad As Double, UnidadOrigen As String, UnidadBase As String) As Double
    Dim CantidadEnBase As Double
    
    ' Primero convertir a la unidad base común
    Select Case LCase(UnidadOrigen)
        ' Conversiones de masa
        Case "kg"
            CantidadEnBase = Cantidad * 1000 ' a gramos
        Case "g"
            CantidadEnBase = Cantidad
        Case "mg"
            CantidadEnBase = Cantidad / 1000
        Case "lb"
            CantidadEnBase = Cantidad * 453.592 ' libras a gramos
        Case "oz"
            CantidadEnBase = Cantidad * 28.3495 ' onzas a gramos
            
        ' Conversiones de volumen
        Case "l"
            CantidadEnBase = Cantidad * 1000 ' a mL
        Case "ml"
            CantidadEnBase = Cantidad
    End Select
    
    ' Ahora convertir de la base común a la unidad base específica
    Select Case LCase(UnidadBase)
        ' Desde gramos a...
        Case "kg"
            ConvertirABase = CantidadEnBase / 1000
        Case "g"
            ConvertirABase = CantidadEnBase
        Case "mg"
            ConvertirABase = CantidadEnBase * 1000
        Case "lb"
            ConvertirABase = CantidadEnBase / 453.592
        Case "oz"
            ConvertirABase = CantidadEnBase / 28.3495
            
        ' Desde mL a...
        Case "l"
            ConvertirABase = CantidadEnBase / 1000
        Case "ml"
            ConvertirABase = CantidadEnBase
            
        Case "u"
            ConvertirABase = Cantidad ' Si la base es unidad, mantener igual
    End Select
End Function
