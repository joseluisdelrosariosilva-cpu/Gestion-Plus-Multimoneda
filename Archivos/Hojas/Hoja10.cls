VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Worksheet_Change(ByVal Target As Range)

    On Error Resume Next  ' Si hay error, continúa
    
    Dim celda_c As Range
    Dim rango_c As Range
    Dim valorD As Double
    Dim valorC As Double
    Dim porcentaje As Double
    
    Set rango_c = Range("D5:D1000")
    
    ' Recorre SOLO las celdas cambiadas en Columna D
    For Each celda_c In rango_c
        ' Si el cambio NO fue en Columna D, ignora
        If celda_c.Column <> 4 Then GoTo SiguienteCelda
        
        ' Si la celda está vacía, quita color y sigue
        If IsEmpty(celda_c) Or celda_c.Value = "" Then
            celda_c.Interior.Color = vbBlack
            GoTo SiguienteCelda
        End If
        
        ' Obtener valores de Columna D (actual) y Columna C (referencia)
        valorD = celda_c.Value          ' Columna D
        valorC = celda_c.Offset(0, -1).Value  ' Columna C (izquierda de D)
        
        ' Si no hay referencia o es cero, quita color
        If valorC = 0 Then
            celda_c.Interior.Color = vbBlack
            GoTo SiguienteCelda
        End If
        
        ' Calcular porcentaje
        porcentaje = valorD / valorC
        
        ' Aplicar colores SEGÚN TUS ESCALAS
        If porcentaje >= 0.8 Then
            celda_c.Interior.Color = RGB(0, 176, 80)   ' Verde - 80% o más
        ElseIf porcentaje >= 0.5 Then
            celda_c.Interior.Color = RGB(255, 255, 0)   ' Amarillo - 50% a 79.9%
        ElseIf porcentaje >= 0.2 Then
            celda_c.Interior.Color = RGB(255, 192, 0)   ' Naranja - 20% a 49.9%
        Else
            celda_c.Interior.Color = RGB(255, 0, 0)     ' Rojo - Menos de 20%
        End If
        
SiguienteCelda:
        Next celda_c

    Dim wsIng As Worksheet
    Dim wsRec As Worksheet
    Dim wsIngRec As Worksheet
    Dim tblIng As ListObject
    Dim tblRecetas As ListObject
    Dim tblIngRec As ListObject
    Dim rngChanged As Range
    Dim celda As Range
    Dim ingrediente As String
    Dim recetasAfectadas As Collection
    Dim receta As Variant
    Dim recetaNombre As String
    Dim cantidadLote As Double
    Dim precioLote As Double
    Dim precioUnidad As Double
    Dim costototal As Double
    Dim costoUnidad As Double
    Dim gananciaLote As Double
    Dim gananciaUnidad As Double
    Dim i As Long
    Dim filaIngRec As ListRow
    Dim filaReceta As ListRow
    
    ' Verificar que el cambio sea en la columna de precio (columna 5) de la tabla INGREDIENTES
    Set wsIng = ThisWorkbook.Sheets("Gestión de Ingredientes")
    Set tblIng = wsIng.ListObjects("INGREDIENTES")
    
    ' Si el cambio no está en la tabla, salir
    Set rngChanged = Intersect(Target, tblIng.DataBodyRange)
    If rngChanged Is Nothing Then Exit Sub
    
    ' Verificar si el cambio es en la columna 5 (precio por unidad base)
    ' Asumimos que la columna 5 es la de precio. Ajusta si es necesario.
    Dim colPrecio As Long
    colPrecio = tblIng.ListColumns("Precio/U.Base").Index  ' Ajusta el nombre de la columna si es diferente
    
    ' Si no se puede obtener la columna por nombre, usar el índice 5
    If colPrecio = 0 Then colPrecio = 5
    
    ' Filtrar solo las celdas cambiadas en la columna de precio
    Set rngChanged = Intersect(rngChanged, tblIng.ListColumns(colPrecio).Range)
    If rngChanged Is Nothing Then Exit Sub
    
    ' Deshabilitar eventos para evitar bucles
    Application.EnableEvents = False
    
    ' Recorrer cada celda cambiada
    For Each celda In rngChanged
        ' Obtener el nombre del ingrediente (columna 1 de la misma fila)
        ingrediente = tblIng.DataBodyRange.Cells(celda.Row - tblIng.HeaderRowRange.Row, 1).Value
        
        ' Buscar todas las recetas que usan este ingrediente
        Set recetasAfectadas = ObtenerRecetasPorIngrediente(ingrediente)
        
        ' Si hay recetas afectadas, actualizarlas
        If recetasAfectadas.Count > 0 Then
            Set wsRec = ThisWorkbook.Sheets("Elaboración de Productos")
            Set wsIngRec = ThisWorkbook.Sheets("Ingredientes")
            Set tblRecetas = wsRec.ListObjects("Recetas")
            Set tblIngRec = wsIngRec.ListObjects("Ingredientes_Receta")
            
            For Each receta In recetasAfectadas
                recetaNombre = receta
                
                ' Buscar la receta en la tabla Recetas
                Dim filaEncontrada As ListRow
                Set filaEncontrada = Nothing
                For Each filaReceta In tblRecetas.ListRows
                    If filaReceta.Range.Cells(1, 1).Value = recetaNombre Then
                        Set filaEncontrada = filaReceta
                        Exit For
                    End If
                Next
                
                If Not filaEncontrada Is Nothing Then
                    ' Obtener datos actuales de la receta
                    cantidadLote = filaEncontrada.Range.Cells(1, 2).Value   ' Cantidad/Lote
                    precioLote = filaEncontrada.Range.Cells(1, 4).Value     ' Precio/Lote
                    precioUnidad = filaEncontrada.Range.Cells(1, 7).Value   ' Precio/Unidad
                    
                    ' Recalcular el costo total de la receta (por lote)
                    costototal = CalcularCostoReceta(recetaNombre)
                    
                    ' Calcular los demás valores
                    costoUnidad = costototal / cantidadLote
                    gananciaLote = precioLote - costototal
                    gananciaUnidad = precioUnidad - costoUnidad
                    
                    ' Actualizar la tabla Recetas
                    filaEncontrada.Range.Cells(1, 3).Value = costototal   ' Inversión/Lote
                    filaEncontrada.Range.Cells(1, 5).Value = gananciaLote ' Ganancia/Lote
                    filaEncontrada.Range.Cells(1, 6).Value = costoUnidad  ' Inversión/Unidad
                    filaEncontrada.Range.Cells(1, 8).Value = gananciaUnidad ' Ganancia/Unidad
                End If
            Next receta
        End If
    Next celda
    
    ' Rehabilitar eventos
    Application.EnableEvents = True
End Sub

' Función para obtener todas las recetas que usan un ingrediente específico
Private Function ObtenerRecetasPorIngrediente(ingrediente As String) As Collection
    Dim wsIngRec As Worksheet
    Dim tblIngRec As ListObject
    Dim fila As ListRow
    Dim Recetas As New Collection
    Dim recetaNombre As String
    
    Set wsIngRec = ThisWorkbook.Sheets("Ingredientes")
    Set tblIngRec = wsIngRec.ListObjects("Ingredientes_Receta")
    
    For Each fila In tblIngRec.ListRows
        If fila.Range.Cells(1, 2).Value = ingrediente Then  ' Columna 2: Ingrediente
            recetaNombre = fila.Range.Cells(1, 1).Value     ' Columna 1: Receta
            ' Agregar la receta a la colección (sin duplicados)
            On Error Resume Next
            Recetas.Add recetaNombre, recetaNombre
            On Error GoTo 0
        End If
    Next
    
    Set ObtenerRecetasPorIngrediente = Recetas
End Function

' Función para calcular el costo total de una receta (por lote)
Private Function CalcularCostoReceta(recetaNombre As String) As Double
    Dim wsIngRec As Worksheet
    Dim wsIng As Worksheet
    Dim tblIngRec As ListObject
    Dim tblIng As ListObject
    Dim fila As ListRow
    Dim ingrediente As String
    Dim Cantidad As Double
    Dim Unidad As String
    Dim cantidadBase As Double
    Dim precioUnitario As Double
    Dim costototal As Double
    
    Set wsIngRec = ThisWorkbook.Sheets("Ingredientes")
    Set wsIng = ThisWorkbook.Sheets("Gestión de Ingredientes")
    Set tblIngRec = wsIngRec.ListObjects("Ingredientes_Receta")
    Set tblIng = wsIng.ListObjects("INGREDIENTES")
    
    costototal = 0
    
    ' Recorrer la tabla Ingredientes_Receta para encontrar los ingredientes de la receta
    For Each fila In tblIngRec.ListRows
        If fila.Range.Cells(1, 1).Value = recetaNombre Then  ' Columna 1: Receta
            ingrediente = fila.Range.Cells(1, 2).Value      ' Columna 2: Ingrediente
            Cantidad = fila.Range.Cells(1, 3).Value         ' Columna 3: Cantidad
            Unidad = fila.Range.Cells(1, 4).Value           ' Columna 4: Unidad
            
            ' Convertir la cantidad a unidad base
            cantidadBase = ConvertirUnidad(ingrediente, Cantidad, Unidad)
            
            ' Obtener el precio por unidad base del ingrediente
            precioUnitario = ObtenerPrecioIngrediente(ingrediente)
            
            ' Acumular el costo
            costototal = costototal + (cantidadBase * precioUnitario)
        End If
    Next
    
    CalcularCostoReceta = costototal
End Function

' Función para obtener el precio por unidad base de un ingrediente
Private Function ObtenerPrecioIngrediente(ingrediente As String) As Double
    Dim wsIng As Worksheet
    Dim tblIng As ListObject
    Dim fila As ListRow
    
    Set wsIng = ThisWorkbook.Sheets("Gestión de Ingredientes")
    Set tblIng = wsIng.ListObjects("INGREDIENTES")
    
    For Each fila In tblIng.ListRows
        If fila.Range.Cells(1, 1).Value = ingrediente Then  ' Columna 1: Ingrediente
            ' Asumimos que el precio está en la columna 5
            ObtenerPrecioIngrediente = fila.Range.Cells(1, 5).Value
            Exit Function
        End If
    Next
    
    ' Si no se encuentra, devolver 0
    ObtenerPrecioIngrediente = 0
End Function
