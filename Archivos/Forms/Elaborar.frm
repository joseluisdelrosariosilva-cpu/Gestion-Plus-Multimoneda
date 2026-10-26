VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Elaborar 
   Caption         =   "Elaborar Producto"
   ClientHeight    =   5240
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   15060
   OleObjectBlob   =   "Elaborar.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Elaborar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' --- VARIABLES DEL FORMULARIO ---
Private recetaSeleccionada As String
Private cantidadPorLote As Double
Private inversionPorLote As Double
Private precioPorUnidad As Double
Private producto As String

Private Sub CargarRecetas()
    Dim ws As Worksheet
    Dim tblRecetas As ListObject
    Dim fila As ListRow
    
    Set ws = ThisWorkbook.Sheets("Elaboración de Productos")
    Set tblRecetas = ws.ListObjects("Recetas")
    
    Me.Recetas.Clear
    Me.Recetas.AddItem "" ' Opción vacía
    
    For Each fila In tblRecetas.ListRows
        Me.Recetas.AddItem fila.Range.Cells(1, 1).Value
    Next fila
End Sub

Private Sub Btn_Active_Click()
    Dim wsIngredientes As Worksheet
    Dim wsInventario As Worksheet
    Dim tblIngredientes As ListObject
    Dim tblInventario As ListObject
    Dim filaIngrediente As ListRow
    Dim nuevaFilaInventario As ListRow
    Dim i As Long
    Dim ingredienteNombre As String
    Dim cantidadNecesaria As Double
    Dim unidadMedida As String
    Dim CantidadEnBase As Double
    Dim stockActual As Double
    Dim fondoActual As Double
    Dim suficienteStock As Boolean
    Dim numLotes As Double
    Dim codigoProducto As String
    Dim cantidadElaborada As Double
    Dim inversionTotal As Double
    Dim ingresoEsperado As Double
    Dim gananciaEsperada As Double
    Dim costoPorUnidad As Double
    
    ' Validar que haya una receta seleccionada
    If Me.Recetas.Value = "" Then
        MsgBox "Por favor, seleccione una receta.", vbExclamation
        Me.Recetas.SetFocus
        Exit Sub
    End If
    
    ' Validar que haya un número de lotes válido
    If Me.Lotes.Value = "" Or Not IsNumeric(Me.Lotes.Value) Then
        MsgBox "Por favor, ingrese una cantidad válida de lotes.", vbExclamation
        Me.Lotes.SetFocus
        Exit Sub
    End If
    
    numLotes = CDbl(Me.Lotes.Value)
    
    ' Validar que haya al menos un ingrediente en la lista
    If Me.Ingredientes.ListCount = 0 Then
        MsgBox "La receta seleccionada no tiene ingredientes.", vbExclamation
        Exit Sub
    End If
    
    ' Obtener hojas de trabajo
    Set wsIngredientes = ThisWorkbook.Sheets("Gestión de Ingredientes")
    Set wsInventario = ThisWorkbook.Sheets("Gestión de Almacén")
    
    ' Obtener tablas
    Set tblIngredientes = wsIngredientes.ListObjects("INGREDIENTES")
    Set tblInventario = wsInventario.ListObjects("INVENTARIO")
    
    ' Variable para controlar si hay suficiente stock
    suficienteStock = True
    Dim ingredientesFaltantes As String
    ingredientesFaltantes = ""
    
    ' PASO 1: VALIDAR STOCK DE TODOS LOS INGREDIENTES
    Application.ScreenUpdating = False
    
    For i = 0 To Me.Ingredientes.ListCount - 1
        ingredienteNombre = Me.Ingredientes.List(i, 0)
        cantidadNecesaria = CDbl(Me.Ingredientes.List(i, 1))
        unidadMedida = Me.Ingredientes.List(i, 2)
        
        ' Buscar el ingrediente en la tabla INGREDIENTES
        Dim encontrado As Boolean
        encontrado = False
        
        For Each filaIngrediente In tblIngredientes.ListRows
            If filaIngrediente.Range.Cells(1, 1).Value = ingredienteNombre Then
                encontrado = True
                
                ' Convertir la cantidad necesaria a unidad base
                CantidadEnBase = ConvertirUnidad(ingredienteNombre, cantidadNecesaria, unidadMedida)
                
                ' Obtener el stock actual (columna 3)
                stockActual = filaIngrediente.Range.Cells(1, 3).Value
                
                ' Verificar si hay suficiente stock
                If stockActual < CantidadEnBase Then
                    suficienteStock = False
                    ingredientesFaltantes = ingredientesFaltantes & "- " & ingredienteNombre & ": " & _
                                            "Necesario: " & Format(CantidadEnBase, "0.00") & " " & _
                                            filaIngrediente.Range.Cells(1, 4).Value & _
                                            " | Disponible: " & Format(stockActual, "0.00") & _
                                            filaIngrediente.Range.Cells(1, 4).Value & vbCrLf
                End If
                
                Exit For
            End If
        Next filaIngrediente
        
        If Not encontrado Then
            MsgBox "El ingrediente '" & ingredienteNombre & "' no existe en la tabla de ingredientes.", vbExclamation
            Application.ScreenUpdating = True
            Exit Sub
        End If
    Next i
    
    ' Si no hay suficiente stock, mostrar mensaje y salir
    If Not suficienteStock Then
        Application.ScreenUpdating = True
        MsgBox "Stock insuficiente para los siguientes ingredientes:" & vbCrLf & vbCrLf & _
               ingredientesFaltantes & vbCrLf & _
               "Por favor, reponga el stock y vuelva a intentar.", vbExclamation
        Exit Sub
    End If
    
    ' PASO 2: ACTUALIZAR STOCK DE INGREDIENTES
    For i = 0 To Me.Ingredientes.ListCount - 1
        ingredienteNombre = Me.Ingredientes.List(i, 0)
        cantidadNecesaria = CDbl(Me.Ingredientes.List(i, 1))
        unidadMedida = Me.Ingredientes.List(i, 2)
        
        ' Buscar el ingrediente en la tabla INGREDIENTES
        For Each filaIngrediente In tblIngredientes.ListRows
            If filaIngrediente.Range.Cells(1, 1).Value = ingredienteNombre Then
                ' Convertir la cantidad necesaria a unidad base
                CantidadEnBase = ConvertirUnidad(ingredienteNombre, cantidadNecesaria, unidadMedida)
                
                ' Restar del stock actual
                stockActual = filaIngrediente.Range.Cells(1, 3).Value
                fondoActual = filaIngrediente.Range.Cells(1, 6).Value
                filaIngrediente.Range.Cells(1, 3).Value = stockActual - CantidadEnBase
                filaIngrediente.Range.Cells(1, 6).Value = fondoActual - (CantidadEnBase * filaIngrediente.Range.Cells(1, 5).Value)
                
                Exit For
            End If
        Next filaIngrediente
    Next i
    
    ' PASO 3: REGISTRAR EN INVENTARIO
    ' Generar código del producto
    Genera_Codigo
    codigoProducto = Codigo_Producto
  
    ' Obtener datos para el inventario
    cantidadElaborada = CDbl(Me.Cantidad.Value)
    inversionTotal = CDbl(Me.Inversion.Value)
    costoPorUnidad = inversionTotal / cantidadElaborada
    ingresoEsperado = CDbl(Me.Ingreso.Value)
    gananciaEsperada = CDbl(Me.Beneficio.Value)
    
    ThisWorkbook.Sheets("Gestión de Almacén").Unprotect password:="theSheetProtectpassisthe011rd"
    
    ' Agregar nueva fila a la tabla INVENTARIO
    Set nuevaFilaInventario = tblInventario.ListRows.Add
    With nuevaFilaInventario
        .Range.Cells(1, 1).Value = codigoProducto  ' Código
        .Range.Cells(1, 2).Value = Me.Recetas.Value  ' Nombre del producto
        .Range.Cells(1, 3).Value = Date  ' Fecha de entrada (hoy)
        .Range.Cells(1, 4).Value = cantidadElaborada  ' Cantidad Inicial
        .Range.Cells(1, 5).Value = cantidadElaborada  ' Cantidad Actual
        .Range.Cells(1, 6).Value = CDbl(Me.Precio_venta.Value)  ' Precio Venta
        .Range.Cells(1, 7).Value = costoPorUnidad  ' Precio Costo
        .Range.Cells(1, 8).Value = costoPorUnidad * cantidadElaborada  ' Fondo
        .Range.Cells(1, 9).Value = ingresoEsperado 'ingreso
        .Range.Cells(1, 10).Value = gananciaEsperada 'ganancia
        
    End With
    
    Confirmar_Codigo_Producto codigoProducto
    
    producto = Me.Recetas.Value
    Agregar_Producto
    Agregar_Fecha
    
   ThisWorkbook.Sheets("Gestión de Almacén").Protect password:="theSheetProtectpassisthe011rd", userinterfaceonly:=True
    
    Application.ScreenUpdating = True
    
    ' PASO 4: MOSTRAR MENSAJE DE ÉXITO
    Dim resumen As String
    Dim ingresoEsperado As Double
    Dim beneficioEsperado As Double
    
    ingresoEsperado = CDbl(Me.Ingreso.Value)
    beneficioEsperado = CDbl(Me.Beneficio.Value)
    
    resumen = "ELABORACIÓN COMPLETADA EXITOSAMENTE" & vbCrLf & vbCrLf & _
              "Receta: " & Me.Recetas.Value & vbCrLf & _
              "Lotes producidos: " & numLotes & vbCrLf & _
              "Unidades totales: " & cantidadElaborada & vbCrLf & _
              "Código del producto: " & codigoProducto & vbCrLf & _
              "Fecha de producción: " & Format(Date, "dd/mm/yyyy") & vbCrLf & _
              "Inversión total: $" & Format(inversionTotal, "0.00") & vbCrLf & _
              "Ingreso esperado: $" & Format(ingresoEsperado, "0.00") & vbCrLf & _
              "Beneficio esperado: $" & Format(beneficioEsperado, "0.00") & vbCrLf & _
              "Costo por unidad: $" & Format(costoPorUnidad, "0.00") & vbCrLf & _
              "Precio de venta: $" & Format(CDbl(Me.Precio_venta.Value), "0.00") & vbCrLf & _
              "Ganancia por unidad: $" & Format(CDbl(Me.Precio_venta.Value) - costoPorUnidad, "0.00")
    
    MsgBox resumen, vbInformation, "Elaboración Exitosa"
    
    ' Limpiar formulario para una nueva elaboración
    Unload Me
End Sub

Private Sub Btn_Reposo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.Btn_Active.ZOrder msoBringToFront
    Me.Txt_Reposo.ZOrder msoBringToFront
End Sub

Private Sub Lotes_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If (KeyAscii >= 46 And KeyAscii <= 57) Then
    KeyAscii = KeyAscii
  Else
    KeyAscii = 0
  End If
End Sub

' --- EVENTO PARA ACTUALIZAR INGRESO Y BENEFICIO CUANDO CAMBIA EL PRECIO ---
Private Sub Precio_venta_Change()
    ' Si hay una receta seleccionada y hay cantidad de lotes, actualizar cálculos
    If Me.Recetas.Value <> "" And Me.Lotes.Value <> "" And IsNumeric(Me.Lotes.Value) Then
        Dim ingresoEsperado As Double
        Dim beneficioEsperado As Double
        
        If Me.Precio_venta.Value <> "" And IsNumeric(Me.Precio_venta.Value) And _
           Me.Cantidad.Value <> "" And IsNumeric(Me.Cantidad.Value) Then
            
            ' Calcular ingreso esperado
            ingresoEsperado = CDbl(Me.Precio_venta.Value) * CDbl(Me.Cantidad.Value)
            Me.Ingreso.Value = Format(ingresoEsperado, "0.00")
            
            ' Calcular beneficio esperado
            If Me.Inversion.Value <> "" And IsNumeric(Me.Inversion.Value) Then
                beneficioEsperado = ingresoEsperado - CDbl(Me.Inversion.Value)
                Me.Beneficio.Value = Format(beneficioEsperado, "0.00")
            End If
        Else
            Me.Ingreso.Value = ""
            Me.Beneficio.Value = ""
        End If
    End If
End Sub

Private Sub Recetas_Change()
    Dim wsRecetas As Worksheet
    Dim wsIngredientes As Worksheet
    Dim tblRecetas As ListObject
    Dim tblIngredientesReceta As ListObject
    Dim filaReceta As ListRow
    Dim filaIngrediente As ListRow
    
    ' Limpiar datos anteriores
    recetaSeleccionada = ""
    cantidadPorLote = 0
    inversionPorLote = 0
    precioPorUnidad = 0
    
    Me.Ingredientes.Clear
    Me.Lotes.Value = ""
    Me.Cantidad.Value = ""
    Me.Inversion.Value = ""
    Me.Precio_venta.Value = ""
    Me.Ingreso.Value = ""
    Me.Beneficio.Value = ""
    
    If Me.Recetas.Value = "" Then
        Exit Sub
    End If
    
    recetaSeleccionada = Me.Recetas.Value
    
    ' Obtener hojas de trabajo
    Set wsRecetas = ThisWorkbook.Sheets("Elaboración de Productos")
    Set wsIngredientes = ThisWorkbook.Sheets("Ingredientes")
    
    ' Obtener tablas
    Set tblRecetas = wsRecetas.ListObjects("Recetas")
    Set tblIngredientesReceta = wsIngredientes.ListObjects("Ingredientes_Receta")
    
    ' Buscar la receta seleccionada en la tabla Recetas
    For Each filaReceta In tblRecetas.ListRows
        If filaReceta.Range.Cells(1, 1).Value = recetaSeleccionada Then
            ' Obtener los datos de la receta
            cantidadPorLote = filaReceta.Range.Cells(1, 2).Value ' Cantidad/Lote
            inversionPorLote = filaReceta.Range.Cells(1, 3).Value ' Inversión/Lote
            precioPorUnidad = filaReceta.Range.Cells(1, 7).Value ' Precio/Unidad
            
            ' Mostrar el precio por unidad en el textbox correspondiente
            Me.Precio_venta.Value = precioPorUnidad
            
            Exit For
        End If
    Next filaReceta
    
    ' Cargar los ingredientes de la receta DESDE LA TABLA
    For Each filaIngrediente In tblIngredientesReceta.ListRows
        If filaIngrediente.Range.Cells(1, 1).Value = recetaSeleccionada Then
            ' Agregar el ingrediente al ListBox
            Me.Ingredientes.AddItem
            
            With Me.Ingredientes
                ' Obtener los datos directamente de la tabla
                .List(.ListCount - 1, 0) = filaIngrediente.Range.Cells(1, 2).Value ' Nombre del ingrediente
                .List(.ListCount - 1, 1) = filaIngrediente.Range.Cells(1, 3).Value ' Cantidad por lote
                .List(.ListCount - 1, 2) = filaIngrediente.Range.Cells(1, 4).Value ' Unidad de medida
            End With
        End If
    Next filaIngrediente
    
    ' Si hay ingredientes, actualizar la lista con 1 lote por defecto
    If Me.Ingredientes.ListCount > 0 And Me.Lotes.Value = "" Then
        Me.Lotes.Value = 1
        ActualizarCalculos
    End If
End Sub

Private Sub Lotes_Change()
    ' Validar que solo se ingresen números
    Dim texto As String
    Dim i As Integer
    
    texto = Me.Lotes.Value
    If texto = "" Then
        ' Limpiar los cálculos si no hay lotes
        Me.Cantidad.Value = ""
        Me.Inversion.Value = ""
        ActualizarListaIngredientes 0
        Exit Sub
    End If
    
    ' Actualizar cálculos
    If IsNumeric(Me.Lotes.Value) Then
        ActualizarCalculos
    End If
End Sub

Private Sub ActualizarCalculos()
    Dim numLotes As Double
    Dim ingresoEsperado As Double
    Dim beneficioEsperado As Double
    
    If Me.Lotes.Value = "" Then
        numLotes = 0
    Else
        numLotes = CDbl(Me.Lotes.Value)
    End If
    
    ' Calcular y mostrar la cantidad total de unidades
    Me.Cantidad.Value = cantidadPorLote * numLotes
    
    ' Calcular y mostrar la inversión total
    Me.Inversion.Value = Format(Format(inversionPorLote, "0.00") * numLotes, "0.00")
    
    ' Calcular y mostrar el ingreso esperado (precio de venta * cantidad)
    If Me.Precio_venta.Value <> "" And IsNumeric(Me.Precio_venta.Value) Then
        ingresoEsperado = CDbl(Me.Precio_venta.Value) * CDbl(Me.Cantidad.Value)
        Me.Ingreso.Value = Format(ingresoEsperado, "0.00")
    Else
        Me.Ingreso.Value = ""
    End If
    
    ' Calcular y mostrar el beneficio esperado (ingreso - inversión)
    If Me.Inversion.Value <> "" And IsNumeric(Me.Inversion.Value) And _
       Me.Ingreso.Value <> "" And IsNumeric(Me.Ingreso.Value) Then
        beneficioEsperado = CDbl(Me.Ingreso.Value) - CDbl(Me.Inversion.Value)
        Me.Beneficio.Value = Format(beneficioEsperado, "0.00")
    Else
        Me.Beneficio.Value = ""
    End If
    
    ' Actualizar la lista de ingredientes con las cantidades multiplicadas
    ActualizarListaIngredientes numLotes
End Sub

Private Sub ActualizarListaIngredientes(numLotes As Double)
    Dim wsIngredientes As Worksheet
    Dim tblIngredientesReceta As ListObject
    Dim filaIngrediente As ListRow
    Dim i As Long
    Dim ingredienteNombre As String
    Dim cantidadOriginal As Double
    Dim encontrado As Boolean
    
    ' Obtener referencia a la hoja y tabla
    Set wsIngredientes = ThisWorkbook.Sheets("Ingredientes")
    Set tblIngredientesReceta = wsIngredientes.ListObjects("Ingredientes_Receta")
    
    ' Recorrer el ListBox y actualizar las cantidades
    For i = 0 To Me.Ingredientes.ListCount - 1
        ingredienteNombre = Me.Ingredientes.List(i, 0)
        encontrado = False
        
        ' Buscar el ingrediente en la tabla Ingredientes_Receta
        For Each filaIngrediente In tblIngredientesReceta.ListRows
            ' Verificar si es el mismo ingrediente y la misma receta
            If filaIngrediente.Range.Cells(1, 1).Value = recetaSeleccionada And _
               filaIngrediente.Range.Cells(1, 2).Value = ingredienteNombre Then
                
                ' Obtener la cantidad original por lote desde la tabla
                cantidadOriginal = CDbl(filaIngrediente.Range.Cells(1, 3).Value)
                
                ' Actualizar la cantidad en el ListBox multiplicando por el número de lotes
                Me.Ingredientes.List(i, 1) = cantidadOriginal * numLotes
                
                encontrado = True
                Exit For
            End If
        Next filaIngrediente
        
        ' Si no se encontró el ingrediente, mantener la cantidad actual
        If Not encontrado Then
            ' Mantener la cantidad que ya está en el ListBox
            ' (esto no debería pasar si la receta está bien configurada)
        End If
    Next i
End Sub

' --- EVENTO PARA VALIDAR EL CAMBIO EN EL TEXTBOX LOTES ---
Private Sub Lotes_AfterUpdate()
    ' Forzar la actualización si el usuario presiona Enter
    ActualizarCalculos
End Sub

' --- FUNCIÓN PARA FORMATO DE NÚMEROS ---
Private Function FormatearNumero(valor As Variant) As String
    If IsNumeric(valor) Then
        FormatearNumero = Format(valor, "0.00")
    Else
        FormatearNumero = ""
    End If
End Function

Private Sub Txt_Reposo_Click()
    Btn_Active_Click
End Sub

'inicializa el listbox con sus columnas
Private Sub UserForm_Initialize()

' Cargar recetas en el combobox
    CargarRecetas
    
    ' Inicializar otros controles
    Me.Lotes.Value = ""
    Me.Cantidad.Value = ""
    Me.Inversion.Value = ""
    Me.Precio_venta.Value = ""
    Me.Ingreso.Value = ""
    Me.Beneficio.Value = ""

    With Me.Ingredientes
        .ColumnCount = 3
        .ColumnWidths = "120;80;80"
    End With
    
End Sub

' --- BOTÓN PARA CONFIRMAR ELABORACIÓN ---
Private Sub Elaborar_Click()
    Dim wsIngredientes As Worksheet
    Dim wsInventario As Worksheet
    Dim tblIngredientes As ListObject
    Dim tblInventario As ListObject
    Dim filaIngrediente As ListRow
    Dim nuevaFilaInventario As ListRow
    Dim i As Long
    Dim ingredienteNombre As String
    Dim cantidadNecesaria As Double
    Dim unidadMedida As String
    Dim CantidadEnBase As Double
    Dim stockActual As Double
    Dim suficienteStock As Boolean
    Dim numLotes As Double
    Dim codigoProducto As String
    Dim cantidadElaborada As Double
    Dim inversionTotal As Double
    Dim costoPorUnidad As Double
    
    ' Validar que haya una receta seleccionada
    If Me.Recetas.Value = "" Then
        MsgBox "Por favor, seleccione una receta.", vbExclamation
        Me.Recetas.SetFocus
        Exit Sub
    End If
    
    ' Validar que haya un número de lotes válido
    If Me.Lotes.Value = "" Or Not IsNumeric(Me.Lotes.Value) Then
        MsgBox "Por favor, ingrese una cantidad válida de lotes.", vbExclamation
        Me.Lotes.SetFocus
        Exit Sub
    End If
    
    numLotes = CDbl(Me.Lotes.Value)
    
    ' Validar que haya al menos un ingrediente en la lista
    If Me.Ingredientes.ListCount = 0 Then
        MsgBox "La receta seleccionada no tiene ingredientes.", vbExclamation
        Exit Sub
    End If
    
    ' Obtener hojas de trabajo
    Set wsIngredientes = ThisWorkbook.Sheets("Gestión de Ingredientes")
    Set wsInventario = ThisWorkbook.Sheets("Gestión de Almacén")
    
    ' Obtener tablas
    Set tblIngredientes = wsIngredientes.ListObjects("INGREDIENTES")
    Set tblInventario = wsInventario.ListObjects("INVENTARIO")
    
    ' Variable para controlar si hay suficiente stock
    suficienteStock = True
    Dim ingredientesFaltantes As String
    ingredientesFaltantes = ""
    
    ' PASO 1: VALIDAR STOCK DE TODOS LOS INGREDIENTES
    Application.ScreenUpdating = False
    
    For i = 0 To Me.Ingredientes.ListCount - 1
        ingredienteNombre = Me.Ingredientes.List(i, 0)
        cantidadNecesaria = CDbl(Me.Ingredientes.List(i, 1))
        unidadMedida = Me.Ingredientes.List(i, 2)
        
        ' Buscar el ingrediente en la tabla INGREDIENTES
        Dim encontrado As Boolean
        encontrado = False
        
        For Each filaIngrediente In tblIngredientes.ListRows
            If filaIngrediente.Range.Cells(1, 1).Value = ingredienteNombre Then
                encontrado = True
                
                ' Convertir la cantidad necesaria a unidad base
                CantidadEnBase = ConvertirUnidad(ingredienteNombre, cantidadNecesaria, unidadMedida)
                
                ' Obtener el stock actual (columna 3)
                stockActual = filaIngrediente.Range.Cells(1, 3).Value
                
                ' Verificar si hay suficiente stock
                If stockActual < CantidadEnBase Then
                    suficienteStock = False
                    ingredientesFaltantes = ingredientesFaltantes & "- " & ingredienteNombre & ": " & _
                                            "Necesario: " & Format(CantidadEnBase, "0.00") & " " & _
                                            filaIngrediente.Range.Cells(1, 4).Value & _
                                            " | Disponible: " & Format(stockActual, "0.00") & _
                                            filaIngrediente.Range.Cells(1, 4).Value & vbCrLf
                End If
                
                Exit For
            End If
        Next filaIngrediente
        
        If Not encontrado Then
            MsgBox "El ingrediente '" & ingredienteNombre & "' no existe en la tabla de ingredientes.", vbExclamation
            Application.ScreenUpdating = True
            Exit Sub
        End If
    Next i
    
    ' Si no hay suficiente stock, mostrar mensaje y salir
    If Not suficienteStock Then
        Application.ScreenUpdating = True
        MsgBox "Stock insuficiente para los siguientes ingredientes:" & vbCrLf & vbCrLf & _
               ingredientesFaltantes & vbCrLf & _
               "Por favor, reponga el stock y vuelva a intentar.", vbExclamation
        Exit Sub
    End If
    
    ' PASO 2: ACTUALIZAR STOCK DE INGREDIENTES
    For i = 0 To Me.Ingredientes.ListCount - 1
        ingredienteNombre = Me.Ingredientes.List(i, 0)
        cantidadNecesaria = CDbl(Me.Ingredientes.List(i, 1))
        unidadMedida = Me.Ingredientes.List(i, 2)
        
        ' Buscar el ingrediente en la tabla INGREDIENTES
        For Each filaIngrediente In tblIngredientes.ListRows
            If filaIngrediente.Range.Cells(1, 1).Value = ingredienteNombre Then
                ' Convertir la cantidad necesaria a unidad base
                CantidadEnBase = ConvertirUnidad(ingredienteNombre, cantidadNecesaria, unidadMedida)
                
                ' Restar del stock actual
                stockActual = filaIngrediente.Range.Cells(1, 3).Value
                filaIngrediente.Range.Cells(1, 3).Value = stockActual - CantidadEnBase
                
                Exit For
            End If
        Next filaIngrediente
    Next i
    
    ' PASO 3: REGISTRAR EN INVENTARIO
    ' Generar código del producto
    Genera_Codigo
    codigoProducto = Codigo_Producto
  
    ' Obtener datos para el inventario
    cantidadElaborada = CDbl(Me.Cantidad.Value)
    inversionTotal = CDbl(Me.Inversion.Value)
    costoPorUnidad = inversionTotal / cantidadElaborada
    
    ThisWorkbook.Sheets("Gestión de Almacén").Unprotect password:="theSheetProtectpassisthe011rd"
    
    ' Agregar nueva fila a la tabla INVENTARIO
    Set nuevaFilaInventario = tblInventario.ListRows.Add
    With nuevaFilaInventario
        .Range.Cells(1, 1).Value = codigoProducto  ' Código
        .Range.Cells(1, 2).Value = Me.Recetas.Value  ' Nombre del producto
        .Range.Cells(1, 3).Value = Date  ' Fecha de entrada (hoy)
        .Range.Cells(1, 4).Value = cantidadElaborada  ' Cantidad Inicial
        .Range.Cells(1, 5).Value = cantidadElaborada  ' Cantidad Actual
        .Range.Cells(1, 6).Value = CDbl(Me.Precio_venta.Value)  ' Precio Venta
        .Range.Cells(1, 7).Value = costoPorUnidad  ' Precio Costo
        .Range.Cells(1, 8).Value = costoPorUnidad * cantidadElaborada  ' Fondo
    End With
    
    Confirmar_Codigo_Producto codigoProducto
    
    producto = Me.Recetas.Value
    Agregar_Producto
    Agregar_Fecha
    
   ThisWorkbook.Sheets("Gestión de Almacén").Protect password:="theSheetProtectpassisthe011rd", userinterfaceonly:=True
    
    Application.ScreenUpdating = True
    
    ' PASO 4: MOSTRAR MENSAJE DE ÉXITO
    Dim resumen As String
    Dim ingresoEsperado As Double
    Dim beneficioEsperado As Double
    
    ingresoEsperado = CDbl(Me.Ingreso.Value)
    beneficioEsperado = CDbl(Me.Beneficio.Value)
    
    resumen = "ELABORACIÓN COMPLETADA EXITOSAMENTE" & vbCrLf & vbCrLf & _
              "Receta: " & Me.Recetas.Value & vbCrLf & _
              "Lotes producidos: " & numLotes & vbCrLf & _
              "Unidades totales: " & cantidadElaborada & vbCrLf & _
              "Código del producto: " & codigoProducto & vbCrLf & _
              "Fecha de producción: " & Format(Date, "dd/mm/yyyy") & vbCrLf & _
              "Inversión total: $" & Format(inversionTotal, "0.00") & vbCrLf & _
              "Ingreso esperado: $" & Format(ingresoEsperado, "0.00") & vbCrLf & _
              "Beneficio esperado: $" & Format(beneficioEsperado, "0.00") & vbCrLf & _
              "Costo por unidad: $" & Format(costoPorUnidad, "0.00") & vbCrLf & _
              "Precio de venta: $" & Format(CDbl(Me.Precio_venta.Value), "0.00") & vbCrLf & _
              "Ganancia por unidad: $" & Format(CDbl(Me.Precio_venta.Value) - costoPorUnidad, "0.00")
    
    MsgBox resumen, vbInformation, "Elaboración Exitosa"
    
    ' Limpiar formulario para una nueva elaboración
    Unload Me
End Sub

' --- FUNCIÓN PARA LIMPIAR EL FORMULARIO DESPUÉS DE ELABORAR ---
Private Sub LimpiarFormularioElaborar()
    Me.Recetas.Value = ""
    Me.Lotes.Value = ""
    Me.Cantidad.Value = ""
    Me.Inversion.Value = ""
    Me.Precio_venta.Value = ""
    Me.Ingreso.Value = ""
    Me.Beneficio.Value = ""
    Me.Ingredientes.Clear
    Me.Recetas.SetFocus
End Sub

'Macro para agregar un producto a la tabla de productos unicos
Private Sub Agregar_Producto()

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim newrow As ListRow
    
    Set ws = ThisWorkbook.Sheets("Fechas")
    Set tbl = ws.ListObjects("Productos_Unicos")
    
    For Each fila In tbl.ListRows
        If UCase(Trim(fila.Range(1))) = UCase(Trim(producto)) Then
            Exit Sub
        End If
    Next fila
    
    ws.Unprotect password:="theSheetProtectpassisthe011rd"
    
    Set newrow = tbl.ListRows.Add
    
    newrow.Range(1) = producto
    
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

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.Btn_Reposo.ZOrder msoBringToFront
    Me.Txt_Reposo.ZOrder msoBringToFront
End Sub
