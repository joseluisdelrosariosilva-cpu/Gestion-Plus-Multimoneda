VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Ingr_Receta 
   Caption         =   "Ingresar Receta"
   ClientHeight    =   6620
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   14820
   OleObjectBlob   =   "Ingr_Receta.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Ingr_Receta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Btn_Active_Click()
    If modoModificacion Then
        ModificarRecetaExistente
    Else
        AgregarNuevaReceta
    End If
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.Btn_Reposo.ZOrder msoBringToFront
    Me.Agregar.ZOrder msoBringToFront
End Sub

Private Sub Btn_Reposo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.Btn_Active.ZOrder msoBringToFront
    Me.Agregar.ZOrder msoBringToFront
End Sub

Private Sub AgregarNuevaReceta()
    Dim wsRecetas As Worksheet
    Dim wsIngredientes As Worksheet
    Dim tblRecetas As ListObject
    Dim tblIngredientesReceta As ListObject
    Dim nuevaFilaReceta As ListRow
    Dim nuevaFilaIngrediente As ListRow
    Dim i As Integer
    Dim recetaNombre As String
    Dim cantidadLote As Double
    Dim inversionLote As Double
    Dim precioLote As Double
    Dim gananciaLote As Double
    Dim inversionUnidad As Double
    Dim precioUnidad As Double
    Dim gananciaUnidad As Double
    
    ' Obtener hojas de trabajo
    Set wsRecetas = ThisWorkbook.Sheets("Elaboración de Productos")
    Set wsIngredientes = ThisWorkbook.Sheets("Ingredientes")
    
    ' Obtener tablas
    Set tblRecetas = wsRecetas.ListObjects("Recetas")
    Set tblIngredientesReceta = wsIngredientes.ListObjects("Ingredientes_Receta")
    
    ' Validar datos básicos
    recetaNombre = Me.Txt_Receta.Value
    If recetaNombre = "" Then
        MsgBox "Por favor, ingrese un nombre para la receta.", vbExclamation
        Me.Txt_Receta.SetFocus
        Exit Sub
    End If
    
    ' Verificar si la receta ya existe
    If RecetaExiste(recetaNombre) Then
        MsgBox "Ya existe una receta con este nombre.", vbExclamation
        Me.Txt_Receta.SetFocus
        Exit Sub
    End If
    
    ' Validar cantidad del lote
    If Me.Txt_Cantidad.Value = "" Then
        MsgBox "Por favor, ingrese la cantidad de unidades que produce el lote.", vbExclamation
        Me.Txt_Cantidad.SetFocus
        Exit Sub
    End If
    
    If Not IsNumeric(Me.Txt_Cantidad.Value) Then
        MsgBox "Por favor, ingrese una cantidad válida para el lote.", vbExclamation
        Me.Txt_Cantidad.SetFocus
        Exit Sub
    End If
    
    cantidadLote = CDbl(Me.Txt_Cantidad.Value)
    
    ' Validar precio por unidad
    If Me.Txt_Precio_u.Value = "" Then
        MsgBox "Por favor, ingrese el precio de venta por unidad.", vbExclamation
        Me.Txt_Precio_u.SetFocus
        Exit Sub
    End If
    
    If Not IsNumeric(Me.Txt_Precio_u.Value) Then
        MsgBox "Por favor, ingrese un precio de venta por unidad válido.", vbExclamation
        Me.Txt_Precio_u.SetFocus
        Exit Sub
    End If
    
    precioUnidad = CDbl(Me.Txt_Precio_u.Value)
    
    ' Obtener los valores calculados
    If Me.Txt_Costo.Value = "" Or Not IsNumeric(Me.Txt_Costo.Value) Then
        inversionLote = 0
    Else
        inversionLote = CDbl(Me.Txt_Costo.Value)
    End If
    
    If Me.Txt_Precio.Value = "" Or Not IsNumeric(Me.Txt_Precio.Value) Then
        precioLote = 0
    Else
        precioLote = CDbl(Me.Txt_Precio.Value)
    End If
    
    If Me.Txt_Costo_u.Value = "" Or Not IsNumeric(Me.Txt_Costo_u.Value) Then
        inversionUnidad = 0
    Else
        inversionUnidad = CDbl(Me.Txt_Costo_u.Value)
    End If
    
    ' Calcular ganancias
    gananciaLote = precioLote - inversionLote
    gananciaUnidad = precioUnidad - inversionUnidad
    
    ' Verificar que haya al menos un ingrediente
    If Not HayIngredientesValidos() Then
        MsgBox "Por favor, agregue al menos un ingrediente a la receta.", vbExclamation
        Exit Sub
    End If
    
    wsRecetas.Unprotect password:="theSheetProtectpassisthe011rd"
    wsIngredientes.Unprotect password:="theSheetProtectpassisthe011rd"
    
    ' Agregar nueva receta a la tabla Recetas
    Set nuevaFilaReceta = tblRecetas.ListRows.Add
    With nuevaFilaReceta
        .Range.Cells(1, 1).Value = recetaNombre  ' Receta
        .Range.Cells(1, 2).Value = cantidadLote  ' Lote
        .Range.Cells(1, 3).Value = inversionLote  ' Inversion/Lote
        .Range.Cells(1, 4).Value = precioLote  ' Precio/Lote
        .Range.Cells(1, 5).Value = gananciaLote  ' Ganancia/Lote
        .Range.Cells(1, 6).Value = inversionUnidad  ' Inversion/Unidad
        .Range.Cells(1, 7).Value = precioUnidad  ' Precio/Unidad
        .Range.Cells(1, 8).Value = gananciaUnidad  ' Ganancia/Unidad
    End With
    
    ' Agregar ingredientes a la tabla Ingredientes_Receta
    For i = 1 To 10
        Dim cbIngrediente As MSForms.ComboBox
        Dim txtCantidad As MSForms.TextBox
        Dim cbUnidad As MSForms.ComboBox
        
        Set cbIngrediente = Me.Controls("Cmb" & i)
        Set txtCantidad = Me.Controls("TextBox" & i)
        Set cbUnidad = Me.Controls("Cmb_" & i)
        
        If cbIngrediente.Value <> "" And cbUnidad.Value <> "" And txtCantidad.Value <> "" Then
            If IsNumeric(txtCantidad.Value) Then
                Set nuevaFilaIngrediente = tblIngredientesReceta.ListRows.Add
                With nuevaFilaIngrediente
                    .Range.Cells(1, 1).Value = recetaNombre
                    .Range.Cells(1, 2).Value = cbIngrediente.Value
                    .Range.Cells(1, 3).Value = CDbl(txtCantidad.Value)
                    .Range.Cells(1, 4).Value = cbUnidad.Value
                End With
            End If
        End If
    Next i
    
    tblRecetas.Sort.Apply
    wsRecetas.Protect password:="theSheetProtectpassisthe011rd", userinterfaceonly:=True
    
    tblIngredientesReceta.Sort.Apply
    wsIngredientes.Protect password:="theSheetProtectpassisthe011rd", userinterfaceonly:=True
    
    MsgBox "Receta '" & recetaNombre & "' guardada exitosamente.", vbInformation
    Unload Me
End Sub

Private Sub Agregar_Click()
    Btn_Active_Click
End Sub

' --- FUNCIÓN PARA MODIFICAR RECETA EXISTENTE ---
Private Sub ModificarRecetaExistente()
    Dim wsRecetas As Worksheet
    Dim wsIngredientes As Worksheet
    Dim tblRecetas As ListObject
    Dim tblIngredientesReceta As ListObject
    Dim rngBusqueda As Range
    Dim celdaEncontrada As Range
    Dim filaReceta As ListRow
    Dim filaIngrediente As ListRow
    Dim i As Integer
    Dim j As Long
    Dim recetaNombre As String
    Dim cantidadLote As Double
    Dim inversionLote As Double
    Dim precioLote As Double
    Dim gananciaLote As Double
    Dim inversionUnidad As Double
    Dim precioUnidad As Double
    Dim gananciaUnidad As Double
    
    ' Validar datos básicos
    recetaNombre = Me.Txt_Receta.Value
    If recetaNombre = "" Then
        MsgBox "Por favor, ingrese un nombre para la receta.", vbExclamation
        Me.Txt_Receta.SetFocus
        Exit Sub
    End If
    
    ' Validar cantidad del lote
    If Me.Txt_Cantidad.Value = "" Then
        MsgBox "Por favor, ingrese la cantidad de unidades que produce el lote.", vbExclamation
        Me.Txt_Cantidad.SetFocus
        Exit Sub
    End If
    
    If Not IsNumeric(Me.Txt_Cantidad.Value) Then
        MsgBox "Por favor, ingrese una cantidad válida para el lote.", vbExclamation
        Me.Txt_Cantidad.SetFocus
        Exit Sub
    End If
    
    cantidadLote = CDbl(Me.Txt_Cantidad.Value)
    
    ' Validar precio por unidad
    If Me.Txt_Precio_u.Value = "" Then
        MsgBox "Por favor, ingrese el precio de venta por unidad.", vbExclamation
        Me.Txt_Precio_u.SetFocus
        Exit Sub
    End If
    
    If Not IsNumeric(Me.Txt_Precio_u.Value) Then
        MsgBox "Por favor, ingrese un precio de venta por unidad válido.", vbExclamation
        Me.Txt_Precio_u.SetFocus
        Exit Sub
    End If
    
    precioUnidad = CDbl(Me.Txt_Precio_u.Value)
    
    ' Obtener los valores calculados
    If Me.Txt_Costo.Value = "" Or Not IsNumeric(Me.Txt_Costo.Value) Then
        inversionLote = 0
    Else
        inversionLote = CDbl(Me.Txt_Costo.Value)
    End If
    
    If Me.Txt_Precio.Value = "" Or Not IsNumeric(Me.Txt_Precio.Value) Then
        precioLote = 0
    Else
        precioLote = CDbl(Me.Txt_Precio.Value)
    End If
    
    If Me.Txt_Costo_u.Value = "" Or Not IsNumeric(Me.Txt_Costo_u.Value) Then
        inversionUnidad = 0
    Else
        inversionUnidad = CDbl(Me.Txt_Costo_u.Value)
    End If
    
    ' Calcular ganancias
    gananciaLote = precioLote - inversionLote
    gananciaUnidad = precioUnidad - inversionUnidad
    
    ' Verificar que haya al menos un ingrediente
    If Not HayIngredientesValidos() Then
        MsgBox "Por favor, agregue al menos un ingrediente a la receta.", vbExclamation
        Exit Sub
    End If
    
    ' Obtener hojas de trabajo
    Set wsRecetas = ThisWorkbook.Sheets("Elaboración de Productos")
    Set wsIngredientes = ThisWorkbook.Sheets("Ingredientes")
    
    ' Obtener tablas
    Set tblRecetas = wsRecetas.ListObjects("Recetas")
    Set tblIngredientesReceta = wsIngredientes.ListObjects("Ingredientes_Receta")
    
    wsRecetas.Unprotect password:="theSheetProtectpassisthe011rd"
    wsIngredientes.Unprotect password:="theSheetProtectpassisthe011rd"
    
   ' ELIMINAR REGISTROS ANTIGUOS - SOLO DE LAS TABLAS
    j = 1
    Dim filaTabla As ListRow
    Do While j <= tblRecetas.ListRows.Count
        Set filaTabla = tblRecetas.ListRows(j)
        If filaTabla.Range(1, 1).Value = recetaAModificar Then
            filaTabla.Delete
            Exit Do
        Else
            j = j + 1
        End If
    Loop
    
    ' Eliminar ingredientes de la tabla Ingredientes_Receta
    j = 1
    Do While j <= tblIngredientesReceta.ListRows.Count
        Set filaIngrediente = tblIngredientesReceta.ListRows(j)
        ' Acceder al valor de la celda de forma más directa
        Dim valorCelda As String
        valorCelda = filaIngrediente.Range(1, 1).Value
        
        If valorCelda = recetaAModificar Then
            filaIngrediente.Delete
        Else
            j = j + 1
        End If
    Loop
    
    ' AGREGAR NUEVOS REGISTROS
    ' Agregar nueva receta a la tabla Recetas
    Dim nuevaFilaReceta As ListRow
    Set nuevaFilaReceta = tblRecetas.ListRows.Add
    With nuevaFilaReceta
        .Range.Cells(1, 1).Value = recetaNombre  ' Receta
        .Range.Cells(1, 2).Value = cantidadLote  ' Lote
        .Range.Cells(1, 3).Value = inversionLote  ' Inversion/Lote
        .Range.Cells(1, 4).Value = precioLote  ' Precio/Lote
        .Range.Cells(1, 5).Value = gananciaLote  ' Ganancia/Lote
        .Range.Cells(1, 6).Value = inversionUnidad  ' Inversion/Unidad
        .Range.Cells(1, 7).Value = precioUnidad  ' Precio/Unidad
        .Range.Cells(1, 8).Value = gananciaUnidad  ' Ganancia/Unidad
    End With
    
    ' Agregar ingredientes a la tabla Ingredientes_Receta
    For i = 1 To 10
        Dim cbIngrediente As MSForms.ComboBox
        Dim txtCantidad As MSForms.TextBox
        Dim cbUnidad As MSForms.ComboBox
        
        Set cbIngrediente = Me.Controls("Cmb" & i)
        Set txtCantidad = Me.Controls("TextBox" & i)
        Set cbUnidad = Me.Controls("Cmb_" & i)
        
        If cbIngrediente.Value <> "" And cbUnidad.Value <> "" And txtCantidad.Value <> "" Then
            If IsNumeric(txtCantidad.Value) Then
                Dim nuevaFilaIngrediente As ListRow
                Set nuevaFilaIngrediente = tblIngredientesReceta.ListRows.Add
                With nuevaFilaIngrediente
                    .Range.Cells(1, 1).Value = recetaNombre
                    .Range.Cells(1, 2).Value = cbIngrediente.Value
                    .Range.Cells(1, 3).Value = CDbl(txtCantidad.Value)
                    .Range.Cells(1, 4).Value = cbUnidad.Value
                End With
            End If
        End If
    Next i
    
    tblRecetas.Sort.Apply
    wsRecetas.Protect password:="theSheetProtectpassisthe011rd", userinterfaceonly:=True
    
    tblIngredientesReceta.Sort.Apply
    wsIngredientes.Protect password:="theSheetProtectpassisthe011rd", userinterfaceonly:=True
    
    MsgBox "Receta '" & recetaNombre & "' modificada exitosamente.", vbInformation
    
    ' Cerrar formulario
    Unload Me
End Sub

' --- MODIFICAR LA FUNCIÓN RecetaExiste PARA QUE NO VALIDE EN MODO MODIFICACIÓN ---
Private Function RecetaExiste(nombreReceta As String) As Boolean
    ' Si estamos en modo modificación y el nombre no cambió, no validar
    If modoModificacion And nombreReceta = recetaAModificar Then
        RecetaExiste = False
        Exit Function
    End If
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim fila As ListRow
    
    Set ws = ThisWorkbook.Sheets("Elaboración de Productos")
    Set tbl = ws.ListObjects("Recetas")
    
    RecetaExiste = False
    
    For Each fila In tbl.ListRows
        If LCase(fila.Range.Cells(1, 1).Value) = LCase(nombreReceta) Then
            RecetaExiste = True
            Exit Function
        End If
    Next fila
End Function

' --- AGREGAR EVENTO PARA EL BOTÓN CERRAR ---
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' Limpiar variables cuando se cierra el formulario
    modoModificacion = False
    recetaAModificar = ""
End Sub

' Función para verificar si hay al menos un ingrediente válido
Private Function HayIngredientesValidos() As Boolean
    Dim i As Integer
    
    HayIngredientesValidos = False
    
    For i = 1 To 10
        Dim cbIngrediente As MSForms.ComboBox
        Dim txtCantidad As MSForms.TextBox
        Dim cbUnidad As MSForms.ComboBox
        
        Set cbIngrediente = Me.Controls("Cmb" & i)
        Set txtCantidad = Me.Controls("TextBox" & i)
        Set cbUnidad = Me.Controls("Cmb_" & i)
        
        ' Verificar si el módulo tiene datos válidos
        If cbIngrediente.Value <> "" And cbUnidad.Value <> "" And txtCantidad.Value <> "" Then
            If IsNumeric(txtCantidad.Value) Then
                HayIngredientesValidos = True
                Exit Function
            End If
        End If
    Next i
End Function

Private Sub LimpiarFormulario()
    Dim i As Integer
    
    ' Limpiar campos principales
    Me.Txt_Receta.Value = ""
    Me.Txt_Cantidad.Value = ""
    Me.Txt_Costo.Value = ""
    Me.Txt_Costo_u.Value = ""
    Me.Txt_Precio_u.Value = ""
    Me.Txt_Precio.Value = ""
    
    ' Limpiar todos los módulos
    For i = 1 To 10
        Dim cbIngrediente As MSForms.ComboBox
        Dim txtCantidad As MSForms.TextBox
        Dim cbUnidad As MSForms.ComboBox
        
        Set cbIngrediente = Me.Controls("Cmb" & i)
        Set txtCantidad = Me.Controls("TextBox" & i)
        Set cbUnidad = Me.Controls("Cmb_" & i)
        
        cbIngrediente.Value = ""
        txtCantidad.Value = ""
        cbUnidad.Value = ""
    Next i
    
    Me.Txt_Receta.SetFocus
End Sub



Private Sub TextBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If (KeyAscii >= 46 And KeyAscii <= 57) Then
    KeyAscii = KeyAscii
  Else
    KeyAscii = 0
  End If
End Sub

Private Sub TextBox2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If (KeyAscii >= 46 And KeyAscii <= 57) Then
    KeyAscii = KeyAscii
  Else
    KeyAscii = 0
  End If
End Sub

Private Sub TextBox3_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If (KeyAscii >= 46 And KeyAscii <= 57) Then
    KeyAscii = KeyAscii
  Else
    KeyAscii = 0
  End If
End Sub

Private Sub TextBox4_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If (KeyAscii >= 46 And KeyAscii <= 57) Then
    KeyAscii = KeyAscii
  Else
    KeyAscii = 0
  End If
End Sub

Private Sub TextBox5_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If (KeyAscii >= 46 And KeyAscii <= 57) Then
    KeyAscii = KeyAscii
  Else
    KeyAscii = 0
  End If
End Sub

Private Sub TextBox6_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If (KeyAscii >= 46 And KeyAscii <= 57) Then
    KeyAscii = KeyAscii
  Else
    KeyAscii = 0
  End If
End Sub

Private Sub TextBox7_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If (KeyAscii >= 46 And KeyAscii <= 57) Then
    KeyAscii = KeyAscii
  Else
    KeyAscii = 0
  End If
End Sub

Private Sub TextBox8_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If (KeyAscii >= 46 And KeyAscii <= 57) Then
    KeyAscii = KeyAscii
  Else
    KeyAscii = 0
  End If
End Sub

Private Sub TextBox9_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If (KeyAscii >= 46 And KeyAscii <= 57) Then
    KeyAscii = KeyAscii
  Else
    KeyAscii = 0
  End If
End Sub

Private Sub TextBox10_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If (KeyAscii >= 46 And KeyAscii <= 57) Then
    KeyAscii = KeyAscii
  Else
    KeyAscii = 0
  End If
End Sub

Private Sub Txt_Cantidad_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If (KeyAscii >= 46 And KeyAscii <= 57) Then
    KeyAscii = KeyAscii
  Else
    KeyAscii = 0
  End If
End Sub

Private Sub Txt_Precio_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If (KeyAscii >= 46 And KeyAscii <= 57) Then
    KeyAscii = KeyAscii
  Else
    KeyAscii = 0
  End If
End Sub

Private Sub Txt_Precio_u_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If (KeyAscii >= 46 And KeyAscii <= 57) Then
    KeyAscii = KeyAscii
  Else
    KeyAscii = 0
  End If
End Sub

Private Sub UserForm_Initialize()
    ' Cargar ingredientes en los combobox
    CargarIngredientes
    
    ' Inicializar combobox de unidades vacíos
    Dim i As Integer
    For i = 1 To 10
        Me.Controls("Cmb_" & i).Clear
    Next i

' Si estamos en modo modificación, cargar los datos
    If modoModificacion Then
        Me.Caption = "Modificar Receta"
        Me.Agregar.Caption = "Modificar"
        CargarDatosReceta
    Else
        Me.Caption = "Nueva Receta"
        Me.Agregar.Caption = "Agregar"
    End If
End Sub

' --- FUNCIÓN PARA CARGAR DATOS DE LA RECETA ---
Private Sub CargarDatosReceta()
    Dim wsRecetas As Worksheet
    Dim wsIngredientes As Worksheet
    Dim tblRecetas As ListObject
    Dim tblIngredientesReceta As ListObject
    Dim filaReceta As Range
    Dim filaIngrediente As ListRow
    Dim i As Integer
    
    ' Obtener hojas de trabajo
    Set wsRecetas = ThisWorkbook.Sheets("Elaboración de Productos")
    Set wsIngredientes = ThisWorkbook.Sheets("Ingredientes")
    
    ' Obtener tablas
    Set tblRecetas = wsRecetas.ListObjects("Recetas")
    Set tblIngredientesReceta = wsIngredientes.ListObjects("Ingredientes_Receta")
    
    ' Buscar la receta en la tabla
    For Each filaReceta In tblRecetas.ListColumns("Receta").DataBodyRange
        If filaReceta.Value = recetaAModificar Then
            ' Cargar datos básicos de la receta
            Dim fila As Long
            fila = filaReceta.Row
            
            Me.Txt_Receta.Value = wsRecetas.Cells(fila, tblRecetas.ListColumns("Receta").Range.Column).Value
            Me.Txt_Cantidad.Value = wsRecetas.Cells(fila, tblRecetas.ListColumns("Lote").Range.Column).Value
            Me.Txt_Precio_u.Value = wsRecetas.Cells(fila, tblRecetas.ListColumns("Precio/Unidad").Range.Column).Value
            
            ' Los campos calculados se cargarán automáticamente
            Exit For
        End If
    Next filaReceta
    
    ' Cargar ingredientes de la receta
    i = 1
    For Each filaIngrediente In tblIngredientesReceta.ListRows
        If filaIngrediente.Range.Cells(1, 1).Value = recetaAModificar Then
            If i <= 10 Then ' Solo cargamos hasta 10 ingredientes
                ' Obtener controles del módulo
                Dim cbIngrediente As MSForms.ComboBox
                Dim txtCantidad As MSForms.TextBox
                Dim cbUnidad As MSForms.ComboBox
                
                Set cbIngrediente = Me.Controls("Cmb" & i)
                Set txtCantidad = Me.Controls("TextBox" & i)
                Set cbUnidad = Me.Controls("Cmb_" & i)
                
                ' Cargar datos del ingrediente
                cbIngrediente.Value = filaIngrediente.Range.Cells(1, 2).Value ' Ingrediente
                txtCantidad.Value = Format(filaIngrediente.Range.Cells(1, 3).Value, "0.00") ' Cantidad
                
                ' Configurar las unidades disponibles para este ingrediente
                Dim wsGestión As Worksheet
                Dim tblIngredientes As ListObject
                Dim filaIng As ListRow
                
                Set wsGestión = ThisWorkbook.Sheets("Gestión de Ingredientes")
                Set tblIngredientes = wsGestión.ListObjects("INGREDIENTES")
                
                ' Buscar el ingrediente para obtener su unidad base
                Dim UnidadBase As String
                For Each filaIng In tblIngredientes.ListRows
                    If filaIng.Range.Cells(1, 1).Value = cbIngrediente.Value Then
                        UnidadBase = filaIng.Range.Cells(1, 4).Value
                        Exit For
                    End If
                Next filaIng
                
                ' Cargar unidades según tipo
                cbUnidad.Clear
                Select Case LCase(UnidadBase)
                    Case "g", "kg", "mg", "lb", "oz"
                        cbUnidad.AddItem "g"
                        cbUnidad.AddItem "kg"
                        cbUnidad.AddItem "mg"
                        cbUnidad.AddItem "lb"
                        cbUnidad.AddItem "oz"
                    Case "ml", "l"
                        cbUnidad.AddItem "ml"
                        cbUnidad.AddItem "l"
                    Case "u"
                        cbUnidad.AddItem "u"
                End Select
                
                ' Seleccionar la unidad guardada
                cbUnidad.Value = filaIngrediente.Range.Cells(1, 4).Value ' Unidad
                
                i = i + 1
            End If
        End If
    Next filaIngrediente
    
    ' Calcular costo total después de cargar todos los ingredientes
    CalcularCostoTotal
End Sub

Private Sub CargarIngredientes()
    Dim tblIngredientes As ListObject
    Dim fila As ListRow
    Dim ws As Worksheet
    Dim cbIngrediente As MSForms.ComboBox
    
    ' Obtener la hoja y tabla de ingredientes
    Set ws = ThisWorkbook.Sheets("Gestión de Ingredientes")
    Set tblIngredientes = ws.ListObjects("INGREDIENTES")
    
    ' Cargar cada combobox de ingredientes
    For i = 1 To 10
        Set cbIngrediente = Me.Controls("Cmb" & i)
        cbIngrediente.Clear
        
        ' Agregar opción vacía
        cbIngrediente.AddItem ""
        
        ' Agregar cada ingrediente de la tabla
        For Each fila In tblIngredientes.ListRows
            cbIngrediente.AddItem fila.Range.Cells(1, 1).Value
        Next fila
    Next i
End Sub

' Función para calcular el costo total de la receta
Private Sub CalcularCostoTotal()
    Dim totalCosto As Double
    Dim i As Integer
    Dim ingrediente As String
    Dim Cantidad As Double
    Dim Unidad As String
    Dim cantidadBase As Double
    Dim precioUnitario As Double
    
    ' Inicializar total
    totalCosto = 0
    
    ' Recorrer cada módulo
    For i = 1 To 10
        ' Obtener controles del módulo actual
        Dim cbIngrediente As MSForms.ComboBox
        Dim txtCantidad As MSForms.TextBox
        Dim cbUnidad As MSForms.ComboBox
        
        Set cbIngrediente = Me.Controls("Cmb" & i)
        Set txtCantidad = Me.Controls("TextBox" & i)
        Set cbUnidad = Me.Controls("Cmb_" & i)
        
        ' Verificar que haya ingrediente seleccionado y unidad seleccionada
        If cbIngrediente.Value <> "" And cbUnidad.Value <> "" Then
            ' Verificar que la cantidad sea un número válido
            If IsNumeric(txtCantidad.Value) And txtCantidad.Value <> "" Then
                ingrediente = cbIngrediente.Value
                Cantidad = CDbl(txtCantidad.Value)
                Unidad = cbUnidad.Value
                
                ' Convertir a unidad base usando la función previa
                cantidadBase = ConvertirUnidad(ingrediente, Cantidad, Unidad)
                
                ' Obtener precio por unidad base desde la tabla
                precioUnitario = ObtenerPrecioPorUnidadBase(ingrediente)
                
                ' Sumar al costo total
                totalCosto = totalCosto + (cantidadBase * precioUnitario)
            End If
        End If
    Next i
    
    ' Mostrar el costo total en el TextBox
    Me.Txt_Costo.Value = Format(totalCosto, "0.00")
    
    ' Calcular costo por unidad si hay cantidad de lote
    If Me.Txt_Cantidad.Value <> "" And IsNumeric(Me.Txt_Cantidad.Value) Then
        Dim cantidadLote As Double
        cantidadLote = CDbl(Me.Txt_Cantidad.Value)
        
        If cantidadLote > 0 Then
            Me.Txt_Costo_u.Value = Format(totalCosto / cantidadLote, "0.00")
        End If
    End If
End Sub

' Función para obtener el precio por unidad base desde la tabla
Private Function ObtenerPrecioPorUnidadBase(ingrediente As String) As Double
    Dim ws As Worksheet
    Dim tblIngredientes As ListObject
    Dim fila As ListRow
    
    Set ws = ThisWorkbook.Sheets("Gestión de Ingredientes")
    Set tblIngredientes = ws.ListObjects("INGREDIENTES")
    
    ' Buscar el ingrediente en la tabla
    For Each fila In tblIngredientes.ListRows
        If fila.Range.Cells(1, 1).Value = ingrediente Then
            ' Asumiendo que el precio está en la columna 5
            ObtenerPrecioPorUnidadBase = fila.Range.Cells(1, 5).Value
            Exit Function
        End If
    Next fila
    
    ' Si no encuentra, devolver 0
    ObtenerPrecioPorUnidadBase = 0
End Function

' Eventos Change para los TextBox de cantidad (del 1 al 10)
Private Sub TextBox1_Change()
    CalcularCostoTotal
End Sub

Private Sub TextBox2_Change()
    CalcularCostoTotal
End Sub

Private Sub TextBox3_Change()
    CalcularCostoTotal
End Sub

Private Sub TextBox4_Change()
    CalcularCostoTotal
End Sub

Private Sub TextBox5_Change()
    CalcularCostoTotal
End Sub

Private Sub TextBox6_Change()
    CalcularCostoTotal
End Sub

Private Sub TextBox7_Change()
    CalcularCostoTotal
End Sub

Private Sub TextBox8_Change()
    CalcularCostoTotal
End Sub

Private Sub TextBox9_Change()
    CalcularCostoTotal
End Sub

Private Sub TextBox10_Change()
    CalcularCostoTotal
End Sub



' Eventos Change para los combobox de ingredientes (modificados para actualizar unidades y calcular costo)
Private Sub Cmb1_Change()
    ActualizarUnidades 1
    CalcularCostoTotal
End Sub

Private Sub Cmb2_Change()
    ActualizarUnidades 2
    CalcularCostoTotal
End Sub

Private Sub Cmb3_Change()
    ActualizarUnidades 3
    CalcularCostoTotal
End Sub

Private Sub Cmb4_Change()
    ActualizarUnidades 4
    CalcularCostoTotal
End Sub

Private Sub Cmb5_Change()
    ActualizarUnidades 5
    CalcularCostoTotal
End Sub

Private Sub Cmb6_Change()
    ActualizarUnidades 6
    CalcularCostoTotal
End Sub

Private Sub Cmb7_Change()
    ActualizarUnidades 7
    CalcularCostoTotal
End Sub

Private Sub Cmb8_Change()
    ActualizarUnidades 8
    CalcularCostoTotal
End Sub

Private Sub Cmb9_Change()
    ActualizarUnidades 9
    CalcularCostoTotal
End Sub

Private Sub Cmb10_Change()
    ActualizarUnidades 10
    CalcularCostoTotal
End Sub

' Eventos Change para los combobox de unidades (también deben calcular costo)
Private Sub Cmb_1_Change()
    CalcularCostoTotal
End Sub

Private Sub Cmb_2_Change()
    CalcularCostoTotal
End Sub

Private Sub Cmb_3_Change()
    CalcularCostoTotal
End Sub

Private Sub Cmb_4_Change()
    CalcularCostoTotal
End Sub

Private Sub Cmb_5_Change()
    CalcularCostoTotal
End Sub

Private Sub Cmb_6_Change()
    CalcularCostoTotal
End Sub

Private Sub Cmb_7_Change()
    CalcularCostoTotal
End Sub

Private Sub Cmb_8_Change()
    CalcularCostoTotal
End Sub

Private Sub Cmb_9_Change()
    CalcularCostoTotal
End Sub

Private Sub Cmb_10_Change()
    CalcularCostoTotal
End Sub

' Modificar la función ActualizarUnidades para inicializar el TextBox de cantidad
Private Sub ActualizarUnidades(numeroModulo As Integer)
    Dim ws As Worksheet
    Dim tblIngredientes As ListObject
    Dim cbIngrediente As MSForms.ComboBox
    Dim cbUnidad As MSForms.ComboBox
    Dim txtCantidad As MSForms.TextBox
    Dim ingredienteSeleccionado As String
    Dim UnidadBase As String
    Dim fila As ListRow
    
    Set ws = ThisWorkbook.Sheets("Gestión de Ingredientes")
    Set tblIngredientes = ws.ListObjects("INGREDIENTES")
    Set cbIngrediente = Me.Controls("Cmb" & numeroModulo)
    Set cbUnidad = Me.Controls("Cmb_" & numeroModulo)
    Set txtCantidad = Me.Controls("TextBox" & numeroModulo)
    
    ' Limpiar combobox de unidades y cantidad
    cbUnidad.Clear
    txtCantidad.Value = ""
    
    ' Verificar si hay ingrediente seleccionado
    ingredienteSeleccionado = cbIngrediente.Value
    If ingredienteSeleccionado = "" Then Exit Sub
    
    ' Buscar la unidad base del ingrediente
    For Each fila In tblIngredientes.ListRows
        If fila.Range.Cells(1, 1).Value = ingredienteSeleccionado Then
            UnidadBase = fila.Range.Cells(1, 4).Value
            Exit For
        End If
    Next fila
    
    ' Cargar unidades según tipo
    Select Case LCase(UnidadBase)
        Case "g", "kg", "mg", "lb", "oz"
            cbUnidad.AddItem "g"
            cbUnidad.AddItem "kg"
            cbUnidad.AddItem "mg"
            cbUnidad.AddItem "Lb"
            cbUnidad.AddItem "oz"
            
        Case "ml", "l"
            cbUnidad.AddItem "mL"
            cbUnidad.AddItem "L"
            
        Case "u"
            cbUnidad.AddItem "u"
    End Select
    
    ' Agregar opción vacía al principio
'    If cbUnidad.ListCount > 0 Then
'        cbUnidad.List(0) = ""
'    End If
End Sub

Private Sub Txt_Cantidad_Change()
    ' Calcular precio total del lote
    If Me.Txt_Cantidad.Value <> "" And IsNumeric(Me.Txt_Cantidad.Value) Then
        
        Dim Cantidad As Double
        Cantidad = CDbl(Me.Txt_Cantidad.Value)
        
        If Me.Txt_Precio_u.Value <> "" And IsNumeric(Me.Txt_Precio_u.Value) Then
            Dim precioUnitario As Double
            precioUnitario = CDbl(Me.Txt_Precio_u.Value)
            
            Me.Txt_Precio.Value = Format(precioUnitario * Cantidad, "0.00")
        End If
        
        ' Calcular costo por unidad
        If Me.Txt_Costo.Value <> "" And IsNumeric(Me.Txt_Costo.Value) And Me.Txt_Cantidad.Value <> 0 Then
            Dim costototal As Double
            costototal = CDbl(Me.Txt_Costo.Value)
            
            Me.Txt_Costo_u.Value = Format(costototal / Cantidad, "0.00")
        End If
    End If
End Sub

Private Sub Txt_Precio_u_Change()
    ' Calcular precio total cuando cambia el precio por unidad
    If Me.Txt_Precio_u.Value <> "" And IsNumeric(Me.Txt_Precio_u.Value) Then
        If Me.Txt_Cantidad.Value <> "" And IsNumeric(Me.Txt_Cantidad.Value) Then
            Dim precioUnitario As Double
            Dim Cantidad As Double
            
            precioUnitario = CDbl(Me.Txt_Precio_u.Value)
            Cantidad = CDbl(Me.Txt_Cantidad.Value)
            
            Me.Txt_Precio.Value = Format(precioUnitario * Cantidad, "0.00")
        End If
    End If
End Sub
