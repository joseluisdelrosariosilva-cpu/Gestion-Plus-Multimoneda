VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ingresar_form 
   Caption         =   "Ingresar nuevo producto"
   ClientHeight    =   5890
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   16870
   OleObjectBlob   =   "ingresar_form.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "ingresar_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public producto As String

Private Sub boton_calendario_covered_Click()
    seleccion_destino_calendario = 1
    frmCalendario.Show
End Sub

Private Sub boton_calendario_reposo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Me.boton_calendario_covered.ZOrder msoBringToFront
End Sub

Private Sub boton_ingresar_covered_Click()

'Esta macro permite registrar un nuevo producto en la base de datos del almacen
'primeramente valida los campos vacíos y luego si no existen campos vacíos
'rellena entonces los campos de la tabla, agregando una nueva fila a la tabla Inventario
'Por último, realiza un ordenamiento de la tabla por nombre

    Call validar_vacios
    If contador > 0 Then
        MsgBox "Rellene los campos vacíos", vbExclamation
    Else:
        Dim HojaBD As Worksheet
        Dim tabla As ListObject
        Dim nuevaFila As ListRow
        
        Set HojaBD = ThisWorkbook.Sheets("Gestión de Almacén")
        Set tabla = HojaBD.ListObjects("INVENTARIO")
        
        HojaBD.Unprotect password:="theSheetProtectpassisthe011rd"
        
        ' Agregar nueva fila a la tabla
        Set nuevaFila = tabla.ListRows.Add
        
        ' Asignar valores a las columnas por índice (comenzando desde 1)
        ' A=1, B=2, C=3, D=4, E=5, F=6, G=7, H=8, I=9, J=10, K=11, L=12, M=13, N=14, O=15, P=16
        nuevaFila.Range(1).Value = ingresar_form.codigo_TextBox.Value                    ' Columna A: Código
        nuevaFila.Range(2).Value = ingresar_form.nombre_de_producto_TextBox.Value        ' Columna B: Producto
        nuevaFila.Range(3).Value = CDate(ingresar_form.fecha_de_entrada_TextBox.Value)   ' Columna C: Fecha Entrada
        nuevaFila.Range(4).Value = ingresar_form.cantidad_TextBox.Value                  ' Columna D: Cantidad
        nuevaFila.Range(5).Value = ingresar_form.cantidad_TextBox.Value                  ' Columna E: Stock
        nuevaFila.Range(6).Value = ingresar_form.precio_venta_TextBox.Value              ' Columna F: Precio Venta
        nuevaFila.Range(7).Value = ingresar_form.Moneda_Venta.Value                      ' Columna G: Moneda Venta
        nuevaFila.Range(8).Value = ingresar_form.precio_costo_TextBox.Value              ' Columna H: Precio Costo
        nuevaFila.Range(9).Value = ingresar_form.Moneda_Costo.Value                      ' Columna I: Moneda Costo
        
        ' Calcular fondo (costo total) - Este sí se calcula porque no está en el formulario
        Dim fondo As Double
        fondo = CDbl(ingresar_form.precio_costo_TextBox.Value) * CDbl(ingresar_form.cantidad_TextBox.Value)
        nuevaFila.Range(10).Value = Format(fondo, "0.00")                                                ' Columna J: Fondo
        nuevaFila.Range(11).Value = ingresar_form.Moneda_Costo.Value                     ' Columna K: Moneda Fondo
        
        ' Tomar ingreso y ganancia directamente del formulario
        nuevaFila.Range(12).Value = Format(ingresar_form.Ingreso_TextBox.Value, "0.00")                  ' Columna L: Ingreso Esperado
        nuevaFila.Range(13).Value = ingresar_form.Moneda_Ingreso.Value                   ' Columna M: Moneda Ingreso
        nuevaFila.Range(14).Value = Format(ingresar_form.Ganancia_TextBox.Value, "0.00")                 ' Columna N: Ganancia
        nuevaFila.Range(15).Value = ingresar_form.Moneda_Ganancia.Value                  ' Columna O: Moneda Ganancia
        
        producto = ingresar_form.nombre_de_producto_TextBox.Value
        Agregar_Producto
        Agregar_Fecha
        Unload ingresar_form 'descargar formulario
        
        ' Ordenar la tabla por nombre (columna B o columna 2)
        tabla.Sort.SortFields.Clear
        tabla.Sort.SortFields.Add Key:=tabla.ListColumns(2).Range, SortOn:=xlSortOnValues, Order:=xlAscending
        tabla.Sort.Apply
        HojaBD.Protect password:="theSheetProtectpassisthe011rd", userinterfaceonly:=True
        
    End If
End Sub

Private Sub boton_ingresar_reposo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.boton_ingresar_covered.ZOrder msoBringToFront
    Me.Label8.ZOrder msoBringToFront
End Sub

Private Sub boton_modificar_covered_Click()
    
    Dim HojaBD As Worksheet
    Dim tabla As ListObject
    Dim filaTabla As ListRow
    
    Set HojaBD = ThisWorkbook.Sheets("Gestión de Almacén")
    Set tabla = HojaBD.ListObjects("INVENTARIO")
    
    HojaBD.Unprotect password:="theSheetProtectpassisthe011rd"
    
    ' Obtener la fila específica de la tabla
    Set filaTabla = tabla.ListRows(fila_a_modificar)
    
    ' Validar campos vacíos
    Call validar_vacios
    If contador > 0 Then
        MsgBox "Rellene los campos vacíos", vbExclamation
        Exit Sub
    End If
    
    ' Actualizar los datos en la tabla usando índices
    ' A=1, B=2, C=3, D=4, E=5, F=6, G=7, H=8, I=9, J=10, K=11, L=12, M=13, N=14, O=15
    filaTabla.Range(1).Value = ingresar_form.codigo_TextBox.Value                    ' Columna A: Código
    filaTabla.Range(2).Value = ingresar_form.nombre_de_producto_TextBox.Value        ' Columna B: Producto
    filaTabla.Range(3).Value = CDate(ingresar_form.fecha_de_entrada_TextBox.Value)   ' Columna C: Fecha Entrada
    filaTabla.Range(4).Value = ingresar_form.cantidad_TextBox.Value                  ' Columna D: Cantidad
    filaTabla.Range(5).Value = ingresar_form.cantidad_TextBox.Value                  ' Columna E: Stock
    filaTabla.Range(6).Value = ingresar_form.precio_venta_TextBox.Value              ' Columna F: Precio Venta
    filaTabla.Range(7).Value = ingresar_form.Moneda_Venta.Value                      ' Columna G: Moneda Venta
    filaTabla.Range(8).Value = ingresar_form.precio_costo_TextBox.Value              ' Columna H: Precio Costo
    filaTabla.Range(9).Value = ingresar_form.Moneda_Costo.Value                      ' Columna I: Moneda Costo
    
    ' Calcular fondo (costo total)
    Dim fondo As Double
    fondo = CDbl(ingresar_form.precio_costo_TextBox.Value) * CDbl(ingresar_form.cantidad_TextBox.Value)
    filaTabla.Range(10).Value = fondo                                                ' Columna J: Fondo
    filaTabla.Range(11).Value = ingresar_form.Moneda_Costo.Value                     ' Columna K: Moneda Fondo
    
    ' Ingreso y ganancia (tomados directamente del formulario)
    filaTabla.Range(12).Value = ingresar_form.Ingreso_TextBox.Value                  ' Columna L: Ingreso Esperado
    filaTabla.Range(13).Value = ingresar_form.Moneda_Ingreso.Value                   ' Columna M: Moneda Ingreso
    filaTabla.Range(14).Value = ingresar_form.Ganancia_TextBox.Value                 ' Columna N: Ganancia
    filaTabla.Range(15).Value = ingresar_form.Moneda_Ganancia.Value                  ' Columna O: Moneda Ganancia
    
    producto = ingresar_form.nombre_de_producto_TextBox.Value
    Agregar_Producto
    
    ' Ordenar la tabla por nombre (columna B o columna 2)
    tabla.Sort.SortFields.Clear
    tabla.Sort.SortFields.Add Key:=tabla.ListColumns(2).Range, SortOn:=xlSortOnValues, Order:=xlAscending
    tabla.Sort.Apply
    
    Unload ingresar_form
    
    Call setear_funcion
    HojaBD.Protect password:="theSheetProtectpassisthe011rd", userinterfaceonly:=True
    
End Sub

'Validacion para evitar que se escriban letras
Private Sub cantidad_TextBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If (KeyAscii >= 46 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
      Else
        KeyAscii = 0
      End If
End Sub

Private Sub fondo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.boton_ingresar_reposo.ZOrder msoBringToFront
    Me.boton_modificar_reposo.ZOrder msoBringToFront
    Me.boton_calendario_reposo.ZOrder msoBringToFront
End Sub

Private Sub boton_modificar_reposo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.boton_modificar_covered.Visible = True
    Me.Label9.Visible = True
    Me.boton_modificar_covered.ZOrder msoBringToFront
    Me.Label9.ZOrder msoBringToFront
End Sub

Private Sub Label8_Click()
    boton_ingresar_covered_Click
End Sub

Private Sub Label9_Click()
    boton_modificar_covered_Click
End Sub

Private Sub precio_costo_TextBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If (KeyAscii >= 46 And KeyAscii <= 57) Then
    KeyAscii = KeyAscii
  Else
    KeyAscii = 0
  End If
End Sub


Private Sub precio_venta_TextBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If (KeyAscii >= 46 And KeyAscii <= 57) Then
    KeyAscii = KeyAscii
  Else
    KeyAscii = 0
  End If
End Sub


Private Sub UserForm_Activate()
    'Esto es para diferenciar el boton registrar y el de modificar para poderlo
    'poner en el mismo formulario. El objetivo es que modificar_registrar_textbox
    'funcione como una variable que compruebe que macro ha sido ejecutada, y poder entonces
    'realiza un ajuste del boton
    
    Dim HojaBD As Worksheet
    Dim tabla As ListObject
    Dim filaTabla As ListRow
    
    Set HojaBD = ThisWorkbook.Sheets("Gestión de Almacén")
    Set tabla = HojaBD.ListObjects("INVENTARIO")
    
    If Me.modificar_registrar_textbox.Value = "R" Then
        
        Me.boton_modificar_reposo.Visible = False
        boton_ingresar_reposo.Visible = True
        Call Genera_Codigo
        Me.codigo_TextBox.Value = Codigo_Producto
        Me.fecha_de_entrada_TextBox.Value = Format(Date, "d-mmm-yyyy")
        
    End If
    
    If Me.modificar_registrar_textbox.Value = "M" Then
        
        Me.boton_modificar_reposo.Visible = True
        Me.boton_ingresar_reposo.Visible = False
        Me.Label9.Visible = True
        
        ' Obtener la fila específica de la tabla
        Set filaTabla = tabla.ListRows(fila_a_modificar)
        
        With ingresar_form
            .codigo_TextBox.Value = filaTabla.Range(1).Value
            .nombre_de_producto_TextBox.Value = filaTabla.Range(2).Value
            .fecha_de_entrada_TextBox.Value = Format(filaTabla.Range(3).Value, "d-mmm-yyyy")
            .cantidad_TextBox.Value = filaTabla.Range(4).Value
            .precio_venta_TextBox.Value = filaTabla.Range(6).Value
            .Moneda_Venta.Value = filaTabla.Range(7).Value
            .precio_costo_TextBox.Value = filaTabla.Range(8).Value
            .Moneda_Costo.Value = filaTabla.Range(9).Value
            .Ingreso_TextBox.Value = filaTabla.Range(12).Value
            .Moneda_Ingreso.Value = filaTabla.Range(13).Value
            .Ganancia_TextBox.Value = filaTabla.Range(14).Value
            .Moneda_Ganancia.Value = filaTabla.Range(15).Value
        End With
        
    End If
    
    ' Cargar opciones de moneda en los ComboBox
    With Me.Moneda_Costo
        .Clear
        .AddItem "MN"
        .AddItem "USD"
    End With
    
    With Me.Moneda_Venta
        .Clear
        .AddItem "MN"
        .AddItem "USD"
    End With
    
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
        If CDate(fila.Range(1)) = CDate(ingresar_form.fecha_de_entrada_TextBox.Value) Then
            Exit Sub
        End If
    Next fila
    
    ws.Unprotect password:="theSheetProtectpassisthe011rd"
    
    Set newrow = tbl.ListRows.Add
    
    newrow.Range(1) = CDate(ingresar_form.fecha_de_entrada_TextBox.Value)
    
    tbl.Sort.Apply
    
    ws.Protect password:="theSheetProtectpassisthe011rd", userinterfaceonly:=True

End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.boton_ingresar_reposo.ZOrder msoBringToFront
    Me.Label8.ZOrder msoBringToFront
    Me.boton_modificar_reposo.ZOrder msoBringToFront
    Me.Label9.ZOrder msoBringToFront
    Me.boton_calendario_reposo.ZOrder msoBringToFront
End Sub

Private Sub cantidad_TextBox_Change()
    If Me.cantidad_TextBox.Value <> "" And Me.precio_venta_TextBox.Value <> "" Then
        Dim ingresoEsperado As Double
        Dim beneficioEsperado As Double
        Dim fondo As Double
        Dim tipoCambio As Double
        
        ' Obtener tipo de cambio de la celda G3
        tipoCambio = ThisWorkbook.Sheets("Gestión de Almacén").Range("G3").Value
        
        ' Obtener monedas seleccionadas
        Dim monedaCosto As String
        Dim monedaVenta As String
        
        monedaCosto = Me.Moneda_Costo.Value
        monedaVenta = Me.Moneda_Venta.Value
        
        ' Calcular fondo (costo total)
        If Me.precio_costo_TextBox.Value <> "" Then
            fondo = CDbl(Me.precio_costo_TextBox.Value) * CDbl(Me.cantidad_TextBox.Value)
        End If
        
        ' Calcular ingreso esperado
        ingresoEsperado = CDbl(Me.precio_venta_TextBox.Value) * CDbl(Me.cantidad_TextBox.Value)
        
        ' Determinar moneda del ingreso según caso
        Select Case monedaCosto & "-" & monedaVenta
            Case "MN-MN"
                ' Caso 1: Costo MN, Venta MN
                Me.Ingreso_TextBox.Value = Format(ingresoEsperado, "0.00")
                Me.Moneda_Ingreso.Value = "MN"
                
                ' Calcular ganancia en MN
                If Me.precio_costo_TextBox.Value <> "" Then
                    beneficioEsperado = ingresoEsperado - fondo
                    Me.Ganancia_TextBox.Value = Format(beneficioEsperado, "0.00")
                    Me.Moneda_Ganancia.Value = "MN"
                End If
                
            Case "MN-USD"
                ' Caso 2: Costo MN, Venta USD
                Me.Ingreso_TextBox.Value = Format(ingresoEsperado, "0.00")
                Me.Moneda_Ingreso.Value = "USD"
                
                ' Calcular ganancia en USD (convertir fondo de MN a USD)
                If Me.precio_costo_TextBox.Value <> "" Then
                    beneficioEsperado = ingresoEsperado - (fondo / tipoCambio)
                    Me.Ganancia_TextBox.Value = Format(beneficioEsperado, "0.00")
                    Me.Moneda_Ganancia.Value = "USD"
                End If
                
            Case "USD-MN"
                ' Caso 3: Costo USD, Venta MN
                Me.Ingreso_TextBox.Value = Format(ingresoEsperado, "0.00")
                Me.Moneda_Ingreso.Value = "MN"
                
                ' Calcular ganancia en MN (convertir fondo de USD a MN)
                If Me.precio_costo_TextBox.Value <> "" Then
                    beneficioEsperado = ingresoEsperado - (fondo * tipoCambio)
                    Me.Ganancia_TextBox.Value = Format(beneficioEsperado, "0.00")
                    Me.Moneda_Ganancia.Value = "MN"
                End If
                
            Case "USD-USD"
                ' Caso 4: Costo USD, Venta USD
                Me.Ingreso_TextBox.Value = Format(ingresoEsperado, "0.00")
                Me.Moneda_Ingreso.Value = "USD"
                
                ' Calcular ganancia en USD
                If Me.precio_costo_TextBox.Value <> "" Then
                    beneficioEsperado = ingresoEsperado - fondo
                    Me.Ganancia_TextBox.Value = Format(beneficioEsperado, "0.00")
                    Me.Moneda_Ganancia.Value = "USD"
                End If
        End Select
        
    Else
        Me.Ingreso_TextBox.Value = ""
        Me.Ganancia_TextBox.Value = ""
    End If
End Sub

Private Sub precio_venta_TextBox_Change()
    If Me.cantidad_TextBox.Value <> "" And Me.precio_venta_TextBox.Value <> "" Then
        Dim ingresoEsperado As Double
        Dim beneficioEsperado As Double
        Dim fondo As Double
        Dim tipoCambio As Double
        
        ' Obtener tipo de cambio de la celda G3
        tipoCambio = ThisWorkbook.Sheets("Gestión de Almacén").Range("G3").Value
        
        ' Obtener monedas seleccionadas
        Dim monedaCosto As String
        Dim monedaVenta As String
        
        monedaCosto = Me.Moneda_Costo.Value
        monedaVenta = Me.Moneda_Venta.Value
        
        ' Calcular fondo (costo total)
        If Me.precio_costo_TextBox.Value <> "" Then
            fondo = CDbl(Me.precio_costo_TextBox.Value) * CDbl(Me.cantidad_TextBox.Value)
        End If
        
        ' Calcular ingreso esperado
        ingresoEsperado = CDbl(Me.precio_venta_TextBox.Value) * CDbl(Me.cantidad_TextBox.Value)
        
        ' Determinar moneda del ingreso según caso
        Select Case monedaCosto & "-" & monedaVenta
            Case "MN-MN"
                ' Caso 1: Costo MN, Venta MN
                Me.Ingreso_TextBox.Value = Format(ingresoEsperado, "0.00")
                Me.Moneda_Ingreso.Value = "MN"
                
                ' Calcular ganancia en MN
                If Me.precio_costo_TextBox.Value <> "" Then
                    beneficioEsperado = ingresoEsperado - fondo
                    Me.Ganancia_TextBox.Value = Format(beneficioEsperado, "0.00")
                    Me.Moneda_Ganancia.Value = "MN"
                End If
                
            Case "MN-USD"
                ' Caso 2: Costo MN, Venta USD
                Me.Ingreso_TextBox.Value = Format(ingresoEsperado, "0.00")
                Me.Moneda_Ingreso.Value = "USD"
                
                ' Calcular ganancia en USD (convertir fondo de MN a USD)
                If Me.precio_costo_TextBox.Value <> "" Then
                    beneficioEsperado = ingresoEsperado - (fondo / tipoCambio)
                    Me.Ganancia_TextBox.Value = Format(beneficioEsperado, "0.00")
                    Me.Moneda_Ganancia.Value = "USD"
                End If
                
            Case "USD-MN"
                ' Caso 3: Costo USD, Venta MN
                Me.Ingreso_TextBox.Value = Format(ingresoEsperado, "0.00")
                Me.Moneda_Ingreso.Value = "MN"
                
                ' Calcular ganancia en MN (convertir fondo de USD a MN)
                If Me.precio_costo_TextBox.Value <> "" Then
                    beneficioEsperado = ingresoEsperado - (fondo * tipoCambio)
                    Me.Ganancia_TextBox.Value = Format(beneficioEsperado, "0.00")
                    Me.Moneda_Ganancia.Value = "MN"
                End If
                
            Case "USD-USD"
                ' Caso 4: Costo USD, Venta USD
                Me.Ingreso_TextBox.Value = Format(ingresoEsperado, "0.00")
                Me.Moneda_Ingreso.Value = "USD"
                
                ' Calcular ganancia en USD
                If Me.precio_costo_TextBox.Value <> "" Then
                    beneficioEsperado = ingresoEsperado - fondo
                    Me.Ganancia_TextBox.Value = Format(beneficioEsperado, "0.00")
                    Me.Moneda_Ganancia.Value = "USD"
                End If
        End Select
        
    Else
        Me.Ingreso_TextBox.Value = ""
        Me.Ganancia_TextBox.Value = ""
    End If
End Sub

Private Sub precio_costo_TextBox_Change()
    If Me.cantidad_TextBox.Value <> "" And Me.precio_costo_TextBox.Value <> "" And Me.precio_venta_TextBox.Value <> "" Then
        Dim fondo As Double
        Dim Ingreso As Double
        Dim beneficioEsperado As Double
        Dim tipoCambio As Double
        
        ' Obtener tipo de cambio de la celda G3
        tipoCambio = ThisWorkbook.Sheets("Gestión de Almacén").Range("G3").Value
        
        ' Obtener monedas seleccionadas
        Dim monedaCosto As String
        Dim monedaVenta As String
        
        monedaCosto = Me.Moneda_Costo.Value
        monedaVenta = Me.Moneda_Venta.Value
        
        ' Calcular fondo (costo total)
        fondo = CDbl(Me.precio_costo_TextBox.Value) * CDbl(Me.cantidad_TextBox.Value)
        
        ' Calcular ingreso esperado
        Ingreso = CDbl(Me.precio_venta_TextBox.Value) * CDbl(Me.cantidad_TextBox.Value)
        
        ' Recalcular ingreso y ganancia según los casos
        Select Case monedaCosto & "-" & monedaVenta
            Case "MN-MN"
                ' Caso 1: Costo MN, Venta MN
                Me.Ingreso_TextBox.Value = Format(Ingreso, "0.00")
                Me.Moneda_Ingreso.Value = "MN"
                
                ' Calcular ganancia en MN
                beneficioEsperado = Ingreso - fondo
                Me.Ganancia_TextBox.Value = Format(beneficioEsperado, "0.00")
                Me.Moneda_Ganancia.Value = "MN"
                
            Case "MN-USD"
                ' Caso 2: Costo MN, Venta USD
                Me.Ingreso_TextBox.Value = Format(Ingreso, "0.00")
                Me.Moneda_Ingreso.Value = "USD"
                
                ' Calcular ganancia en USD (convertir fondo de MN a USD)
                beneficioEsperado = Ingreso - (fondo / tipoCambio)
                Me.Ganancia_TextBox.Value = Format(beneficioEsperado, "0.00")
                Me.Moneda_Ganancia.Value = "USD"
                
            Case "USD-MN"
                ' Caso 3: Costo USD, Venta MN
                Me.Ingreso_TextBox.Value = Format(Ingreso, "0.00")
                Me.Moneda_Ingreso.Value = "MN"
                
                ' Calcular ganancia en MN (convertir fondo de USD a MN)
                beneficioEsperado = Ingreso - (fondo * tipoCambio)
                Me.Ganancia_TextBox.Value = Format(beneficioEsperado, "0.00")
                Me.Moneda_Ganancia.Value = "MN"
                
            Case "USD-USD"
                ' Caso 4: Costo USD, Venta USD
                Me.Ingreso_TextBox.Value = Format(Ingreso, "0.00")
                Me.Moneda_Ingreso.Value = "USD"
                
                ' Calcular ganancia en USD
                beneficioEsperado = Ingreso - fondo
                Me.Ganancia_TextBox.Value = Format(beneficioEsperado, "0.00")
                Me.Moneda_Ganancia.Value = "USD"
        End Select
    End If
End Sub

Private Sub Moneda_Costo_Change()
    ' Cuando cambia la moneda del costo, se deben recalcular los valores
    ' que dependen de esta moneda (fondo y ganancia)
    
    If Me.cantidad_TextBox.Value <> "" And Me.precio_costo_TextBox.Value <> "" And Me.precio_venta_TextBox.Value <> "" Then
        Dim fondo As Double
        Dim Ingreso As Double
        Dim beneficioEsperado As Double
        Dim tipoCambio As Double
        
        ' Obtener tipo de cambio de la celda G3
        tipoCambio = ThisWorkbook.Sheets("Gestión de Almacén").Range("G3").Value
        
        ' Obtener monedas seleccionadas
        Dim monedaCosto As String
        Dim monedaVenta As String
        
        monedaCosto = Me.Moneda_Costo.Value
        monedaVenta = Me.Moneda_Venta.Value
        
        ' Calcular fondo (costo total)
        fondo = CDbl(Me.precio_costo_TextBox.Value) * CDbl(Me.cantidad_TextBox.Value)
        
        ' Calcular ingreso esperado
        Ingreso = CDbl(Me.precio_venta_TextBox.Value) * CDbl(Me.cantidad_TextBox.Value)
        
        ' Recalcular ingreso y ganancia según los casos
        Select Case monedaCosto & "-" & monedaVenta
            Case "MN-MN"
                ' Caso 1: Costo MN, Venta MN
                Me.Ingreso_TextBox.Value = Format(Ingreso, "0.00")
                Me.Moneda_Ingreso.Value = "MN"
                
                ' Calcular ganancia en MN
                beneficioEsperado = Ingreso - fondo
                Me.Ganancia_TextBox.Value = Format(beneficioEsperado, "0.00")
                Me.Moneda_Ganancia.Value = "MN"
                
            Case "MN-USD"
                ' Caso 2: Costo MN, Venta USD
                Me.Ingreso_TextBox.Value = Format(Ingreso, "0.00")
                Me.Moneda_Ingreso.Value = "USD"
                
                ' Calcular ganancia en USD (convertir fondo de MN a USD)
                beneficioEsperado = Ingreso - (fondo / tipoCambio)
                Me.Ganancia_TextBox.Value = Format(beneficioEsperado, "0.00")
                Me.Moneda_Ganancia.Value = "USD"
                
            Case "USD-MN"
                ' Caso 3: Costo USD, Venta MN
                Me.Ingreso_TextBox.Value = Format(Ingreso, "0.00")
                Me.Moneda_Ingreso.Value = "MN"
                
                ' Calcular ganancia en MN (convertir fondo de USD a MN)
                beneficioEsperado = Ingreso - (fondo * tipoCambio)
                Me.Ganancia_TextBox.Value = Format(beneficioEsperado, "0.00")
                Me.Moneda_Ganancia.Value = "MN"
                
            Case "USD-USD"
                ' Caso 4: Costo USD, Venta USD
                Me.Ingreso_TextBox.Value = Format(Ingreso, "0.00")
                Me.Moneda_Ingreso.Value = "USD"
                
                ' Calcular ganancia en USD
                beneficioEsperado = Ingreso - fondo
                Me.Ganancia_TextBox.Value = Format(beneficioEsperado, "0.00")
                Me.Moneda_Ganancia.Value = "USD"
        End Select
    End If
End Sub

Private Sub Moneda_Venta_Change()
    ' Cuando cambia la moneda de venta, se deben recalcular los valores
    ' que dependen de esta moneda (ingreso y ganancia)
    
    If Me.cantidad_TextBox.Value <> "" And Me.precio_costo_TextBox.Value <> "" And Me.precio_venta_TextBox.Value <> "" Then
        Dim fondo As Double
        Dim Ingreso As Double
        Dim beneficioEsperado As Double
        Dim tipoCambio As Double
        
        ' Obtener tipo de cambio de la celda G3
        tipoCambio = ThisWorkbook.Sheets("Gestión de Almacén").Range("G3").Value
        
        ' Obtener monedas seleccionadas
        Dim monedaCosto As String
        Dim monedaVenta As String
        
        monedaCosto = Me.Moneda_Costo.Value
        monedaVenta = Me.Moneda_Venta.Value
        
        ' Calcular fondo (costo total)
        fondo = CDbl(Me.precio_costo_TextBox.Value) * CDbl(Me.cantidad_TextBox.Value)
        
        ' Calcular ingreso esperado
        Ingreso = CDbl(Me.precio_venta_TextBox.Value) * CDbl(Me.cantidad_TextBox.Value)
        
        ' Recalcular ingreso y ganancia según los casos
        Select Case monedaCosto & "-" & monedaVenta
            Case "MN-MN"
                ' Caso 1: Costo MN, Venta MN
                Me.Ingreso_TextBox.Value = Format(Ingreso, "0.00")
                Me.Moneda_Ingreso.Value = "MN"
                
                ' Calcular ganancia en MN
                beneficioEsperado = Ingreso - fondo
                Me.Ganancia_TextBox.Value = Format(beneficioEsperado, "0.00")
                Me.Moneda_Ganancia.Value = "MN"
                
            Case "MN-USD"
                ' Caso 2: Costo MN, Venta USD
                Me.Ingreso_TextBox.Value = Format(Ingreso, "0.00")
                Me.Moneda_Ingreso.Value = "USD"
                
                ' Calcular ganancia en USD (convertir fondo de MN a USD)
                beneficioEsperado = Ingreso - (fondo / tipoCambio)
                Me.Ganancia_TextBox.Value = Format(beneficioEsperado, "0.00")
                Me.Moneda_Ganancia.Value = "USD"
                
            Case "USD-MN"
                ' Caso 3: Costo USD, Venta MN
                Me.Ingreso_TextBox.Value = Format(Ingreso, "0.00")
                Me.Moneda_Ingreso.Value = "MN"
                
                ' Calcular ganancia en MN (convertir fondo de USD a MN)
                beneficioEsperado = Ingreso - (fondo * tipoCambio)
                Me.Ganancia_TextBox.Value = Format(beneficioEsperado, "0.00")
                Me.Moneda_Ganancia.Value = "MN"
                
            Case "USD-USD"
                ' Caso 4: Costo USD, Venta USD
                Me.Ingreso_TextBox.Value = Format(Ingreso, "0.00")
                Me.Moneda_Ingreso.Value = "USD"
                
                ' Calcular ganancia en USD
                beneficioEsperado = Ingreso - fondo
                Me.Ganancia_TextBox.Value = Format(beneficioEsperado, "0.00")
                Me.Moneda_Ganancia.Value = "USD"
        End Select
    End If
End Sub
