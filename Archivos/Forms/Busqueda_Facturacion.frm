VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Busqueda_Facturacion 
   Caption         =   "Buscar"
   ClientHeight    =   5520
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   14880
   OleObjectBlob   =   "Busqueda_Facturacion.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Busqueda_Facturacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private filasOriginales As Collection 'esta variable se usa para guardar la filaIndex de las coincidencias

Private Sub boton_calendario_covered_Click()
    seleccion_destino_calendario = 7
    frmCalendario.Show
End Sub

Private Sub boton_calendario_covered2_Click()
    seleccion_destino_calendario = 8
    frmCalendario.Show
End Sub



Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.Btn_Reposo.ZOrder msoBringToFront
    Me.boton_calendario_reposo.ZOrder msoBringToFront
    Me.boton_calendario_reposo2.ZOrder msoBringToFront
End Sub

Private Sub Btn_Reposo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.Btn_Active.ZOrder msoBringToFront
End Sub

Private Sub boton_calendario_reposo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.boton_calendario_covered.ZOrder msoBringToFront
End Sub

Private Sub boton_calendario_reposo2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.boton_calendario_covered2.ZOrder msoBringToFront
End Sub

Private Sub Btn_Active_Click()

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim fila As ListRow
    Dim fechaInicio As Date, fechaFin As Date
    Dim descripcionBusqueda As String
    Dim resultados() As Variant
    Dim contador As Integer
    Dim i As Integer
    
    Set ws = ThisWorkbook.Sheets("Registro de facturas")
    Set tbl = ws.ListObjects("HISTORIALVENTAS4")
    Set filasOriginales = New Collection
    
    ' Obtener valores (ya validados)
    fechaInicio = CDate(Trim(Me.Txt_Fecha))
    fechaFin = CDate(Trim(Me.Txt_Fecha_Final))
    
    ' Obtener descripción del combobox
    descripcionBusqueda = IIf(Me.Txt_Descripcion.ListIndex > 0, Trim(Me.Txt_Descripcion.Value), "")
    
    ' Contar resultados
    contador = 0
    For Each fila In tbl.ListRows
        If CumpleCriteriosFacturas(fila, fechaInicio, fechaFin, descripcionBusqueda) Then
            contador = contador + 1
        End If
    Next fila
    
    ' Llenar array de resultados
    If contador > 0 Then
        ReDim resultados(1 To contador, 1 To 5)
        i = 1
        
        For Each fila In tbl.ListRows
            If CumpleCriteriosFacturas(fila, fechaInicio, fechaFin, descripcionBusqueda) Then
                resultados(i, 1) = fila.Range(1)
                resultados(i, 2) = Format(fila.Range(2), "d-mmm-yyyy")
                resultados(i, 3) = Format(fila.Range(3), "#,##0.00")
                resultados(i, 4) = Format(fila.Range(4), "#,##0.00")
                resultados(i, 5) = Format(fila.Range(5), "#,##0.00")
                
                filasOriginales.Add fila.Range.Row, CStr(i)
                i = i + 1
            End If
        Next fila
        
        Me.lstResultados.List = resultados
    Else
        Me.lstResultados.Clear
    End If

End Sub

' Función para verificar si una fila cumple con los criterios en FACTURAS
Private Function CumpleCriteriosFacturas(fila As ListRow, fechaInicio As Date, fechaFin As Date, descripcionBusqueda As String) As Boolean
    
    Dim fechaTabla As Date
    fechaTabla = fila.Range(2).Value ' Columna 2 es la fecha en la tabla de facturas
    
    ' Verificar rango de fechas
    If fechaTabla < fechaInicio Or fechaTabla > fechaFin Then
        CumpleCriteriosFacturas = False
        Exit Function
    End If
    
    ' Verificar descripción (si se especificó)
    If descripcionBusqueda <> "" Then
        If InStr(1, UCase(CStr(fila.Range(1))), UCase(descripcionBusqueda), vbTextCompare) = 0 Then
            CumpleCriteriosFacturas = False
            Exit Function
        End If
    End If
    
    CumpleCriteriosFacturas = True
End Function

' Función para cargar las descripciones únicas en el combobox para FACTURAS
Private Sub CargarDescripcionesUnicasFacturas()

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim fila As ListRow
    Dim dict As Object
    
    If Me.Txt_Descripcion Is Nothing Then Exit Sub
    
    Set ws = ThisWorkbook.Sheets("Registro de facturas")
    Set tbl = ws.ListObjects("HISTORIALVENTAS4")
    Set dict = CreateObject("Scripting.Dictionary")
    
    Me.Txt_Descripcion.Clear
    Me.Txt_Descripcion.AddItem ""
    
    For Each fila In tbl.ListRows
        Dim descripcion As String
        descripcion = Trim(CStr(fila.Range(1).Value)) ' Columna 1 es la descripción en facturas
        
        If descripcion <> "" And Not dict.Exists(descripcion) Then
            dict.Add descripcion, descripcion
            Me.Txt_Descripcion.AddItem descripcion
        End If
    Next fila
    
    If Me.Txt_Descripcion.ListCount > 0 Then
        Me.Txt_Descripcion.ListIndex = 0
    End If
    
    Set dict = Nothing
End Sub


'aqui se maneja para el doble clic, mostrar el dato en la tabla original
Private Sub lstResultados_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    Dim filaSeleccionada As Integer
    Dim numFilaOriginal As Integer
    Dim ws As Worksheet
    Dim tbl As ListObject
    
    If Me.lstResultados.ListIndex = -1 Then Exit Sub
    
    filaSeleccionada = Me.lstResultados.ListIndex + 1
    numFilaOriginal = filasOriginales.Item(CStr(filaSeleccionada))
    
    Set ws = ThisWorkbook.Sheets("Registro de facturas")
    Set tbl = ws.ListObjects("HISTORIALVENTAS4")
    
    ws.Activate
    Application.GoTo Reference:=ws.Cells(numFilaOriginal + 5, 1), Scroll:=True
    ws.Rows(numFilaOriginal).Select


    Unload Me

End Sub
'inicializa el listbox con sus columnas
Private Sub UserForm_Initialize()

    CargarDescripcionesUnicasFacturas
    
    Set filasOriginales = New Collection
    
    With Me.lstResultados
        .ColumnCount = 5
        .ColumnWidths = "95;90;110;110;110"
    End With
    
    Me.Txt_Fecha = Format(Date, "d-mmm-yyyy")
    Me.Txt_Fecha_Final = Format(Date, "d-mmm-yyyy")
    
End Sub


