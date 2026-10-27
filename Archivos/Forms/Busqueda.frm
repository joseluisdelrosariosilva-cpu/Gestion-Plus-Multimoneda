VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Busqueda 
   Caption         =   "Buscar"
   ClientHeight    =   5540
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   14570
   OleObjectBlob   =   "Busqueda.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Busqueda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private filasOriginales As Collection 'esta variable se usa para guardar la filaIndex de las coincidencias

Private Sub boton_calendario_covered_Click()
    seleccion_destino_calendario = 3
    frmCalendario.Show
End Sub

Private Sub boton_calendario_covered2_Click()
    seleccion_destino_calendario = 4
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

'macro para realizar una busqueda por rango de fechas - versión optimizada
Private Sub Btn_Active_Click()

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim fila As ListRow
    Dim fechaInicio As Date, fechaFin As Date
    Dim fechaTabla As Date
    Dim descripcionBusqueda As String
    Dim resultados() As Variant
    Dim contador As Integer
    Dim i As Integer
    Dim cumpleCondiciones As Boolean
    
    Set ws = ThisWorkbook.Sheets("Historial de Gastos")
    Set tbl = ws.ListObjects("Gastos")
    Set filasOriginales = New Collection
    
    ' Obtener fechas (ya validadas)
    fechaInicio = CDate(Trim(Me.Txt_Fecha))
    fechaFin = CDate(Trim(Me.Txt_Fecha_Final))
    
    ' Obtener descripción del combobox
    If Me.Txt_Descripcion.ListIndex >= 0 Then
        descripcionBusqueda = Trim(Me.Txt_Descripcion.Value)
    Else
        descripcionBusqueda = ""
    End If
    
    ' Primera pasada: contar resultados
    contador = 0
    For Each fila In tbl.ListRows
        fechaTabla = fila.Range(1).Value
        cumpleCondiciones = True
        
        ' Validar rango de fechas
        If fechaTabla < fechaInicio Then cumpleCondiciones = False
        If fechaTabla > fechaFin Then cumpleCondiciones = False
        
        ' Validar descripción (si se especificó)
        If descripcionBusqueda <> "" Then
            If InStr(1, UCase(CStr(fila.Range(3))), UCase(descripcionBusqueda), vbTextCompare) = 0 Then
                cumpleCondiciones = False
            End If
        End If
        
        If cumpleCondiciones Then contador = contador + 1
    Next fila
    
    ' Segunda pasada: llenar array de resultados
    If contador > 0 Then
        ReDim resultados(1 To contador, 1 To 5)
        i = 1
        
        For Each fila In tbl.ListRows
            fechaTabla = fila.Range(1).Value
            cumpleCondiciones = True
            
            ' Validar rango de fechas
            If fechaTabla < fechaInicio Then cumpleCondiciones = False
            If fechaTabla > fechaFin Then cumpleCondiciones = False
            
            ' Validar descripción
            If descripcionBusqueda <> "" Then
                If InStr(1, UCase(CStr(fila.Range(3))), UCase(descripcionBusqueda), vbTextCompare) = 0 Then
                    cumpleCondiciones = False
                End If
            End If
            
            If cumpleCondiciones Then
                resultados(i, 1) = Format(fila.Range(1), "d-mmm-yyyy")
                resultados(i, 2) = fila.Range(2)
                resultados(i, 3) = fila.Range(3)
                resultados(i, 4) = Format(fila.Range(4), "#,##0.00")
                resultados(i, 5) = fila.Range(5)
                
                filasOriginales.Add fila.Range.Row, CStr(i)
                i = i + 1
            End If
        Next fila
        
        Me.lstResultados.List = resultados
    Else
        Me.lstResultados.Clear
    End If

End Sub

' Función auxiliar para mostrar mensaje con resultados
Private Sub MostrarMensajeResultados(tieneInicio As Boolean, tieneFin As Boolean, _
                                     fechaInicio As Date, fechaFin As Date, _
                                     descripcion As String, cantidad As Integer)
    
    Dim mensaje As String
    
    If tieneInicio And tieneFin Then
        mensaje = "Mostrando resultados del " & Format(fechaInicio, "d-mmm-yyyy") & _
                  " al " & Format(fechaFin, "d-mmm-yyyy")
    ElseIf tieneInicio Then
        mensaje = "Mostrando resultados desde el " & Format(fechaInicio, "d-mmm-yyyy")
    ElseIf tieneFin Then
        mensaje = "Mostrando resultados hasta el " & Format(fechaFin, "d-mmm-yyyy")
    Else
        mensaje = "Mostrando todos los resultados"
    End If
    
    If descripcion <> "" Then
        mensaje = mensaje & " con descripción: '" & descripcion & "'"
    End If
    
    Me.lblMensaje.Caption = mensaje & " (" & cantidad & " registros)"
End Sub

' Función auxiliar para mostrar mensaje sin resultados
Private Sub MostrarMensajeSinResultados(tieneInicio As Boolean, tieneFin As Boolean, _
                                        fechaInicio As Date, fechaFin As Date, _
                                        descripcion As String)
    
    Dim criterios As String
    criterios = ""
    
    If tieneInicio Then criterios = criterios & " desde " & Format(fechaInicio, "d-mmm-yyyy")
    If tieneFin Then criterios = criterios & " hasta " & Format(fechaFin, "d-mmm-yyyy")
    If descripcion <> "" Then criterios = criterios & " con descripción: '" & descripcion & "'"
    
    If criterios <> "" Then
        MsgBox "No se encontraron gastos que coincidan con los criterios:" & criterios, vbInformation, "Sin resultados"
    Else
        MsgBox "No se encontraron gastos", vbInformation, "Sin resultados"
    End If
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

Set ws = ThisWorkbook.Sheets("Historial de Gastos")
Set tbl = ws.ListObjects("Gastos")

ws.Activate
Application.GoTo Reference:=ws.Cells(numFilaOriginal + 5, 1), Scroll:=True
ws.Rows(numFilaOriginal).Select


Unload Me

End Sub
'inicializa el listbox con sus columnas
Private Sub UserForm_Initialize()

    Set filasOriginales = New Collection
    
    With Me.lstResultados
        .ColumnCount = 5
        .ColumnWidths = "103;85;173;90;60"
    End With
    Me.Txt_Fecha = Format(Date, "d-mmm-yyyy")
    Me.Txt_Fecha_Final = Format(Date, "d-mmm-yyyy")
    
    CargarDescripcionesUnicas

End Sub

' Version mas simple de cargar descripciones (sin ordenar)
Private Sub CargarDescripcionesUnicas()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim fila As ListRow
    Dim dict As Object
    Dim descripcion As String
    Dim clave As String

    If Me.Txt_Descripcion Is Nothing Then Exit Sub

    Set ws = ThisWorkbook.Sheets("Historial de Gastos")
    Set tbl = ws.ListObjects("Gastos")
    Set dict = CreateObject("Scripting.Dictionary")

    Me.Txt_Descripcion.Clear
    Me.Txt_Descripcion.AddItem "" ' Opcion vacia para no filtrar

    For Each fila In tbl.ListRows
        descripcion = Trim(CStr(fila.Range(3).Value))
        clave = NormalizarDescripcion(descripcion)

        If clave <> "" And Not dict.Exists(clave) Then
            dict.Add clave, clave
            Me.Txt_Descripcion.AddItem clave
        End If
    Next fila

    Me.Txt_Descripcion.ListIndex = 0

    Set dict = Nothing
End Sub

' Un producto aparece como "Refresco" cuando se vendio y como
' "Refresco x 10" cuando se saco por merma. En el combo se muestran los dos
' como "Refresco" para no repetir el mismo producto una vez por cantidad.
' La busqueda es por coincidencia parcial, asi que elegir "Refresco"
' devuelve todas sus filas, con y sin cantidad.
Private Function NormalizarDescripcion(ByVal valor As String) As String
    Dim texto As String
    Dim pos As Long
    Dim sufijo As String

    texto = Trim$(valor)
    pos = InStrRev(texto, " x ")

    If pos = 0 Then
        NormalizarDescripcion = texto
        Exit Function
    End If

    sufijo = Trim$(Mid$(texto, pos + 3))

    If sufijo <> "" And IsNumeric(sufijo) Then
        NormalizarDescripcion = Trim$(Left$(texto, pos - 1))
    Else
        NormalizarDescripcion = texto
    End If
End Function
