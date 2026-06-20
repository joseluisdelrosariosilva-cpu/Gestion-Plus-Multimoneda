VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    
    crearRespaldoFinal
    
    desprotegerLibro

    ThisWorkbook.Sheets("Blanco").Visible = xlSheetVisible
    ThisWorkbook.Sheets("Historial de Gastos").Visible = xlSheetVeryHidden
    ThisWorkbook.Sheets("Registro de facturas").Visible = xlSheetVeryHidden
    ThisWorkbook.Sheets("Historial de ventas").Visible = xlSheetVeryHidden
    ThisWorkbook.Sheets("Gestión de Almacén").Visible = xlSheetVeryHidden
    ThisWorkbook.Sheets("Dashboard").Visible = xlSheetVeryHidden
    ThisWorkbook.Sheets("Gestión de Trabajadores").Visible = xlSheetVeryHidden
    ThisWorkbook.Sheets("Gestión de Ingredientes").Visible = xlSheetVeryHidden
    ThisWorkbook.Sheets("Elaboración de Productos").Visible = xlSheetVeryHidden
    ThisWorkbook.Save

    protegerLibro
    
    Application.Quit
    
End Sub

Private Sub Workbook_Open()
    Application.Visible = False
'verifica el dispositivo
    If Not verificarDispositivo() Then
        MsgBox "Este libro no está autorizado para su uso en este equipo", vbCritical
        ThisWorkbook.Close SaveChanges:=False
    End If
'verifica que no se haya cambiado la fecha en el dispositivo
verificarIntegridadFecha
    
'verifica la vigencia
'    If Not verificarVigencia() Then
'        If Not solicitarCodigoRenovacion() Then
'            MsgBox "El libro ha expirado y no se ha proporcionado un código válido.", vbCritical
'            ThisWorkbook.Close SaveChanges:=False
'        End If
'    End If
    
    protegerHojas
    
    facturacion.Show
    
End Sub

'genera el id unico del dispositivo
Private Function ObtenerIDDispositivo() As String

    Dim id As String
    
    id = Environ("COMPUTERNAME") & "|" & Environ("USERNAME") & "|" & _
            CreateObject("Scripting.FileSystemObject").GetDrive("C:").SerialNumber
    ObtenerIDDispositivo = id
    
End Function
'macro para proteger todas las hojas cuando se abre el libro
Private Sub protegerHojas()

    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        ws.Protect password:="theSheetProtectpassisthe011rd", _
        userinterfaceonly:=True
    Next ws
    
    ThisWorkbook.Worksheets("Dashboard").Protect password:="theSheetProtectpassisthe011rd", _
        contents:=True, _
        AllowFiltering:=True, _
        AllowSorting:=True, _
        AllowUsingPivotTables:=True, _
        userinterfaceonly:=True
        
End Sub
'Macro para verificar la vigencia
Private Function verificarVigencia() As Boolean
    
    Dim ws As Worksheet
    Dim fechaExpiracion As Date
    On Error GoTo ErrorHandler
    
    Set ws = ThisWorkbook.Sheets("Hidden")
    fechaExpiracion = ws.Range("A1")
        If Date >= fechaExpiracion - 3 And Date <= fechaExpiracion Then
            MsgBox "Este libro está pronto a expirar. Contacte con su proveedor.", vbInformation
        End If
    verificarVigencia = (Date <= fechaExpiracion)
    Exit Function
    
ErrorHandler:

    ws.Unprotect password:="theSheetProtectpassisthe011rd"
    fechaExpiracion = DateAdd("d", 1, Date)
    ThisWorkbook.Sheets("Hidden").Range("A1") = fechaExpiracion
    ThisWorkbook.Sheets("Hidden").Visible = xlSheetVeryHidden
    verificarVigencia = True
    ws.Protect password:="theSheetProtectpassisthe011rd", userinterfaceonly:=True
    
End Function
'macro para solicitar codigo de licencia
Private Function solicitarCodigoRenovacion() As Boolean
    
    Dim ws As Worksheet
    Dim codigoIngresado As String
    Dim codigoValido_mes As String
    Dim codigoValido_anual As String
    
    Set ws = ThisWorkbook.Sheets("Hidden")
    
    codigoValido_mes = GenerarCodigoMensual(Day(Date), Month(Date), year(Date))
    codigoValido_anual = GenerarCodigoAnual(Day(Date), Month(Date), year(Date))
    
    codigoIngresado = InputBox("El libro ha expirado. Por favor, ingrese el código de renovación:", "Renovación requerida")
    
    If codigoIngresado = codigoValido_mes Then
        ws.Unprotect password:="theSheetProtectpassisthe011rd"
        ThisWorkbook.Sheets("Hidden").Range("A1") = DateAdd("d", 30, Date)
        solicitarCodigoRenovacion = True
        ws.Protect password:="theSheetProtectpassisthe011rd", userinterfaceonly:=True
        MsgBox "Gracias. Se han añadido 30 días a la vigencia del libro", vbInformation
        
    ElseIf codigoIngresado = codigoValido_anual Then
        ws.Unprotect password:="theSheetProtectpassisthe011rd"
        ThisWorkbook.Sheets("Hidden").Range("A1") = DateAdd("d", 180, Date)
        solicitarCodigoRenovacion = True
        ws.Protect password:="theSheetProtectpassisthe011rd", userinterfaceonly:=True
        MsgBox "Gracias. Se han añadido 180 días a la vigencia del libro", vbInformation
        
    Else
        solicitarCodigoRenovacion = False
    End If

End Function
'genera el codigo mensual para validar la licencia
Private Function GenerarCodigoMensual(dia As Integer, mes As Integer, Ano As Integer) As String

    Dim semilla As String
    semilla = "theMonthCodegenisthe100th" & CStr(dia * 149) & CStr(mes * 149) & CStr(Ano)
    
    GenerarCodigoMensual = HashString(semilla)

End Function

'genera el codigo anual para validar la licencia
Private Function GenerarCodigoAnual(dia As Integer, mes As Integer, Ano As Integer) As String

    Dim semilla As String
    semilla = "theYearCodegenisthe101th" & CStr(dia * 149) & CStr(mes * 149) & CStr(Ano)
    
    GenerarCodigoAnual = HashString(semilla)

End Function
'funcion para encryptar
Private Function HashString(texto As String) As String

    Dim i As Integer
    Dim hash As Long
    hash = 1

    For i = 1 To Len(texto)
        hash = (hash * 31 + Asc(Mid(texto, i, 1))) Mod 1000000
    Next i
    
    HashString = CStr(hash)

End Function
'verifica que el cliente no le haya cambiado la fecha al dispositivo
Sub verificarIntegridadFecha()

    Dim ultimaFecha As Date
    Dim FechaActual As Date
    
    ThisWorkbook.Sheets("Hidden").Unprotect password:="theSheetProtectpassisthe011rd"
    
    FechaActual = Date
    ultimaFecha = ThisWorkbook.Sheets("Hidden").Range("A2")
    
    If ultimaFecha = 0 Then
        ThisWorkbook.Sheets("Hidden").Range("A2") = FechaActual
        ThisWorkbook.Sheets("Hidden").Protect password:="theSheetProtectpassisthe011rd", userinterfaceonly:=True
        Exit Sub
    End If
    
    If FechaActual < ultimaFecha Then
        MsgBox "Se detectó un retroceso en la fecha del sistema. El libro se cerrará.", vbCritical
        ThisWorkbook.Close SaveChanges:=False
    End If
    If FechaActual > ultimaFecha Then
        If Not ThisWorkbook.Name Like "Respaldo_*" Then
            crear_respaldo_diario
        End If
        ThisWorkbook.Sheets("Hidden").Range("A2") = FechaActual
    End If
    
    ThisWorkbook.Sheets("Hidden").Protect password:="theSheetProtectpassisthe011rd", userinterfaceonly:=True
End Sub

'macro para tomar un id del dispositivo y verificar que esta permitido
Private Function verificarDispositivo() As Boolean

    Dim dispositivosPermitidos As Collection
    Set dispositivosPermitidos = New Collection
    
    dispositivosPermitidos.Add "DESKTOP-FFH10GP|Jose|279994388"
    dispositivosPermitidos.Add "DESKTOP-27I02D0|Aris|351847289"
    dispositivosPermitidos.Add "DESKTOP-6Q2K16M|Noel|1088704218"
    
    Dim idActual As String
    idActual = ObtenerIDDispositivo()
    
    Dim i As Integer
    For i = 1 To dispositivosPermitidos.Count
        If dispositivosPermitidos(i) = idActual Then
            verificarDispositivo = True
            Exit Function
        End If
    Next i
    
    verificarDispositivo = False
End Function
