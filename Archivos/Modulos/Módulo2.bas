Attribute VB_Name = "Módulo2"
'macro para limpiar la base de datos y crear respaldo
Sub new_respaldo()

    Dim wbNuevo As Workbook
    Dim wsOrigenVentas As Worksheet, wsOrigenGastos As Worksheet
    Dim wsDestinoVentas As Worksheet, wsDestinoGastos As Worksheet
    Dim tblventas As ListObject, tblGastos As ListObject
    Dim rangoFechasVentas As String, rangoFechasGastos As String
    Dim nombreArchivo As String
    Dim fechaMin As Date, fechaMax As Date
    Dim rutaRespaldos As String
    Dim ActualToInicial As ListRow
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    On Error GoTo ErrorHandler
    
    Set wsOrigenVentas = ThisWorkbook.Sheets("Historial de ventas")
    Set tblventas = wsOrigenVentas.ListObjects("HISTORIALVENTAS")
    
    Set wsOrigenGastos = ThisWorkbook.Sheets("Historial de Gastos")
    Set tblGastos = wsOrigenGastos.ListObjects("Gastos")
    
    fechaMin = Application.WorksheetFunction.Min(ThisWorkbook.Sheets("Fechas").ListObjects("FechasUnicas").ListColumns("Fecha").DataBodyRange)
    fechaMax = Application.WorksheetFunction.Max(ThisWorkbook.Sheets("Fechas").ListObjects("FechasUnicas").ListColumns("Fecha").DataBodyRange)
    
    nombreArchivo = "Respaldo_" & Format(fechaMin, "d-mmm-yyyy") & "_" & Format(fechaMax, "d-mmm-yyyy") & ".xlsx"
    
    rutaRespaldos = Environ("Jose") & "D:\Work\Respaldos\"
    If Dir(rutaRespaldos, vbDirectory) = "" Then MkDir rutaRespaldos
    
    Set wbNuevo = Workbooks.Add
    
    tblventas.HeaderRowRange.Copy
    Set wsDestinoVentas = wbNuevo.Sheets(1)
    wsDestinoVentas.Name = "Historial de ventas"
    wsDestinoVentas.Range("A1").PasteSpecial Paste:=xlPasteAll
    
    If Not tblventas.DataBodyRange Is Nothing Then
        tblventas.DataBodyRange.Copy
        wsDestinoVentas.Range("A2").PasteSpecial Paste:=xlPasteAll
    End If
    
    tblGastos.HeaderRowRange.Copy
    Set wsDestinoGastos = wbNuevo.Sheets.Add(After:=wbNuevo.Sheets(wbNuevo.Sheets.Count))
    wsDestinoGastos.Name = "Historial de Gastos"
    wsDestinoGastos.Range("A1").PasteSpecial Paste:=xlPasteAll
    
    If Not tblGastos.DataBodyRange Is Nothing Then
        tblGastos.DataBodyRange.Copy
        wsDestinoGastos.Range("A2").PasteSpecial Paste:=xlPasteAll
    End If
    
    Application.CutCopyMode = False
    
    wbNuevo.SaveAs rutaRespaldos & nombreArchivo, FileFormat:=xlOpenXMLWorkbook
    wbNuevo.Close SaveChanges:=False
    
    For Each ActualToInicial In ThisWorkbook.Worksheets("Gestión de Almacén").ListObjects("INVENTARIO").ListRows
        ActualToInicial.Range(4) = ActualToInicial.Range(5)
    Next ActualToInicial
    
    LimpiarTabla tblventas
    LimpiarTabla tblGastos
    LimpiarTabla ThisWorkbook.Sheets("Fechas").ListObjects("FechasUnicas")
    'LimpiarTabla ThisWorkbook.Sheets("Fechas").ListObjects("Productos_Unicos")
    
    MsgBox "Respaldo completado exitosamente." & vbCrLf & _
                    "Archivo guardado en: " & rutaRespaldos & nombreArchivo & vbCrLf & _
                    "Base de datos original ha sido limpiada. ", vbInformation
                    
exitsub:

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    Resume exitsub
    
End Sub

Sub LimpiarTabla(tbl As ListObject)

    If Not tbl.DataBodyRange Is Nothing Then
        tbl.DataBodyRange.Delete
    End If
    
End Sub
'macro para limpiar la base de datos y crear respaldo v2
Sub crear_respaldo()

    Dim wbRespaldo As Workbook
    Dim wbOriginal As Workbook
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim nombreRespaldo As String
    Dim fechaMin As Date, fechaMax As Date
    Dim rutaRespaldos As String
    Dim i As Long
    Dim response As VbMsgBoxResult
    
    response = MsgBox("Está seguro de que desea crear una copia de respaldo y limpiar la base de datos actual?", _
                                vbYesNo + vbQuestion)

    If Not response = vbYes Then
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    On Error GoTo ErrorHandler
    
    Set wbOriginal = ThisWorkbook
    
    fechaMin = Application.WorksheetFunction.Min(wbOriginal.Sheets("Fechas").ListObjects("FechasUnicas").ListColumns("Fecha").DataBodyRange)
    fechaMax = Application.WorksheetFunction.Max(wbOriginal.Sheets("Fechas").ListObjects("FechasUnicas").ListColumns("Fecha").DataBodyRange)
    
    nombreRespaldo = "Respaldo_" & Format(fechaMin, "d-mmm-yyyy") & "_" & Format(fechaMax, "d-mmm-yyyy") & ".xlsm"
    
    rutaRespaldos = wbOriginal.Path & "\Respaldos\"
    If Dir(rutaRespaldos, vbDirectory) = "" Then MkDir rutaRespaldos
    
    wbOriginal.SaveCopyAs rutaRespaldos & nombreRespaldo
    
    Set wbRespaldo = Workbooks.Open(rutaRespaldos & nombreRespaldo)
    
    For Each ws In wbRespaldo.Worksheets
        ws.Unprotect password:="theSheetProtectpassisthe011rd"
        For i = ws.Shapes.Count To 1 Step -1
            If ws.Shapes(i).Name Like "Boton_*" Then
                ws.Shapes(i).Delete
            End If
        Next i
        ws.Protect password:="theSheetProtectpassisthe011rd", userinterfaceonly:=True
    Next ws
    
    wbRespaldo.Worksheets("Dashboard").Protect password:="theSheetProtectpassisthe011rd", _
        contents:=True, _
        AllowFiltering:=True, _
        AllowSorting:=True, _
        AllowUsingPivotTables:=True, _
        userinterfaceonly:=True
    
    wbRespaldo.Worksheets("Hidden").Unprotect password:="theSheetProtectpassisthe011rd"
    wbRespaldo.Worksheets("Hidden").Range("A1") = DateAdd("y", 10000, Date)
    wbRespaldo.Worksheets("Hidden").Protect password:="theSheetProtectpassisthe011rd", userinterfaceonly:=True
    
    wbRespaldo.Save
    wbRespaldo.Close
    
    
    
    LimpiarTabla wbOriginal.Sheets("Fechas").ListObjects("FechasUnicas")
    LimpiarTabla wbOriginal.Sheets("Historial de Gastos").ListObjects("Gastos")
    LimpiarTabla wbOriginal.Sheets("Historial de ventas").ListObjects("HISTORIALVENTAS")
    LimpiarTabla wbOriginal.Sheets("Registro de facturas").ListObjects("HISTORIALVENTAS4")
    
    MsgBox "Respaldo completado exitosamente." & vbCrLf & _
                    "Archivo guardado en: " & rutaRespaldos & nombreRespaldo & vbCrLf & _
                    "Base de datos original ha sido limpiada. ", vbInformation
                    
exitsub:

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    Resume exitsub
    
End Sub
'macro para crear respaldo diario
Public Sub crear_respaldo_diario()

    Dim wbOriginal As Workbook
    Dim nombreRespaldo As String
    Dim fechaRespaldo As Date
    Dim rutaRespaldos As String
    Dim respaldoAnterior As String
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    On Error GoTo ErrorHandler
    
    Set wbOriginal = ThisWorkbook
    
    fechaRespaldo = Date
    
    nombreRespaldo = "Respaldo_diario_" & Format(fechaRespaldo, "d-mmm-yyyy") & ".xlsm"
    
    rutaRespaldos = wbOriginal.Path & "\Respaldos\"
    If Dir(rutaRespaldos, vbDirectory) = "" Then MkDir rutaRespaldos
    
    respaldoAnterior = Dir(rutaRespaldos & "Respaldo_diario_*.xlsm")
    If respaldoAnterior <> "" Then
        Kill rutaRespaldos & respaldoAnterior
    End If
    wbOriginal.SaveCopyAs rutaRespaldos & nombreRespaldo
                    
exitsub:

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    Resume exitsub
    
End Sub
'macro para crear respaldo antes de cerrar
Public Sub crearRespaldoFinal()

    Dim wb_Original As Workbook
    Dim nombre_Respaldo As String
    Dim fecha_Respaldo As Date
    Dim ruta_Respaldos As String
    Dim respaldo_Anterior As String
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    On Error GoTo ErrorHandler
    
    Set wb_Original = ThisWorkbook
    
    fecha_Respaldo = Date
    
    nombre_Respaldo = "Ultimo_respaldo_" & Format(fecha_Respaldo, "d-mmm-yyyy") & "_" & Format(Time, "hh-mm AM/PM") & ".xlsm"
    
    ruta_Respaldos = wb_Original.Path & "\Respaldos\"
    If Dir(ruta_Respaldos, vbDirectory) = "" Then MkDir ruta_Respaldos
    
    respaldo_Anterior = Dir(ruta_Respaldos & "Ultimo_respaldo_*.xlsm")
    If respaldo_Anterior <> "" Then
        Kill ruta_Respaldos & respaldo_Anterior
    End If
    wb_Original.SaveCopyAs ruta_Respaldos & nombre_Respaldo
                    
exitsub:

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    Resume exitsub
    
End Sub

Sub limpiar()
    LimpiarTabla ThisWorkbook.Sheets("Fechas").ListObjects("FechasUnicas")
    LimpiarTabla ThisWorkbook.Sheets("Fechas").ListObjects("Productos_Unicos")
End Sub
