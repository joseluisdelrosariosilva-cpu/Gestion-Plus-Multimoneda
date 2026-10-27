Attribute VB_Name = "Módulo2"
Private Const PASS_HOJA As String = "theSheetProtectpassisthe011rd"

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
        ws.Protect password:="theSheetProtectpassisthe011rd", UserInterfaceOnly:=True
    Next ws
    
    wbRespaldo.Worksheets("Dashboard").Protect password:="theSheetProtectpassisthe011rd", _
        contents:=True, _
        AllowFiltering:=True, _
        AllowSorting:=True, _
        AllowUsingPivotTables:=True, _
        UserInterfaceOnly:=True
    
    wbRespaldo.Worksheets("Hidden").Unprotect password:="theSheetProtectpassisthe011rd"
    wbRespaldo.Worksheets("Hidden").Range("A1") = DateAdd("y", 10000, Date)
    wbRespaldo.Worksheets("Hidden").Protect password:="theSheetProtectpassisthe011rd", UserInterfaceOnly:=True
    
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

' ===========================================================================
' IMPORTAR DATOS DESDE OTRO LIBRO
'
' Recorre TODAS las tablas del libro actual y, para cada una, busca una tabla
' con el mismo nombre en el libro que elijas. Si la encuentra y los
' encabezados coinciden, REEMPLAZA el contenido local por el del origen.
'
' Reemplaza, no agrega. Las tablas locales que no existan en el origen no se
' tocan. Si los encabezados o la cantidad de columnas no coinciden, esa tabla
' se saltea y se informa al final.
' ===========================================================================

Public Sub ImportarDatosDesdeExcel()
    Dim srcFile    As Variant
    Dim srcWB      As Workbook
    Dim ws         As Worksheet
    Dim lo         As ListObject
    Dim srcLO      As ListObject
    Dim importadas As Long
    Dim omitidas   As Long
    Dim incompat   As Long
    Dim mensaje    As String

    Const TITULO As String = "Importar datos desde otro libro"

    ' ---- 1. Elegir el libro origen ----
    srcFile = Application.GetOpenFilename( _
        FileFilter:="Libros de Excel (*.xls;*.xlsx;*.xlsm;*.xlsb), *.xls; *.xlsx; *.xlsm; *.xlsb", _
        Title:="Selecciona el libro del cual importar los datos", _
        MultiSelect:=False)

    If srcFile = False Then Exit Sub
    If StrComp(ThisWorkbook.FullName, CStr(srcFile), vbTextCompare) = 0 Then
        MsgBox "Seleccionaste el mismo libro que esta abierto. Elegi otro.", _
               vbExclamation, TITULO
        Exit Sub
    End If

    If MsgBox("Se van a REEMPLAZAR los datos de todas las tablas locales que " & _
              "existan tambien en el libro elegido." & vbCrLf & vbCrLf & _
              "¿Continuar?", vbYesNo + vbQuestion, TITULO) <> vbYes Then
        Exit Sub
    End If

    ' ---- 2. Abrir el origen sin ejecutar sus macros ----
    On Error GoTo ErrorAbrir

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.AutomationSecurity = msoAutomationSecurityForceDisable

    Set srcWB = Workbooks.Open(Filename:=CStr(srcFile), ReadOnly:=True, UpdateLinks:=0)

    On Error GoTo ErrorTrabajo

    ' ---- 3. Desproteger el destino ----
    DesprotegerTodo

    ' ---- 4. Copiar tabla por tabla ----
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            Set srcLO = BuscarTabla(srcWB, lo.Name)

            If srcLO Is Nothing Then
                omitidas = omitidas + 1
            ElseIf CopiarTabla(lo, srcLO) Then
                importadas = importadas + 1
            Else
                incompat = incompat + 1
            End If
        Next lo
    Next ws

    ' Si trajiste el inventario de otro libro, el contador tiene que alcanzar
    ' al codigo mas alto importado o el proximo producto lo repetiria.
    Call Sincronizar_Contador

    mensaje = "Importacion terminada." & vbCrLf & vbCrLf & _
              "Tablas importadas: " & importadas & vbCrLf & _
              "Tablas no encontradas en el origen: " & omitidas

    If incompat > 0 Then
        mensaje = mensaje & vbCrLf & _
                  "Tablas con columnas distintas (no se tocaron): " & incompat
    End If

Salir:
    On Error Resume Next
    If Not srcWB Is Nothing Then
        srcWB.Close SaveChanges:=False
        Set srcWB = Nothing
    End If
    ProtegerTodo
    RestaurarAplicacion
    On Error GoTo 0

    If Len(mensaje) > 0 Then MsgBox mensaje, vbInformation, TITULO
    Exit Sub

ErrorAbrir:
    RestaurarAplicacion
    MsgBox "No se pudo abrir el archivo:" & vbCrLf & srcFile & vbCrLf & vbCrLf & _
           "Error: " & Err.Description, vbCritical, TITULO
    Exit Sub

ErrorTrabajo:
    mensaje = ""
    MsgBox "Error durante la importacion:" & vbCrLf & vbCrLf & _
           Err.Description & vbCrLf & vbCrLf & _
           "Revisa el libro antes de seguir usando.", vbCritical, TITULO
    Resume Salir
End Sub


' ===========================================================================
' LIMPIAR TODAS LAS TABLAS
'
' Deja TODAS las tablas del libro solo con su fila de encabezado.
' Uso previsto: preparar el libro para un cliente nuevo.
' No toca formulas, formatos, validaciones ni nombres definidos.
' ===========================================================================

Public Sub LimpiarTodasLasTablas()
    Dim ws        As Worksheet
    Dim lo        As ListObject
    Dim respuesta As VbMsgBoxResult
    Dim tablas    As Long
    Dim mensaje   As String

    Const TITULO As String = "Limpiar todas las tablas"

    respuesta = MsgBox( _
        "Esto borra TODAS las filas de TODAS las tablas del libro:" & vbCrLf & _
        "inventario, ingredientes, recetas, ventas, facturas, gastos," & vbCrLf & _
        "trabajadores, fechas y productos unicos." & vbCrLf & vbCrLf & _
        "Las tablas quedan solo con los encabezados." & vbCrLf & _
        "NO se puede deshacer. ¿Continuar?", _
        vbYesNo + vbExclamation, TITULO)

    If respuesta <> vbYes Then Exit Sub

    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    DesprotegerTodo

    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            LimpiarTabla lo
            tablas = tablas + 1
        Next lo
    Next ws

    ' Descomenta esta linea si ademas queres que el proximo producto sea Pr_00001.
    Call Resetear_Contador_Codigos

    mensaje = "Listo: " & tablas & " tablas quedaron vacias (solo encabezados)."

Salir:
    On Error Resume Next
    ProtegerTodo
    RestaurarAplicacion
    On Error GoTo 0

    If Len(mensaje) > 0 Then MsgBox mensaje, vbInformation, TITULO
    Exit Sub

ErrorHandler:
    mensaje = "Se produjo un error y la limpieza quedo a medias:" & vbCrLf & vbCrLf & _
              Err.Description & vbCrLf & vbCrLf & _
              "Revisa el libro antes de seguir usando."
    Resume Salir
End Sub


' ===========================================================================
' AUXILIARES
' ===========================================================================

' Busca una tabla por nombre en todo el libro. Comparacion sin distinguir
' mayusculas: VBA compara texto como binario y "INVENTARIO" <> "Inventario".
Private Function BuscarTabla(wb As Workbook, nombre As String) As ListObject
    Dim ws As Worksheet
    Dim lo As ListObject

    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, nombre, vbTextCompare) = 0 Then
                Set BuscarTabla = lo
                Exit Function
            End If
        Next lo
    Next ws
End Function


' Reemplaza el contenido de destLO con el de srcLO.
' Devuelve False si las dos tablas no tienen las mismas columnas en el mismo
' orden, para no escribir datos en columnas equivocadas.
Private Function CopiarTabla(destLO As ListObject, srcLO As ListObject) As Boolean
    Dim vDatos As Variant
    Dim i As Long
    Dim j As Long

    If srcLO.ListColumns.Count <> destLO.ListColumns.Count Then Exit Function

    For j = 1 To srcLO.ListColumns.Count
        If StrComp(Trim$(srcLO.ListColumns(j).Name), _
                   Trim$(destLO.ListColumns(j).Name), vbTextCompare) <> 0 Then
            Exit Function
        End If
    Next j

    LimpiarTabla destLO

    If srcLO.DataBodyRange Is Nothing Then
        CopiarTabla = True
        Exit Function
    End If

    vDatos = srcLO.DataBodyRange.Value

    ' Con una sola fila y una sola columna Excel devuelve un escalar, no una matriz.
    If Not IsArray(vDatos) Then
        destLO.ListRows.Add
        destLO.DataBodyRange.Value = vDatos
        CopiarTabla = True
        Exit Function
    End If

    For i = 1 To UBound(vDatos, 1)
        destLO.ListRows.Add
    Next i

    destLO.DataBodyRange.Value = vDatos
    CopiarTabla = True
End Function

Private Sub DesprotegerTodo()
    Dim ws As Worksheet

    On Error Resume Next
    desprotegerLibro
    For Each ws In ThisWorkbook.Worksheets
        ws.Unprotect password:=PASS_HOJA
    Next ws
    On Error GoTo 0
End Sub


Private Sub ProtegerTodo()
    Dim ws As Worksheet

    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        ProtegerHoja ws
    Next ws
    protegerLibro
    On Error GoTo 0
End Sub


' El Dashboard tiene permisos extra (filtrar, ordenar, tablas dinamicas).
' Si se lo protege con la llamada generica, el usuario los pierde.
Private Sub ProtegerHoja(ws As Worksheet)
    If ws.Name = "Dashboard" Then
        ws.Protect password:=PASS_HOJA, contents:=True, _
                   AllowFiltering:=True, _
                   AllowSorting:=True, _
                   AllowUsingPivotTables:=True, _
                   UserInterfaceOnly:=True
    Else
        ws.Protect password:=PASS_HOJA, UserInterfaceOnly:=True
    End If
End Sub


Private Sub RestaurarAplicacion()
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.AutomationSecurity = msoAutomationSecurityByUI
    Application.CutCopyMode = False
End Sub

