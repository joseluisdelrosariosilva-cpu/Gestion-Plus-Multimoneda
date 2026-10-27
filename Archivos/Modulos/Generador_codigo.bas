Attribute VB_Name = "Generador_codigo"
Public Codigo_Producto As String

' --- CONTADOR PERSISTENTE DE CODIGOS ---------------------------------------
' El ultimo numero usado se guarda en la hoja oculta:
'   Hidden!A1 = fecha de expiracion | A2 = ultima fecha usada | A4 = pass admin
'   Hidden!B1 = ultimo numero de codigo de producto  <-- ESTE MODULO
' El contador nunca retrocede y solo se mueve al GUARDAR el producto.
' ---------------------------------------------------------------------------

Private Const CELDA_ULTIMO_CODIGO As String = "B1"
Private Const HOJA_OCULTA As String = "Hidden"
Private Const PASS_HOJA_OCULTA As String = "theSheetProtectpassisthe011rd"


' --- 1. GENERAR (solo sugiere, NO persiste) --------------------------------
Sub Genera_Codigo()
    Codigo_Producto = Formatear_Codigo(Siguiente_Numero_Disponible())
End Sub


' --- 2. CONFIRMAR (llamar al guardar, despues de escribir la fila) ---------
Public Sub Confirmar_Codigo_Producto(codigo As String)
    Ajustar_Contador Extraer_Numero(codigo)
End Sub


' --- 3. SINCRONIZAR (opcional, despues de importar productos) --------------
Public Sub Sincronizar_Contador()
    Ajustar_Contador 0
End Sub


' --- 4. ¿Sigue libre ese codigo? (opcional, antes de guardar) --------------
Public Function Codigo_Disponible(codigo As String) As Boolean
    Codigo_Disponible = Not Existe_Codigo(Tabla_Inventario(), codigo)
End Function

' --- 5. RESETEAR (SOLO para preparar un libro nuevo) -----------------------
' Deshace a proposito la garantia de "el codigo nunca se recicla": deja el
' contador en 0, asi el proximo producto vuelve a Pr_00001.
' Llamalo SOLO despues de vaciar la tabla INVENTARIO.
Public Sub Resetear_Contador_Codigos()
    If Guardar_Ultimo_Numero(0) Then Exit Sub

    MsgBox "No se pudo reiniciar el contador de codigos en la hoja oculta """ & _
           HOJA_OCULTA & """ (celda " & CELDA_ULTIMO_CODIGO & ")." & vbCrLf & vbCrLf & _
           "Revise la proteccion de la hoja.", vbExclamation, "Generador de codigos"
End Sub

' ===========================================================================
' LOGICA INTERNA
' ===========================================================================

' Sube el contador a MAX(contador, numeroMinimo, maximo real de la tabla).
Private Sub Ajustar_Contador(numeroMinimo As Long)
    Dim nuevo As Long
    Dim maxTabla As Long
    
    nuevo = Leer_Ultimo_Numero()
    If numeroMinimo > nuevo Then nuevo = numeroMinimo
    
    maxTabla = Obtener_Maximo_En_Tabla(Tabla_Inventario())
    If maxTabla > nuevo Then nuevo = maxTabla
    
    If nuevo <= 0 Then Exit Sub
    
    If Not Guardar_Ultimo_Numero(nuevo) Then
        MsgBox "No se pudo actualizar el contador de códigos en la hoja oculta """ & _
               HOJA_OCULTA & """ (celda " & CELDA_ULTIMO_CODIGO & ")." & vbCrLf & vbCrLf & _
               "El producto quedó guardado, pero el próximo código podría repetirse." & vbCrLf & _
               "Revise la protección de la hoja.", vbExclamation, "Generador de códigos"
    End If
End Sub

' Proximo numero libre: MAYOR entre contador y maximo real de la tabla, mas uno.
Private Function Siguiente_Numero_Disponible() As Long
    Dim guardado As Long
    Dim enTabla As Long
    Dim candidato As Long
    
    guardado = Leer_Ultimo_Numero()
    enTabla = Obtener_Maximo_En_Tabla(Tabla_Inventario())
    
    If enTabla > guardado Then guardado = enTabla
    
    candidato = guardado + 1
    If candidato < 1 Then candidato = 1
    
    Do While Existe_Codigo(Tabla_Inventario(), Formatear_Codigo(candidato))
        candidato = candidato + 1
    Loop
    
    Siguiente_Numero_Disponible = candidato
End Function

Private Function Formatear_Codigo(numero As Long) As String
    Formatear_Codigo = "Pr_" & Format(numero, "00000")
End Function

Private Function Tabla_Inventario() As ListObject
    Set Tabla_Inventario = ThisWorkbook.Sheets("Gestión de Almacén").ListObjects("INVENTARIO")
End Function

' Mayor numero "Pr_#####" presente en la tabla de inventario.
Private Function Obtener_Maximo_En_Tabla(tbl As ListObject) As Long
    Dim i As Long
    Dim numero As Long
    Dim maximo As Long
    Dim celda As Variant
    
    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function
    
    For i = 1 To tbl.ListRows.Count
        celda = tbl.ListRows(i).Range.Cells(1, 1).Value
        numero = Extraer_Numero(Valor_Texto(celda))
        If numero > maximo Then maximo = numero
    Next i
    
    Obtener_Maximo_En_Tabla = maximo
End Function

' Extrae el numero de un codigo "Pr_#####". Devuelve 0 si no cumple el formato,
' asi un codigo escrito a mano no altera el contador.
Private Function Extraer_Numero(codigo As String) As Long
    Dim digitos As String
    Dim i As Long
    Dim caracter As String
    
    codigo = Trim(codigo)
    If Len(codigo) < 4 Then Exit Function
    If UCase(Left(codigo, 3)) <> "PR_" Then Exit Function
    
    digitos = Mid(codigo, 4)
    For i = 1 To Len(digitos)
        caracter = Mid(digitos, i, 1)
        If caracter < "0" Or caracter > "9" Then Exit Function
    Next i
    
    If Len(digitos) = 0 Then Exit Function
    If Len(digitos) > 9 Then Exit Function
    
    Extraer_Numero = CLng(digitos)
End Function

' ¿Ya existe ese codigo en el inventario?
Private Function Existe_Codigo(tbl As ListObject, codigo As String) As Boolean
    Dim rango As Range
    Dim encontrado As Range
    
    If tbl Is Nothing Then Exit Function
    
    On Error Resume Next
    Set rango = tbl.ListColumns(1).DataBodyRange
    On Error GoTo 0
    
    If rango Is Nothing Then Exit Function
    
    Set encontrado = rango.Find(What:=codigo, LookIn:=xlValues, LookAt:=xlWhole)
    Existe_Codigo = Not (encontrado Is Nothing)
End Function

' Convierte a texto sin fallar con celdas con error o Null.
Private Function Valor_Texto(v As Variant) As String
    If IsError(v) Then Exit Function
    If IsNull(v) Then Exit Function
    Valor_Texto = CStr(v)
End Function

' Lee el contador. Devuelve 0 si esta vacio o inaccesible.
Private Function Leer_Ultimo_Numero() As Long
    Dim valor As Variant
    
    On Error GoTo Fallo
    valor = ThisWorkbook.Sheets(HOJA_OCULTA).Range(CELDA_ULTIMO_CODIGO).Value
    If IsNumeric(valor) Then Leer_Ultimo_Numero = CLng(valor)
    Exit Function
Fallo:
    Leer_Ultimo_Numero = 0
End Function

' Guarda el contador. True si se pudo escribir.
Private Function Guardar_Ultimo_Numero(numero As Long) As Boolean
    Dim ws As Worksheet
    
    On Error GoTo Fallo
    
    Set ws = ThisWorkbook.Sheets(HOJA_OCULTA)
    ws.Unprotect password:=PASS_HOJA_OCULTA
    ws.Range(CELDA_ULTIMO_CODIGO).NumberFormat = "0"
    ws.Range(CELDA_ULTIMO_CODIGO).Value = numero
    ws.Protect password:=PASS_HOJA_OCULTA, UserInterfaceOnly:=True
    
    Guardar_Ultimo_Numero = True
    Exit Function
Fallo:
    On Error Resume Next
    If Not ws Is Nothing Then ws.Protect password:=PASS_HOJA_OCULTA, UserInterfaceOnly:=True
    Guardar_Ultimo_Numero = False
End Function

