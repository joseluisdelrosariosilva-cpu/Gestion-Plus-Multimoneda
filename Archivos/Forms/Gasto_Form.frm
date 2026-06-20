VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Gasto_Form 
   Caption         =   "Nuevo Gasto Variable"
   ClientHeight    =   4670
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   10730
   OleObjectBlob   =   "Gasto_Form.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Gasto_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Ingresar_Click()
    Btn_Active_Click
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.Btn_Reposo.ZOrder msoBringToFront
    Me.Ingresar.ZOrder msoBringToFront
End Sub

Private Sub Btn_Reposo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.Btn_Active.ZOrder msoBringToFront
    Me.Ingresar.ZOrder msoBringToFront
End Sub

'Macro para validar los campos del formulario
Private Sub Btn_Active_Click()

    If Trim(Me.Txt_Fecha.Value) = "" Then
        MsgBox "Por favor ingrese una fecha", vbExclamation, "Campo requerido"
        Me.Txt_Fecha.SetFocus
        Exit Sub
    End If
    
    If Trim(Me.Txt_Monto.Value) = "" Then
        MsgBox "Por favor ingrese un monto", vbExclamation, "Campo requerido"
        Me.Txt_Monto.SetFocus
        Exit Sub
    End If
    
    If Trim(Me.Txt_Descripcion.Value) = "" Then
        MsgBox "Por favor ingrese una descripción", vbExclamation, "Campo requerido"
        Me.Txt_Descripcion.SetFocus
        Exit Sub
    End If
    
    If Trim(Me.Moneda_gasto.Value) = "" Then
        MsgBox "Por favor ingrese una moneda", vbExclamation, "Campo requerido"
        Me.Moneda_gasto.SetFocus
        Exit Sub
    End If
    
    If Not IsDate(Me.Txt_Fecha.Value) Then
        MsgBox "Por favor ingrese una fecha válida", vbExclamation, "Fecha inválida"
        Me.Txt_Fecha.SetFocus
        Exit Sub
    End If
    
    If Not IsNumeric(Me.Txt_Monto.Value) Or Me.Txt_Monto.Value < 0 Then
        MsgBox "Por favor ingrese un valor numérico positivo en el monto", vbExclamation, "Monto inválido"
        Me.Txt_Monto.SetFocus
        Exit Sub
    End If
    
    Agregar
    'Agregar_Fecha
    MsgBox "Gasto agregado correctamente", vbInformation, "Listo!"
    Unload Gasto_Form

End Sub
'Macro para agregar el gasto ingresado a la tabla
Private Sub Agregar()

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim newrow As ListRow
    
    Set ws = ThisWorkbook.Sheets("Historial de Gastos")
    Set tbl = ws.ListObjects("Gastos")
    
    ws.Unprotect password:="theSheetProtectpassisthe011rd"
    
    Set newrow = tbl.ListRows.Add
    
    With newrow
        .Range(1) = CDate(Me.Txt_Fecha.Value)
        .Range(2) = "Variable"
        .Range(3) = Me.Txt_Descripcion.Value
        .Range(4) = CDbl(Me.Txt_Monto.Value)
        .Range(5) = Me.Moneda_gasto.Value
    End With
    
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
        If CDate(fila.Range(1)) = CDate(Me.Txt_Fecha.Value) Then
            Exit Sub
        End If
    Next fila
    
    ws.Unprotect password:="theSheetProtectpassisthe011rd"
    
    Set newrow = tbl.ListRows.Add
    
    newrow.Range(1) = CDate(Me.Txt_Fecha.Value)
    
    tbl.Sort.Apply
    
    ws.Protect password:="theSheetProtectpassisthe011rd", userinterfaceonly:=True

End Sub
Private Sub Txt_Monto_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If (KeyAscii >= 46 And KeyAscii <= 57) Then
    KeyAscii = KeyAscii
  Else
    KeyAscii = 0
  End If
End Sub

Private Sub UserForm_Initialize()

    Txt_Fecha.Value = Format(Date, "d-mmm-yyyy")
    ' Cargar opciones de moneda en los ComboBox
    With Me.Moneda_gasto
        .Clear
        .AddItem "MN"
        .AddItem "USD"
    End With
End Sub
