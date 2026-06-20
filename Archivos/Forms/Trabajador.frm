VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Trabajador 
   Caption         =   "Nuevo Trabajador"
   ClientHeight    =   4770
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   9160.001
   OleObjectBlob   =   "Trabajador.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Trabajador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Ingresar_Click()

    If Trim(Me.Txt_Nombre.Value) = "" Then
        MsgBox "Por favor ingrese un nombre", vbExclamation, "Campo requerido"
        Me.Txt_Nombre.SetFocus
        Exit Sub
    End If
    
    If Trim(Me.Txt_Cargo.Value) = "" Then
        MsgBox "Por favor ingrese un cargo", vbExclamation, "Campo requerido"
        Me.Txt_Cargo.SetFocus
        Exit Sub
    End If
    
    If Trim(Me.Txt_sal.Value) = "" Then
        MsgBox "Por favor ingrese un salario", vbExclamation, "Campo requerido"
        Me.Txt_sal.SetFocus
        Exit Sub
    End If
    
    If Trim(Me.Txt_sal_min.Value) = "" Then
        MsgBox "Por favor ingrese un salario mínimo", vbExclamation, "Campo requerido"
        Me.Txt_sal_min.SetFocus
        Exit Sub
    End If
    
    If Not IsNumeric(Me.Txt_sal.Value) Or Me.Txt_sal.Value < 0 Or Me.Txt_sal.Value > 100 Then
        MsgBox "Por favor ingrese un valor numérico entre 0 y 100 en el Salario %", vbExclamation, "Porcentaje inválido"
        Me.Txt_sal.SetFocus
        Exit Sub
    End If
    
    If Not IsNumeric(Me.Txt_sal_min.Value) Or Me.Txt_sal_min.Value < 0 Then
        MsgBox "Por favor ingrese un valor numérico positivo en el salario mínimo", vbExclamation, "Valor inválido"
        Me.Txt_sal_min.SetFocus
        Exit Sub
    End If
    
    Agregar
    MsgBox "Trabajador agregado correctamente", vbInformation, "Listo!"
    Unload Trabajador

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
    Ingresar_Click
End Sub
'Macro para agregar el gasto ingresado a la tabla
Private Sub Agregar()

Dim ws As Worksheet
Dim tbl As ListObject
Dim newrow As ListRow

Set ws = ThisWorkbook.Sheets("Gestión de Trabajadores")
Set tbl = ws.ListObjects("Trabajadores")

ws.Unprotect password:="theSheetProtectpassisthe011rd"

Set newrow = tbl.ListRows.Add

With newrow
    .Range(1) = Me.Txt_Nombre.Value
    .Range(2) = Me.Txt_Cargo.Value
    .Range(3) = CDbl(Me.Txt_sal.Value / 100)
    .Range(4) = CDbl(Me.Txt_sal_min.Value)
End With

tbl.Sort.Apply

ws.Protect password:="theSheetProtectpassisthe011rd", userinterfaceonly:=True

End Sub

Private Sub Txt_sal_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If (KeyAscii >= 46 And KeyAscii <= 57) Then
    KeyAscii = KeyAscii
  Else
    KeyAscii = 0
  End If
End Sub
Private Sub Txt_sal_min_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If (KeyAscii >= 46 And KeyAscii <= 57) Then
    KeyAscii = KeyAscii
  Else
    KeyAscii = 0
  End If
End Sub
