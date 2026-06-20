VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Trabajador_Mod 
   Caption         =   "Modificar Trabajador"
   ClientHeight    =   4770
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   9160.001
   OleObjectBlob   =   "Trabajador_Mod.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Trabajador_Mod"
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
    
    Modificar
    MsgBox "Trabajador modificado correctamente", vbInformation, "Listo!"
    Unload Trabajador_Mod

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
'Macro para modificar al trabajador de la tabla
Private Sub Modificar()

Dim ws As Worksheet
Dim tbl As ListObject
Dim sel_row As ListRow

Set ws = ThisWorkbook.Sheets("Gestión de Trabajadores")
Set tbl = ws.ListObjects("Trabajadores")

ws.Unprotect password:="theSheetProtectpassisthe011rd"

Set sel_row = tbl.ListRows(fila_sel)

With sel_row
    .Range(1) = Me.Txt_Nombre.Value
    .Range(2) = Me.Txt_Cargo.Value
    .Range(3) = CDbl(Me.Txt_sal.Value / 100)
    .Range(4) = CDbl(Me.Txt_sal_min.Value)
End With

tbl.Sort.Apply

ws.Protect password:="theSheetProtectpassisthe011rd", userinterfaceonly:=True

End Sub

'inicializa el formulario con los datos del trabajador que selecciono

Private Sub UserForm_Initialize()

Dim tbl As ListObject
Dim sel_row As ListRow

Set tbl = ThisWorkbook.Sheets("Gestión de Trabajadores").ListObjects("Trabajadores")
Set sel_row = tbl.ListRows(fila_sel)

With sel_row
    Me.Txt_Nombre = .Range(1)
    Me.Txt_Cargo = .Range(2)
    Me.Txt_sal = .Range(3) * 100
    Me.Txt_sal_min = .Range(4)
End With

End Sub
