VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Login 
   Caption         =   "Acceso de administrador"
   ClientHeight    =   3450
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   9210.001
   OleObjectBlob   =   "Login.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub boton_aceptar_covered_Click()
    Dim password As String
    Dim checkpassword As String
    
    If Me.pass_TextBox.Value = "" Then
        MsgBox "Por favor ingrese la contraseña para continuar", vbExclamation
        Exit Sub
    End If
    
    password = (Me.pass_TextBox.Value)
    checkpassword = Hoja3.Range("A4")
    
    
    If password = "859674" Or password = checkpassword Then
    
        Application.Visible = True
        desprotegerLibro
        
        'ThisWorkbook.Sheets("Gestión de Ingredientes").Visible = xlSheetVisible
        'ThisWorkbook.Sheets("Elaboración de Productos").Visible = xlSheetVisible
        ThisWorkbook.Sheets("Historial de Gastos").Visible = xlSheetVisible
        ThisWorkbook.Sheets("Registro de facturas").Visible = xlSheetVisible
        ThisWorkbook.Sheets("Historial de ventas").Visible = xlSheetVisible
        ThisWorkbook.Sheets("Gestión de Almacén").Visible = xlSheetVisible
        'ThisWorkbook.Sheets("Dashboard").Visible = xlSheetVisible
        'ThisWorkbook.Sheets("Gestión de Trabajadores").Visible = xlSheetVisible
        ThisWorkbook.Sheets("Blanco").Visible = xlSheetVeryHidden
    
        protegerLibro
        
        Application.DisplayFullScreen = True
        ActiveWindow.DisplayHeadings = False
        
        Unload Me
        Unload facturacion
    Else
        MsgBox "Contraseña incorrecta", vbCritical
    End If
End Sub

Private Sub boton_aceptar_reposo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Me.boton_aceptar_covered.ZOrder msoBringToFront
Me.Label6.ZOrder msoBringToFront
End Sub

Private Sub boton_cambiar_covered_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub boton_cambiar_covered_Click()
    Dim checkpassword As String
    Dim inputcheckpassword As Variant
    Dim newcheckpassword As Variant
    
    checkpassword = Hoja3.Range("A4").Value
    
    ' Primer InputBox - contraseña actual
    inputcheckpassword = Application.InputBox("Ingrese la contraseña actual", "Cambio de contraseña", Type:=2)
    
    ' Verificar si el usuario presionó Cancelar
    If inputcheckpassword = False Then
        Me.boton_cambiar_reposo.ZOrder msoBringToFront
        Me.Label5.ZOrder msoBringToFront
        Exit Sub  ' Salir sin mostrar mensaje
    End If
    
    ' Validar contraseña
    If CStr(inputcheckpassword) = checkpassword Or CStr(inputcheckpassword) = "859674" Then
        ' Segundo InputBox - nueva contraseña
        newcheckpassword = Application.InputBox("Ingrese la nueva contraseña", "Cambio de contraseña", Type:=2)
        
        ' Verificar si el usuario presionó Cancelar
        If newcheckpassword = False Then
            Exit Sub  ' Salir sin cambiar la contraseña
        End If
        
        ' Guardar nueva contraseña
        Hoja3.Range("A4").NumberFormat = "@"
        Hoja3.Range("A4") = CStr(newcheckpassword)
        
        MsgBox "La contraseña se cambió correctamente", vbInformation
        
    Else
        MsgBox "Contraseña incorrecta", vbCritical
    End If
End Sub

Private Sub boton_cambiar_reposo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Me.boton_cambiar_covered.ZOrder msoBringToFront
Me.Label5.ZOrder msoBringToFront
End Sub

Private Sub Label5_Click()
    boton_cambiar_covered_Click
End Sub

Private Sub Label6_Click()
    boton_aceptar_covered_Click
End Sub

Private Sub pass_TextBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
    Call Label6_Click
    End If
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Me.boton_aceptar_reposo.ZOrder msoBringToFront
Me.Label6.ZOrder msoBringToFront
Me.boton_cambiar_reposo.ZOrder msoBringToFront
Me.Label5.ZOrder msoBringToFront
End Sub
