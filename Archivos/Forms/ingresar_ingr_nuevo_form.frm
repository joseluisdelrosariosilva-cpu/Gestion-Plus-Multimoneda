VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ingresar_ingr_nuevo_form 
   Caption         =   "Ingresar Nuevo Ingrediente"
   ClientHeight    =   5050
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   10760
   OleObjectBlob   =   "ingresar_ingr_nuevo_form.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "ingresar_ingr_nuevo_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Cantidad_ingresar_nuevo_producto_TextBox_Change()
    If Me.Cantidad_ingresar_nuevo_producto_TextBox.Value <> "" And Me.Costo_Total_Ingr_Nuevo_TextBox.Value <> "" Then

        Dim costototal As Currency
        Dim Cantidad As Single
        Dim precioUbase As Currency
    
    
        costototal = Me.Costo_Total_Ingr_Nuevo_TextBox.Value
        Cantidad = Me.Cantidad_ingresar_nuevo_producto_TextBox.Value
        precioUbase = costototal / Cantidad
    
        Me.Precio_UBase_TextBox.Value = Round(precioUbase, 2)
    Else
        Me.Precio_UBase_TextBox.Value = ""
    End If
End Sub

Private Sub Cantidad_ingresar_nuevo_producto_TextBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'Validacion para evitar que se escriban letras
If (KeyAscii >= 46 And KeyAscii <= 57) Then
    KeyAscii = KeyAscii
  Else
    KeyAscii = 0
  End If
End Sub

Private Sub Cantidad_referencia_TextBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If (KeyAscii >= 46 And KeyAscii <= 57) Then
    KeyAscii = KeyAscii
  Else
    KeyAscii = 0
  End If
End Sub

Private Sub Costo_Total_Ingr_Nuevo_TextBox_Change()
    
    If Me.Cantidad_ingresar_nuevo_producto_TextBox.Value <> "" And Me.Costo_Total_Ingr_Nuevo_TextBox.Value <> "" Then

        Dim costototal As Currency
        Dim Cantidad As Single
        Dim precioUbase As Currency
    
    
        costototal = Me.Costo_Total_Ingr_Nuevo_TextBox.Value
        Cantidad = Me.Cantidad_ingresar_nuevo_producto_TextBox.Value
        precioUbase = costototal / Cantidad
    
        Me.Precio_UBase_TextBox.Value = Round(precioUbase, 2)
        
    Else
        Me.Precio_UBase_TextBox.Value = ""
    End If
    
End Sub

Private Sub Costo_Total_Ingr_Nuevo_TextBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If (KeyAscii >= 46 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub Btn_Active_Click()

    Dim fila_activa As Integer
    Dim HojaBD As Worksheet
    Dim ultima_fila As Integer
    
   If Me.modificar_ingresar_nuevo_ingr_textbox = "I" Then
   
    Call validar_vacios_ingredientes
    Call comprobar_repeticion_2
    
      If contador2 > 0 Then
          MsgBox "Rellene los campos vacíos", vbExclamation
      ElseIf repeticion2 = True Then
          MsgBox "No puede añadir un producto añadido anteriormente. Por favor modifíquelo", vbExclamation
      Else
          
          Set HojaBD = ThisWorkbook.Sheets("Gestión de Ingredientes")
          
          ultima_fila = HojaBD.Range("B" & Rows.Count).End(xlUp).Row
          
          If ultima_fila = 5 And Range("B" & ultima_fila).Value = "" Then
              HojaBD.Range("B5").Value = ingresar_ingr_nuevo_form.Nombre_ingrediente_Nuevo_TextBox.Value
              HojaBD.Range("C5").Value = ingresar_ingr_nuevo_form.Cantidad_referencia_TextBox.Value
              HojaBD.Range("D5").Value = ingresar_ingr_nuevo_form.Cantidad_ingresar_nuevo_producto_TextBox.Value
              HojaBD.Range("E5").Value = ingresar_ingr_nuevo_form.UBase_nuevo_ComboBox.Value
              HojaBD.Range("G5").Value = ingresar_ingr_nuevo_form.Costo_Total_Ingr_Nuevo_TextBox.Value
          Else
              HojaBD.Range("B" & (ultima_fila + 1)).Value = ingresar_ingr_nuevo_form.Nombre_ingrediente_Nuevo_TextBox.Value
              HojaBD.Range("C" & (ultima_fila + 1)).Value = ingresar_ingr_nuevo_form.Cantidad_referencia_TextBox.Value
              HojaBD.Range("E" & (ultima_fila + 1)).Value = ingresar_ingr_nuevo_form.UBase_nuevo_ComboBox.Value
              HojaBD.Range("D" & (ultima_fila + 1)).Value = ingresar_ingr_nuevo_form.Cantidad_ingresar_nuevo_producto_TextBox.Value
              HojaBD.Range("G" & (ultima_fila + 1)).Value = ingresar_ingr_nuevo_form.Costo_Total_Ingr_Nuevo_TextBox.Value
          End If
              Call Calcular_UBase
              HojaBD.Range("B5").Sort key1:=Range("B5"), Header:=xlYes
              Unload ingresar_ingr_nuevo_form
              MsgBox "Ingrediente añadido correctamente", vbInformation
      End If
      
    ElseIf Me.modificar_ingresar_nuevo_ingr_textbox.Value = "M" Then
        
        Set HojaBD = ThisWorkbook.Sheets("Gestión de Ingredientes")
        fila_activa = Application.ActiveCell.Row
        
        Call validar_vacios_ingredientes
        Call comprobar_repeticion_3
            
        If contador2 > 0 Then
            MsgBox "Rellene los campos vacíos", vbExclamation
        ElseIf repeticion3 = True Then
            MsgBox "No puede añadir un producto añadido anteriormente", vbExclamation
        Else
            Range("B" & fila_activa).Value = ingresar_ingr_nuevo_form.Nombre_ingrediente_Nuevo_TextBox.Value
            Range("C" & fila_activa).Value = ingresar_ingr_nuevo_form.Cantidad_referencia_TextBox.Value
            Range("D" & fila_activa).Value = ingresar_ingr_nuevo_form.Cantidad_ingresar_nuevo_producto_TextBox.Value
            Range("E" & fila_activa).Value = ingresar_ingr_nuevo_form.UBase_nuevo_ComboBox.Value
            Range("G" & fila_activa).Value = ingresar_ingr_nuevo_form.Costo_Total_Ingr_Nuevo_TextBox.Value
            
            
            Call Calcular_UBase
            HojaBD.Range("B5").Sort key1:=Range("B5"), Header:=xlYes
            Unload ingresar_ingr_nuevo_form
            MsgBox "Ingrediente modificado correctamente", vbInformation
        End If
    End If
End Sub

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


Private Sub UserForm_Activate()
 'Relleno del combobox
    Me.UBase_nuevo_ComboBox.AddItem "u"
    Me.UBase_nuevo_ComboBox.AddItem "Kg"
    Me.UBase_nuevo_ComboBox.AddItem "g"
    Me.UBase_nuevo_ComboBox.AddItem "mg"
    Me.UBase_nuevo_ComboBox.AddItem "Lb"
    Me.UBase_nuevo_ComboBox.AddItem "oz"
    Me.UBase_nuevo_ComboBox.AddItem "L"
    Me.UBase_nuevo_ComboBox.AddItem "mL"
    
    
    If Me.modificar_ingresar_nuevo_ingr_textbox = "I" Then
    
        Me.Ingresar.Caption = "Ingresar"
    
    ElseIf Me.modificar_ingresar_nuevo_ingr_textbox.Value = "M" Then

        Me.Ingresar.Caption = "Modificar"
 
        With ingresar_ingr_nuevo_form
                
            .Nombre_ingrediente_Nuevo_TextBox.Value = Range("B" & fila_a_modificar).Value
            .Cantidad_referencia_TextBox.Value = Range("C" & fila_a_modificar).Value
            .Cantidad_ingresar_nuevo_producto_TextBox.Value = Range("D" & fila_a_modificar).Value
            .UBase_nuevo_ComboBox.Value = Range("E" & fila_a_modificar).Value
            .Costo_Total_Ingr_Nuevo_TextBox.Value = Range("G" & fila_a_modificar).Value
                
        End With
    End If


End Sub
