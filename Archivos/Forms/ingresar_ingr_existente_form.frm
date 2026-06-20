VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ingresar_ingr_existente_form 
   Caption         =   "Ingresar Ingrediente Existente"
   ClientHeight    =   4600
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   10010
   OleObjectBlob   =   "ingresar_ingr_existente_form.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "ingresar_ingr_existente_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub boton_ingresar_covered_Click()
    Call validar_vacios_ingredientes_existentes
    If contador3 > 0 Then
        MsgBox "Rellene los campos vacíos", vbExclamation
        Me.boton_ingresar_reposo.ZOrder msoBringToFront
        Me.Label27.ZOrder msoBringToFront
    Else
        Dim celdaEncontrada As Range
        Dim fila As Integer
        Dim HojaBD As Worksheet
        Dim ingrediente As String
        Dim Unidad As String
        Dim cantidadSinConvertir As Double
        Dim cantidadConvertidaUBase As Double
        Dim cantidadActual As Double
        Dim fondoActual As Double
        Dim costo As Double
        
        Set HojaBD = ThisWorkbook.Sheets("Gestión de Ingredientes")
        
        ingrediente = Me.Nombre_ingr_ComboBox.Value
        cantidadSinConvertir = Me.Cantidad_ingresar_producto_existente_TextBox.Value
        Unidad = Me.Unidad_ComboBox.Value
        costo = Me.Costo_TextBox.Value
        
        
        Application.ScreenUpdating = False
        Set celdaEncontrada = HojaBD.Columns("B").Find(ingrediente, LookIn:=xlValues, LookAt:=xlWhole)
        Application.ScreenUpdating = True
        
        If Not celdaEncontrada Is Nothing Then
            fila = celdaEncontrada.Row
        End If
    
        cantidadConvertidaUBase = ConvertirUnidad(ingrediente, cantidadSinConvertir, Unidad)
        cantidadActual = HojaBD.Range("D" & fila).Value
        fondoActual = HojaBD.Range("G" & fila).Value
        
        HojaBD.Range("D" & fila).Value = cantidadConvertidaUBase + cantidadActual
        HojaBD.Range("G" & fila).Value = fondoActual + costo
        
        
        Call Calcular_UBase
        Unload ingresar_ingr_existente_form
        MsgBox "Cantidad agregada correctamente", vbInformation
    End If
End Sub

Private Sub boton_ingresar_reposo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Me.boton_ingresar_covered.ZOrder msoBringToFront
Me.Label27.ZOrder msoBringToFront
End Sub

Private Sub Cantidad_ingresar_producto_existente_TextBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If (KeyAscii >= 46 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
      Else
        KeyAscii = 0
      End If
End Sub

Private Sub Costo_TextBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If (KeyAscii >= 46 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
      Else
        KeyAscii = 0
      End If
End Sub

Private Sub Label27_Click()
    boton_ingresar_covered_Click
End Sub

Private Sub Nombre_ingr_ComboBox_Change()
  
    Dim valorSeleccionado As String
    Dim Unidad_base_valor As String
    Dim celdaEncontrada As Range
    Dim HojaBD As Worksheet
    Dim filaEncontrada As Integer
    
    valorSeleccionado = Me.Nombre_ingr_ComboBox.Value
    Set HojaBD = ThisWorkbook.Sheets("Gestión de Ingredientes")
    Set celdaEncontrada = HojaBD.Columns("B").Find(valorSeleccionado, LookIn:=xlValues, LookAt:=xlWhole)
    filaEncontrada = celdaEncontrada.Row
    Unidad_base_valor = HojaBD.Range("E" & filaEncontrada).Value
    
    ' Limpiar el ComboBox de unidades primero
    Me.Unidad_ComboBox.Clear
    
    ' Verificar qué se seleccionó
    Select Case LCase(Unidad_base_valor)
        Case "u"
            Me.Unidad_ComboBox.AddItem "u"
            Me.Unidad_ComboBox.ListIndex = 0 ' Seleccionar el primer item
        
        Case "kg", "g", "mg", "lb", "oz"
            Me.Unidad_ComboBox.AddItem "Kg"
            Me.Unidad_ComboBox.AddItem "g"
            Me.Unidad_ComboBox.AddItem "mg"
            Me.Unidad_ComboBox.AddItem "Lb"
            Me.Unidad_ComboBox.AddItem "oz"
            Me.Unidad_ComboBox.ListIndex = 0 ' Seleccionar el primero
        
        Case "l", "ml"
            Me.Unidad_ComboBox.AddItem "L"
            Me.Unidad_ComboBox.AddItem "mL"
            Me.Unidad_ComboBox.ListIndex = 0 ' Seleccionar el primero
    End Select
End Sub

Private Sub Nombre_ingr_ComboBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = 0
End Sub

Private Sub UserForm_Activate()
    Dim ultima_fila As Integer
    Dim HojaBD As Worksheet
    Dim ingrediente As String
    
    Set HojaBD = ThisWorkbook.Sheets("Gestión de Ingredientes")
    ultima_fila = HojaBD.Range("B" & Rows.Count).End(xlUp).Row
    
    For i = 5 To ultima_fila
        ingrediente = HojaBD.Range("B" & i)
        Me.Nombre_ingr_ComboBox.AddItem (ingrediente)
    Next i
    
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Me.boton_ingresar_reposo.ZOrder msoBringToFront
Me.Label27.ZOrder msoBringToFront
End Sub
