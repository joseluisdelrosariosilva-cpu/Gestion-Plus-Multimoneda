Attribute VB_Name = "busqueda_factura"
Sub buscar_factura_avanzada()
    Dim codigoBuscar As String
    Dim celdaEncontrada As Range
    Dim HojaBD As Worksheet
    
    Set HojaBD = ThisWorkbook.Sheets("Registro de facturas")
    
    codigoBuscar = Trim(InputBox("Ingrese el código de factura:", "Buscar Factura"))
    If codigoBuscar = "" Then Exit Sub
    
    Application.ScreenUpdating = False
    
    Set celdaEncontrada = HojaBD.Columns("B").Find(codigoBuscar, LookIn:=xlValues, LookAt:=xlWhole)
    
    Application.ScreenUpdating = True
    
    If Not celdaEncontrada Is Nothing Then
        celdaEncontrada.Select
        Application.GoTo celdaEncontrada, True
    Else
        MsgBox "Factura no encontrada.", vbExclamation
    End If
End Sub

