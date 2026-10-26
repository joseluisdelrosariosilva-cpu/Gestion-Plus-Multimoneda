VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Codigo para que al abrir el dashboard se visualice en fullscreen y para que se actualicen todas las tablas
Private Sub Worksheet_Activate()

    'Application.DisplayFullScreen = True
    'ActiveWindow.DisplayHeadings = False

End Sub
'codigo para que se actualicen las formas cada vez que se aplique un filtro
Private Sub Actualizar_Formas()

    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim valor_Ingresos As Variant
    Dim valor_Gastos As Variant
    Dim valor_Beneficio As Variant
    Dim valor_Margen As Double
    Dim valor_Fondo As Variant
    Dim forma_Ingresos As Shape
    Dim forma_Gastos As Shape
    Dim forma_Beneficio As Shape
    Dim forma_Margen As Shape
    Dim forma_Fondo As Shape
    
    Set ws = ThisWorkbook.Sheets("Dashboard")
    Set pt = ws.PivotTables("Ingresos_Din")
    
    Set forma_Ingresos = ws.Shapes("Shape_Ingresos")
    Set forma_Gastos = ws.Shapes("Shape_Gastos")
    Set forma_Beneficio = ws.Shapes("Shape_Beneficio")
    Set forma_Margen = ws.Shapes("Shape_Margen")
    Set forma_Fondo = ws.Shapes("Shape_Fondo")
    
    Set valor_Ingresos = pt.TableRange2.Cells(pt.TableRange2.Rows.Count, 2)
    Set valor_Gastos = pt.TableRange2.Cells(pt.TableRange2.Rows.Count, 3)
    Set valor_Beneficio = pt.TableRange2.Cells(pt.TableRange2.Rows.Count, 4)
    Set valor_Fondo = ws.PivotTables("Fondo_Total").TableRange2.Cells(ws.PivotTables("Fondo_Total").TableRange2.Rows.Count, 1)
    
    forma_Ingresos.TextFrame.Characters.Text = "Ingreso Total: " & vbCrLf & " CUP   " & Format(valor_Ingresos, "#,##0")
    'forma_Ingresos.TextFrame.AutoSize = msoAutoSizeShapeToFitText
    forma_Fondo.TextFrame.Characters.Text = "Fondo Total: " & vbCrLf & " CUP   " & Format(valor_Fondo, "#,##0")
    'forma_Fondo.TextFrame.AutoSize = msoAutoSizeShapeToFitText
    forma_Gastos.TextFrame.Characters.Text = "Gasto Total: " & vbCrLf & " CUP   " & Format(valor_Gastos, "#,##0")
    'forma_Gastos.TextFrame.AutoSize = msoAutoSizeShapeToFitText
    forma_Beneficio.TextFrame.Characters.Text = "Beneficio: " & vbCrLf & " CUP   " & Format(valor_Beneficio, "#,##0")
    'forma_Beneficio.TextFrame.AutoSize = msoAutoSizeShapeToFitText
    
    On Error GoTo division_cero
    valor_Margen = CDbl(valor_Beneficio) * 100 / CDbl(valor_Ingresos)
    
    forma_Margen.TextFrame.Characters.Text = "Margen:" & vbCrLf & CInt(valor_Margen) & " %"
    'forma_Margen.TextFrame.AutoSize = msoAutoSizeShapeToFitText
Exit Sub
    
division_cero:
    valor_Margen = 0
    forma_Margen.TextFrame.Characters.Text = "Margen:  " & vbCrLf & CInt(valor_Margen) & " %"
    'forma_Margen.TextFrame.AutoSize = msoAutoSizeShapeToFitText
    
End Sub

Private Sub Worksheet_Deactivate()

    'Application.DisplayFullScreen = False
    'ActiveWindow.DisplayHeadings = True
    
End Sub

Private Sub Worksheet_PivotTableUpdate(ByVal Target As PivotTable)

Actualizar_Formas

End Sub
