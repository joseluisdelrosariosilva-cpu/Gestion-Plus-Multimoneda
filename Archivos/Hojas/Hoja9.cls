VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    
    Dim tbl As ListObject
    Dim selectedRow As ListRow
    Dim facturaID As String
    Dim selectedCell As Range
    Dim si As SlicerItem
    
    Set tbl = ThisWorkbook.Sheets("Registro de Facturas").ListObjects("HISTORIALVENTAS4")
   
    If Not Application.Intersect(Target, tbl.Range) Is Nothing Then
        Set selectedCell = Target.Cells(1, 1)
        
        On Error Resume Next
        Set selectedRow = tbl.ListRows(selectedCell.Row - tbl.HeaderRowRange.Row)
        On Error GoTo 0

        If Not selectedRow Is Nothing Then
            facturaID = selectedRow.Range(1)
            Me.Range("I3") = facturaID
            
        End If
    End If
            
End Sub
