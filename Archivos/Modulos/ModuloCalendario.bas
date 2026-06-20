Attribute VB_Name = "ModuloCalendario"
Option Explicit
Option Private Module
Private SenalCambioMes As Long

Public Sub RecibeLaFecha(dia As Long, mes As Long, Ano As Long)
    Dim FechaRecibida As Date
    FechaRecibida = VBA.DateSerial((VBA.CInt(Ano)), (VBA.CInt(mes)), (VBA.CInt(dia)))
    
'INDICAR UBICACIÓN DE LA FECHA
    If seleccion_destino_calendario = 1 Then
        ingresar_form.fecha_de_entrada_TextBox.Value = Format(FechaRecibida, "d-mmm-yyyy")
    ElseIf seleccion_destino_calendario = 2 Then
        facturacion.fecha_de_orden_TextBox.Value = Format(FechaRecibida, "d-mmm-yyyy")
    ElseIf seleccion_destino_calendario = 3 Then
        Busqueda.Txt_Fecha.Value = Format(FechaRecibida, "d-mmm-yyyy")
        If CDate(Busqueda.Txt_Fecha.Value) > CDate(Busqueda.Txt_Fecha_Final.Value) Then
            Busqueda.Txt_Fecha_Final.Value = Format(FechaRecibida, "d-mmm-yyyy")
        End If
    ElseIf seleccion_destino_calendario = 4 Then
        Busqueda.Txt_Fecha_Final.Value = Format(FechaRecibida, "d-mmm-yyyy")
        If CDate(Busqueda.Txt_Fecha.Value) > CDate(Busqueda.Txt_Fecha_Final.Value) Then
            Busqueda.Txt_Fecha.Value = Format(FechaRecibida, "d-mmm-yyyy")
        End If
    ElseIf seleccion_destino_calendario = 5 Then
        Busqueda_Ventas.Txt_Fecha.Value = Format(FechaRecibida, "d-mmm-yyyy")
        If CDate(Busqueda_Ventas.Txt_Fecha.Value) > CDate(Busqueda_Ventas.Txt_Fecha_Final.Value) Then
            Busqueda_Ventas.Txt_Fecha_Final.Value = Format(FechaRecibida, "d-mmm-yyyy")
        End If
    ElseIf seleccion_destino_calendario = 6 Then
        Busqueda_Ventas.Txt_Fecha_Final.Value = Format(FechaRecibida, "d-mmm-yyyy")
        If CDate(Busqueda_Ventas.Txt_Fecha.Value) > CDate(Busqueda_Ventas.Txt_Fecha_Final.Value) Then
            Busqueda_Ventas.Txt_Fecha.Value = Format(FechaRecibida, "d-mmm-yyyy")
        End If
     ElseIf seleccion_destino_calendario = 7 Then
        Busqueda_Facturacion.Txt_Fecha.Value = Format(FechaRecibida, "d-mmm-yyyy")
        If CDate(Busqueda_Facturacion.Txt_Fecha.Value) > CDate(Busqueda_Facturacion.Txt_Fecha_Final.Value) Then
            Busqueda_Facturacion.Txt_Fecha_Final.Value = Format(FechaRecibida, "d-mmm-yyyy")
        End If
     ElseIf seleccion_destino_calendario = 8 Then
        Busqueda_Facturacion.Txt_Fecha_Final.Value = Format(FechaRecibida, "d-mmm-yyyy")
        If CDate(Busqueda_Facturacion.Txt_Fecha.Value) > CDate(Busqueda_Facturacion.Txt_Fecha_Final.Value) Then
            Busqueda_Facturacion.Txt_Fecha.Value = Format(FechaRecibida, "d-mmm-yyyy")
        End If
    End If
End Sub

'********************************** NO MODIFICAR SI NO SABE **********************************

Public Sub InicializaFormularioCalendario()
    SenalCambioMes = 1
    
    With frmCalendario.cboMes
        .AddItem 1
        .List(0, 1) = "enero"
        .AddItem 2
        .List(1, 1) = "febrero"
        .AddItem 3
        .List(2, 1) = "marzo"
        .AddItem 4
        .List(3, 1) = "abril"
        .AddItem 5
        .List(4, 1) = "mayo"
        .AddItem 6
        .List(5, 1) = "junio"
        .AddItem 7
        .List(6, 1) = "julio"
        .AddItem 8
        .List(7, 1) = "agosto"
        .AddItem 9
        .List(8, 1) = "septiembre"
        .AddItem 10
        .List(9, 1) = "octubre"
        .AddItem 11
        .List(10, 1) = "noviembre"
        .AddItem 12
        .List(11, 1) = "diciembre"
    End With
    
    frmCalendario.cboMes.ListIndex = VBA.Month(VBA.Date) - 1
    
    frmCalendario.spbAño.Value = VBA.year(VBA.Date)
    
    frmCalendario.lblAno.Caption = VBA.year(VBA.Date)
    
    Dim Ano As Long, mes As Long
    Ano = VBA.year(VBA.Date)
    mes = VBA.Month(VBA.Date)
    Call ModuloCalendario.CargarLosDias(Ano, mes)
    
    frmCalendario.lblHoy.Caption = Format(VBA.Date, "d-mmm-yyyy")
End Sub

Public Sub CargarLosDias(Ano As Long, mes As Long)
    Dim FechaDelPrimerDia As Date
    Dim FechaDelUltimoDia As Date
    Dim DiaSemanaPrimerDia As Long
    Dim VariableControl As Control
    Dim contador As Long
    
    FechaDelPrimerDia = VBA.DateSerial(Ano, mes, 1)
    FechaDelUltimoDia = Application.WorksheetFunction.EoMonth(VBA.DateSerial(Ano, mes, 1), 0)
    DiaSemanaPrimerDia = Application.WorksheetFunction.Weekday(FechaDelPrimerDia, 2)
    contador = 1
    
    For Each VariableControl In frmCalendario.mrcDias.Controls
        VariableControl.Caption = "-"
        If VariableControl.Tag >= DiaSemanaPrimerDia And contador <= VBA.Day(FechaDelUltimoDia) Then
            VariableControl.Caption = contador
            contador = contador + 1
        End If
    Next VariableControl
End Sub

Public Sub CambioDeMes()
    If SenalCambioMes > 1 Then
        Dim MesEnElCombo As Long, AnoEnElLabel As Long
        
        If Not (IsNull(frmCalendario.cboMes.Value)) And Not (IsNull(frmCalendario.lblAno.Caption)) Then
            MesEnElCombo = VBA.CLng(frmCalendario.cboMes.Value)
            AnoEnElLabel = VBA.CLng(frmCalendario.lblAno.Caption)
            Call ModuloCalendario.DesmarcarDias
            Call ModuloCalendario.CargarLosDias(AnoEnElLabel, MesEnElCombo)
        End If
    End If
    SenalCambioMes = SenalCambioMes + 1
End Sub

Public Sub CambioDeAno()
    Dim MesEnElCombo As Long, AnoEnElLabel As Long
    
    frmCalendario.lblAno.Caption = frmCalendario.spbAño.Value
    
    MesEnElCombo = VBA.CLng(frmCalendario.cboMes.Value)
    AnoEnElLabel = VBA.CLng(frmCalendario.lblAno.Caption)
    Call ModuloCalendario.DesmarcarDias
    Call ModuloCalendario.CargarLosDias(AnoEnElLabel, MesEnElCombo)
    
End Sub

Public Sub UnClickEnHoyEs()
    Dim mes As Long, Ano As Long
    Dim FechaActual As Date
    
    FechaActual = VBA.CDate(frmCalendario.lblHoy.Caption)
    mes = VBA.CLng(VBA.Month(FechaActual))
    Ano = VBA.CLng(VBA.year(FechaActual))
    
    frmCalendario.lblAno.Caption = Ano
    frmCalendario.cboMes.ListIndex = mes - 1
    frmCalendario.spbAño.Value = Ano
    frmCalendario.spbAño.SetFocus
    
    Call ModuloCalendario.DesmarcarDias
    Call ModuloCalendario.CargarLosDias(Ano, mes)
    
End Sub

Sub SalirConEscape()
    Unload frmCalendario
End Sub

Sub MarcarDia(ControlDeEtiqueta As Control)
    Call ModuloCalendario.DesmarcarDias
    ControlDeEtiqueta.Font.Bold = True
    ControlDeEtiqueta.ForeColor = VBA.RGB(255, 0, 0)
End Sub

Sub DesmarcarDias()
    Dim ControlEtiqueta As Control
    
    For Each ControlEtiqueta In frmCalendario.mrcDias.Controls
        ControlEtiqueta.Font.Bold = False
        ControlEtiqueta.ForeColor = VBA.RGB(0, 0, 0)
    Next ControlEtiqueta
End Sub

'********************************** NO MODIFICAR SI NO SABE **********************************

' Nota del autor -----------------------------------------------------------------------------

' Creado por Andrés Rojas Moncada - Autor del canal Excel Hecho Fácil en YouTube

' Versión 1.0 - 20 de julio de 2015


