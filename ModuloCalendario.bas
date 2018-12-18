Attribute VB_Name = "ModuloCalendario"
Option Explicit
Option Private Module
Private SenalCambioMes As Long

Public Sub RecibeLaFecha(Dia As Long, Mes As Long, Ano As Long)
    Dim FechaRecibida As Date
    FechaRecibida = VBA.DateSerial((VBA.CInt(Ano)), (VBA.CInt(Mes)), (VBA.CInt(Dia)))
    
    'DIRECCIONE LA FECHA AL CONTROL O CELDA QUE REQUIERA
    
       Call InsertarFecha(FechaRecibida)
    
End Sub

'********************************** NO MODIFICAR SI NO SABE **********************************
'*************************************|||||||||||||||||||*************************************
'***************************************|||||||||||||||***************************************
'*****************************************|||||||||||*****************************************
'*******************************************|||||||*******************************************
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
    
    frmCalendario.spbA�o.Value = VBA.Year(VBA.Date)
    
    frmCalendario.lblAno.Caption = VBA.Year(VBA.Date)
    
    Dim Ano As Long, Mes As Long
    Ano = VBA.Year(VBA.Date)
    Mes = VBA.Month(VBA.Date)
    Call ModuloCalendario.CargarLosDias(Ano, Mes)
    
    frmCalendario.lblHoy.Caption = VBA.Date
End Sub

Public Sub CargarLosDias(Ano As Long, Mes As Long)
    Dim FechaDelPrimerDia As Date
    Dim FechaDelUltimoDia As Date
    Dim DiaSemanaPrimerDia As Long
    Dim VariableControl As control
    Dim Contador As Long
    
    FechaDelPrimerDia = VBA.DateSerial(Ano, Mes, 1)
    FechaDelUltimoDia = Application.WorksheetFunction.EoMonth(VBA.DateSerial(Ano, Mes, 1), 0)
    DiaSemanaPrimerDia = Application.WorksheetFunction.Weekday(FechaDelPrimerDia, 2)
    Contador = 1
    
    For Each VariableControl In frmCalendario.mrcDias.Controls
        VariableControl.Caption = "-"
        If VariableControl.Tag >= DiaSemanaPrimerDia And Contador <= VBA.Day(FechaDelUltimoDia) Then
            VariableControl.Caption = Contador
            Contador = Contador + 1
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
    
    frmCalendario.lblAno.Caption = frmCalendario.spbA�o.Value
    
    MesEnElCombo = VBA.CLng(frmCalendario.cboMes.Value)
    AnoEnElLabel = VBA.CLng(frmCalendario.lblAno.Caption)
    Call ModuloCalendario.DesmarcarDias
    Call ModuloCalendario.CargarLosDias(AnoEnElLabel, MesEnElCombo)
    
End Sub

Public Sub UnClickEnHoyEs()
    Dim Mes As Long, Ano As Long
    Dim FechaActual As Date
    
    FechaActual = VBA.CDate(frmCalendario.lblHoy.Caption)
    Mes = VBA.CLng(VBA.Month(FechaActual))
    Ano = VBA.CLng(VBA.Year(FechaActual))
    
    frmCalendario.lblAno.Caption = Ano
    frmCalendario.cboMes.ListIndex = Mes - 1
    frmCalendario.spbA�o.Value = Ano
    frmCalendario.spbA�o.SetFocus
    
    Call ModuloCalendario.DesmarcarDias
    Call ModuloCalendario.CargarLosDias(Ano, Mes)
    
End Sub

Sub SalirConEscape()
    Unload frmCalendario
End Sub

Sub MarcarDia(ControlDeEtiqueta As control)
    Call ModuloCalendario.DesmarcarDias
    ControlDeEtiqueta.Font.Bold = True
    ControlDeEtiqueta.ForeColor = VBA.RGB(255, 0, 0)
End Sub

Sub DesmarcarDias()
    Dim ControlEtiqueta As control
    
    For Each ControlEtiqueta In frmCalendario.mrcDias.Controls
        ControlEtiqueta.Font.Bold = False
        ControlEtiqueta.ForeColor = VBA.RGB(0, 0, 0)
    Next ControlEtiqueta
End Sub

'*******************************************|||||||*******************************************
'*****************************************|||||||||||*****************************************
'***************************************|||||||||||||||***************************************
'*************************************||||||||||||||||||**************************************
'********************************** NO MODIFICAR SI NO SABE **********************************

' Nota del autor -----------------------------------------------------------------------------

' Creado por Andr�s Rojas Moncada - Autor del canal Excel Hecho F�cil en YouTube

' Versi�n 1.0 - 20 de julio de 2015

' URL del canal: www.youtube.com/jarmoncada01

' Si quieres usarlo, solo copia y pega el presente m�dulo en conjunto con el UserForm y listo.

' Para ver algunos ejemplos sobre el uso de este calendario, observa este video.

' Enlace: |||||||||||||||||||| https://www.youtube.com/watch?v=FkjsuN2zqSU ||||||||||||||||||||

' ! Muchas gracias y espero lo disfruten ! ---------------------------------------------------


