VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCalendario 
   Caption         =   "Seleccione una fecha"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   2670
   OleObjectBlob   =   "frmCalendario.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmCalendario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Sub cboMes_Click()
    ModuloCalendario.CambioDeMes
End Sub

Private Sub lbl1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl1.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frmCalendario.lbl1.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)
        
        Unload frmCalendario
        Call ModuloCalendario.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Dim Control_lbl1 As control
    Set Control_lbl1 = frmCalendario.lbl1
    
    Call ModuloCalendario.MarcarDia(Control_lbl1)
End Sub

Private Sub lbl10_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl10.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frmCalendario.lbl10.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)
        
        Unload frmCalendario
        Call ModuloCalendario.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl10_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Dim Control_lbl10 As control
    Set Control_lbl10 = frmCalendario.lbl10
    
    Call ModuloCalendario.MarcarDia(Control_lbl10)
End Sub

Private Sub lbl11_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl11.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frmCalendario.lbl11.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)
        
        Unload frmCalendario
        Call ModuloCalendario.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl11_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Dim Control_lbl11 As control
    Set Control_lbl11 = frmCalendario.lbl11
    
    Call ModuloCalendario.MarcarDia(Control_lbl11)
End Sub

Private Sub lbl12_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl12.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frmCalendario.lbl12.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)
        
        Unload frmCalendario
        Call ModuloCalendario.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl12_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Dim Control_lbl12 As control
    Set Control_lbl12 = frmCalendario.lbl12
    
    Call ModuloCalendario.MarcarDia(Control_lbl12)
End Sub

Private Sub lbl13_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl13.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frmCalendario.lbl13.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)
        
        Unload frmCalendario
        Call ModuloCalendario.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl13_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Dim Control_lbl13 As control
    Set Control_lbl13 = frmCalendario.lbl13
    
    Call ModuloCalendario.MarcarDia(Control_lbl13)
End Sub

Private Sub lbl14_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl14.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frmCalendario.lbl14.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)
        
        Unload frmCalendario
        Call ModuloCalendario.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl14_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Dim Control_lbl14 As control
    Set Control_lbl14 = frmCalendario.lbl14
    
    Call ModuloCalendario.MarcarDia(Control_lbl14)
End Sub

Private Sub lbl15_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl15.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frmCalendario.lbl15.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)
        
        Unload frmCalendario
        Call ModuloCalendario.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl15_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Dim Control_lbl15 As control
    Set Control_lbl15 = frmCalendario.lbl15
    
    Call ModuloCalendario.MarcarDia(Control_lbl15)
End Sub

Private Sub lbl16_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl16.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frmCalendario.lbl16.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)
        
        Unload frmCalendario
        Call ModuloCalendario.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl16_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Dim Control_lbl16 As control
    Set Control_lbl16 = frmCalendario.lbl16
    
    Call ModuloCalendario.MarcarDia(Control_lbl16)
End Sub

Private Sub lbl17_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl17.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frmCalendario.lbl17.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)
        
        Unload frmCalendario
        Call ModuloCalendario.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl17_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Dim Control_lbl17 As control
    Set Control_lbl17 = frmCalendario.lbl17
    
    Call ModuloCalendario.MarcarDia(Control_lbl17)
End Sub

Private Sub lbl18_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl18.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frmCalendario.lbl18.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)
        
        Unload frmCalendario
        Call ModuloCalendario.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl18_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Dim Control_lbl18 As control
    Set Control_lbl18 = frmCalendario.lbl18
    
    Call ModuloCalendario.MarcarDia(Control_lbl18)
End Sub

Private Sub lbl19_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl19.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frmCalendario.lbl19.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)
        
        Unload frmCalendario
        Call ModuloCalendario.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl19_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Dim Control_lbl19 As control
    Set Control_lbl19 = frmCalendario.lbl19
    
    Call ModuloCalendario.MarcarDia(Control_lbl19)
End Sub

Private Sub lbl2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl2.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frmCalendario.lbl2.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)
        
        Unload frmCalendario
        Call ModuloCalendario.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Dim Control_lbl2 As control
    Set Control_lbl2 = frmCalendario.lbl2
    
    Call ModuloCalendario.MarcarDia(Control_lbl2)
End Sub

Private Sub lbl20_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl20.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frmCalendario.lbl20.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)
        
        Unload frmCalendario
        Call ModuloCalendario.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl20_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Dim Control_lbl20 As control
    Set Control_lbl20 = frmCalendario.lbl20
    
    Call ModuloCalendario.MarcarDia(Control_lbl20)
End Sub

Private Sub lbl21_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl21.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frmCalendario.lbl21.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)
        
        Unload frmCalendario
        Call ModuloCalendario.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl21_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Dim Control_lbl21 As control
    Set Control_lbl21 = frmCalendario.lbl21
    
    Call ModuloCalendario.MarcarDia(Control_lbl21)
End Sub

Private Sub lbl22_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl22.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frmCalendario.lbl22.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)
        
        Unload frmCalendario
        Call ModuloCalendario.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl22_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Dim Control_lbl22 As control
    Set Control_lbl22 = frmCalendario.lbl22
    
    Call ModuloCalendario.MarcarDia(Control_lbl22)
End Sub

Private Sub lbl23_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl23.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frmCalendario.lbl23.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)
        
        Unload frmCalendario
        Call ModuloCalendario.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl23_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Dim Control_lbl23 As control
    Set Control_lbl23 = frmCalendario.lbl23
    
    Call ModuloCalendario.MarcarDia(Control_lbl23)
End Sub

Private Sub lbl24_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl24.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frmCalendario.lbl24.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)
        
        Unload frmCalendario
        Call ModuloCalendario.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl24_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Dim Control_lbl24 As control
    Set Control_lbl24 = frmCalendario.lbl24
    
    Call ModuloCalendario.MarcarDia(Control_lbl24)
End Sub

Private Sub lbl25_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl25.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frmCalendario.lbl25.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)
        
        Unload frmCalendario
        Call ModuloCalendario.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl25_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Dim Control_lbl25 As control
    Set Control_lbl25 = frmCalendario.lbl25
    
    Call ModuloCalendario.MarcarDia(Control_lbl25)
End Sub

Private Sub lbl26_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl26.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frmCalendario.lbl26.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)
        
        Unload frmCalendario
        Call ModuloCalendario.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl26_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Dim Control_lbl26 As control
    Set Control_lbl26 = frmCalendario.lbl26
    
    Call ModuloCalendario.MarcarDia(Control_lbl26)
End Sub

Private Sub lbl27_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl27.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frmCalendario.lbl27.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)
        
        Unload frmCalendario
        Call ModuloCalendario.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl27_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Dim Control_lbl27 As control
    Set Control_lbl27 = frmCalendario.lbl27
    
    Call ModuloCalendario.MarcarDia(Control_lbl27)
End Sub

Private Sub lbl28_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl28.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frmCalendario.lbl28.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)
        
        Unload frmCalendario
        Call ModuloCalendario.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl28_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Dim Control_lbl28 As control
    Set Control_lbl28 = frmCalendario.lbl28
    
    Call ModuloCalendario.MarcarDia(Control_lbl28)
End Sub

Private Sub lbl29_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl29.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frmCalendario.lbl29.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)
        
        Unload frmCalendario
        Call ModuloCalendario.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl29_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Dim Control_lbl29 As control
    Set Control_lbl29 = frmCalendario.lbl29
    
    Call ModuloCalendario.MarcarDia(Control_lbl29)
End Sub

Private Sub lbl3_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl3.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frmCalendario.lbl3.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)
        
        Unload frmCalendario
        Call ModuloCalendario.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Dim Control_lbl3 As control
    Set Control_lbl3 = frmCalendario.lbl3
    
    Call ModuloCalendario.MarcarDia(Control_lbl3)
End Sub

Private Sub lbl30_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl30.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frmCalendario.lbl30.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)
        
        Unload frmCalendario
        Call ModuloCalendario.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl30_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Dim Control_lbl30 As control
    Set Control_lbl30 = frmCalendario.lbl30
    
    Call ModuloCalendario.MarcarDia(Control_lbl30)
End Sub

Private Sub lbl31_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl31.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frmCalendario.lbl31.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)
        
        Unload frmCalendario
        Call ModuloCalendario.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl31_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Dim Control_lbl31 As control
    Set Control_lbl31 = frmCalendario.lbl31
    
    Call ModuloCalendario.MarcarDia(Control_lbl31)
End Sub

Private Sub lbl32_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl32.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frmCalendario.lbl32.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)
        
        Unload frmCalendario
        Call ModuloCalendario.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl32_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Dim Control_lbl32 As control
    Set Control_lbl32 = frmCalendario.lbl32
    
    Call ModuloCalendario.MarcarDia(Control_lbl32)
End Sub

Private Sub lbl33_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl33.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frmCalendario.lbl33.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)
        
        Unload frmCalendario
        Call ModuloCalendario.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl33_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Dim Control_lbl33 As control
    Set Control_lbl33 = frmCalendario.lbl33
    
    Call ModuloCalendario.MarcarDia(Control_lbl33)
End Sub

Private Sub lbl34_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl34.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frmCalendario.lbl34.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)
        
        Unload frmCalendario
        Call ModuloCalendario.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl34_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Dim Control_lbl34 As control
    Set Control_lbl34 = frmCalendario.lbl34
    
    Call ModuloCalendario.MarcarDia(Control_lbl34)
End Sub

Private Sub lbl35_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl35.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frmCalendario.lbl35.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)
        
        Unload frmCalendario
        Call ModuloCalendario.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl35_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Dim Control_lbl35 As control
    Set Control_lbl35 = frmCalendario.lbl35
    
    Call ModuloCalendario.MarcarDia(Control_lbl35)
End Sub

Private Sub lbl36_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl36.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frmCalendario.lbl36.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)
        
        Unload frmCalendario
        Call ModuloCalendario.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl36_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Dim Control_lbl36 As control
    Set Control_lbl36 = frmCalendario.lbl36
    
    Call ModuloCalendario.MarcarDia(Control_lbl36)
End Sub

Private Sub lbl37_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl37.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frmCalendario.lbl37.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)
        
        Unload frmCalendario
        Call ModuloCalendario.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl37_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Dim Control_lbl37 As control
    Set Control_lbl37 = frmCalendario.lbl37
    
    Call ModuloCalendario.MarcarDia(Control_lbl37)
End Sub

Private Sub lbl38_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl38.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frmCalendario.lbl38.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)
        
        Unload frmCalendario
        Call ModuloCalendario.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl38_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Dim Control_lbl38 As control
    Set Control_lbl38 = frmCalendario.lbl38
    
    Call ModuloCalendario.MarcarDia(Control_lbl38)
End Sub

Private Sub lbl39_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl39.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frmCalendario.lbl39.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)
        
        Unload frmCalendario
        Call ModuloCalendario.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl39_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Dim Control_lbl39 As control
    Set Control_lbl39 = frmCalendario.lbl39
    
    Call ModuloCalendario.MarcarDia(Control_lbl39)
End Sub

Private Sub lbl4_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl4.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frmCalendario.lbl4.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)
        
        Unload frmCalendario
        Call ModuloCalendario.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Dim Control_lbl4 As control
    Set Control_lbl4 = frmCalendario.lbl4
    
    Call ModuloCalendario.MarcarDia(Control_lbl4)
End Sub

Private Sub lbl40_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl40.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frmCalendario.lbl40.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)
        
        Unload frmCalendario
        Call ModuloCalendario.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl40_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Dim Control_lbl40 As control
    Set Control_lbl40 = frmCalendario.lbl40
    
    Call ModuloCalendario.MarcarDia(Control_lbl40)
End Sub

Private Sub lbl41_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl41.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frmCalendario.lbl41.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)
        
        Unload frmCalendario
        Call ModuloCalendario.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl41_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Dim Control_lbl41 As control
    Set Control_lbl41 = frmCalendario.lbl41
    
    Call ModuloCalendario.MarcarDia(Control_lbl41)
End Sub

Private Sub lbl42_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl42.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frmCalendario.lbl42.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)
        
        Unload frmCalendario
        Call ModuloCalendario.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl42_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Dim Control_lbl42 As control
    Set Control_lbl42 = frmCalendario.lbl42
    
    Call ModuloCalendario.MarcarDia(Control_lbl42)
End Sub

Private Sub lbl5_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl5.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frmCalendario.lbl5.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)
        
        Unload frmCalendario
        Call ModuloCalendario.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl5_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Dim Control_lbl5 As control
    Set Control_lbl5 = frmCalendario.lbl5
    
    Call ModuloCalendario.MarcarDia(Control_lbl5)
End Sub

Private Sub lbl6_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl6.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frmCalendario.lbl6.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)
        
        Unload frmCalendario
        Call ModuloCalendario.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl6_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Dim Control_lbl6 As control
    Set Control_lbl6 = frmCalendario.lbl6
    
    Call ModuloCalendario.MarcarDia(Control_lbl6)
End Sub

Private Sub lbl7_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl7.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frmCalendario.lbl7.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)
        
        Unload frmCalendario
        Call ModuloCalendario.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl7_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Dim Control_lbl7 As control
    Set Control_lbl7 = frmCalendario.lbl7
    
    Call ModuloCalendario.MarcarDia(Control_lbl7)
End Sub

Private Sub lbl8_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl8.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frmCalendario.lbl8.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)
        
        Unload frmCalendario
        Call ModuloCalendario.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl8_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Dim Control_lbl8 As control
    Set Control_lbl8 = frmCalendario.lbl8
    
    Call ModuloCalendario.MarcarDia(Control_lbl8)
End Sub

Private Sub lbl9_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl9.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long
        
        Dia = VBA.CLng(frmCalendario.lbl9.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)
        
        Unload frmCalendario
        Call ModuloCalendario.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl9_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Dim Control_lbl9 As control
    Set Control_lbl9 = frmCalendario.lbl9
    
    Call ModuloCalendario.MarcarDia(Control_lbl9)
End Sub

Private Sub cmdSalirConEscape_Click()
    Call ModuloCalendario.SalirConEscape
End Sub

Private Sub lblHoy_Click()
    ModuloCalendario.UnClickEnHoyEs
End Sub

Private Sub spbA�o_Change()
    ModuloCalendario.CambioDeAno
End Sub

Private Sub UserForm_Initialize()
    Call ModuloCalendario.InicializaFormularioCalendario
End Sub
