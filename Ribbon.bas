Attribute VB_Name = "Ribbon"
Option Explicit
Option Base 1
Public Cinta As IRibbonUI
Public RetVal(18) As Boolean

#If VBA7 And Win64 Then
    Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As LongPtr, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As LongPtr
#Else
    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
#End If


Sub CargarCinta(CintaDeExcel As IRibbonUI)
    Set Cinta = CintaDeExcel
    frm_Login.Show
End Sub

'////////////////////// Llamadas desde la Cinta para ejectuar cada formulario ///////////////////////////////////

Sub Boton1(control As IRibbonControl)
    frm_Setup_CodProd.Show
End Sub

Sub Boton2(control As IRibbonControl)
    frm_TipoMoneda.Show
End Sub

Sub Boton3(control As IRibbonControl)
    frm_RegistrarProducto.Show
End Sub

Sub Boton4(control As IRibbonControl)
    frm_ModificarProducto.Show
End Sub

Sub Boton5(control As IRibbonControl)
    frm_fCompras.Show
End Sub

Sub Boton6(control As IRibbonControl)
    frm_RegistrarProveedor.Show
End Sub

Sub Boton7(control As IRibbonControl)
    frm_EliminarProveedor.Show
End Sub

Sub Boton8(control As IRibbonControl)
    frm_DEV_Compras.Show
End Sub

Sub Boton9(control As IRibbonControl)
    frm_Factura.Show
End Sub

Sub Boton10(control As IRibbonControl)
    frm_RegistrarClientes.Show
End Sub

Sub Boton11(control As IRibbonControl)
    frm_EliminarCliente.Show
End Sub

Sub Boton12(control As IRibbonControl)
    frm_DEV_Ventas.Show
End Sub

Sub Boton13(control As IRibbonControl)
    frm_Transferencias.Show
End Sub

Sub Boton14(control As IRibbonControl)
    frm_Consulta_menu.Show
End Sub

Sub Boton15(control As IRibbonControl)
    frm_Movimientos.Show
End Sub

Sub Boton16(control As IRibbonControl)
    frm_NuevoUsuario.Show
End Sub

Sub Boton17(control As IRibbonControl)
    frm_EliminarUsuario.Show
End Sub

Sub Boton18(control As IRibbonControl)
    frm_Modificar_Permisos.Show
End Sub

Sub Boton19(control As IRibbonControl)
    ThisWorkbook.Save
End Sub

Sub Boton20(control As IRibbonControl)
    frm_Login.Show
End Sub

Sub Boton21(control As IRibbonControl)
    ShellExecute 0, "Open", "https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=6EL5B4TGHQATW", "", "", 1
End Sub

Sub Boton22(control As IRibbonControl)
    ShellExecute 0, "Open", "http://www.youtube.com/user/ottojaviergonzalez?sub_confirmation=1", "", "", 1
End Sub

Sub Boton23(control As IRibbonControl)
    ShellExecute 0, "Open", "https://www.facebook.com/ottojaviergonzalez", "", "", 1
End Sub

Sub Boton24(control As IRibbonControl)
    ShellExecute 0, "Open", "https://twitter.com/ottojgonzalez", "", "", 1
End Sub



'//////////////////// Retornos del estado de cada botón ////////////////////////


Sub RetornoDelBoton1(control As IRibbonControl, ByRef ValorDevuelto)
    ValorDevuelto = RetVal(1)
End Sub

Sub RetornoDelBoton2(control As IRibbonControl, ByRef ValorDevuelto)
    ValorDevuelto = RetVal(2)
End Sub

Sub RetornoDelBoton3(control As IRibbonControl, ByRef ValorDevuelto)
    ValorDevuelto = RetVal(3)
End Sub

Sub RetornoDelBoton4(control As IRibbonControl, ByRef ValorDevuelto)
    ValorDevuelto = RetVal(4)
End Sub

Sub RetornoDelBoton5(control As IRibbonControl, ByRef ValorDevuelto)
    ValorDevuelto = RetVal(5)
End Sub

Sub RetornoDelBoton6(control As IRibbonControl, ByRef ValorDevuelto)
    ValorDevuelto = RetVal(6)
End Sub

Sub RetornoDelBoton7(control As IRibbonControl, ByRef ValorDevuelto)
    ValorDevuelto = RetVal(7)
End Sub

Sub RetornoDelBoton8(control As IRibbonControl, ByRef ValorDevuelto)
    ValorDevuelto = RetVal(8)
End Sub

Sub RetornoDelBoton9(control As IRibbonControl, ByRef ValorDevuelto)
    ValorDevuelto = RetVal(9)
End Sub

Sub RetornoDelBoton10(control As IRibbonControl, ByRef ValorDevuelto)
    ValorDevuelto = RetVal(10)
End Sub

Sub RetornoDelBoton11(control As IRibbonControl, ByRef ValorDevuelto)
    ValorDevuelto = RetVal(11)
End Sub

Sub RetornoDelBoton12(control As IRibbonControl, ByRef ValorDevuelto)
    ValorDevuelto = RetVal(12)
End Sub

Sub RetornoDelBoton13(control As IRibbonControl, ByRef ValorDevuelto)
    ValorDevuelto = RetVal(13)
End Sub

Sub RetornoDelBoton14(control As IRibbonControl, ByRef ValorDevuelto)
    ValorDevuelto = RetVal(14)
End Sub

Sub RetornoDelBoton15(control As IRibbonControl, ByRef ValorDevuelto)
    ValorDevuelto = RetVal(15)
End Sub

Sub RetornoDelBoton16(control As IRibbonControl, ByRef ValorDevuelto)
    ValorDevuelto = RetVal(16)
End Sub

Sub RetornoDelBoton17(control As IRibbonControl, ByRef ValorDevuelto)
    ValorDevuelto = RetVal(17)
End Sub

Sub RetornoDelBoton18(control As IRibbonControl, ByRef ValorDevuelto)
    ValorDevuelto = RetVal(18)
End Sub







