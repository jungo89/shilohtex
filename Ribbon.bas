Option Explicit
Option Base 1
Public Cinta As IRibbonUI
Public RetVal(44) As Boolean

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
    'frm_.Show
    MsgBox ("Boton 1")
End Sub

Sub Boton2(control As IRibbonControl)
    'frm_.Show
    MsgBox ("Boton 2")
End Sub

Sub Boton3(control As IRibbonControl)
    'frm_.Show
    MsgBox ("Boton 3")
End Sub

Sub Boton4(control As IRibbonControl)
    'frm_.Show
    MsgBox ("Boton 4")
End Sub

Sub Boton5(control As IRibbonControl)
    'frm_.Show
    MsgBox ("Boton 5")
End Sub

Sub Boton6(control As IRibbonControl)
    'frm_.Show
    MsgBox ("Boton 6")
End Sub

Sub Boton7(control As IRibbonControl)
    'frm_.Show
    MsgBox ("Boton 7")
End Sub

Sub Boton8(control As IRibbonControl)
    'frm_.Show
    MsgBox ("Boton 8")
End Sub

Sub Boton9(control As IRibbonControl)
    'frm_.Show
    MsgBox ("Boton 9")
End Sub

Sub Boton10(control As IRibbonControl)
    'frm_.Show
    MsgBox ("Boton 10")
End Sub

Sub Boton11(control As IRibbonControl)
    'frm_.Show
    MsgBox ("Boton 11")
End Sub

Sub Boton12(control As IRibbonControl)
    'frm_.Show
    MsgBox ("Boton 12")
End Sub

Sub Boton13(control As IRibbonControl)
    'frm_.Show
    MsgBox ("Boton 13")
End Sub

Sub Boton14(control As IRibbonControl)
    'frm_.Show
    MsgBox ("Boton 14")
End Sub

Sub Boton15(control As IRibbonControl)
    'frm_.Show
    MsgBox ("Boton 15")
End Sub

Sub Boton16(control As IRibbonControl)
    'frm_.Show
    MsgBox ("Boton 16")
End Sub

Sub Boton17(control As IRibbonControl)
    'frm_.Show
    MsgBox ("Boton 17")
End Sub

Sub Boton18(control As IRibbonControl)
    'frm_.Show
    MsgBox ("Boton 18")
End Sub

Sub Boton19(control As IRibbonControl)
    'frm_.Show
    MsgBox ("Boton 19")
End Sub

Sub Boton20(control As IRibbonControl)
    'frm_.Show
    MsgBox ("Boton 20")
End Sub

Sub Boton21(control As IRibbonControl)
    'frm_.Show
    MsgBox ("Boton 21")
End Sub

Sub Boton22(control As IRibbonControl)
    'frm_.Show
    MsgBox ("Boton 22")
End Sub

Sub Boton23(control As IRibbonControl)
    'frm_.Show
    MsgBox ("Boton 23")
End Sub

Sub Boton24(control As IRibbonControl)
    'frm_.Show
    MsgBox ("Boton 24")
End Sub

Sub Boton25(control As IRibbonControl)
    'frm_.Show
    MsgBox ("Boton 25")
End Sub

Sub Boton26(control As IRibbonControl)
    'frm_.Show
    MsgBox ("Boton 26")
End Sub

Sub Boton27(control As IRibbonControl)
    'frm_.Show
    MsgBox ("Boton 27")
End Sub

Sub Boton28(control As IRibbonControl)
    'frm_.Show
    MsgBox ("Boton 28")
End Sub

Sub Boton29(control As IRibbonControl)
    'frm_.Show
    MsgBox ("Boton 29")
End Sub

Sub Boton30(control As IRibbonControl)
    'frm_.Show
    MsgBox ("Boton 30")
End Sub

Sub Boton31(control As IRibbonControl)
    'frm_.Show
    MsgBox ("Boton 31")
End Sub

Sub Boton32(control As IRibbonControl)
    'frm_.Show
    MsgBox ("Boton 32")
End Sub

Sub Boton33(control As IRibbonControl)
    'frm_.Show
    MsgBox ("Boton 33")
End Sub

Sub Boton34(control As IRibbonControl)
    'frm_.Show
    MsgBox ("Boton 34")
End Sub

Sub Boton35(control As IRibbonControl)
    'frm_.Show
    MsgBox ("Boton 35")
End Sub

Sub Boton36(control As IRibbonControl)
    'frm_.Show
    MsgBox ("Boton 36")
End Sub

Sub Boton37(control As IRibbonControl)
    'frm_.Show
    MsgBox ("Boton 37")
End Sub

Sub Boton38(control As IRibbonControl)
    'frm_.Show
    MsgBox ("Boton 38")
End Sub

Sub Boton39(control As IRibbonControl)
    'frm_NuevoUsuario.Show
    MsgBox ("Boton 39")
End Sub

Sub Boton40(control As IRibbonControl)
    'frm_Modificar_Permisos.Show
    MsgBox ("Boton 40")
End Sub

Sub Boton41(control As IRibbonControl)
    'frm.Show
    MsgBox ("Boton 41")
End Sub

Sub Boton42(control As IRibbonControl)
    'ThisWorkbook.Save
    MsgBox ("Boton 42")
End Sub

Sub Boton43(control As IRibbonControl)

'Dim sRuta As String
'Dim sNombreFolder As String
'Dim sSeparador As String
'Dim sRutaDestino As String
'Dim sBackUp As String
'
'    If MsgBox("Seguro que quiere crear una copia de seguridad", vbYesNo + vbQuestion, "Copia de Seguridad") = vbYes Then
'
'            sRuta = Application.ActiveWorkbook.Path
'            sSeparador = Application.PathSeparator
'
'            sBackUp = "Gestor_de_Inventarios_" & CStr(Format(Date, "yyyymmdd")) _
'            & "_" & CStr(Format(Time, "hh-mm-ss")) & ".xlsm"
'
'            sNombreFolder = "BackUp_" & CStr(Format(Date, "yyyy-mm"))
'
'            sRutaDestino = sRuta & sSeparador & sNombreFolder
'
'            If Dir(sRutaDestino, vbDirectory) = Empty Then
'                MkDir (sRutaDestino)
'            End If
'
'            Application.ActiveWorkbook.SaveCopyAs FileName:=sRutaDestino & sSeparador & sBackUp
'
'            MsgBox "La copia se guardó en la siguiente ruta: " & sRutaDestino & sSeparador _
'            & sBackUp, vbInformation, "Gestor de Inventarios"
'
'        Else
'            Exit Sub
'    End If

'frm.Show
    MsgBox ("Boton 43")
    
End Sub

Sub Boton44(control As IRibbonControl)
    'frm_Login.Show
    MsgBox ("Boton 44")
End Sub





'//////////////////// Retornos del estado de cada Boton ////////////////////////


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

Sub RetornoDelBoton19(control As IRibbonControl, ByRef ValorDevuelto)
    ValorDevuelto = RetVal(19)
End Sub

Sub RetornoDelBoton20(control As IRibbonControl, ByRef ValorDevuelto)
    ValorDevuelto = RetVal(20)
End Sub

Sub RetornoDelBoton21(control As IRibbonControl, ByRef ValorDevuelto)
    ValorDevuelto = RetVal(21)
End Sub

Sub RetornoDelBoton22(control As IRibbonControl, ByRef ValorDevuelto)
    ValorDevuelto = RetVal(22)
End Sub

Sub RetornoDelBoton23(control As IRibbonControl, ByRef ValorDevuelto)
    ValorDevuelto = RetVal(23)
End Sub

Sub RetornoDelBoton24(control As IRibbonControl, ByRef ValorDevuelto)
    ValorDevuelto = RetVal(24)
End Sub

Sub RetornoDelBoton25(control As IRibbonControl, ByRef ValorDevuelto)
    ValorDevuelto = RetVal(25)
End Sub

Sub RetornoDelBoton26(control As IRibbonControl, ByRef ValorDevuelto)
    ValorDevuelto = RetVal(26)
End Sub

Sub RetornoDelBoton27(control As IRibbonControl, ByRef ValorDevuelto)
    ValorDevuelto = RetVal(27)
End Sub

Sub RetornoDelBoton28(control As IRibbonControl, ByRef ValorDevuelto)
    ValorDevuelto = RetVal(28)
End Sub

Sub RetornoDelBoton29(control As IRibbonControl, ByRef ValorDevuelto)
    ValorDevuelto = RetVal(29)
End Sub

Sub RetornoDelBoton30(control As IRibbonControl, ByRef ValorDevuelto)
    ValorDevuelto = RetVal(30)
End Sub

Sub RetornoDelBoton31(control As IRibbonControl, ByRef ValorDevuelto)
    ValorDevuelto = RetVal(31)
End Sub

Sub RetornoDelBoton32(control As IRibbonControl, ByRef ValorDevuelto)
    ValorDevuelto = RetVal(32)
End Sub

Sub RetornoDelBoton33(control As IRibbonControl, ByRef ValorDevuelto)
    ValorDevuelto = RetVal(33)
End Sub

Sub RetornoDelBoton34(control As IRibbonControl, ByRef ValorDevuelto)
    ValorDevuelto = RetVal(34)
End Sub

Sub RetornoDelBoton35(control As IRibbonControl, ByRef ValorDevuelto)
    ValorDevuelto = RetVal(35)
End Sub

Sub RetornoDelBoton36(control As IRibbonControl, ByRef ValorDevuelto)
    ValorDevuelto = RetVal(36)
End Sub

Sub RetornoDelBoton37(control As IRibbonControl, ByRef ValorDevuelto)
    ValorDevuelto = RetVal(37)
End Sub

Sub RetornoDelBoton38(control As IRibbonControl, ByRef ValorDevuelto)
    ValorDevuelto = RetVal(38)
End Sub

Sub RetornoDelBoton39(control As IRibbonControl, ByRef ValorDevuelto)
    ValorDevuelto = RetVal(39)
End Sub

Sub RetornoDelBoton40(control As IRibbonControl, ByRef ValorDevuelto)
    ValorDevuelto = RetVal(40)
End Sub

Sub RetornoDelBoton41(control As IRibbonControl, ByRef ValorDevuelto)
    ValorDevuelto = RetVal(41)
End Sub

Sub RetornoDelBoton42(control As IRibbonControl, ByRef ValorDevuelto)
    ValorDevuelto = RetVal(42)
End Sub

Sub RetornoDelBoton43(control As IRibbonControl, ByRef ValorDevuelto)
    ValorDevuelto = RetVal(43)
End Sub

Sub RetornoDelBoton44(control As IRibbonControl, ByRef ValorDevuelto)
    ValorDevuelto = RetVal(44)
End Sub

