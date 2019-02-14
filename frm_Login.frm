VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Login 
   Caption         =   "Gestor de Inventarios"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5640
   OleObjectBlob   =   "frm_Login.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Option Base 1
Private Sub btn_Ingresar_Click()
Dim Usuario As String
Dim Fila, Final As Long
Dim password As String, UsuarioEncontrado As String, yaExiste As Byte, Status As String
Dim Rango As Range
Dim Titulo As String
Dim Hoja As Worksheet
Dim vHoja(13) As String
Dim vBoton(18) As String
Dim i As Byte
Dim x As Byte

On Error GoTo Salir

Application.ScreenUpdating = False

Titulo = "Gestor de Inventarios"


yaExiste = Application.WorksheetFunction.CountIf(Hoja6.Range("tbl_Usuario[Usuario]"), Me.txtUsuario.Text)
Set Rango = Hoja6.Range("tbl_Usuario[Usuario]")

If Me.txtUsuario.Text = "" Or Me.txtPassword.Text = "" Then
    MsgBox "Introduce usuario y contraseña", vbExclamation, Titulo
    Me.txtUsuario.SetFocus

            ElseIf yaExiste = 0 Then
                MsgBox "El usuario '" & Me.txtUsuario.Text & "' no existe", vbExclamation, Titulo
            
            ElseIf yaExiste = 1 Then
                UsuarioEncontrado = Rango.Find(What:=Me.txtUsuario.Text, after:=Rango.Range("A1"), _
                                                LookAt:=xlWhole, MatchCase:=False).Address
                
                password = Hoja6.Range(UsuarioEncontrado).Offset(0, 1).Value
                Status = Hoja6.Range(UsuarioEncontrado).Offset(0, 2).Value
                
                'Permisos y restricciones
                vHoja(2) = Hoja6.Range(UsuarioEncontrado).Offset(0, 3).Value
                vHoja(3) = Hoja6.Range(UsuarioEncontrado).Offset(0, 4).Value
                vHoja(4) = Hoja6.Range(UsuarioEncontrado).Offset(0, 5).Value
                vHoja(5) = Hoja6.Range(UsuarioEncontrado).Offset(0, 6).Value
                vHoja(6) = Hoja6.Range(UsuarioEncontrado).Offset(0, 7).Value
                vHoja(7) = Hoja6.Range(UsuarioEncontrado).Offset(0, 8).Value
                vHoja(8) = Hoja6.Range(UsuarioEncontrado).Offset(0, 9).Value
                vHoja(9) = Hoja6.Range(UsuarioEncontrado).Offset(0, 10).Value
                vHoja(10) = Hoja6.Range(UsuarioEncontrado).Offset(0, 11).Value
                vHoja(11) = Hoja6.Range(UsuarioEncontrado).Offset(0, 12).Value
                vHoja(12) = Hoja6.Range(UsuarioEncontrado).Offset(0, 13).Value
                vHoja(13) = Hoja6.Range(UsuarioEncontrado).Offset(0, 14).Value
                
                vBoton(1) = Hoja6.Range(UsuarioEncontrado).Offset(0, 15).Value
                vBoton(2) = Hoja6.Range(UsuarioEncontrado).Offset(0, 16).Value
                vBoton(3) = Hoja6.Range(UsuarioEncontrado).Offset(0, 17).Value
                vBoton(4) = Hoja6.Range(UsuarioEncontrado).Offset(0, 18).Value
                vBoton(5) = Hoja6.Range(UsuarioEncontrado).Offset(0, 19).Value
                vBoton(6) = Hoja6.Range(UsuarioEncontrado).Offset(0, 20).Value
                vBoton(7) = Hoja6.Range(UsuarioEncontrado).Offset(0, 21).Value
                vBoton(8) = Hoja6.Range(UsuarioEncontrado).Offset(0, 22).Value
                vBoton(9) = Hoja6.Range(UsuarioEncontrado).Offset(0, 23).Value
                vBoton(10) = Hoja6.Range(UsuarioEncontrado).Offset(0, 24).Value
                vBoton(11) = Hoja6.Range(UsuarioEncontrado).Offset(0, 25).Value
                vBoton(12) = Hoja6.Range(UsuarioEncontrado).Offset(0, 26).Value
                vBoton(13) = Hoja6.Range(UsuarioEncontrado).Offset(0, 27).Value
                vBoton(14) = Hoja6.Range(UsuarioEncontrado).Offset(0, 28).Value
                vBoton(15) = Hoja6.Range(UsuarioEncontrado).Offset(0, 29).Value
                vBoton(16) = Hoja6.Range(UsuarioEncontrado).Offset(0, 30).Value
                vBoton(17) = Hoja6.Range(UsuarioEncontrado).Offset(0, 31).Value
                vBoton(18) = Hoja6.Range(UsuarioEncontrado).Offset(0, 32).Value
                
                
                
    
    
            If Hoja6.Range(UsuarioEncontrado).Value = Me.txtUsuario.Text And password = Me.txtPassword.Text Then
            
                        'Validando los permisos y restricciones en las hojas de cálculo
                        For i = 2 To 13
                            For Each Hoja In Worksheets
                            If Hoja.CodeName = "Hoja" & i Then
                                If vHoja(i) = False Then
                                    Hoja.Visible = xlSheetVeryHidden
                                Else
                                    Hoja.Visible = xlSheetVisible
                                End If
                            End If
                            Next Hoja
                        Next i
                                        
                        
                        'Habilitar o deshabilitar botones en la Ribbon
                        
                        For x = 1 To 18
                        
                            If vBoton(x) = True Then
                                RetVal(x) = True
                                Cinta.InvalidateControl ("Button" & x)
                            Else
                                RetVal(x) = False
                                Cinta.InvalidateControl ("Button" & x)
                            End If
                        Next x
            
            
                          ' Registrar al usuario en la hoja Logs
                              
                              Final = GetNuevoR(Hoja8)
                                  Hoja8.Cells(Final, 1) = "=NOW()"
                                  Hoja8.Cells(Final, 1).Copy
                                  Hoja8.Cells(Final, 1).PasteSpecial Paste:=xlPasteValues
                                  Application.CutCopyMode = False
                                  
                                  Hoja8.Cells(Final, 2) = Me.txtUsuario.Text
                                  
                                  Hoja1.txt_UsuarioActual.Text = "Usuario actual: " & UCase(Me.txtUsuario.Text)
                                  
                                  Hoja8.Cells(Final, 3) = Status
                                  
                    
                                 
                                  Hoja8.Range("G1") = Me.txtUsuario.Text
                                  Hoja8.Range("H1") = Status
                                  
                                  
                                '------------Configuración Regional--------------------
                                    Call infoSeparadorDecimal
   
                                        If Hoja12.Range("C5") = "," Then
                                             Formato_Europeo
                                        Else
                                             Formato_Americano
                                        End If
                                '------------------------------------------------------
                                  
                                  ThisWorkbook.Save
                              
                              
                                  Unload Me
                                  Hoja1.Activate
                                  Call CopiaSeguridad
                        Else
                     MsgBox "La contraseña es incorrecta", vbExclamation, Titulo
            End If
End If

Application.ScreenUpdating = True

Salir:
 If Err <> 0 Then
    MsgBox Err.Description, vbExclamation, Titulo
 End If

End Sub
Private Sub btn_Salir_Click()
    ThisWorkbook.Application.DisplayAlerts = False
    Application.ActiveWorkbook.Close
    Unload Me
End Sub

Private Sub imgLogo_Click()

End Sub

Private Sub UserForm_Activate()
Dim Conteo As Long
Dim nFilas As Long
Dim nColumnas As Long
Dim f As Long
Dim c As Long
Dim Porcentaje As Double



imgLock.Visible = False
frame_Lock.Visible = False
btn_Ingresar.Visible = False
btn_Salir.Visible = False


    Conteo = 1
    nFilas = 700
    nColumnas = 700

        For f = 1 To nFilas
            For c = 1 To nColumnas
                Conteo = Conteo + 1
            Next c
                Porcentaje = Conteo / (nFilas * nColumnas)
                Me.Caption = Format(Porcentaje, "0%")
                Me.Label1.Width = Porcentaje * Me.frame_ProgressBar.Width
                DoEvents
        Next f


imgLock.Visible = True
frame_Lock.Visible = True
btn_Ingresar.Visible = True
btn_Salir.Visible = True

frame_ProgressBar.Visible = False
imgLogo.Visible = False
lbl_Titulo.Visible = False
Me.Caption = "Gestor de Inventarios"
Me.Height = 140.25

End Sub

Private Sub UserForm_Initialize()
Hoja1.txt_UsuarioActual.Text = Empty
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        ThisWorkbook.Application.DisplayAlerts = False
        Application.ActiveWorkbook.Close
        Unload Me
    End If
End Sub



