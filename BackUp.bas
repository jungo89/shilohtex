Attribute VB_Name = "BackUp"
Option Explicit

Sub CopiaSeguridad()
Dim sRuta As String
Dim sNombreFolder As String
Dim sSeparador As String
Dim sRutaDestino As String
Dim sBackUp As String

On Error GoTo Salir

            sRuta = Application.ActiveWorkbook.Path
            sSeparador = Application.PathSeparator
            
            sBackUp = "Gestor_de_Inventarios_" & CStr(Format(Date, "yyyymmdd")) _
            & "_" & CStr(Format(Time, "hh-mm-ss")) & ".xlsm"
            
            sNombreFolder = "BackUp_" & CStr(Format(Date, "yyyy-mm"))
            
            sRutaDestino = sRuta & sSeparador & sNombreFolder
            
            If Dir(sRutaDestino, vbDirectory) = Empty Then
                MkDir (sRutaDestino)
            End If
            
            Application.ActiveWorkbook.SaveCopyAs FileName:=sRutaDestino & sSeparador & sBackUp

Salir:
 If Err <> 0 Then
    MsgBox Err.Description, vbExclamation, "Gestor de Inventarios"
 End If
            
End Sub
