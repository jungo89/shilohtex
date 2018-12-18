Attribute VB_Name = "llamarControlCalendario"
Option Explicit
Public banderaCalendario As Long

Public Function LanzarCalendario(CualquierFormulario As Object, xTextBox As String)
Dim xCtrl As control

     Load frmCalendario
    
    For Each xCtrl In CualquierFormulario.Controls
        If xCtrl.Name = xTextBox Then
            frmCalendario.StartUpPosition = 0
            frmCalendario.Left = CualquierFormulario.Left + xCtrl.Left + 5
            frmCalendario.Top = CualquierFormulario.Top + xCtrl.Top + xCtrl.Height + 25
        End If
    Next
    
    frmCalendario.Show

End Function

Public Function InsertarFecha(Fecha As Date)
    Select Case banderaCalendario
        Case 1
            frm_Factura.txtFecha.Text = Fecha
                
        Case 2
            frm_fCompras.txtFecha.Text = Fecha
            
        Case 3
            frm_Consulta2.txtFecha1.Text = Fecha
        
        Case 4
            frm_Consulta2.txtFecha2.Text = Fecha
            
        Case 5
            frm_Transferencias.txt_FechaSal.Text = Fecha
            
        Case Else
            MsgBox "La petición solicitada, aún no se ha establecido dentro de la declaración SELECT CASE", vbCritical
    End Select
End Function
