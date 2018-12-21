VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "EXCELeINFO - Reporte de ventas"
   ClientHeight    =   3120
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   5520
   OleObjectBlob   =   "UserForm3.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'EXCELeINFO
'MVP Sergio Alejandro Campos
'http://www.exceleinfo.com
'https://www.youtube.com/user/sergioacamposh
'http://blogs.itpro.es/exceleinfo

Private Sub CommandButton1_Click()

Control1 = Me.TextBox1.Name
frmCalendario.Show

End Sub

Private Sub CommandButton2_Click()

Control1 = Me.TextBox2.Name
frmCalendario.Show

End Sub

Private Sub CommandButton3_Click()
Dim Conn As ADODB.Connection
Dim MiConexion
Dim Rs As ADODB.Recordset
Dim MiBase As String
Dim Query As String
Dim i, j
Dim Fecha1
Dim Fecha2
Dim Estado As String

MiBase = "MiBase.accdb"

Set Conn = New ADODB.Connection
MiConexion = Application.ThisWorkbook.Path & Application.PathSeparator & MiBase

With Conn
    .Provider = "Microsoft.ACE.OLEDB.12.0"
    .Open MiConexion
End With

Fecha1 = Me.TextBox1.Value
Fecha2 = Me.TextBox2.Value
Estado = Me.TextBox3.Value
Query = "SELECT * FROM Ventas WHERE [Fecha] >= #" & Fecha1 & "# AND [Fecha] <= #" & Fecha2 & "# AND [Estado o provincia] = '" & Estado & "' ORDER BY [Fecha]"

Set Rs = New ADODB.Recordset
Rs.CursorLocation = adUseServer
Rs.Open Source:=Query, _
ActiveConnection:=Conn

'Valir si la consulta devuelve resultados
If Rs.EOF And Rs.BOF Then
    'Borrar la conexión al Recordset
    Rs.Close
    Conn.Close
    'Borrar la memoria
    Set Rs = Nothing
    Set Conn = Nothing
    
    MsgBox "No hay resultados para la consulta", vbInformation, "EXCELeINFO"
    Sheets("Reporte").Range("A1").CurrentRegion.Clear
    Exit Sub
End If

Sheets("Reporte").Range("A1").CurrentRegion.Clear
For i = 0 To Rs.Fields.Count - 1

    Cells(1, i + 1).Value = Rs.Fields(i).Name

Next i

Sheets("Reporte").Range("A2").CopyFromRecordset Rs

'Cerrar la conexión
Rs.Close
Conn.Close
Set Rs = Nothing
Set Conn = Nothing
End Sub

Private Sub CommandButton4_Click()
Unload Me
End Sub

