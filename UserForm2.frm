VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "EXCELeINFO - Buscar registros"
   ClientHeight    =   3720
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   7470
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "UserForm2"
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
Dim Conn As ADODB.Connection
Dim MiConexion
Dim Rs As ADODB.Recordset
Dim MiBase As String
Dim Query As String
Dim i, j

MiBase = "MiBase.accdb"

Set Conn = New ADODB.Connection
MiConexion = Application.ThisWorkbook.Path & Application.PathSeparator & MiBase

With Conn
    .Provider = "Microsoft.ACE.OLEDB.12.0"
    .Open MiConexion
End With

Query = "SELECT * FROM MiTabla WHERE nombre LIKE '%" & Me.TextBox1.Value & "%'"

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
    Me.ListBox1.Clear
    Exit Sub
End If

'Asignar número de columnas
With Me.ListBox1
    .ColumnCount = Rs.Fields.Count
End With

'Recorrer el Recordset
Rs.MoveFirst
i = 1

With Me.ListBox1
    .Clear
    
    'Asignar los encabezados
    .AddItem
    
    For j = 0 To 4
        .List(0, j) = Rs.Fields(j).Name
    Next j
    
    Do
        .AddItem
        .List(i, 0) = Rs![ID]
        .List(i, 1) = Rs![Fecha]
        .List(i, 2) = Rs![Nombre]
        .List(i, 3) = Rs![Ventas]
        .List(i, 4) = Rs![Comentarios]
        i = i + 1
        Rs.MoveNext
        
    Loop Until Rs.EOF
End With

'Cerrar la conexión
Rs.Close
Conn.Close
Set Rs = Nothing
Set Conn = Nothing

End Sub

Private Sub CommandButton2_Click()
Unload Me
End Sub

Private Sub CommandButton3_Click()
Dim Conn As ADODB.Connection
Dim MiConexion
Dim Rs As ADODB.Recordset
Dim MiBase As String
Dim Query As String
Dim i, j
Dim Cuenta As Integer
Dim Numero As Integer
Dim ValorElegido As Integer

'Recorrer el listbox y detectar el item elegido
'''''''''''''''''''''''''''''''
Cuenta = Me.ListBox1.ListCount

'Validamos que haya un elemento seleccionado
For i = 0 To Cuenta - 1
    If Me.ListBox1.Selected(i) = True Then
        Numero = Numero + 1
    End If
Next i

If Numero = 0 Then MsgBox "Debes elegir un elemento", vbExclamation, "EXCELeINFO": Exit Sub

For j = 0 To Cuenta - 1
    If Me.ListBox1.Selected(j) = True Then
        If Me.ListBox1.ListIndex = 0 Then MsgBox "Encabezado!", vbCritical, "EXCELeINFO": Exit Sub
    ValorElegido = Me.ListBox1.List(j)
    'MsgBox ValorElegido
    End If
Next j

'''''''''''''''''''''''''''''''

MiBase = "MiBase.accdb"

Set Conn = New ADODB.Connection
MiConexion = Application.ThisWorkbook.Path & Application.PathSeparator & MiBase

With Conn
    .Provider = "Microsoft.ACE.OLEDB.12.0"
    .Open MiConexion
End With

Query = "DELETE * FROM MiTabla WHERE Id = " & ValorElegido

Set Rs = New ADODB.Recordset
Rs.CursorLocation = adUseServer
Rs.Open Source:=Query, _
ActiveConnection:=Conn

'Cerrar la conexión
'Rs.Close
Conn.Close
Set Rs = Nothing
Set Conn = Nothing

Call CommandButton1_Click

End Sub



























