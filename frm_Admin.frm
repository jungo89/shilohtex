VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Admin 
   Caption         =   "Administrador"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3630
   OleObjectBlob   =   "frm_Admin.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_Admin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
Call MostrarHojas
End Sub
Private Sub CommandButton2_Click()
Call OcultarHojas
End Sub
Private Sub CommandButton3_Click()
frm_NuevoUsuario.Show
End Sub
Private Sub CommandButton4_Click()
frm_EliminarUsuario.Show
End Sub

