Public Function GetUltimoR(Hoja As Worksheet) As Integer
    GetUltimoR = GetNuevoR(Hoja) - 1
End Function

Public Function GetNuevoR(Hoja As Worksheet) As Integer
    
    Dim Fila As Long
    Fila = 2
    
    Do While Hoja.Cells(Fila, 1) <> ""
        Fila = Fila + 1
    Loop
    
    GetNuevoR = Fila
    
End Function

Public Sub Agregar(cmbBox As ComboBox, sItem As String)

'Agrega los item únicos y en orden alfabético

For i = 0 To cmbBox.ListCount - 1
    Select Case StrComp(cmbBox.List(i), sItem, vbTextCompare)
        Case 0: Exit Sub 'ya existe en el combo y ya no lo agrega
        Case 1: cmbBox.AddItem sItem, i: Exit Sub 'Es menor, lo agrega antes del comparado
    End Select
Next

cmbBox.AddItem sItem 'Es mayor lo agrega al final

End Sub
