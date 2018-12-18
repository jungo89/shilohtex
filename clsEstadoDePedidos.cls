option explicit

private pIdDeSituacion as integer
private pNombreDeEstado as string

Public Property let IdDeSituacion (Value as integer)
    pIdDeSituacion = Value
End Property

Public Property get IdDeSituacion() as integer
    IdDeSituacion = pIdDeSituacion
End Property

Public Property let NombreDeEstado (Value as string)
    pNombreDeEstado = Value
End Property

Public Property get NombreDeEstado() as string
    NombreDeEstado = pNombreDeEstado
End Property

Public Sub RegistrarEstadoDePedidos()
'fila vacia por debajo del último registro del listado
    Hoja12.Select 'estado_de_pedidos
    Range("A" & Rows.Count).End(xlUp).Offset(1, 0).Select
    ActiveCell.value = Me.IdDeSituacion
    ActiveCell.Offset(0, 1).value = Me.NombreDeEstado
 
End Sub