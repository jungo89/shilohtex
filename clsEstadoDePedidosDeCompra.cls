option explicit 

private pIdDeSituacion as integer
private pEstado as string

Public Property let IdDeSituacion (Value as integer)
    pIdDeSituacion = Value
End Property

Public Property get IdDeSituacion() as integer
    IdDeSituacion = pIdDeSituacion
End Property

Public Property let Estado (Value as string)
    pEstado = Value
End Property

Public Property get Estado() as string
    Estado = pEstado
End Property

Public Sub RegistrarEstadoDePedidosDeCompra()
'fila vacia por debajo del último registro del listado
    Hoja13.Select 'estado_de_pedidos_de_compra
    Range("A" & Rows.Count).End(xlUp).Offset(1, 0).Select
    ActiveCell.value = Me.IdDeSituacion
    ActiveCell.Offset(0, 1).value = Me.Estado

End Sub