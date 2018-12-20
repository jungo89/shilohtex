option explicit

private pId as integer
private pIdPedido as integer
private pIdProducto as integer
private pCantidad as double
private pPrecio as currency
private pFlete as currency
private pIdSituacion as integer
private pFechaAsignacion as date
private pIdPedidoCompra as integer
private pIdDeInventario  as integer

Public Property let Id (Value as integer)
    pId = Value
End Property

Public Property get Id() as integer
    Id = pId
End Property

Public Property let IdPedido (Value as integer)
    pIdPedido = Value
End Property

Public Property get IdPedido() as integer
    IdPedido = pIdPedido
End Property

Public Property let IdProducto (Value as integer)
    pIdProducto = Value
End Property

Public Property get IdProducto() as integer
    IdProducto = pIdProducto
End Property

Public Property let Cantidad (Value as double)
    pCantidad = Value
End Property

Public Property get Cantidad() as double
    Cantidad = pCantidad
End Property

Public Property let Precio (Value as currency)
    pPrecio = Value
End Property

Public Property get Precio() as currency
    Precio = pPrecio
End Property

Public Property let Flete (Value as currency)
    pFlete = Value
End Property

Public Property get Flete() as currency
    Flete = pFlete
End Property

Public Property let IdSituacion (Value as integer)
    pIdSituacion = Value
End Property

Public Property get IdSituacion() as integer
    IdSituacion = pIdSituacion
End Property

Public Property let FechaAsignacion (Value as date)
    pFechaAsignacion = Value
End Property

Public Property get FechaAsignacion() as date
    FechaAsignacion = pFechaAsignacion
End Property

Public Property let IdPedidoCompra (Value as integer)
    pIdPedidoCompra = Value
End Property

Public Property get IdPedidoCompra() as integer
    IdPedidoCompra = pIdPedidoCompra
End Property

Public Property let IdInventario (Value as integer)
    pIdInventario = Value
End Property

Public Property get IdInventario() as integer
    IdInventario = pIdInventario
End Property

Public Sub RegistrarDetallesDePedido()
'fila vacia por debajo del último registro del listado
    Hoja7.select 'detalles_pedido
    Range("A" & Rows.Count).End(xlUp).Offset(1, 0).Select
    ActiveCell.value = Me.Id
    ActiveCell.Offset(0, 1).value = Me.IdPedido
    ActiveCell.Offset(0, 2).value = Me.IdProducto
    ActiveCell.Offset(0, 3).value = Me.Cantidad
    ActiveCell.Offset(0, 4).value = Me.Precio
    ActiveCell.Offset(0, 5).value = Me.Flete
    ActiveCell.Offset(0, 6).value = Me.IdSituacion
    ActiveCell.Offset(0, 7).value = Me.FechaAsignacion
    ActiveCell.Offset(0, 8).value = Me.IdPedidoCompra
    ActiveCell.Offset(0, 9).value = Me.Id
End Sub