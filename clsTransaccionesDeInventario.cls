option explicit

private pIdTransaccionesDeInventario as integer
private pIdProducto as integer
private pIdPedidoCompra as integer
private pIdPedidoCliente as integer
private pTipoDeTransaccion as integer
private pFechaDeCreacionDeTransaccion as date
private pFechaDeModificacionDeInventrio as date
private pCantidad as double
private pComentarios as string
private pIdUbicacion as integer

Public Property let IdTransaccionesDeInventario (Value as integer)
    pIdTransaccionesDeInventario = Value
End Property

Public Property get IdTransaccionesDeInventario() as integer
    IdTransaccionesDeInventario = pIdTransaccionesDeInventario
End Property

Public Property let IdProducto (Value as integer)
    pIdProducto = Value
End Property

Public Property get IdProducto() as integer
    IdProducto = pIdProducto
End Property

Public Property let IdPedidoCompra (Value as integer)
    pIdPedidoCompra = Value
End Property

Public Property get IdPedidoCompra() as integer
    IdPedidoCompra = pIdPedidoCompra
End Property

Public Property let IdPedidoCliente (Value as integer)
    pIdPedidoCliente = Value
End Property

Public Property get IdPedidoCliente() as integer
    IdPedidoCliente = pIdPedidoCliente
End Property

Public Property let TipoDeTransaccion (Value as integer)
    pTipoDeTransaccion = Value
End Property

Public Property get TipoDeTransaccion() as integer
    TipoDeTransaccion = pTipoDeTransaccion
End Property

Public Property let FechaDeCreacionDeTransaccion (Value as date)
    pFechaDeCreacionDeTransaccion = Value
End Property

Public Property get FechaDeCreacionDeTransaccion() as date
    FechaDeCreacionDeTransaccion = pFechaDeCreacionDeTransaccion
End Property

Public Property let FechaDeModificacionDeTransaccion (Value as date)
    pFechaDeModificacionDeTransaccion = Value
End Property

Public Property get FechaDeModificacionDeTransaccion() as date
    FechaDeModificacionDeTransaccion = pFechaDeModificacionDeTransaccion
End Property

Public Property let Cantidad (Value as double)
    pCantidad = Value
End Property

Public Property get Cantidad() as double
    Cantidad = pCantidad
End Property

Public Property let Comentarios (Value as string)
    pComentarios = Value
End Property

Public Property get Comentarios() as string
    Comentarios = pComentarios
End Property

Public Property let IdUbicacion (Value as integer)
    pIdUbicacion = Value
End Property

Public Property get IdUbicacion() as integer
    IdUbicacion = pIdUbicacion
End Property

Public Sub RegistrarTransaccionesDeInventario()
'fila vacia por debajo del último registro del listado
    Hoja18.select 'transacciones_de_inventario
    Range("A" & Rows.Count).End(xlUp).Offset(1, 0).Select
    ActiveCell.value = Me.IdTransaccionesDeInventario
    ActiveCell.Offset(0, 1).value = Me.IdProducto
    ActiveCell.Offset(0, 2).value = Me.IdPedidoCompra
    ActiveCell.Offset(0, 3).value = Me.IdPedidoCliente
    ActiveCell.Offset(0, 4).value = Me.TipoDeTransaccion
    ActiveCell.Offset(0, 5).value = Me.FechaDeCreacionDeTransaccion
    ActiveCell.Offset(0, 6).value = Me.FechaDeModificacionDeTransaccion
    ActiveCell.Offset(0, 7).value = Me.Cantidad
    ActiveCell.Offset(0, 8).value = Me.Comentarios
    ActiveCell.Offset(0, 9).value = Me.IdUbicacion
End Sub