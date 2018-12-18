option explicit

private pId as integer
private pIdPedidoDeCompra as integer
private pIdproducto as integer
private pCantidad as double
private pCostoUnitario as currency
private pFechaRecepcion as date
private pPublicadoEnInventario as boolean
private pIdInventario as integer

Public Property let Id (Value as integer)
    pId = Value
End Property

Public Property get Id() as integer
    Id = pId
End Property

Public Property let IdPedidoCompra (Value as integer)
    pIdPedidoCompra = Value
End Property

Public Property get IdPedidoCompra() as integer
    IdPedidoCompra = pIdPedidoCompra
End Property

Public Property let Idproducto (Value as integer)
    pIdproducto = Value
End Property

Public Property get Idproducto() as integer
    Idproducto = pIdproducto
End Property

Public Property let Cantidad (Value as double)
    pCantidad = Value
End Property

Public Property get Cantidad() as double
    Cantidad = pCantidad
End Property

Public Property let CostoUnitario (Value as currency)
    pCostoUnitario = Value
End Property

Public Property get CostoUnitario() as currency
    CostoUnitario = pCostoUnitario
End Property

Public Property let FechaRecepcion (Value as date)
    pFechaRecepcion = Value
End Property

Public Property get FechaRecepcion() as date
    FechaRecepcion = pFechaRecepcion
End Property

Public Property let PublicadoEnInventario (Value as boolean)
    pPublicadoEnInventario = Value
End Property

Public Property get PublicadoEnInventario() as boolean
    PublicadoEnInventario = pPublicadoEnInventario
End Property

Public Property let IdInventario (Value as integer)
    pIdInventario = Value
End Property

Public Property get IdInventario() as integer
    IdInventario = pIdInventario
End Property

Public Sub RegistrarDetallesDePedidosDeCompra()
'fila vacia por debajo del último registro del listado
    Hoja8.Select 'detalles_pedido_compra
    Range("A" & Rows.Count).End(xlUp).Offset(1, 0).Select
    ActiveCell.value = Me.Id
    ActiveCell.Offset(0, 1).value = Me.IdPedidoCompra
    ActiveCell.Offset(0, 2).value = Me.Idproducto
    ActiveCell.Offset(0, 3).value = Me.Cantidad
    ActiveCell.Offset(0, 4).value = Me.CostoUnitario
    ActiveCell.Offset(0, 5).value = Me.FechaRecepcion
    ActiveCell.Offset(0, 6).value = Me.PublicadoEnInventario
    ActiveCell.Offset(0, 7).value = Me.IdInventario

End Sub