option explicit

private pId as integer
private pIdPedido as integer
private pFecha as date
private pFechaVencimiento as date
private pReteFuente as currency
private pReteica as currency
private pIva as currency
private pImporteDebido as currency
private pInteres as double

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

Public Property let Fecha (Value as date)
    pFecha = Value
End Property

Public Property get Fecha() as date
    Fecha = pFecha
End Property

Public Property let FechaVencimiento (Value as date)
    pFechaVencimiento = Value
End Property

Public Property get FechaVencimiento() as date
    FechaVencimiento = pFechaVencimiento
End Property

Public Property let ReteFuente (Value as currency)
    pReteFuente = Value
End Property

Public Property get ReteFuente() as currency
    ReteFuente = pReteFuente
End Property

Public Property let ReteIca (Value as currency)
    pReteIca = Value
End Property

Public Property get ReteIca() as currency
    ReteIca = pReteIca
End Property

Public Property let Iva (Value as currency)
    pIva = Value
End Property

Public Property get Iva() as currency
    Iva = pIva
End Property

Public Property let ImporteDebido (Value as currency)
    pImporteDebido = Value
End Property

Public Property get ImporteDebido() as currency
    ImporteDebido = pImporteDebido
End Property

Public Property let Interes (Value as double)
    pInteres = Value
End Property

Public Property get Interes() as double
    Interes = pInteres
End Property

Public Sub RegistrarFactura()
'fila vacia por debajo del último registro del listado
    Hoja14.Select 'factura
    Range("A" & Rows.Count).End(xlUp).Offset(1, 0).Select
    ActiveCell.value = Me.Id
    ActiveCell.Offset(0, 1).value = Me.IdPedido
    ActiveCell.Offset(0, 2).value = Me.Fecha
    ActiveCell.Offset(0, 3).value = Me.FechaVencimiento
    ActiveCell.Offset(0, 4).value = Me.ReteFuente
    ActiveCell.Offset(0, 5).value = Me.ReteIca
    ActiveCell.Offset(0, 6).value = Me.Iva
    ActiveCell.Offset(0, 7).value = Me.ImporteDebido
    ActiveCell.Offset(0, 8).value = Me.Interes
End Sub