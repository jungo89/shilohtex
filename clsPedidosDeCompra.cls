option explicit

private pId as integer
private pIdProveedor as integer
private pCreadoPor as integer
private pFechaDeEnvio as date
private pFechaDeCreacion as date
private pIdSituacion as integer
private pFechaPrevista as date
private pGastosEnvio as currency
private pImpuestos as currency
private pFechaDePago as date
private pImporteDePago as currency
private pModoDePago as string
private pNotas as string
private pAprobadoPor as integer
private pFechaDeAprobacion as date
private pEnviadoPor as integer

Public Property let Id (Value as integer)
    pId = Value
End Property

Public Property get Id() as integer
    Id = pId
End Property

Public Property let IdProveedor (Value as integer)
    pIdProveedor = Value
End Property

Public Property get IdProveedor() as integer
    IdProveedor = pIdProveedor
End Property

Public Property let CreadoPor (Value as integer)
    pCreadoPor = Value
End Property

Public Property get CreadoPor() as integer
    CreadoPor = pCreadoPor
End Property

Public Property let FechaDeEnvio (Value as date)
    pFechaDeEnvio = Value
End Property

Public Property get FechaDeEnvio() as date
    FechaDeEnvio = pFechaDeEnvio
End Property

Public Property let FechaDeCreacion (Value as date)
    pFechaDeCreacion = Value
End Property

Public Property get FechaDeCreacion() as date
    FechaDeCreacion = pFechaDeCreacion
End Property

Public Property let IdSituacion (Value as integer)
    pIdSituacion = Value
End Property

Public Property get IdSituacion() as integer
    IdSituacion = pIdSituacion
End Property

Public Property let FechaPrevista (Value as date)
    pFechaPrevista = Value
End Property

Public Property get FechaPrevista() as date
    FechaPrevista = pFechaPrevista
End Property

Public Property let GastosEnvio (Value as currency)
    pGastosEnvio = Value
End Property

Public Property get GastosEnvio() as currency
    GastosEnvio = pGastosEnvio
End Property

Public Property let Impuestos (Value as currency)
    pImpuestos = Value
End Property

Public Property get Impuestos() as currency
    Impuestos = pImpuestos
End Property

Public Property let FechaDePago (Value as date)
    pFechaDePago = Value
End Property

Public Property get FechaDePago() as date
    FechaDePago = pFechaDePago
End Property

Public Property let ImporteDePago (Value as currency)
    pImporteDePago = Value
End Property

Public Property get ImporteDePago() as currency
    ImporteDePago = pImporteDePago
End Property

Public Property let ModoDePago (Value as string)
    pModoDePago = Value
End Property

Public Property get ModoDePago() as string
    ModoDePago = pModoDePago
End Property

Public Property let Notas (Value as string)
    pNotas = Value
End Property

Public Property get Notas() as string
    Notas = pNotas
End Property

Public Property let AprobadoPor (Value as integer)
    pAprobadoPor = Value
End Property

Public Property get AprobadoPor() as integer
    AprobadoPor = pAprobadoPor
End Property

Public Property let FechaDeAprobacion (Value as date)
    pFechaDeAprobacion = Value
End Property

Public Property get FechaDeAprobacion() as date
    FechaDeAprobacion = pFechaDeAprobacion
End Property

Public Property let EnviadoPor (Value as integer)
    pEnviadoPor = Value
End Property

Public Property get EnviadoPor() as integer
    EnviadoPor = pEnviadoPor
End Property

Public Sub RegistrarPedidosDeCompra()
'fila vacia por debajo del último registro del listado
    Hoja16.select 'pedidos_de_compra
    Range("A" & Rows.Count).End(xlUp).Offset(1, 0).Select
    ActiveCell.value = Me.Id
    ActiveCell.Offset(0, 1).value = Me.IdProveedor
    ActiveCell.Offset(0, 2).value = Me.CreadoPor
    ActiveCell.Offset(0, 3).value = Me.FechaDeEnvio
    ActiveCell.Offset(0, 4).value = Me.FechaDeCreacion
    ActiveCell.Offset(0, 5).value = Me.IdSituacion
    ActiveCell.Offset(0, 6).value = Me.FechaPrevista
    ActiveCell.Offset(0, 7).value = Me.GastosEnvio
    ActiveCell.Offset(0, 8).value = Me.Impuestos
    ActiveCell.Offset(0, 9).value = Me.FechaDePago
    ActiveCell.Offset(0, 10).value = Me.ImporteDePago
    ActiveCell.Offset(0, 11).value = Me.ModoDePago
    ActiveCell.Offset(0, 12).value = Me.Notas
    ActiveCell.Offset(0, 13).value = Me.AprobadoPor
    ActiveCell.Offset(0, 14).value = Me.FechaDeAprobacion
    ActiveCell.Offset(0, 15).value = Me.EnviadoPor
End Sub