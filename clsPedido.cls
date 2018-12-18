option explicit

private pIdDePedido as integer
private pIdDeEmpleado as integer
private pIdDeCliente as integer
private pIdDeTransportista as integer
private pIdDeSituacion as integer
private pFechaDePedido as date
private pFechaDeEnvio as date
private pNombreDeEnvio as string
private pDireccionDeEnvio as string
private pCiudadDestino as string
private pGastosDeEnvio as currency
private pImpuestos as currency
private pTipoDePago as integer
private pFechaDePago as date
private pPrioridad as string
private pNotas as string
private pEstadoDeImpuestos as integer
 
Public Property let IdDePedido (Value as integer)
    pIdDePedido = Value
End Property

Public Property get IdDePedido() as integer
    IdDePedido = pIdDePedido
End Property

Public Property let IdDeEmpleado (Value as integer)
    pIdDeEmpleado = Value
End Property

Public Property get IdDeEmpleado() as integer
    IdDeEmpleado = pIdDeEmpleado
End Property

Public Property let IdDeCliente (Value as integer)
    pIdDeCliente = Value
End Property

Public Property get IdDeCliente() as integer
    IdDeCliente = pIdDeCliente
End Property

Public Property let IdDeTransportista (Value as integer)
    pIdDeTransportista = Value
End Property

Public Property get IdDeTransportista() as integer
    IdDeTransportista = pIdDeTransportista
End Property

Public Property let IdDeSituacion (Value as integer)
    pIdDeSituacion = Value
End Property

Public Property get IdDeSituacion() as integer
    IdDeSituacion = pIdDeSituacion
End Property

Public Property let FechaDePedido (Value as date)
    pFechaDePedido = Value
End Property

Public Property get FechaDePedido() as date
    FechaDePedido = pFechaDePedido
End Property

Public Property let FechaDeEnvio (Value as date)
    pFechaDeEnvio = Value
End Property

Public Property get FechaDeEnvio() as date
    FechaDeEnvio = pFechaDeEnvio
End Property

Public Property let NombreDeEnvio (Value as string)
    pNombreDeEnvio = Value
End Property

Public Property get NombreDeEnvio() as string
    NombreDeEnvio = pNombreDeEnvio
End Property

Public Property let DireccionDeEnvio (Value as string)
    pDireccionDeEnvio = Value
End Property

Public Property get DireccionDeEnvio() as string
    DireccionDeEnvio = pDireccionDeEnvio
End Property

Public Property let CiudadDestino (Value as string)
    pCiudadDestino = Value
End Property

Public Property get CiudadDestino() as string
    CiudadDestino = pCiudadDestino
End Property

Public Property let GastosDeEnvio (Value as currency)
    pGastosDeEnvio = Value
End Property

Public Property get GastosDeEnvio() as currency
    GastosDeEnvio = pGastosDeEnvio
End Property

Public Property let Impuestos (Value as currency)
    pImpuestos = Value
End Property

Public Property get Impuestos() as currency
    Impuestos = pImpuestos
End Property

Public Property let TipoDePago (Value as integer)
    pTipoDePago = Value
End Property

Public Property get TipoDePago() as integer
    TipoDePago = pTipoDePago
End Property

Public Property let FechaDePago (Value as date)
    pFechaDePago = Value
End Property

Public Property get FechaDePago() as date
    FechaDePago = pFechaDePago
End Property

Public Property let Prioridad (Value as string)
    pPrioridad = Value
End Property

Public Property get Prioridad() as string
    Prioridad = pPrioridad
End Property

Public Property let Notas (Value as string)
    pNotas = Value
End Property

Public Property get Notas() as string
    Notas = pNotas
End Property

Public Property let EstadoDeImpuestos (Value as integer)
    pEstadoDeImpuestos = Value
End Property

Public Property get EstadoDeImpuestos() as integer
    EstadoDeImpuestos = pEstadoDeImpuestos
End Property

Public Sub RegistrarPedido()
'fila vacia por debajo del último registro del listado
    Hoja15.Select 'pedidos
    Range("A" & Rows.Count).End(xlUp).Offset(1, 0).Select
    ActiveCell.value = Me.IdDePedido
    ActiveCell.Offset(0, 1).value = Me.IdDeEmpleado
    ActiveCell.Offset(0, 2).value = Me.IdDeCliente
    ActiveCell.Offset(0, 3).value = Me.IdDeTransportista
    ActiveCell.Offset(0, 4).value = Me.pIdDeSituacion
    ActiveCell.Offset(0, 5).value = Me.FechaDePedido
    ActiveCell.Offset(0, 6).value = Me.FechaDeEnvio
    ActiveCell.Offset(0, 7).value = Me.NombreDeEnvio
    ActiveCell.Offset(0, 8).value = Me.DireccionDeEnvio
    ActiveCell.Offset(0, 9).value = Me.CiudadDestino
    ActiveCell.Offset(0, 10).value = Me.GastosDeEnvio
    ActiveCell.Offset(0, 11).value = Me.Impuestos
    ActiveCell.Offset(0, 12).value = Me.TipoDePago
    ActiveCell.Offset(0, 13).value = Me.FechaDePago
    ActiveCell.Offset(0, 14).value = Me.Prioridad
    ActiveCell.Offset(0, 15).value = Me.Notas
    ActiveCell.Offset(0, 16).value = Me.EstadoDeImpuestos