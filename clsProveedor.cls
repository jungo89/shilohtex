option explicit

private pId as integer
private pNombre  as string
private pTipoDocumento  as string
private pDocumento as string
private pRazonSocial as string
private pNit as string
private pFormaPago as string

Public Property let Id (Value as integer)
    pId = Value
End Property

Public Property get Id() as integer
    Id = pId
End Property

Public Property let Nombre (Value as string)
    pNombre = Value
End Property

Public Property get Nombre() as string
    Nombre = pNombre
End Property

Public Property let TipoDocumento (Value as string)
    pTipoDocumento = Value
End Property

Public Property get TipoDocumento() as string
    TipoDocumento = pTipoDocumento
End Property

Public Property let Documento (Value as string)
    pDocumento = Value
End Property

Public Property get Documento() as string
    Documento = pDocumento
End Property

Public Property let RazonSocial (Value as string)
    pRazonSocial = Value
End Property

Public Property get RazonSocial() as string
    RazonSocial = pRazonSocial
End Property

Public Property let Nit (Value as string)
    pNit = Value
End Property

Public Property get Nit() as string
    Nit = pNit
End Property

Public Property let FormaPago (Value as string)
    pFormaPago = Value
End Property

Public Property get FormaPago() as string
    FormaPago = pFormaPago
End Property

Public Sub RegistrarProveedor()
'fila vacia por debajo del último registro del listado
    Hoja4.select 'proveedores
    Range("A" & Rows.Count).End(xlUp).Offset(1, 0).Select
    ActiveCell.value = Me.Id
    ActiveCell.Offset(0, 1).value = Me.Id
    ActiveCell.Offset(0, 2).value = Me.Nombre
    ActiveCell.Offset(0, 3).value = Me.TipoDocumento
    ActiveCell.Offset(0, 4).value = Me.Documento
    ActiveCell.Offset(0, 5).value = Me.RazonSocial
    ActiveCell.Offset(0, 6).value = Me.Nit
    ActiveCell.Offset(0, 7).value = Me.FormaPago
End Sub