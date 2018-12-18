
option explicit

private pId as integer
private pTipoDocumento as string
private pDocumento as string
private pNombreContacto as string
private pNit as string
private pRazonSocial as string
private pComercio as string
private pNicho as string
private pSegmentacion as string
private pProducto as string
private pDistribucion as string
private pCupo as currency
private pCredito as currency
private pSaldo as currency
private pCategoria as string

'propiedad de escritura
Public Property Let Id(value As integer)
    pId = value
End Property

'propiedad de lectura
Public Property Get Id() As Integer
    Id = pId
End Property

public Property let TipoDocumento(value as string)
    pTipoDocumento = value
End property

public property Get TipoDocumento () as string
    TipoDocumento = pTipoDocumento
End property

public property let Documento(value as string)
    pDocumento = value
End property   

public property get Documento() as string
    Documento = pDocumento
End property

public property let NombreContacto(value as string)
    pNombreContacto = value
End property

public property get NombreContacto() as string
    NombreContacto = pNombreContacto
End property

public property let Nit(value as string)
    pNit = value
End Property

public property get Nit() as string
    Nit = pNit
End Property

public property let RazonSocial(value as string)
    pRazonSocial = value
End Property

public property get RazonSocial() as string
    RazonSocial = pRazonSocial
End Property

public property let Comercio(value as string)
    pComercio =  value
End Property

public property get Comercio() as string
    Comercio = pComercio
End property

public property let Nicho(value as string)
    pNicho = value
End Property

public property get Nicho() as string
    Nicho = pNicho
End Property

public property let Segmentacion(value as string)
    pSegmentacion = value
End property

public property get Segmentacion() as string
    Segmentacion = pSegmentacion
End property

public property let Producto(value as string)
    pProducto = value
end property

public property get Producto() as string
    Producto = pProducto
End property

public property let Distribucion(value as string)
    pDistribucion = value
End property

public property get Distribucion() as string
    Distribucion = pDistribucion
End Property

Public property let Cupo( value as currency)
    pCupo = value
end property

public property get Cupo() as currency
    Cupo = pCupo
End property

public property let Credito(value as currency)
    pCredito = value
End property

public property get Credito ()  as currency
    Credito = pCredito
End property

public property let Saldo(value as currency)
    pSaldo = value
End property

public property get Saldo() as currency
    Saldo = pSaldo
End property

Public Property let Categoria (Value as string)
    pCategoria = Value
End Property

Public Property get Categoria() as string
    Categoria = pCategoria
End Property

Public Sub RegistrarCliente()
'fila vacia por debajo del último registro del listado
    Hoja1.select 'Clientes
    Range("A" & Rows.Count).End(xlUp).Offset(1, 0).Select
    ActiveCell.value = Me.Id
    ActiveCell.Offset(0, 1).value = Me.TipoDocumento
    ActiveCell.Offset(0, 2).value = Me.Documento
    ActiveCell.Offset(0, 3).value = Me.NombreContacto
    ActiveCell.Offset(0, 4).value = Me.Nit
    ActiveCell.Offset(0, 5).value = Me.RazonSocial
    ActiveCell.Offset(0, 6).value = Me.Comercio
    ActiveCell.Offset(0, 7).value = Me.Nicho
    ActiveCell.Offset(0, 8).value = Me.Segmentacion
    ActiveCell.Offset(0, 9).value = Me.Producto
    ActiveCell.Offset(0, 10).value = Me.Distribucion
    ActiveCell.Offset(0, 11).value = Me.Cupo
    ActiveCell.Offset(0, 12).value= Me.Credito
    ActiveCell.Offset(0, 13).value = Me.Saldo
    ActiveCell.Offset(0, 14).value = Me.Categoria
End Sub