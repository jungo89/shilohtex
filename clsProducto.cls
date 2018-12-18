option explicit

private pId as integer
private pIdProveedor as integer
private pProducto as string
private pColor as string
private pMedida as string
private pCantidad as double
private pPresentacion as string
private pCosto as currency
private pUtilidad as double
private pVenta as currency
private pIva as double
private pVentaIva as currency
private pCategoria as string
private pDisponible as double
private pStock as double
private pPedir as double

public property let Id (value as integer)
    pId = value
End property

public property get Id () as integer
    Id = pId
End  property

public property let IdProveedor(value as integer)
    pIdProveedor = value
End property

public property get IdProveedor () as integer
    IdProveedor = pIdProveedor
End property

public property let Producto(value as string)
    pProducto = value
End property

public property get Producto() as string
    Producto = pProducto
End property

public property let Color(value as string)
    pColor = value
End property

public property get Color() as string
    Color = pColor
End property

public property let Medida(value as string)
    pMedida = value
End property

public property get Medida() as string
    Medida = pMedida
End property

public property let Cantidad(value as double)
    pCantidad = value
End property

public property get Cantidad() as double
    Cantidad = pCantidad
End property

public property let Presentacion(value as string)
    pPresentacion = value
End property

public property get Presentacion() as string
    Presentacion = pPresentacion
End property

public property let Costo(value as currency)
    pCosto = value
end property

public property get Costo() as currency
    Costo = pCosto
End property

Public Property let Utilidad (Value as double)
    pUtilidad = Value
End Property

Public Property get Utilidad() as double
    Utilidad = pUtilidad
End Property

Public Property let Venta (Value as currency)
    pVenta = Value
End Property

Public Property get Venta() as currency
    Venta = pVenta
End Property

Public Property let Iva (Value as double)
    pIva = Value
End Property

Public Property get Iva() as double
    Iva = pIva
End Property

Public Property let VentaIva (Value as currency)
    pVentaIva = Value
End Property

Public Property get VentaIva() as currency
    VentaIva = pVentaIva
End Property

Public Property let Categoria (Value as string)
    pCategoria = Value
End Property

Public Property get Categoria() as string
    Categoria = pCategoria
End Property

Public Property let Disponible (Value as double)
    pDisponible = Value
End Property

Public Property get Disponible() as double
    Disponible = pDisponible
End Property

Public Property let Stock (Value as double)
    pStock = Value
End Property

Public Property get Stock() as double
    Stock = pStock
End Property

Public Property let Pedir (Value as double)
    pPedir = Value
End Property

Public Property get Pedir() as double
    Pedir = pPedir
End Property

Public Sub RegistrarProducto()
'fila vacia por debajo del último registro del listado
    Hoja2.select 'productos
    Range("A" & Rows.Count).End(xlUp).Offset(1, 0).Select
    ActiveCell.value = Me.Id
    ActiveCell.Offset(0, 1).value = Me.IdProveedor
    ActiveCell.Offset(0, 2).value = Me.Producto
    ActiveCell.Offset(0, 3).value = Me.Color
    ActiveCell.Offset(0, 4).value = Me.Medida
    ActiveCell.Offset(0, 5).value = Me.Cantidad
    ActiveCell.Offset(0, 6).value = Me.Presentacion
    ActiveCell.Offset(0, 7).value = Me.Costo
    ActiveCell.Offset(0, 8).value = Me.Utilidad
    ActiveCell.Offset(0, 9).value = Me.Venta
    ActiveCell.Offset(0, 10).value = Me.Iva
    ActiveCell.Offset(0, 11).value = Me.VentaIva
    ActiveCell.Offset(0, 12).value= Me.Categoria
    ActiveCell.Offset(0, 13).value = Me.Disponible
    ActiveCell.Offset(0, 14).value = Me.Stock
    ActiveCell.Offset(0, 15).value = Me.Pedir
   
End Sub

