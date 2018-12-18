option explicit

private pId as integer
private pNombreDelTipo as string

Public Property let Id (Value as integer)
    pId = Value
End Property

Public Property get Id() as integer
    Id = pId
End Property

Public Property let NombreDelTipo (Value as string)
    pNombreDelTipo = Value
End Property

Public Property get NombreDelTipo() as string
    NombreDelTipo = pNombreDelTipo
End Property

Public Sub RegistrarTipoDeTransaccionesDeInvent()
'fila vacia por debajo del último registro del listado
    Hoja17.select 'tipo_de_transacciones_de_invent
    Range("A" & Rows.Count).End(xlUp).Offset(1, 0).Select
    ActiveCell.value = Me.Id
    ActiveCell.Offset(0, 1).value = Me.NombreDelTipo
 
End Sub