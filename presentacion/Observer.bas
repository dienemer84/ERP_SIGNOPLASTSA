Attribute VB_Name = "Observer"
Private oListaFCProveedor As Collection
Private oListaProveedores As Collection
Private oListaMateriales As Collection
Private oListaReques As Collection
'Private oListaClientes As Collection
Private oListaPO As Collection


Private oListaPresupuestos_lista As Collection
Private oListaPresupuestos_grilla As GridEX
Private oListaPresupuestos As Boolean





Public Property Let ListaPO(nValue As Collection)
    Set oListaClientes = nValue
End Property
Public Property Get ListaPO() As Collection
    Set ListaPO = oListaPO
End Property



Public Property Let ListaProveedores(nValue As Collection)
    Set oListaProveedores = nValue
End Property
Public Property Get ListaProveedores() As Collection
    Set ListaProveedores = oListaProveedores
End Property
Public Property Let ListaReques(nValue As Collection)
    Set oListaReques = nValue
End Property
Public Property Get ListaReques() As Collection
    Set ListaReques = oListaReques
End Property

Public Property Let ListaMateriales(nValue As Collection)
    Set oListaMateriales = nValue
End Property
Public Property Get ListaMateriales() As Collection
    Set ListaMateriales = oListaMateriales
End Property


Public Property Let ListaFacturasProveedor(nValue As Collection)
    Set oListaFCProveedor = nValue
End Property

Public Property Get ListaFacturasProveedor() As Collection
    Set ListaFacturaProveedor = oListaFCProveedor
End Property

