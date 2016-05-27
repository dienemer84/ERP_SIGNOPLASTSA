VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRequeMateriales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private tmpEntrega As clsRequeEntregas
Private mProveedores As New Collection    '(of proveedores)
Private vObservaciones As String
Private mMaterial As clsMaterial
Private mEntregas As New Collection    '(of RequeEntregaItem)
Private vCantidad As Double
Private vLargo As Double
Private vAncho As Double
Private vId As Long
Private iKg As Double
Private iUn As Double
Private iM2 As Double
Private iMl As Double
Public estado As EstadoRequeCompra
Public RequeId As Long

Public Function TotalCantidad(unidad As Unidades) As Double
    If unidad = kg_ Then
        TotalCantidad = Kg
    ElseIf unidad = m2_ Then
        TotalCantidad = m2
    ElseIf unidad = Ml_ Then
        TotalCantidad = ML
    Else
        TotalCantidad = Cantidad
    End If
End Function
Public Property Let Largo(nvalue As Double)
    vLargo = nvalue
End Property
Public Property Let Ancho(nvalue As Double)
    vAncho = nvalue
End Property

Public Property Get Largo() As Double
    Largo = vLargo
End Property
Public Property Get Ancho() As Double
    Ancho = vAncho
End Property



Public Property Get Kg() As Double
    datosMaterial
    Kg = iKg * vCantidad
End Property
Public Property Get m2() As Double
    datosMaterial
    m2 = iM2 * vCantidad
End Property
Public Property Get ML() As Double
    datosMaterial
    ML = iMl * vCantidad
End Property
Public Property Get UN() As Double
    UN = iUn * vCantidad
End Property
'datosMaterial
Public Property Let ListaProveedores(nvalue As Collection)
    Set mProveedores = nvalue
End Property
Public Property Get ListaProveedores() As Collection
    Set ListaProveedores = mProveedores
End Property
Public Property Let id(nvalue As Long)
    vId = nvalue
End Property
Public Property Get id() As Long
    id = vId
End Property
Public Property Let Observaciones(nvalue As String)
    vObservaciones = nvalue
End Property
Public Property Let Material(nvalue As clsMaterial)
    Set mMaterial = nvalue
End Property
Public Property Let Cantidad(nvalue As Double)
    vCantidad = nvalue
End Property
Public Property Get Cantidad() As Double
    Cantidad = vCantidad
End Property
Public Property Let Entregas(nvalue As Collection)
    Set mEntregas = nvalue
End Property
Public Property Get Observaciones() As String
    Observaciones = vObservaciones
End Property
Public Property Get Material() As clsMaterial
    Set Material = mMaterial
End Property
Public Property Get Entregas() As Collection
    Set Entregas = mEntregas
End Property
Public Function TotalEntregas() As Double
    Dim tot_entregas As Double
    tot_entregas = 0
    Dim x As Long
    For x = 1 To mEntregas.count
        Set tmpEntrega = mEntregas(x)
        tot_entregas = tmpEntrega.Cantidad + tot_entregas
    Next x
    TotalEntregas = tot_entregas
End Function



Public Function FechasValidas() As Boolean
    Dim ok As Boolean
    ok = True
    Dim x As Long
    For x = 1 To mEntregas.count
        Set tmpEntrega = mEntregas(x)
        If tmpEntrega.FEcha < Date Then
            ok = False
            Exit For
        End If
    Next x
    FechasValidas = ok
End Function


Private Sub datosMaterial()
    If mMaterial.unidad = Unidades.Ml_ Then
        iKg = Math.Round(mMaterial.PesoXUnidad * (vLargo / 1000), 3)
        iMl = Math.Round(vLargo / 1000, 3)
        iM2 = 0
        iUn = 0
    ElseIf mMaterial.unidad = Unidades.kg_ Then
        iKg = Math.Round(mMaterial.PesoXUnidad, 3)
        iM2 = 0
        iMl = 0
        iUn = 0
    ElseIf mMaterial.unidad = Unidades.m2_ Then
        iKg = Math.Round(mMaterial.PesoXUnidad * ((vAncho * vLargo) / 1000000), 3)
        iMl = 0
        iM2 = Math.Round((vAncho * vLargo) / 1000000, 3)
        iUn = 0
    ElseIf mMaterial.unidad = Unidades.un_ Then
        iM2 = 0
        iMl = 0
        iKg = 0
        iUn = 1    'cantidad unitaria
    End If
End Sub


