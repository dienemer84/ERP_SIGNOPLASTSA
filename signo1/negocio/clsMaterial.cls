VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMaterial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Tipo As TipoMaterial
Public Altura As Double

Private iKg As Double
Private iM2 As Double
Private iUn As Double
Private iMl As Double
Private vHistorico As Collection
Private mMoneda As clsMoneda
Private vValor As Double
Private vFechaValor As Date
Private vId As Long
Private vcodigo As String
Private vdescripcion As String
Private vespesor As Double
Private vpesoxunidad As Double
Private vunidad As Unidades

Public UnidadPedido As Unidades
Public UnidadCompra As Unidades
Public ValorCompra As Double

Private vEstado As EstadoMaterial
Private vAncho As Double
Private vLargo As Double
Private vCantidad As Double
Private mGrupo As clsGrupo
Private mAlmacen As clsAlmacen


Public StockMinimo As Double
Public PuntoReposicion As Double
Public Aprobado As Boolean

Public Property Let historico(nValor As Collection)
    Set vHistorico = nValor
End Property
Public Property Get historico() As Collection
    Set historico = vHistorico
End Property
Public Property Let Valor(nValor As Double)
    vValor = nValor
End Property
Public Property Let FechaValor(nValor As Date)
    vFechaValor = nValor
End Property
Public Property Let Id(nvalue As Long)
    vId = nvalue
End Property
Public Property Get Id() As Long
    Id = vId
End Property
Public Property Let moneda(nvalue As clsMoneda)
    Set mMoneda = nvalue
End Property
Public Property Get moneda() As clsMoneda
    Set moneda = mMoneda
End Property
Public Property Let codigo(nvalue As String)
    vcodigo = nvalue
End Property
Public Property Let descripcion(nvalue As String)
    vdescripcion = nvalue
End Property
Public Property Let Espesor(nvalue As Double)
    vespesor = nvalue
End Property
Public Property Let PesoXUnidad(nvalue As Double)
    vpesoxunidad = nvalue
End Property
Public Property Let unidad(nvalue As Unidades)
    vunidad = nvalue
End Property
Public Property Let estado(nvalue As EstadoMaterial)
    vEstado = nvalue
End Property
Public Property Let Ancho(nvalue As Double)
    vAncho = nvalue
End Property
Public Property Let Largo(nvalue As Double)
    vLargo = nvalue
End Property
Public Property Let Cantidad(nvalue As Double)
    vCantidad = nvalue
End Property
Public Property Let Grupo(nvalue As clsGrupo)
    Set mGrupo = nvalue
End Property
Public Property Let almacen(nvalue As clsAlmacen)
    Set mAlmacen = nvalue
End Property

Public Property Get FechaValor() As Date
    FechaValor = vFechaValor
End Property
Public Property Get Valor() As Double
    Valor = vValor
End Property
Public Property Get codigo() As String
    codigo = vcodigo
End Property
Public Property Get descripcion() As String
    descripcion = vdescripcion
End Property
Public Property Get Espesor() As Double
    Espesor = vespesor
End Property
Public Property Get PesoXUnidad() As Double
    PesoXUnidad = vpesoxunidad
End Property
Public Property Get unidad() As Unidades
    unidad = vunidad
End Property
Public Property Get estado() As EstadoMaterial
    estado = vEstado
End Property
Public Property Get Ancho() As Double
    Ancho = vAncho
End Property
Public Property Get Largo() As Double
    Largo = vLargo
End Property
Public Property Get Cantidad() As Double
    Cantidad = vCantidad
End Property
Public Property Get Grupo() As clsGrupo
    Set Grupo = mGrupo
End Property
Public Property Get almacen() As clsAlmacen
    Set almacen = mAlmacen
End Property

Public Property Get m2() As Double
    datosMaterial
    m2 = iM2
End Property
Public Property Get ML() As Double
    datosMaterial
    ML = iMl
End Property
Public Property Get UN() As Double
    datosMaterial
    UN = iUn
End Property
Public Property Get Kg() As Double
    datosMaterial
    Kg = iKg
End Property






Private Sub datosMaterial()
    If vunidad = Ml_ Then
        iKg = Math.Round(vpesoxunidad * (vLargo / 1000), 3)
        iMl = Math.Round(vLargo / 1000, 3)
        iM2 = 0
        iUn = 0
    ElseIf vunidad = kg_ Then
        iKg = Math.Round(vpesoxunidad, 3)
        iM2 = 0
        iMl = 0
        iUn = 0
    ElseIf vunidad = m2_ Then
        iKg = Math.Round(vpesoxunidad * ((vAncho * vLargo) / 1000000), 3)
        iMl = 0
        iM2 = Math.Round((vAncho * vLargo) / 1000000, 3)
        iUn = 0
    ElseIf vunidad = un_ Then
        iM2 = 0
        iMl = 0
        iKg = 0
        iUn = vCantidad
    End If
End Sub

Public Property Get Atributos() As Collection
    Dim attrs As New Collection

    If Me.Tipo = TM_UnidadKilo Then
        If Me.unidad = kg_ Then
            attrs.Add "Kg: " & Me.PesoXUnidad
        ElseIf Me.unidad = un_ Then
            attrs.Add "Uni: " & Me.PesoXUnidad
        ElseIf Me.unidad = litro_ Then
            attrs.Add "Litro: " & Me.PesoXUnidad
        End If
    Else
        attrs.Add "Largo: " & Me.Largo
        attrs.Add "Ancho: " & Me.Ancho
        attrs.Add "Espesor: " & Me.Espesor

        If Me.Tipo = TM_PerfilRectangular Then
            attrs.Add "Altura: " & Me.Altura
        End If
    End If

    Set Atributos = attrs
End Property
