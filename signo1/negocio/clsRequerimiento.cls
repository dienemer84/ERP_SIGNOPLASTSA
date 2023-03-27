VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRequerimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vId As Long
Private vDestinoOT As Long
Private vTipo As destino
Private mUsuario_creador As clsUsuario
Private mUsuario_aprobador As clsUsuario
Private mSector As clsSector
Private vEstado As EstadoRequeCompra
Private vFechaCreado As Date
Private mMateriales As New Collection    'los detalles del req, cada detalle tiene un material
'Private mConceptos As Collection
Private vGuardado As Date
Public Property Let Guardado(nvalue As Date)
    vGuardado = nvalue
End Property
Public Property Get Guardado() As Date
    Guardado = vGuardado
End Property
'Public Property Let Conceptos(nValue As Collection)
'Set mConceptos = nValue
'End Property
'Public Property Get Conceptos() As Collection
'Set Conceptos = mConceptos
'End Property
Public Property Let Materiales(nvalue As Collection)
    Set mMateriales = nvalue
End Property
Public Property Get Materiales() As Collection
    Set Materiales = mMateriales
End Property
Public Property Let Tipo(nvalue As destino)
    vTipo = nvalue
End Property
Public Property Get Tipo() As destino
    Tipo = vTipo
End Property
Public Property Let DestinoOT(nvalue As Long)
    vDestinoOT = nvalue
End Property
Public Property Get DestinoOT() As Long
    DestinoOT = vDestinoOT
End Property
Public Property Let Id(nvalue As Long)
    vId = nvalue
End Property
Public Property Get Id() As Long
    Id = vId
End Property
Public Property Let Sector(nvalue As clsSector)
    Set mSector = nvalue
End Property
Public Property Let estado(nvalue As EstadoRequeCompra)
    vEstado = nvalue
End Property
Public Property Let fechaCreado(nvalue As Date)
    vFechaCreado = nvalue
End Property
Public Property Get Sector() As clsSector
    Set Sector = mSector
End Property
Public Property Get estado() As EstadoRequeCompra
    estado = vEstado
End Property
Public Property Get fechaCreado() As Date
    fechaCreado = vFechaCreado
End Property
Public Property Let Usuario_creador(nvalue As clsUsuario)
    Set mUsuario_creador = nvalue
End Property
Public Property Get Usuario_creador() As clsUsuario
    Set Usuario_creador = mUsuario_creador
End Property
Public Property Let Usuario_aprobador(nvalue As clsUsuario)
    Set mUsuario_aprobador = nvalue
End Property
Public Property Get Usuario_aprobador() As clsUsuario
    Set Usuario_aprobador = mUsuario_aprobador
End Property
Public Function ValidarEntregas() As Boolean

    Dim tmpDetalle As clsRequeMateriales
    Dim ok As Boolean
    'primnero de materiales
    ok = True
    Dim i As Long

    For i = 1 To mMateriales.count
        Set tmpDetalle = mMateriales(i)
        ok = tmpDetalle.FechasValidas And (tmpDetalle.TotalEntregas = tmpDetalle.Cantidad)
        If Not ok Then Exit For
    Next i

    Set tmpDetalle = Nothing
    ValidarEntregas = ok
End Function

Public Function ValidarProveedores() As Boolean
    Dim ok As Boolean
    ok = True

    Dim tmpMaterial As clsRequeMateriales

    Dim x As Long
    For x = 1 To mMateriales.count
        Set tmpMaterial = mMateriales(x)
        If tmpMaterial.ListaProveedores.count = 0 Then
            ok = False
            Exit For
        End If
    Next

    ValidarProveedores = ok
End Function

Public Property Get StringDestino() As String
    Dim destinity As String
    Dim Ot As OrdenTrabajo

    If vTipo = ot_ Then

        If IsSomething(Ot) Then
            destinity = "OT " & vDestinoOT
        Else
            destinity = "OT " & vDestinoOT
        End If

    Else
        destinity = "Stock"
    End If
    StringDestino = destinity
End Property
