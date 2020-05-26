VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Remito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public id As Long
Public numero As Long
Public Cliente As clsCliente
Public FEcha As Date
Public detalle As String
Public estado As EstadoRemito
Public EstadoFacturado As EstadoRemitoFacturado
Public contacto As clsContacto
Public UsuarioCreador As clsUsuario
Public UsuarioAprobador As clsUsuario
Public Detalles As New Collection    '(of RemitoDetalle)

Public CantidadBultos As Integer
Public ControlCargaImpresiones As Integer



Public Function CantidadDeLineasActuales() As Double
    Dim Cant As Integer
    Dim deta As remitoDetalle
    Cant = 0
    If IsSomething(Detalles) Then
        For Each deta In Detalles
            If LenB(deta.Observaciones) > 0 Then
                Cant = Cant + 2
            Else
                Cant = Cant + 1
            End If

        Next deta

    End If
    CantidadDeLineasActuales = Cant
End Function



Public Property Get VerEstadoFacturado() As String
    If estado = RemitoAnulado Then
        VerEstadoFacturado = "Anulado"
    Else
        VerEstadoFacturado = enums.EnumEstadoRemitoFacturado(EstadoFacturado)
    End If
End Property


