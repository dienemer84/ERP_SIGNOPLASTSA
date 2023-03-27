VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTipoFacturaDiscriminado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public TipoFactura As clsTipoFactura
Public TipoDoc As tipoDocumentoContable
Public Numeracion As Long
Public PuntoVenta As PuntoVenta
Public Id As Long
Public Property Get descripcion() As String
    If IsSomething(PuntoVenta) Then
        descripcion = Format(Me.PuntoVenta.PuntoVenta, "000") & "-" & Me.PuntoVenta.descripcion
    Else
        descripcion = "Error"
    End If
End Property
