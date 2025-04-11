VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDetalleComprobante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clsDetalleAbonado.cls
Option Explicit

' Propiedades
Private mIdOrdenPago As Long
Private mTotal As Double
Private mNetoGravado As Double
Private mOtros As Double
Private mFechaEmision As Date

' Propiedad IdOrdenPago
Public Property Get IdOrdenPago() As Long
    IdOrdenPago = mIdOrdenPago
End Property

Public Property Let IdOrdenPago(ByVal Value As Long)
    mIdOrdenPago = Value
End Property

' Propiedad Total
Public Property Get Total() As Double
    Total = mTotal
End Property

Public Property Let Total(ByVal Value As Double)
    mTotal = Value
End Property

' Propiedad NetoGravado
Public Property Get NetoGravado() As Double
    NetoGravado = mNetoGravado
End Property

Public Property Let NetoGravado(ByVal Value As Double)
    mNetoGravado = Value
End Property

' Propiedad Otros
Public Property Get Otros() As Double
    Otros = mOtros
End Property

Public Property Let Otros(ByVal Value As Double)
    mOtros = Value
End Property

' Propiedad FechaEmision
Public Property Get FechaEmision() As Date
    FechaEmision = mFechaEmision
End Property

Public Property Let FechaEmision(ByVal Value As Date)
    mFechaEmision = Value
End Property
