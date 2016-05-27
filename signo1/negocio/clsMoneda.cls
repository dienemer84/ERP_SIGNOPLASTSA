VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMoneda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private vNombreCorto As String
Private vNombreLargo As String
Private vCambio As Double
Private vpatron As Boolean
Private vfechaActual As Date
Private vId As Long
Public MonedaCambio As clsMoneda

Public Property Let id(nvalue As Long)
    vId = nvalue
End Property
Public Property Get id() As Long
    id = vId
End Property
Public Property Let NombreCorto(nvalue As String)
    vNombreCorto = nvalue
End Property
Public Property Let NombreLargo(nvalue As String)
    vNombreLargo = nvalue
End Property
Public Property Let Cambio(nvalue As Double)
    vCambio = nvalue
End Property
Public Property Let Patron(nvalue As Boolean)
    vpatron = nvalue
End Property
Public Property Let FechaActual(nvalue As Date)
    vfechaActual = nvalue
End Property

Public Property Get NombreCorto() As String
    NombreCorto = vNombreCorto
End Property
Public Property Get NombreLargo() As String
    NombreLargo = vNombreLargo
End Property
Public Property Get Cambio() As Double
    Cambio = vCambio
End Property
Public Property Get Patron() As Boolean
    Patron = vpatron
End Property
Public Property Get FechaActual() As Date
    FechaActual = vfechaActual
End Property


