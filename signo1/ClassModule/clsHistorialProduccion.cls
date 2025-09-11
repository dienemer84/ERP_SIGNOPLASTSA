VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHistorialProduccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vId As Long
Private vIdPedido As Long
Private vIdDetalle As Long
Private vIdPieza As Long
Private vSector As String

Private vUsuario As clsUsuario
Private vUsuarioRecibio As clsUsuario

Private vFecha As Date
Private vAccion As String
Private vCantRecibida As Double
Private vCantFabricada As Double
Private vCantScrap As Double
Private vNota As String

'--- Ids de referencia
Public Property Let Id(nValue As Long)
    vId = nValue
End Property
Public Property Get Id() As Long
    Id = vId
End Property

Public Property Let IdPedido(nValue As Long)
    vIdPedido = nValue
End Property
Public Property Get IdPedido() As Long
    IdPedido = vIdPedido
End Property

Public Property Let IdDetalle(nValue As Long)
    vIdDetalle = nValue
End Property
Public Property Get IdDetalle() As Long
    IdDetalle = vIdDetalle
End Property

Public Property Let idPieza(nValue As Long)
    vIdPieza = nValue
End Property
Public Property Get idPieza() As Long
    idPieza = vIdPieza
End Property

'--- Sector
Public Property Let Sector(nValue As String)
    vSector = nValue
End Property
Public Property Get Sector() As String
    Sector = vSector
End Property

'--- Usuario
Public Property Set Usuario(nValue As clsUsuario)
    Set vUsuario = nValue
End Property

Public Property Get Usuario() As clsUsuario
    Set Usuario = vUsuario
End Property

Public Property Set UsuarioRecibio(nValue As clsUsuario)
    Set vUsuarioRecibio = nValue
End Property

Public Property Get UsuarioRecibio() As clsUsuario
    Set UsuarioRecibio = vUsuarioRecibio
End Property

'--- Fecha y acción
Public Property Let FEcha(nValue As Date)
    vFecha = nValue
End Property
Public Property Get FEcha() As Date
    FEcha = vFecha
End Property

Public Property Let Accion(nValue As String)
    vAccion = nValue
End Property
Public Property Get Accion() As String
    Accion = vAccion
End Property

'--- Cantidades
Public Property Let CantRecibida(nValue As Double)
    vCantRecibida = nValue
End Property
Public Property Get CantRecibida() As Double
    CantRecibida = vCantRecibida
End Property

Public Property Let CantFabricada(nValue As Double)
    vCantFabricada = nValue
End Property
Public Property Get CantFabricada() As Double
    CantFabricada = vCantFabricada
End Property

Public Property Let CantScrap(nValue As Double)
    vCantScrap = nValue
End Property
Public Property Get CantScrap() As Double
    CantScrap = vCantScrap
End Property

'--- Nota opcional
Public Property Let Nota(nValue As String)
    vNota = nValue
End Property
Public Property Get Nota() As String
    Nota = vNota
End Property
