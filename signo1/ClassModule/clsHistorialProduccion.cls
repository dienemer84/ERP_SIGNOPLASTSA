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
Private vIdSector As Long

Private vUsuarioOperacion As clsUsuario
Private vUsuarioRecibio As clsUsuario

' Cantidades old / new
Private vCantRecibidaOld As Double
Private vCantRecibidaNew As Double
Private vCantFabricadaOld As Double
Private vCantFabricadaNew As Double
Private vCantScrapOld As Double
Private vCantScrapNew As Double

' Fechas old / new
Private vFechaInicioOld As Date
Private vFechaInicioNew As Date
Private vFechaFinOld As Date
Private vFechaFinNew As Date

' Procesos
Private vProcesoOld As clsSector
Private vProcesoNew As clsSector

' Datos generales
Private vAccion As String
Private vNota As String
Private vFecha As Date

'--- Ids de referencia
Public Property Let Id(nValue As Long): vId = nValue: End Property
Public Property Get Id() As Long: Id = vId: End Property

Public Property Let IdPedido(nValue As Long): vIdPedido = nValue: End Property
Public Property Get IdPedido() As Long: IdPedido = vIdPedido: End Property

Public Property Let IdDetalle(nValue As Long): vIdDetalle = nValue: End Property
Public Property Get IdDetalle() As Long: IdDetalle = vIdDetalle: End Property

Public Property Let idPieza(nValue As Long): vIdPieza = nValue: End Property
Public Property Get idPieza() As Long: idPieza = vIdPieza: End Property

Public Property Let IdSector(nValue As Long): vIdSector = nValue: End Property
Public Property Get IdSector() As Long: IdSector = vIdSector: End Property

'--- Usuario
Public Property Set UsuarioOperacion(nValue As clsUsuario): Set vUsuarioOperacion = nValue: End Property
Public Property Get UsuarioOperacion() As clsUsuario: Set UsuarioOperacion = vUsuarioOperacion: End Property

Public Property Set UsuarioRecibio(nValue As clsUsuario): Set vUsuarioRecibio = nValue: End Property
Public Property Get UsuarioRecibio() As clsUsuario: Set UsuarioRecibio = vUsuarioRecibio: End Property

'--- Cantidades
Public Property Let CantRecibidaOld(nValue As Double): vCantRecibidaOld = nValue: End Property
Public Property Get CantRecibidaOld() As Double: CantRecibidaOld = vCantRecibidaOld: End Property

Public Property Let CantRecibidaNew(nValue As Double): vCantRecibidaNew = nValue: End Property
Public Property Get CantRecibidaNew() As Double: CantRecibidaNew = vCantRecibidaNew: End Property

Public Property Let CantFabricadaOld(nValue As Double): vCantFabricadaOld = nValue: End Property
Public Property Get CantFabricadaOld() As Double: CantFabricadaOld = vCantFabricadaOld: End Property

Public Property Let CantFabricadaNew(nValue As Double): vCantFabricadaNew = nValue: End Property
Public Property Get CantFabricadaNew() As Double: CantFabricadaNew = vCantFabricadaNew: End Property

Public Property Let CantScrapOld(nValue As Double): vCantScrapOld = nValue: End Property
Public Property Get CantScrapOld() As Double: CantScrapOld = vCantScrapOld: End Property

Public Property Let CantScrapNew(nValue As Double): vCantScrapNew = nValue: End Property
Public Property Get CantScrapNew() As Double: CantScrapNew = vCantScrapNew: End Property

'--- Fechas
Public Property Let FechaInicioOld(nValue As Date): vFechaInicioOld = nValue: End Property
Public Property Get FechaInicioOld() As Date: FechaInicioOld = vFechaInicioOld: End Property

Public Property Let FechaInicioNew(nValue As Date): vFechaInicioNew = nValue: End Property
Public Property Get FechaInicioNew() As Date: FechaInicioNew = vFechaInicioNew: End Property

Public Property Let FechaFinOld(nValue As Date): vFechaFinOld = nValue: End Property
Public Property Get FechaFinOld() As Date: FechaFinOld = vFechaFinOld: End Property

Public Property Let FechaFinNew(nValue As Date): vFechaFinNew = nValue: End Property
Public Property Get FechaFinNew() As Date: FechaFinNew = vFechaFinNew: End Property


'--- Procesos

Public Property Set ProcesoOld(nValue As clsSector)
    Set vProcesoOld = nValue
End Property
Public Property Get ProcesoOld() As clsSector
    Set ProcesoOld = vProcesoOld
End Property

Public Property Set ProcesoNew(nValue As clsSector)
    Set vProcesoNew = nValue
End Property
Public Property Get ProcesoNew() As clsSector
    Set ProcesoNew = vProcesoNew
End Property

'--- Accion
Public Property Let Accion(nValue As String): vAccion = nValue: End Property
Public Property Get Accion() As String: Accion = vAccion: End Property

'--- Nota
Public Property Let Nota(nValue As String): vNota = nValue: End Property
Public Property Get Nota() As String: Nota = vNota: End Property

'--- Fecha registro
Public Property Let FEcha(nValue As Date): vFecha = nValue: End Property
Public Property Get FEcha() As Date: FEcha = vFecha: End Property
