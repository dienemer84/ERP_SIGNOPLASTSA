VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NotaNoConformidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Id As Long
Public TiempoProceso As PlaneamientoTiempoProceso
Public FechaCreacion As Date
Public FechaResolucion As Date
Public usuarioCreador As clsUsuario
Public Operario As clsEmpleado
Public Encargado As clsEmpleado
Public descripcion As String
Public AccionTomada As String
Public UsuarioResolucionador As clsUsuario
Public TareaOrigen As clsTarea
Public estado As EstadoNotaNoConformidad
Public Incidencias As String

Public Property Get Resuelta() As Boolean
    Resuelta = (Me.FechaResolucion <> 0)
End Property

Public Property Get EsInterna() As Boolean
    EsInterna = (Me.TareaOrigen.Sector.Id = TiempoProceso.Tarea.Sector.Id)
End Property

Public Property Get numero() As String
    If EsInterna Then numero = "I-" & Format(Id, "0000") Else numero = "E-" & Format(Id, "0000")
End Property
