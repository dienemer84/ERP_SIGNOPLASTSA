VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFilaPlanoRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public item As String
Public nombre As String

Public Nivel As Integer

Public IdTabla As Long 'id de tabla

Public IdPedido As Long
Public idPiezaPedido As Long
Public cantpedida As Long
Public CantRecibida As Long
Public CantFabricada As Long
Public CantScrap As Long
    
Public FechaInicio As Variant
Public HoraInicio As Variant
Public FechaFin As Variant
Public HoraFin As Variant

Public UsuarioRecibio As Long

Public Almacen As Long

Public ProcesoSiguiente As String

Public Observaciones As String

Public UnidadMedida As String
Public EsConjunto As Boolean

Public IdSector As Long

Public Historial As New Collection
