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
Public Id As Integer
Public UnidadMedida As String

Public CantPedida As Double
Public CantRecibida As Double
Public CantFabricada As Double

Public CantScrap As Double
    
Public FechaInicio As Date
Public FechaFin As Date
    
Public UsuarioRecibio As Long
    
Public ProcesoSiguiente As String
    
Public FechaMod As Date

Public EsConjunto As Boolean
