VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InformeAccidente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public id As Long

Public Puesto As String
Public NombreTestigos As String
Public DescripcionHecho As String    'como, donde, con que
Public HsExtras As Boolean

'factores contribuyentes
Public FallaMaquinasEquipos As String
Public FaltaElementosProteccionPersonal As String
Public ActoInseguro As String
Public Otros As String

'completar por responsable
Public NaturalezaLesion As String
Public UbicacionLesion As String
Public FormaAccidente As String
Public AgenteMaterial As String
Public RecomendacionParaEvitarRepeticion As String


