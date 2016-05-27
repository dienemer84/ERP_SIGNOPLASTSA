VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDTOPadronIIBB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public FechaPublicacion As String
Public FechaDesde As String
Public FechaHasta As String
Public Cuit As String
Public tipo As String
Public AltaBaja As String
Public Cambio As String
'Public Percepcion As Double
' Public Retencion As Double
'Public GrupoRetencion As String
'Public GrupoPercepcion As String
Public Grupo As String
Public Alicuota As Double
Public Discriminador As String
