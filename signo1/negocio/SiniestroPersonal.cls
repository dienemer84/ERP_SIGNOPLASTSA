VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SiniestroPersonal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public id As Long
Public NroSiniestro As String
Public FechaHoraOcurrido As Date
Public Asegurado As clsEmpleado
Public Supervisor As clsEmpleado
Public Diagnostico As String
Public PrestadorMedico As String
Public TipoAccidente As TipoAccidenteSiniestro
Public TipoTratamiento As TipoTratamientoSiniestro
Public TipoGravedad As TipoGravedadSiniestro
Public RenaudaTareas As Date
Public Gestor As String
Public ART As ART
Public InformeAccidente As InformeAccidente
Public Sector As clsSector

Public Property Get InformeAccidenteConfeccionado() As Boolean
    InformeAccidenteConfeccionado = IsSomething(Me.InformeAccidente)
End Property

Private Sub Class_Initialize()

End Sub
