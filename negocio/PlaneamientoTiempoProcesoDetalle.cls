VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PlaneamientoTiempoProcesoDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public id As Long
Public IdPlaneamientoTiempoProceso As Long
Public legajo As Long
Public FechaCarga As Date
Public FechaInicioTarea As Date
Public FechaFinTarea As Date
Public CantidadProcesada As Double
Public usuario As clsUsuario
Public Empleado As clsEmpleado
Public PlaneamientoTiempoProceso As PlaneamientoTiempoProceso

Public Property Get DiferenciaTiempos() As Double    'en horas
    If CDbl(FechaInicioTarea) = 0 Or CDbl(FechaFinTarea) = 0 Then
        DiferenciaTiempos = 0
    Else
        DiferenciaTiempos = RedondearDecimales(DateDiff("n", FechaInicioTarea, FechaFinTarea) / 60, 2)
    End If
End Property

Public Property Get DiferenciaTiemposMinutos() As Double
    DiferenciaTiemposMinutos = RedondearDecimales(DiferenciaTiempos * 60, 2)
End Property

Public Property Get DiferenciaTiempoHorasMinutos() As String
    On Error GoTo E

    Dim horas As Double
    Dim minutos As Double

    Dim tmp As Double: tmp = DiferenciaTiempos
    horas = Fix(tmp)
    minutos = tmp - Fix(tmp)
    minutos = Fix(minutos * 60)

    DiferenciaTiempoHorasMinutos = horas & ":" & Format(minutos, "00")
    Exit Property
E:
    DiferenciaTiempoHorasMinutos = "0:00"
End Property


