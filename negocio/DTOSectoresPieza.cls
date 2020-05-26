VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DTOSectoresTiempo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sector As clsSector
Public ListaDtoTareaTiempoPendiente As New Collection
Public ListaDtoTareaTiempo As New Collection
Public Property Get TotalTareas() As Integer
    Dim c As Integer
    c = 0
    Dim dto As DTOTareaTiempo

    For Each dto In ListaDtoTareaTiempo
        c = c + dto.CantidadTareas

    Next

    TotalTareas = c
End Property
Public Property Get TotalTareasFinalizadas() As Integer
    Dim c As Integer
    c = 0
    Dim dto As DTOTareaTiempo

    For Each dto In ListaDtoTareaTiempo
        c = c + dto.CantidadTareasFinalizadas

    Next

    TotalTareasFinalizadas = c
End Property


Public Property Get Tiempo() As Double
    Dim dto As DTOTareaTiempo
    Dim c As Double
    For Each dto In ListaDtoTareaTiempo
        c = c + dto.Tiempo
    Next
    Tiempo = c
End Property

Public Property Get TiempoPendiente() As Double
    Dim dto As DTOTareaTiempo
    Dim c As Double
    For Each dto In ListaDtoTareaTiempoPendiente
        c = c + dto.TiempoPendiente
    Next
    TiempoPendiente = c
End Property

