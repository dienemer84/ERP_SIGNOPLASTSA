VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 1  'vbSimpleBound
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTarea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vId As Long
Private vSector As clsSector
Private vTarea As String
Private vCantXProc As Double

'agregados de valores_MDO
Public IdValorMDO As Long
Public Valor As Double
Public FEcha As Date
Public Moneda As clsMoneda
Public descripcion As String

Public CategoriaSueldo As CategoriaSueldo
Public SectorID As Long

Public Property Let id(nvalue As Long)
    vId = nvalue
End Property

Public Property Set Sector(nvalue As clsSector)
    Set vSector = nvalue
End Property

Public Property Let Tarea(nvalue As String)

    vTarea = nvalue
End Property

Public Property Let CantPorProc(nvalue As Double)
    vCantXProc = nvalue
End Property


Public Property Get id() As Long
    id = vId
End Property

Public Property Get Sector() As clsSector
    Set Sector = vSector
End Property

Public Property Get Tarea() As String
    Tarea = vTarea
End Property

Public Property Get CantPorProc() As Double
    CantPorProc = vCantXProc
End Property

Public Property Get Description() As String
    Description = Me.id & " - " & Me.Tarea
End Property


Public Property Get CantPorProcSmartProperty() As String
    If Me.CantPorProc = -1 Then
        CantPorProcSmartProperty = "Cambio"
    ElseIf Me.CantPorProc = 0 Then
        CantPorProcSmartProperty = "Fijo"
    Else
        CantPorProcSmartProperty = Me.CantPorProc
    End If
End Property

