VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private vSector As String
Private vId As Long
Private vSectorizacion As Long
Private vModulo As String

Public Property Let Sector(nvalue As String)
    vSector = nvalue
End Property

Public Property Get Sector() As String
    Sector = vSector
End Property


Public Property Let Id(nvalue As Long)
    vId = nvalue
End Property

Public Property Get Id() As Long
    Id = vId
End Property


Public Property Let Sectorizacion(nvalue As Long)
    vSectorizacion = nvalue
End Property

Public Property Get Sectorizacion() As Long
    Sectorizacion = vSectorizacion
End Property


Public Property Let Modulo(nvalue As String)
    vModulo = nvalue
End Property

Public Property Get Modulo() As String
    Modulo = vModulo
End Property



