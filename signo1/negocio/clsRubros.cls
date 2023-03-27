VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRubros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private vId As Long
Private vRubro As String
Private vIniciales As String
Private vContador As Long
'Private vGrupos As Collection


'Public Property Let grupos(nValue As Collection)
'    Set vGrupos = nValue
'End Property
'Public Property Get grupos() As Collection
'Set grupos = vGrupos
'End Property
Public Property Let Id(nvalue As Long)
    vId = nvalue
End Property
Public Property Let rubro(nvalue As String)
    vRubro = UCase(nvalue)
End Property
Public Property Let iniciales(nvalue As String)
    vIniciales = UCase(nvalue)
End Property
Public Property Let Contador(nvalue As Long)
    vContador = nvalue
End Property
Public Property Get Id() As Long
    Id = vId
End Property
Public Property Get rubro() As String
    rubro = vRubro
End Property
Public Property Get iniciales() As String
    iniciales = vIniciales
End Property
Public Property Get Contador() As Long
    Contador = vContador
End Property

