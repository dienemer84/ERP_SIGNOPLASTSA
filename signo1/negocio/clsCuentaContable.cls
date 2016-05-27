VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCuentaContable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim vId As Long
Dim vcodigo As String
Dim vNombre As String
Public TotalAcumulado As Double
Public Property Let id(nvalue As Long)
    vId = nvalue
End Property
Public Property Get id() As Long
    id = vId
End Property
Public Property Let codigo(nvalue As String)
    vcodigo = nvalue
End Property
Public Property Get codigo() As String
    codigo = vcodigo
End Property
Public Property Let nombre(nvalue As String)
    vNombre = nvalue
End Property
Public Property Get nombre() As String
    nombre = vNombre
End Property




