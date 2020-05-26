VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHistorial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim vUsuario As clsUsuario
Dim vFecha As Date
Dim vMensaje As String

Public Property Let FEcha(nvalue As Date)
    vFecha = nvalue
End Property

Public Property Let mensaje(nvalue As String)
    vMensaje = nvalue
End Property


Public Property Get FEcha() As Date
    FEcha = vFecha
End Property
Public Property Get mensaje() As String
    mensaje = vMensaje
End Property

Public Property Let usuario(nvalue As clsUsuario)
    Set vUsuario = nvalue
End Property
Public Property Get usuario() As clsUsuario
    Set usuario = vUsuario
End Property
