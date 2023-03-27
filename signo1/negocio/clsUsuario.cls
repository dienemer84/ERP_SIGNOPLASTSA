VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private vUsuario As String
Private vId As Long
Private vPassword As String
'Private vEstado As EstadoUsuario
Private mEmpleado As clsEmpleado
Private m_EventosSuscriptos As Dictionary
Public Memo As String

Public Property Set EventosSuscriptos(value As Dictionary)
    Set m_EventosSuscriptos = value
End Property

Public Property Get EventosSuscriptos() As Dictionary
    If Not IsSomething(m_EventosSuscriptos) Then
        Set m_EventosSuscriptos = DAOEvento.GetEventBroadCastTypesSuscribedForUser(Me.Id)
    End If

    Set EventosSuscriptos = m_EventosSuscriptos
End Property


Public Property Let usuario(nvalue As String)
    vUsuario = nvalue
End Property
Public Property Let Id(nvalue As Long)
    vId = nvalue
End Property
Public Property Let PassWord(nvalue As String)
    vPassword = nvalue
End Property
'Public Property Let estado(nvalue As EstadoUsuario)
'
'    vEstado = nvalue '
'End Property
Public Property Let Empleado(nvalue As clsEmpleado)
    Set mEmpleado = nvalue
End Property
Public Property Get usuario() As String
    usuario = vUsuario
End Property
Public Property Get Id() As Long
    Id = vId
End Property
Public Property Get PassWord() As String
    PassWord = vPassword
End Property
'Public Property Get estado() As EstadoUsuario
'   estado = vEstado
'End Property
Public Property Get Empleado() As clsEmpleado
    Set Empleado = mEmpleado
End Property
