VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGrupo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private vId As Long
Private vGrupo As String
Private mRubros As clsRubros


Public Property Let Id(nvalue As Long)
    vId = nvalue
End Property
Public Property Let Grupo(nvalue As String)
    vGrupo = UCase(nvalue)
End Property
Public Property Let rubros(nvalue As clsRubros)
    Set mRubros = nvalue
End Property

Public Property Get Id() As Long
    Id = vId
End Property
Public Property Get Grupo() As String
    Grupo = vGrupo
End Property
Public Property Get rubros() As clsRubros
    Set rubros = mRubros
End Property
