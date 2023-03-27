VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMaterialHistorico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim vFecha As Date
Dim vValor As String
Dim vMoneda As clsMoneda
Public Property Let moneda(nvalue As clsMoneda)
    Set vMoneda = nvalue
End Property
Public Property Let FEcha(nvalue As Date)
    vFecha = nvalue
End Property
Public Property Let Valor(nvalue As String)
    vValor = nvalue
End Property
Public Property Get FEcha() As Date
    FEcha = vFecha
End Property
Public Property Get Valor() As String
    Valor = vValor
End Property
Public Property Get moneda() As clsMoneda
    Set moneda = vMoneda
End Property


