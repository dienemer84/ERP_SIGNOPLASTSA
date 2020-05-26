VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Chequera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Public id As Long
Public Banco As Banco
Public NumeroDesde As Long
Public NumeroHasta As Long
Public FechaCreacion As Date
Public Moneda As clsMoneda
Public numero As Long
Public Observaciones As String
Public Cheques As New Collection

Public Property Get Description() As String
    Dim ret As String
    If IsSomething(Me.Banco) Then
        ret = Me.Banco.nombre
    End If
    ret = ret & " (Nº " & Me.NumeroDesde & " a Nº " & Me.NumeroHasta & ") [" & Me.Moneda.NombreCorto & "]"
    Description = ret
End Property
