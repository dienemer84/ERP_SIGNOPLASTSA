VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAESolicitar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public CAE As String
Public CAEVencimiento As String
Public Comprobante As String
Public FechaEmision As String
Public Resultado As String
Public FechaProceso As String
Public Errores As String
Public observaciones As String



Public Function getFechaFromString(str As String) As Date
    If LenB(str) = 16 Then
        Dim mes As String
        Dim dia As String
        Dim ano As String

        ano = Left(str, 4)
        dia = Right(str, 2)
        mes = Mid(str, 5, 2)
        getFechaFromString = CDate(ano & "-" & mes & "-" & dia)
    Else
        getFechaFromString = Null
    End If
End Function
