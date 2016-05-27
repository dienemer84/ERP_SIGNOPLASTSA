VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FECAEDetRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Concepto As String
Public DocTipo As String
Public DocNro As String
Public CbteDesde As String
Public CbteHasta As String
Public CbteFch As String
Public ImpTotal As String
Public ImpTotConc As String
Public ImpNeto As String
Public ImpOpEx As String
Public ImpTrib As String
Public ImpIVA As String
Public FchServDesde As String
Public FchServHasta As String
Public FchVtoPago As String
Public MonId As String
Public MonCotiz As String
Public CbtesAsoc As New Collection    'of cbteAsoc

Public Tributos As New Collection    'of tributo
Public Iva As New Collection    'of AlicIva
Public Opcionales As New Collection    'of opcional
