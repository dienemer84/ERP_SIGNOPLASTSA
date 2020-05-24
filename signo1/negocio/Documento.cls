VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Documento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public id As Long
Public nombre As String
Public Ancho As Double
Public Alto As Double
Public Imagen As Long
Public Detalles As Collection    'of DocumentoDetalle
Public estado As Boolean
Public TipoDocumento As TipoDocumentoImpresion



