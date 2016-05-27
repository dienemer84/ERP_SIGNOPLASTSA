VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DocumentoDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Id As Long
Public PosX As Long
Public PosY As Long
Public Alto As Long
Public Ancho As Long
Public Fijo As Boolean
Public Alineacion As AlignmentConstants
Public Negrita As Boolean
Public Cursiva As Boolean
Public Tachado As Boolean
Public Subrayado As Boolean
Public nombreFuente As String
Public Tamano As Double
Public Tag As String
Public documento As documento
