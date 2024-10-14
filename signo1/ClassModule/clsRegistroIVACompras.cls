VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRegistroIVACompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public FEcha As String
Public tipodecomprobante As String
Public tipodoccontable As String
Public idconfigfactura As String
Public puntoDeVenta As String
Public numerodecomprobante As String
Public vacio1 As String
Public codigodevendedor As String
Public Cuit As String
Public denominacionvendedor As String
Public quincecaracteresUno As String
Public quincecaracteresDos As String
Public quincecaracteresTres As String
Public moneda As String
Public diezcaracteres As String
Public cantidaddealicuotas As String
Public codigodeoperacionXXX As String
Public veintiseiscaracteres As String
Public quincecaracteresCuatro As String
Public vacio2 As String

Public importettoperacionnuevo As Double
Public redondeoiva As Double
Public percepcionesSoloIva As Double
Public percepcionessSinIva As Double
Public impuestosinternos As Double
Public valoriva As Double
Public percepcionesvalor As Double
Public ivavalor As Double
Public montoneto As Double
Public valorAlicuota As Double

Public contadoridIVA As Integer
Public idIVA As Integer
Public cantidadidIVADistintas As Integer
