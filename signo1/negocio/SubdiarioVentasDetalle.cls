VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SubdiarioVentasDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Id As Long
Public FEcha As String

Public ListaPercepciones As Collection
Public Comprobante As String
Public RazonSocial As String
Public Cuit As String
Public CondicionIva As String
Public NetoGravado As Double
Public Iva As Double
Public percepciones As Double
' Public PercepcionesIVA As Double
Public Exento As Double
Public total As Double
Public estado As EstadoFacturaCliente
Public FacturaId As Long
Public LiquidacionId As Long
Public AlicuotasIva As New Collection
Public NetosGravado As New Collection
Public ImpuestoInterno As Double
Public redondeo As Double

Public Function TienePercepcionesIVA() As Boolean
    Dim P As clsPercepcionesAplicadas
   
    For Each P In ListaPercepciones
        If P.Percepcion.Id = Configurar.IdPercepcionIVA Then
            TienePercepcionesIVA = True
            Exit Function
        End If
    Next P

    TienePercepcionesIVA = False
End Function

Public Function TienePercepcionesIIBB() As Boolean
    Dim P As clsPercepcionesAplicadas
    Dim tp As clsPercepciones

    For Each P In ListaPercepciones
        For Each tp In Configurar.idPercepcionesIIBB
            If tp.Id = P.Percepcion.Id Then
                TienePercepcionesIIBB = True
                Exit Function
            End If
        Next
    Next P

    TienePercepcionesIIBB = False
End Function

Public Property Get ComprobanteTipoLetra() As String
    Dim s As String
    Dim pos As Long

    s = Trim$(Me.Comprobante)
    pos = InStr(1, s, "-")

    If pos > 0 Then
        ComprobanteTipoLetra = Mid$(s, pos + 1, 1)
    Else
        ComprobanteTipoLetra = ""
    End If
End Property

Public Property Get ComprobanteUltimoBloque() As String
    Dim s As String
    Dim pos As Long

    s = Trim$(Me.Comprobante)

    pos = InStrRev(s, "|")
    If pos > 0 Then
        ComprobanteUltimoBloque = Trim$(Mid$(s, pos + 1))
    Else
        ComprobanteUltimoBloque = s
    End If
End Property

Public Property Get ComprobantePuntoVenta() As String
    Dim s As String
    Dim pos As Long

    s = Trim$(Me.ComprobanteUltimoBloque)
    pos = InStr(1, s, "-")

    If pos > 0 Then
        ComprobantePuntoVenta = Trim$(Left$(s, pos - 1))
    Else
        ComprobantePuntoVenta = ""
    End If
End Property

Public Property Get ComprobantePuntoVenta5() As String
    Dim pv As String

    pv = Trim$(Me.ComprobantePuntoVenta)

    If Len(pv) > 0 Then
        ComprobantePuntoVenta5 = Right$("00000" & pv, 5)
    Else
        ComprobantePuntoVenta5 = "00000"
    End If
End Property

Public Property Get ComprobanteNro() As String
    Dim s As String
    Dim pos As Long

    s = Trim$(Me.ComprobanteUltimoBloque)
    pos = InStr(1, s, "-")

    If pos > 0 Then
        ComprobanteNro = Trim$(Mid$(s, pos + 1))
    Else
        ComprobanteNro = ""
    End If
End Property

Public Property Get ComprobanteNro8() As String
    Dim NRO As String

    NRO = Trim$(Me.ComprobanteNro)

    If Len(NRO) > 0 Then
        ComprobanteNro8 = Right$("00000000" & NRO, 8)
    Else
        ComprobanteNro8 = "00000000"
    End If
End Property

Public Property Get EstadoForCompras() As EstadoFacturaProveedor
    If Me.estado = EstadoFacturaCliente.Anulada Then
        EstadoForCompras = EstadoFacturaProveedor.Saldada
    Else
        If Me.estado = EstadoFacturaCliente.Aprobada Then
            EstadoForCompras = EstadoFacturaProveedor.Aprobada
        ElseIf Me.estado = EstadoFacturaCliente.EnProceso Then
            EstadoForCompras = EstadoFacturaProveedor.EnProceso
        End If
    End If
End Property
