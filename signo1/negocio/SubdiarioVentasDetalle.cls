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
Public Redondeo As Double




Public Function TienePercepcionesIVA() As Boolean
    Dim P As clsPercepcionesAplicadas
   
    For Each P In ListaPercepciones

        If P.Percepcion.Id = Configurar.IdPercepcionIVA Then
            TienePercepcionesIVA = True
            Exit Function
        End If


    Next P
    TienePercepcionesIVA = False


    Exit Function
End Function

Public Function TienePercepcionesIIBB() As Boolean
    Dim P As clsPercepcionesAplicadas
    Dim tp As clsPercepciones
    Dim IIBB As Integer
    For Each P In ListaPercepciones
        Dim esta As Boolean
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
    Dim pos As Long
    Dim tipoLetra As String
    pos = InStr(1, Me.Comprobante, "-")
    If pos > 0 Then
        tipoLetra = Mid$(Me.Comprobante, pos + 1, 1)
    End If

    ComprobanteTipoLetra = tipoLetra
End Property

Public Property Get ComprobanteNro() As String
    Dim pos As Long
    Dim nro As String
    pos = InStrRev(Me.Comprobante, "-")
    If pos > 0 Then
        nro = Mid$(Me.Comprobante, pos + 1)
    End If

    ComprobanteNro = nro
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



