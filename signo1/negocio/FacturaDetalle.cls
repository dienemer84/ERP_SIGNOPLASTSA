VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FacturaDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public id As Long
Private m_DetalleRemito As remitoDetalle
Public DetalleRemitoId As Long
Public idFactura As Long
Public estado As EstadoDetalleFacturaCliente
Public DescuentoAnticipo As Boolean
Public Bruto As Double  'bruto, sin descuento


Public idprovincia As Long
Public detalle As String
Public Cantidad As Double
Public IvaAplicado As Boolean
Public IBAplicado As Boolean
Public AplicadoARemito As Boolean
Public CantidadRemitosAplicados As Integer
Public ListaRemitosAplicados As String
Public PorcentajeDescuento As Double
Public Observacion As String

Public OtIdAnticipo As Long

Public Factura As Factura

Public Property Get detalleRemito() As remitoDetalle
    'lazy load simulation
    If m_DetalleRemito Is Nothing And Me.id > 0 And DetalleRemitoId <> -1 Then    ' creo que DetalleRemitoId tiene que ser <> -1 si no es concepto
        Set m_DetalleRemito = DAORemitoSDetalle.FindById(DetalleRemitoId)
    End If

    Set detalleRemito = m_DetalleRemito
End Property

Public Property Set detalleRemito(value As remitoDetalle)
    Set m_DetalleRemito = value
End Property

'neto gravado con iva por cantidad
Public Property Get Total() As Double
    Total = funciones.RedondearDecimales(Me.SubTotal * Me.Cantidad)    'CAMBIADO X NICO 5-3
End Property

'neto gravado con iva
Public Property Get SubTotal() As Double

    If (Me.IvaAplicado And Factura.EstaDiscriminada) Or Not Me.IvaAplicado Then
        SubTotal = (Me.NetoGravado)
    Else
        SubTotal = (Me.Factura.AlicuotaCorrespondiente * (Me.NetoGravado))
    End If

End Property

'no tiene iva ni percep
Public Property Get NetoGravado() As Double
    'Debug.Assert factura.numero <> 9126

    NetoGravado = funciones.RedondearDecimales((Me.Bruto) * ((100 - Me.PorcentajeDescuento) / 100))
End Property

Public Property Get VerOrigen() As String
    Dim rto As Remito


    If IsSomething(Me.detalleRemito) Then
        Set rto = DAORemitoS.FindById(Me.detalleRemito.Remito)
    End If

    If Not Me.AplicadoARemito Then
        If DetalleRemitoId <= 0 Then
            VerOrigen = "CONCEPTO"
        Else
            If IsSomething(Me.detalleRemito) Then
                VerOrigen = "RTO " & rto.numero & " | " & Me.detalleRemito.VerOrigen(True)
            Else
                VerOrigen = vbNullString
            End If
        End If
    Else
        Dim r As Recordset
        Dim Cant As Long
        Set r = conectar.RSFactory("SELECT COUNT(0) as cant FROM AdminFacturasDetalleAplicacionRemitos WHERE idFacturaDetalle = " & Me.id)
        If Not r.EOF Then Cant = r!Cant

        If Cant = 1 Then
            If IsSomething(rto) Then
                VerOrigen = "APLICADO RTO " & rto.numero & " DE " & Me.detalleRemito.VerOrigen(True)
            Else
                VerOrigen = "EL DETALLE ESTA ENLAZADO A UNA ENTREGA INEXISTENTE"
            End If
        ElseIf Cant > 1 Then
            VerOrigen = "APLICADO A VARIOS RTOS"
        End If

    End If
End Property

Public Property Get OrigenEsConcepto() As Boolean
    OrigenEsConcepto = (Me.DetalleRemitoId <= 0)
End Property

Private Sub Class_Initialize()
    Me.DetalleRemitoId = -1
End Sub


