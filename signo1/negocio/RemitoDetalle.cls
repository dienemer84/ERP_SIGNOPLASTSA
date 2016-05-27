VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RemitoDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public id As Long
Public DetallePedido As DetalleOrdenTrabajo
Public idDetallePedido As Long
Public idpedido As Long
Public Remito As Long    'id remito o nro?
Public Cantidad As Double
Public FEcha As Date
Public Origen As OrigenRemito
Public Facturado As Boolean
Private m_concepto As String
Public facturable As Boolean
Public Valor As Double
Public ValorModificado As Boolean
Public EstadoRemito As EstadoRemito
Public RemitoAlQuePertenece As Remito
Public Observaciones As String

Public Property Let Concepto(Valor As String)
    If Origen = OrigenRemitoConcepto Or Origen = OrigenRemitoAplicado Then
        m_concepto = Valor
    Else
        m_concepto = vbNullString
    End If
End Property

Public Property Get Concepto() As String
    Concepto = m_concepto
End Property
Public Property Get VerOrigen(Optional withOTItem As Boolean = False) As String
    Dim ori As String
    If Origen = OrigenRemitoConcepto Then
        ori = "CONCEPTO"
    ElseIf Origen = OrigenRemitooe Then
        ori = "OE"
    ElseIf Origen = OrigenRemitoOt Then
        ori = "OT-" & idpedido
    ElseIf Origen = OrigenRemitoAplicado Then
        ori = "CONCEPTO APLICADO A OT-" & idpedido
        If withOTItem Then
            Dim detaOT As DetalleOrdenTrabajo
            Set detaOT = DAODetalleOrdenTrabajo.FindById(idDetallePedido)
            If IsSomething(detaOT) Then
                ori = ori & " (Item " & detaOT.item & ")"
            End If
        End If
    End If
    VerOrigen = ori
End Property

Public Property Get VerFacturado() As String

    If Not facturable Then
        VerFacturado = "No Facturable"
    Else
        If Facturado Then
            VerFacturado = "Facturado"
        Else
            VerFacturado = "No Facturado"
        End If
    End If
End Property

Public Property Get VerElemento() As String
    If Origen = OrigenRemitoConcepto Then
        VerElemento = Concepto
    ElseIf Origen = OrigenRemitoOt Then
        VerElemento = DetallePedido.Pieza.nombre
    ElseIf Origen = OrigenRemitooe Then
        Dim sql As String
        sql = "SELECT s.detalle as piezaaa FROM detallesPedidosEntregas dp INNER JOIN stock s ON s.id=dp.idPieza WHERE dp.id=" & Me.idDetallePedido
        Dim rs As Recordset
        Set rs = conectar.RSFactory(sql)
        If Not rs.EOF And Not rs.BOF Then
            VerElemento = rs!piezaaa
        Else
            VerElemento = "show oe error "
        End If

    ElseIf Origen = OrigenRemitoAplicado Then
        VerElemento = Concepto
    End If
End Property
