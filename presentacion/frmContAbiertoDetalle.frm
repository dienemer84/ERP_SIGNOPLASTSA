VERSION 5.00
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#12.0#0"; "CODEJO~1.OCX"
Begin VB.Form frmContAbiertoDetalle 
   Caption         =   "Contrato Abierto Nº "
   ClientHeight    =   6540
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9930
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6540
   ScaleWidth      =   9930
   Begin XtremeReportControl.ReportControl ReportControl 
      Height          =   5385
      Left            =   -15
      TabIndex        =   0
      Top             =   -15
      Width           =   7905
      _Version        =   786432
      _ExtentX        =   13944
      _ExtentY        =   9499
      _StockProps     =   64
      BorderStyle     =   2
      PreviewMode     =   -1  'True
   End
End
Attribute VB_Name = "frmContAbiertoDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public OTMarco As OrdenTrabajo

Private deta As DetalleOrdenTrabajo

Private Sub Form_Load()
    Customize Me
    Dim lista As New Collection


    'If OTMarco.Detalles Is Nothing Or OTMarco.Detalles.count = 0 Then
    Set OTMarco.Detalles = DAODetalleOrdenTrabajo.FindAllByOrdenTrabajo(OTMarco.id, True, True, True)
    'End If

    Me.caption = "Contrato Abierto Nº " & OTMarco.id

    Me.ReportControl.Columns.DeleteAll
    Me.ReportControl.Records.DeleteAll

    AddColumn "Item", True, , 190
    AddColumn "Cant Pedida", , xtpAlignmentRight, 75
    AddColumn "Cant Fabricada", , xtpAlignmentRight, 88
    AddColumn "Cant Facturada", , xtpAlignmentRight, 88
    AddColumn "Cant Entregada", , xtpAlignmentRight, 88

    Dim rec As ReportRecord
    Dim rec2 As ReportRecord
    Dim deta2 As DetalleOrdenTrabajo
    Dim strsql As String
    For Each deta In OTMarco.Detalles
        Set rec = Me.ReportControl.Records.Add
        rec.AddItem deta.item
        rec.PreviewText = deta.Pieza.nombre
        rec.AddItem deta.CantidadPedida & "  (" & deta.CantidadConsumida & ")"
        rec.AddItem deta.Cantidad_Fabricada
        rec.AddItem deta.Cantidad_Facturada
        rec.AddItem deta.Cantidad_Entregada
        Dim rs As Recordset
        For Each deta2 In deta.DetallesHijasMarcoPadre


            Set rec2 = New ReportRecord

            If deta2.OrdenTrabajo.estado = EstadoOT_EnProceso Or deta2.OrdenTrabajo.estado = EstadoOT_Finalizado Then
                ' If deta2.item = "002" And deta2.OrdenTrabajo.id = 972 Then Stop

                strsql = "Select descripcion from pedidos where id=" & deta2.OrdenTrabajo.id
                Set rs = conectar.RSFactory(strsql)

                deta2.OrdenTrabajo.descripcion = rs!descripcion

                rec2.AddItem "OT" & deta2.OrdenTrabajo.id & " | " & deta2.item & " | " & deta2.OrdenTrabajo.descripcion
                rec2.AddItem deta2.CantidadPedida
                rec2.AddItem deta2.Cantidad_Fabricada
                rec2.AddItem deta2.Cantidad_Facturada
                rec2.AddItem deta2.Cantidad_Entregada
                Me.ReportControl.AddRecordEx2 rec2, rec
            End If
        Next deta2
    Next deta

    Me.ReportControl.Populate
End Sub

Private Sub AddColumn(caption As String, Optional IsTree As Boolean = False, Optional align As XTPColumnAlignment = xtpAlignmentIconLeft, Optional ByVal Ancho As Double = 0)
    Static idx As Long
    Dim Re As ReportColumn
    Set Re = Me.ReportControl.Columns.Add(idx, caption, 50, True)
    Re.TreeColumn = IsTree
    Re.Alignment = align
    If Ancho <> 0 Then Re.Width = Ancho

    idx = idx + 1
End Sub

Private Sub Form_Resize()
    Me.ReportControl.Width = Me.ScaleWidth
    Me.ReportControl.Height = Me.ScaleHeight
End Sub
