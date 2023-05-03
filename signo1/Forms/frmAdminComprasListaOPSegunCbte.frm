VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmAdminComprasListaOPSegunCbte 
   Caption         =   "Ordenes de Pago vinculadas"
   ClientHeight    =   4755
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6750
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4755
   ScaleWidth      =   6750
   Begin XtremeSuiteControls.GroupBox GroupBox 
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   3840
      Width           =   6495
      _Version        =   786432
      _ExtentX        =   11456
      _ExtentY        =   1508
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.PushButton btnCerrar 
         Height          =   495
         Left            =   2400
         TabIndex        =   3
         Top             =   240
         Width           =   1575
         _Version        =   786432
         _ExtentX        =   2778
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Cerrar"
         UseVisualStyle  =   -1  'True
      End
   End
   Begin GridEX20.GridEX gridOP 
      Height          =   2775
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   4895
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      ColumnAutoResize=   -1  'True
      MethodHoldFields=   -1  'True
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      ColumnsCount    =   3
      Column(1)       =   "frmAdminComprasListaOPSegunCbte.frx":0000
      Column(2)       =   "frmAdminComprasListaOPSegunCbte.frx":016C
      Column(3)       =   "frmAdminComprasListaOPSegunCbte.frx":02AC
      FormatStylesCount=   9
      FormatStyle(1)  =   "frmAdminComprasListaOPSegunCbte.frx":03F4
      FormatStyle(2)  =   "frmAdminComprasListaOPSegunCbte.frx":052C
      FormatStyle(3)  =   "frmAdminComprasListaOPSegunCbte.frx":05DC
      FormatStyle(4)  =   "frmAdminComprasListaOPSegunCbte.frx":0690
      FormatStyle(5)  =   "frmAdminComprasListaOPSegunCbte.frx":0768
      FormatStyle(6)  =   "frmAdminComprasListaOPSegunCbte.frx":0820
      FormatStyle(7)  =   "frmAdminComprasListaOPSegunCbte.frx":0900
      FormatStyle(8)  =   "frmAdminComprasListaOPSegunCbte.frx":09B8
      FormatStyle(9)  =   "frmAdminComprasListaOPSegunCbte.frx":0A6C
      ImageCount      =   0
      PrinterProperties=   "frmAdminComprasListaOPSegunCbte.frx":0B20
   End
   Begin XtremeSuiteControls.Label lblInstrucciones 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   6495
      _Version        =   786432
      _ExtentX        =   11456
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "* Doble click para abrir la OP / LIQ seleccionada."
   End
   Begin XtremeSuiteControls.Label lblNumeroCbte 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6495
      _Version        =   786432
      _ExtentX        =   11456
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Comprobante N°"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmAdminComprasListaOPSegunCbte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vFactura As clsFacturaProveedor
Dim ordenes As New Collection
Private Orden As OrdenPago


Public Property Let Factura(nFactura As clsFacturaProveedor)
    If IsSomething(nFactura) Then
        Set vFactura = DAOFacturaProveedor.FindById(nFactura.Id)
    End If
End Property

Private Sub btnCerrar_Click()
    Unload Me
    
End Sub

Private Sub Form_Load()

    Me.Height = 5265
    Me.Width = 6870
    
    Me.Left = frmPrincipal.ScaleWidth / 3
    Me.Top = frmPrincipal.ScaleHeight / 4
    
    FormHelper.Customize Me
    
    Me.lblNumeroCbte.caption = "Comprobante: " & vFactura.NumeroFormateado & " | " & UCase(vFactura.Proveedor.RazonSocial)
    
    MostrarOPyLiquidaciones
    
End Sub

Public Sub MostrarOPyLiquidaciones()

   Dim filter As String
    filter = "1 = 1"

    filter = filter & " AND  AdminComprasFacturasProveedores.id  = " & vFactura.Id

    Me.gridOP.ItemCount = 0
    Set ordenes = DAOOrdenPago.FindAll(filter, "ordenes_pago.id DESC")
    Me.gridOP.ItemCount = ordenes.count

End Sub

Private Sub gridOP_ColumnHeaderClick(ByVal Column As GridEX20.JSColumn)
    GridEXHelper.ColumnHeaderClick Me.gridOP, Column
End Sub

Private Sub gridOP_DblClick()
    gridOP_SelectionChange
    verOP
End Sub

Private Sub gridOP_RowFormat(RowBuffer As GridEX20.JSRowData)
    If RowBuffer.RowIndex > 0 And ordenes.count > 0 Then
        Set Orden = ordenes.item(RowBuffer.RowIndex)
        If Orden.estado = EstadoOrdenPago.EstadoOrdenPago_Aprobada Then
            RowBuffer.CellStyle(3) = "aprobada"
        ElseIf Orden.estado = EstadoOrdenPago_Anulada Then
            RowBuffer.RowStyle = "anulada2"

            RowBuffer.CellStyle(3) = "anulada"
        ElseIf Orden.estado = EstadoOrdenPago_pendiente Then
            RowBuffer.CellStyle(3) = "pendiente"
        End If
    End If
End Sub

Private Sub gridOP_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex > 0 And ordenes.count > 0 Then
    
        Set Orden = ordenes.item(RowIndex)
        
        Values(1) = Orden.Id
        Values(2) = Orden.FEcha
        Values(3) = enums.EnumEstadoOrdenPago(Orden.estado)
        
    End If
End Sub

Private Sub verOP()
    Dim f22 As New frmAdminPagosCrearOrdenPago
    f22.Show
    f22.ReadOnly = True
    f22.Cargar Orden
    
End Sub

Private Sub gridOP_SelectionChange()
    On Error Resume Next
    Set Orden = ordenes.item(gridOP.RowIndex(gridOP.row))
    
End Sub
