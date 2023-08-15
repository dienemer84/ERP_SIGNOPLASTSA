VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmAdminComprasListaLIQSegunCbte 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Liquidacion de Caja vinculadas"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6810
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   6810
   ShowInTaskbar   =   0   'False
   Begin GridEX20.GridEX gridLIQ 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
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
      Column(1)       =   "frmAdminComprasListaLIQSegunCbte.frx":0000
      Column(2)       =   "frmAdminComprasListaLIQSegunCbte.frx":016C
      Column(3)       =   "frmAdminComprasListaLIQSegunCbte.frx":02AC
      FormatStylesCount=   9
      FormatStyle(1)  =   "frmAdminComprasListaLIQSegunCbte.frx":03F4
      FormatStyle(2)  =   "frmAdminComprasListaLIQSegunCbte.frx":052C
      FormatStyle(3)  =   "frmAdminComprasListaLIQSegunCbte.frx":05DC
      FormatStyle(4)  =   "frmAdminComprasListaLIQSegunCbte.frx":0690
      FormatStyle(5)  =   "frmAdminComprasListaLIQSegunCbte.frx":0768
      FormatStyle(6)  =   "frmAdminComprasListaLIQSegunCbte.frx":0820
      FormatStyle(7)  =   "frmAdminComprasListaLIQSegunCbte.frx":0900
      FormatStyle(8)  =   "frmAdminComprasListaLIQSegunCbte.frx":09B8
      FormatStyle(9)  =   "frmAdminComprasListaLIQSegunCbte.frx":0A6C
      ImageCount      =   0
      PrinterProperties=   "frmAdminComprasListaLIQSegunCbte.frx":0B20
   End
   Begin XtremeSuiteControls.GroupBox GroupBox 
      Height          =   855
      Left            =   120
      TabIndex        =   3
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
         TabIndex        =   4
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
   Begin XtremeSuiteControls.Label lblNumeroCbte 
      Height          =   375
      Left            =   240
      TabIndex        =   2
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
   Begin XtremeSuiteControls.Label lblInstrucciones 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   6495
      _Version        =   786432
      _ExtentX        =   11456
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "* Doble click para abrir la LIQ seleccionada."
   End
End
Attribute VB_Name = "frmAdminComprasListaLIQSegunCbte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vFactura As clsFacturaProveedor
Dim liquidaciones As New Collection
Private Liquidacion As clsLiquidacionCaja

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

    Me.gridLIQ.ItemCount = 0
   
    Set liquidaciones = DAOLiquidacionCaja.FindAll(filter, "liquidaciones_caja.id DESC")
    Me.gridLIQ.ItemCount = liquidaciones.count

End Sub

Private Sub gridLIQ_ColumnHeaderClick(ByVal Column As GridEX20.JSColumn)
    GridEXHelper.ColumnHeaderClick Me.gridLIQ, Column
End Sub

Private Sub gridLIQ_DblClick()
    gridLIQ_SelectionChange
    verLIQ
End Sub


Private Sub gridLIQ_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If rowIndex > 0 And liquidaciones.count > 0 Then
    
        Set Liquidacion = liquidaciones.item(rowIndex)
        
        Values(1) = Liquidacion.NumeroLiq
        Values(2) = Liquidacion.FEcha
        
        If Liquidacion.estado = EstadoLiquidacionCaja_Aprobada Then
            Values(3) = "Aprobada"
        ElseIf Liquidacion.estado = EstadoLiquidacionCaja_Anulada Then
            Values(3) = "Anulada"
        ElseIf Liquidacion.estado = EstadoLiquidacionCaja_pendiente Then
            Values(3) = "Pendiente"
        End If
    
  End If
    
End Sub

Private Sub verLIQ()
    Dim f22 As New frmAdminPagosLiqCajaListaDG
    f22.Show
    f22.ReadOnly = True
    f22.Cargar Liquidacion
    
End Sub

Private Sub gridLIQ_SelectionChange()
    On Error Resume Next
    Set Liquidacion = liquidaciones.item(gridLIQ.rowIndex(gridLIQ.row))
    
End Sub

