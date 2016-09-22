VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~3.OCX"
Begin VB.Form frmPlaneamientoPedidosSeleccion 
   Caption         =   "Lista OT"
   ClientHeight    =   5220
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9840
   Icon            =   "frmPlaneamientoPedidosSeleccion.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5220
   ScaleWidth      =   9840
   StartUpPosition =   3  'Windows Default
   Begin GridEX20.GridEX GridEX1 
      Height          =   4080
      Left            =   75
      TabIndex        =   1
      Top             =   1065
      Width           =   9690
      _ExtentX        =   17092
      _ExtentY        =   7197
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      PreviewColumn   =   "cliente"
      PreviewRowLines =   1
      ColumnAutoResize=   -1  'True
      MethodHoldFields=   -1  'True
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   6
      Column(1)       =   "frmPlaneamientoPedidosSeleccion.frx":000C
      Column(2)       =   "frmPlaneamientoPedidosSeleccion.frx":0114
      Column(3)       =   "frmPlaneamientoPedidosSeleccion.frx":0218
      Column(4)       =   "frmPlaneamientoPedidosSeleccion.frx":0324
      Column(5)       =   "frmPlaneamientoPedidosSeleccion.frx":0418
      Column(6)       =   "frmPlaneamientoPedidosSeleccion.frx":0540
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmPlaneamientoPedidosSeleccion.frx":0650
      FormatStyle(2)  =   "frmPlaneamientoPedidosSeleccion.frx":0788
      FormatStyle(3)  =   "frmPlaneamientoPedidosSeleccion.frx":0838
      FormatStyle(4)  =   "frmPlaneamientoPedidosSeleccion.frx":08EC
      FormatStyle(5)  =   "frmPlaneamientoPedidosSeleccion.frx":09C4
      FormatStyle(6)  =   "frmPlaneamientoPedidosSeleccion.frx":0A7C
      ImageCount      =   0
      PrinterProperties=   "frmPlaneamientoPedidosSeleccion.frx":0B5C
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   885
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   9705
      _Version        =   786432
      _ExtentX        =   17119
      _ExtentY        =   1561
      _StockProps     =   79
      Caption         =   "Parámetros de búsqueda"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.PushButton cmdBuscar 
         Height          =   390
         Left            =   8235
         TabIndex        =   5
         Top             =   300
         Width           =   1140
         _Version        =   786432
         _ExtentX        =   2011
         _ExtentY        =   688
         _StockProps     =   79
         Caption         =   "Buscar"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton CMDsINCliente 
         Height          =   285
         Left            =   6345
         TabIndex        =   4
         Top             =   360
         Width           =   330
         _Version        =   786432
         _ExtentX        =   582
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "X"
         BackColor       =   12632256
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cboClientes 
         Height          =   315
         Left            =   870
         TabIndex        =   3
         Top             =   345
         Width           =   5415
         _Version        =   786432
         _ExtentX        =   9551
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin VB.Label lblCliente 
         Caption         =   "Cliente"
         Height          =   225
         Left            =   255
         TabIndex        =   2
         Top             =   375
         Width           =   585
      End
   End
End
Attribute VB_Name = "frmPlaneamientoPedidosSeleccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cliente As clsCliente
Private ots As New Collection
Private Ot As OrdenTrabajo
Public MostrarAnticipo As Boolean
Dim q As String

Private Sub cmdBuscar_Click()
    llenarLista
End Sub

Private Sub CMDsINCliente_Click()
    Me.cboClientes.ListIndex = -1
End Sub
Private Sub Form_Load()
    Customize Me
    GridEXHelper.CustomizeGrid Me.GridEX1, True, False
    DAOCliente.llenarComboXtremeSuite Me.cboClientes, False, True, False
    CMDsINCliente_Click

    If IsSomething(cliente) Then
        'Me.GroupBox1.Enabled = False
        Me.cboClientes.ListIndex = funciones.PosIndexCbo(cliente.id, Me.cboClientes)
        Me.cboClientes.Enabled = False
        Me.CMDsINCliente.Enabled = False
    End If

    Me.GridEX1.Columns(5).Visible = MostrarAnticipo
    Me.GridEX1.ItemCount = 0
End Sub
Private Sub llenarLista()
    q = "{pedido}.{activo}=1"
    If Me.cboClientes.ListIndex <> -1 Then
        q = q & " AND {pedido}.idClienteFacturar = " & Me.cboClientes.ItemData(Me.cboClientes.ListIndex)
        'q = Replace$(q, "{cliente_id}", DAOOrdenTrabajo.CAMPO_CLIENTE_ID)
    End If
    If MostrarAnticipo Then
        q = q & " and {pedido}." & DAOOrdenTrabajo.CAMPO_ANTICIPO_FACTURADO & "=0 and {pedido}." & DAOOrdenTrabajo.CAMPO_ANTICIPO & ">0"
    End If

    q = Replace$(q, "{pedido}", DAOOrdenTrabajo.TABLA_PEDIDO)
    q = Replace$(q, "{activo}", DAOOrdenTrabajo.CAMPO_ACTIVO)
    Set ots = DAOOrdenTrabajo.FindAll(q)
    Me.GridEX1.ItemCount = 0
    Me.GridEX1.ItemCount = ots.count


End Sub
Private Sub Form_Resize()
    Me.GroupBox1.Height = Me.ScaleHeight
End Sub


Private Sub Form_Terminate()
    Set Selecciones.OrdenTrabajo = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Selecciones.OrdenTrabajo = Ot
End Sub

Private Sub GridEX1_ColumnHeaderClick(ByVal Column As GridEX20.JSColumn)
    GridEXHelper.ColumnHeaderClick Me.GridEX1, Column
End Sub

Private Sub GridEX1_DblClick()
    GridEX1_SelectionChange
    Unload Me
End Sub

Private Sub GridEX1_SelectionChange()
    Set Ot = ots.item(Me.GridEX1.RowIndex(Me.GridEX1.row))
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Set Ot = ots.item(RowIndex)

    Values(1) = Ot.id
    Values(2) = Ot.descripcion
    Values(3) = Ot.FechaEntrega
    Values(4) = funciones.estado_pedido(Ot.estado)
    Values(5) = Ot.Anticipo
    Values(6) = Ot.cliente.razon

End Sub
