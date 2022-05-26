VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmPlaneamientoPedidosSeleccion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " "
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9960
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmPlaneamientoPedidosSeleccion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   9960
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.GroupBox GroupBox2 
      Height          =   855
      Left            =   120
      TabIndex        =   6
      Top             =   5280
      Width           =   9735
      _Version        =   786432
      _ExtentX        =   17171
      _ExtentY        =   1508
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.PushButton PushButtonCancelar 
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1575
         _Version        =   786432
         _ExtentX        =   2778
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Cerrar"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButtonAceptar 
         Height          =   495
         Left            =   8040
         TabIndex        =   7
         Top             =   240
         Width           =   1575
         _Version        =   786432
         _ExtentX        =   2778
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Aceptar"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   495
         Left            =   2400
         TabIndex        =   9
         Top             =   240
         Width           =   4935
         _Version        =   786432
         _ExtentX        =   8705
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Seleccione la OT y luego Acepte para ingresarla en el Comprobante."
      End
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   3840
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   9690
      _ExtentX        =   17092
      _ExtentY        =   6773
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
      ColumnsCount    =   7
      Column(1)       =   "frmPlaneamientoPedidosSeleccion.frx":000C
      Column(2)       =   "frmPlaneamientoPedidosSeleccion.frx":0168
      Column(3)       =   "frmPlaneamientoPedidosSeleccion.frx":0298
      Column(4)       =   "frmPlaneamientoPedidosSeleccion.frx":03F8
      Column(5)       =   "frmPlaneamientoPedidosSeleccion.frx":0540
      Column(6)       =   "frmPlaneamientoPedidosSeleccion.frx":0694
      Column(7)       =   "frmPlaneamientoPedidosSeleccion.frx":07B4
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmPlaneamientoPedidosSeleccion.frx":08C4
      FormatStyle(2)  =   "frmPlaneamientoPedidosSeleccion.frx":09FC
      FormatStyle(3)  =   "frmPlaneamientoPedidosSeleccion.frx":0AAC
      FormatStyle(4)  =   "frmPlaneamientoPedidosSeleccion.frx":0B60
      FormatStyle(5)  =   "frmPlaneamientoPedidosSeleccion.frx":0C38
      FormatStyle(6)  =   "frmPlaneamientoPedidosSeleccion.frx":0CF0
      ImageCount      =   0
      PrinterProperties=   "frmPlaneamientoPedidosSeleccion.frx":0DD0
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9735
      _Version        =   786432
      _ExtentX        =   17171
      _ExtentY        =   1508
      _StockProps     =   79
      Caption         =   "Parámetros de búsqueda"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.PushButton cmdActualizarBusqueda 
         Height          =   495
         Left            =   8040
         TabIndex        =   5
         Top             =   240
         Width           =   1500
         _Version        =   786432
         _ExtentX        =   2646
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Actualizar"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton CMDsINCliente 
         Height          =   285
         Left            =   6345
         TabIndex        =   4
         Top             =   345
         Visible         =   0   'False
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
         Top             =   330
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
   Begin XtremeSuiteControls.Label LabeSinResultados 
      Height          =   375
      Left            =   7080
      TabIndex        =   10
      Top             =   4920
      Width           =   2655
      _Version        =   786432
      _ExtentX        =   4683
      _ExtentY        =   661
      _StockProps     =   79
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
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

'Private Sub cmdBuscar_Click()
'    llenarLista
'End Sub

Private Sub cmdActualizarBusqueda_Click()
    llenarLista
End Sub

Private Sub CMDsINCliente_Click()
' ESTE BOTON QUEDA DESACTIVADO
    Me.cboClientes.ListIndex = -1
End Sub

Private Sub Form_Load()
    Customize Me
    GridEXHelper.CustomizeGrid Me.GridEX1, True, False
    DAOCliente.llenarComboXtremeSuite Me.cboClientes, False, True, False
    
    'CMDsINCliente_Click

    If IsSomething(cliente) Then
        'Me.GroupBox1.Enabled = False
        Me.cboClientes.ListIndex = funciones.PosIndexCbo(cliente.Id, Me.cboClientes)
        Me.cboClientes.Enabled = False
        'Me.CMDsINCliente.Enabled = False
    End If

    Me.GridEX1.Columns(5).Visible = MostrarAnticipo
    Me.GridEX1.ItemCount = 0
    
    Me.caption = "Lista OT (" & Name & ")"

' SE COMPLETA EL GRID AUTOMATICAMENTE AL CARGAR EL FORM
    Set Ot = Nothing
    
    llenarLista
    
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
    
    If Me.GridEX1.ItemCount = 0 Then
        Me.LabeSinResultados.caption = "No hay resultados para mostrar..."
        
    End If
    


End Sub

Private Sub GridEX1_ColumnHeaderClick(ByVal Column As GridEX20.JSColumn)
    GridEXHelper.ColumnHeaderClick Me.GridEX1, Column
    
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Set Ot = ots.item(RowIndex)

    Values(1) = Ot.Id
    Values(2) = Ot.descripcion
    Values(3) = Ot.FechaEntrega
    Values(4) = funciones.estado_pedido(Ot.estado)
    Values(5) = Ot.Anticipo
    Values(6) = Ot.moneda.NombreCorto & "- " & Ot.moneda.NombreLargo
    Values(7) = Ot.cliente.razon

End Sub

Private Sub PushButtonAceptar_Click()
    If IsSomething(Ot) Then
        Set Ot = ots.item(Me.GridEX1.RowIndex(Me.GridEX1.row))
        Set Selecciones.OrdenTrabajo = Ot
    Else
        Set Selecciones.OrdenTrabajo = Nothing
    End If
    
    Unload Me
    
End Sub

Private Sub PushButtonCancelar_Click()
    Set Selecciones.OrdenTrabajo = Nothing
    
    Unload Me
    
End Sub
