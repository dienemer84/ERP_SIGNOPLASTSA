VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmAdminCobranzasSeleccionarReciboAnticipo 
   Caption         =   "Recibos de Anticipo"
   ClientHeight    =   6330
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9915
   LinkTopic       =   "Form1"
   ScaleHeight     =   6330
   ScaleWidth      =   9915
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.GroupBox GroupBox2 
      Height          =   855
      Left            =   120
      TabIndex        =   0
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
         TabIndex        =   1
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
         TabIndex        =   2
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
         Left            =   2520
         TabIndex        =   3
         Top             =   240
         Width           =   4935
         _Version        =   786432
         _ExtentX        =   8705
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Seleccione el Recibo y luego Acepte para aplicar al Comprobante."
      End
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   3840
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   9690
      _ExtentX        =   17092
      _ExtentY        =   6773
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      PreviewRowLines =   1
      ColumnAutoResize=   -1  'True
      MethodHoldFields=   -1  'True
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   6
      Column(1)       =   "frmAdminCobranzasRemitosAnticipo.frx":0000
      Column(2)       =   "frmAdminCobranzasRemitosAnticipo.frx":0164
      Column(3)       =   "frmAdminCobranzasRemitosAnticipo.frx":02C4
      Column(4)       =   "frmAdminCobranzasRemitosAnticipo.frx":0428
      Column(5)       =   "frmAdminCobranzasRemitosAnticipo.frx":0568
      Column(6)       =   "frmAdminCobranzasRemitosAnticipo.frx":06B0
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminCobranzasRemitosAnticipo.frx":07F8
      FormatStyle(2)  =   "frmAdminCobranzasRemitosAnticipo.frx":0930
      FormatStyle(3)  =   "frmAdminCobranzasRemitosAnticipo.frx":09E0
      FormatStyle(4)  =   "frmAdminCobranzasRemitosAnticipo.frx":0A94
      FormatStyle(5)  =   "frmAdminCobranzasRemitosAnticipo.frx":0B6C
      FormatStyle(6)  =   "frmAdminCobranzasRemitosAnticipo.frx":0C24
      ImageCount      =   0
      PrinterProperties=   "frmAdminCobranzasRemitosAnticipo.frx":0D04
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   855
      Left            =   120
      TabIndex        =   5
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
         TabIndex        =   6
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
         TabIndex        =   7
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
         TabIndex        =   8
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
         TabIndex        =   9
         Top             =   375
         Width           =   585
      End
   End
End
Attribute VB_Name = "frmAdminCobranzasSeleccionarReciboAnticipo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Cliente As clsCliente
Private rcbos As New Collection
Private ots As New Collection
Private Ot As OrdenTrabajo
Private ReciboA As Recibo
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

    If IsSomething(Cliente) Then
        'Me.GroupBox1.Enabled = False
        Me.cboClientes.ListIndex = funciones.PosIndexCbo(Cliente.Id, Me.cboClientes)
        Me.cboClientes.Enabled = False
        'Me.CMDsINCliente.Enabled = False
    End If

    'Me.GridEX1.Columns(5).Visible = MostrarAnticipo
    'Me.GridEX1.ItemCount = 0
    
    Me.caption = "Lista de Recibos de Anticipos (" & Name & ")"

' SE COMPLETA EL GRID AUTOMATICAMENTE AL CARGAR EL FORM
    Set Ot = Nothing
    
    llenarLista
    
End Sub

Private Sub llenarLista()
'    q = "{pedido}.{activo}=1"
    
'    If Me.cboClientes.ListIndex <> -1 Then
'        q = q & " AND {pedido}.idClienteFacturar = " & Me.cboClientes.ItemData(Me.cboClientes.ListIndex)
'    End If
'
'    If MostrarAnticipo Then
'        q = q & " and {pedido}." & DAOOrdenTrabajo.CAMPO_ANTICIPO_FACTURADO & "=0 and {pedido}." & DAOOrdenTrabajo.CAMPO_ANTICIPO & ">0"
'    End If
'
'    q = Replace$(q, "{pedido}", DAOOrdenTrabajo.TABLA_PEDIDO)
'    q = Replace$(q, "{activo}", DAOOrdenTrabajo.CAMPO_ACTIVO)

    If Me.cboClientes.ListIndex <> -1 Then
        q = "idCliente = " & Me.cboClientes.ItemData(Me.cboClientes.ListIndex) & " AND rec.pagoACuenta = 1 AND rec.estado=2"
    End If
    
    
    
    Set rcbos = DAORecibo.FindAll(q)
    
    Me.GridEX1.ItemCount = 0
    
    Me.GridEX1.ItemCount = rcbos.count
    
    If Me.GridEX1.ItemCount = 0 Then
        'Me.LabeSinResultados.caption = "No hay resultados para mostrar..."
        
    End If
    


End Sub

Private Sub GridEX1_ColumnHeaderClick(ByVal Column As GridEX20.JSColumn)
    GridEXHelper.ColumnHeaderClick Me.GridEX1, Column
    
End Sub

Private Sub GridEX1_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    
    Set ReciboA = rcbos.item(rowIndex)
    
    Values(1) = ReciboA.Id
    Values(2) = Format(ReciboA.FEcha, "yyyy/mm/dd", vbSunday)
    Values(3) = Format(ReciboA.FechaCreacion, "yyyy/mm/dd", vbSunday)
    Values(6) = ReciboA.moneda.NombreCorto
    Values(4) = Replace(FormatCurrency(funciones.FormatearDecimales(ReciboA.ACuentaDisponible)), "$", "")
    Values(5) = enums.EnumEstadoRecibo(ReciboA.estado)

End Sub

Private Sub PushButtonAceptar_Click()
    If IsSomething(Ot) Then
        Set Ot = ots.item(Me.GridEX1.rowIndex(Me.GridEX1.row))
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

