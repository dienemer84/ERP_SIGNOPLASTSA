VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmAdminPagosLiquidaciondeCajaCrear 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Crear Liquidación de Caja"
   ClientHeight    =   10545
   ClientLeft      =   5940
   ClientTop       =   1470
   ClientWidth     =   17280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10545
   ScaleWidth      =   17280
   Begin XtremeSuiteControls.GroupBox grpCbtesConfirmados 
      Height          =   5655
      Index           =   1
      Left            =   8760
      TabIndex        =   38
      Top             =   2160
      Width           =   8535
      _Version        =   786432
      _ExtentX        =   15055
      _ExtentY        =   9975
      _StockProps     =   79
      Caption         =   "2- Comprobantes Confirmados a Pagar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Begin VB.TextBox txtInstruccionDos 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   44
         Text            =   "frmAdminPagosCrearLiquidaciondeCaja.frx":0000
         Top             =   240
         Width           =   8175
      End
      Begin XtremeSuiteControls.PushButton btnSacarTodos 
         Height          =   495
         Left            =   2040
         TabIndex        =   39
         Top             =   4920
         Width           =   1695
         _Version        =   786432
         _ExtentX        =   2990
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Quitar todos"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnSacarSeleccionado 
         Height          =   495
         Left            =   120
         TabIndex        =   40
         Top             =   4920
         Width           =   1695
         _Version        =   786432
         _ExtentX        =   2990
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Quitar seleccionado"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ListBox lstFacturasFiltradas 
         Height          =   3015
         Left            =   120
         TabIndex        =   41
         Top             =   1680
         Width           =   8205
         _Version        =   786432
         _ExtentX        =   14473
         _ExtentY        =   5318
         _StockProps     =   77
         BackColor       =   -2147483643
         Appearance      =   6
         Style           =   1
      End
      Begin XtremeSuiteControls.Label lblCantidadComprobantesConfirmados 
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   1320
         Width           =   3735
         _Version        =   786432
         _ExtentX        =   6588
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Cbtes. Confirmados: 0"
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
   Begin XtremeSuiteControls.GroupBox grpFiltros 
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8535
      _Version        =   786432
      _ExtentX        =   15055
      _ExtentY        =   3413
      _StockProps     =   79
      Caption         =   "Parametros de Búsqueda"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.PushButton btnLimpiarProveedor 
         Height          =   375
         Left            =   4800
         TabIndex        =   36
         Top             =   480
         Width           =   375
         _Version        =   786432
         _ExtentX        =   661
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "X"
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnLimpiarNúmero 
         Height          =   375
         Left            =   2520
         TabIndex        =   35
         Top             =   1200
         Width           =   375
         _Version        =   786432
         _ExtentX        =   661
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "X"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cboProveedor 
         Height          =   360
         Left            =   120
         TabIndex        =   33
         Top             =   480
         Width           =   4575
         _Version        =   786432
         _ExtentX        =   8070
         _ExtentY        =   635
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Text            =   "cboProveedor"
      End
      Begin XtremeSuiteControls.PushButton btnFiltrarResultados 
         Height          =   495
         Left            =   6000
         TabIndex        =   29
         Top             =   360
         Width           =   2295
         _Version        =   786432
         _ExtentX        =   4048
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Mostrar comprobantes"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
      End
      Begin VB.TextBox txtFiltroNumero 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   1200
         Width           =   2295
      End
      Begin VB.Label lblProveedor 
         Caption         =   "Filtro por Proveedor"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label lblNúmero 
         Caption         =   "Filtro por Número de comprobante"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   960
         Width           =   4095
      End
   End
   Begin XtremeSuiteControls.GroupBox grpGuardar 
      Height          =   1935
      Index           =   0
      Left            =   8760
      TabIndex        =   22
      Top             =   120
      Width           =   8580
      _Version        =   786432
      _ExtentX        =   15134
      _ExtentY        =   3413
      _StockProps     =   79
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Begin VB.TextBox txtNumerodeLiquidacion 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   45
         Top             =   502
         Width           =   2295
      End
      Begin XtremeSuiteControls.DateTimePicker dtpFecha 
         Height          =   375
         Left            =   2760
         TabIndex        =   25
         Top             =   502
         Width           =   1365
         _Version        =   786432
         _ExtentX        =   2408
         _ExtentY        =   661
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   1
         CurrentDate     =   40183.7263657407
      End
      Begin XtremeSuiteControls.PushButton btnGuardar 
         Height          =   495
         Left            =   6000
         TabIndex        =   26
         Top             =   1200
         Width           =   2295
         _Version        =   786432
         _ExtentX        =   4048
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Guardar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
      End
      Begin VB.Label lblNumero 
         Caption         =   "Número de Liquidación:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   46
         Top             =   210
         Width           =   2415
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2760
         TabIndex        =   27
         Tag             =   "Total: "
         Top             =   240
         Width           =   600
      End
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         Caption         =   "Total Pagos: $ 0,00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   24
         Tag             =   "Total: "
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label lblTotalFacturas 
         AutoSize        =   -1  'True
         Caption         =   "Total Comprobantes: $ 0,00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   23
         Top             =   1440
         Width           =   2370
      End
   End
   Begin VB.TextBox txtOtrosDescuentos 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   17760
      TabIndex        =   21
      Top             =   11760
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.TextBox txtDifCambioNG1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   19080
      TabIndex        =   20
      Top             =   11760
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.TextBox txtDifCambioTOTAL1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   20280
      TabIndex        =   19
      Top             =   11760
      Visible         =   0   'False
      Width           =   960
   End
   Begin XtremeSuiteControls.GroupBox grpOrigen 
      Height          =   2535
      Left            =   120
      TabIndex        =   2
      Top             =   7920
      Width           =   17175
      _Version        =   786432
      _ExtentX        =   30295
      _ExtentY        =   4471
      _StockProps     =   79
      Caption         =   "3- Valores de pago"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.TabControl TabControl 
         Height          =   2100
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   16740
         _Version        =   786432
         _ExtentX        =   29527
         _ExtentY        =   3704
         _StockProps     =   68
         Appearance      =   10
         Color           =   32
         PaintManager.ShowIcons=   -1  'True
         ItemCount       =   2
         SelectedItem    =   1
         Item(0).Caption =   "Banco"
         Item(0).ControlCount=   2
         Item(0).Control(0)=   "gridCompensatorios"
         Item(0).Control(1)=   "gridDepositosOperaciones"
         Item(1).Caption =   "Caja"
         Item(1).ControlCount=   1
         Item(1).Control(0)=   "gridCajaOperaciones"
         Begin GridEX20.GridEX gridCompensatorios 
            Height          =   4710
            Left            =   -1.39895e5
            TabIndex        =   4
            Top             =   435
            Visible         =   0   'False
            Width           =   9330
            _ExtentX        =   16457
            _ExtentY        =   8308
            Version         =   "2.0"
            BoundColumnIndex=   ""
            ReplaceColumnIndex=   ""
            PreviewColumn   =   "observacion"
            PreviewRowLines =   1
            ColumnAutoResize=   -1  'True
            MethodHoldFields=   -1  'True
            ContScroll      =   -1  'True
            AllowColumnDrag =   0   'False
            AllowDelete     =   -1  'True
            GroupByBoxVisible=   0   'False
            RowHeaders      =   -1  'True
            DataMode        =   99
            ColumnHeaderHeight=   285
            IntProp1        =   0
            IntProp2        =   0
            IntProp7        =   0
            ColumnsCount    =   5
            Column(1)       =   "frmAdminPagosCrearLiquidaciondeCaja.frx":0012
            Column(2)       =   "frmAdminPagosCrearLiquidaciondeCaja.frx":015A
            Column(3)       =   "frmAdminPagosCrearLiquidaciondeCaja.frx":0266
            Column(4)       =   "frmAdminPagosCrearLiquidaciondeCaja.frx":0352
            Column(5)       =   "frmAdminPagosCrearLiquidaciondeCaja.frx":0456
            FormatStylesCount=   6
            FormatStyle(1)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":0596
            FormatStyle(2)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":06CE
            FormatStyle(3)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":077E
            FormatStyle(4)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":0832
            FormatStyle(5)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":090A
            FormatStyle(6)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":09C2
            ImageCount      =   0
            PrinterProperties=   "frmAdminPagosCrearLiquidaciondeCaja.frx":0AA2
         End
         Begin GridEX20.GridEX gridDepositosOperaciones 
            Height          =   1455
            Left            =   -69880
            TabIndex        =   37
            Top             =   480
            Visible         =   0   'False
            Width           =   9210
            _ExtentX        =   16245
            _ExtentY        =   2566
            Version         =   "2.0"
            BoundColumnIndex=   ""
            ReplaceColumnIndex=   ""
            ColumnAutoResize=   -1  'True
            MethodHoldFields=   -1  'True
            ContScroll      =   -1  'True
            AllowDelete     =   -1  'True
            GroupByBoxVisible=   0   'False
            RowHeaders      =   -1  'True
            DataMode        =   99
            AllowAddNew     =   -1  'True
            ColumnHeaderHeight=   285
            IntProp1        =   0
            IntProp2        =   0
            IntProp7        =   0
            ColumnsCount    =   5
            Column(1)       =   "frmAdminPagosCrearLiquidaciondeCaja.frx":0C7A
            Column(2)       =   "frmAdminPagosCrearLiquidaciondeCaja.frx":0DDA
            Column(3)       =   "frmAdminPagosCrearLiquidaciondeCaja.frx":0F16
            Column(4)       =   "frmAdminPagosCrearLiquidaciondeCaja.frx":104A
            Column(5)       =   "frmAdminPagosCrearLiquidaciondeCaja.frx":118E
            FormatStylesCount=   6
            FormatStyle(1)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":1292
            FormatStyle(2)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":13CA
            FormatStyle(3)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":147A
            FormatStyle(4)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":152E
            FormatStyle(5)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":1606
            FormatStyle(6)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":16BE
            ImageCount      =   0
            PrinterProperties=   "frmAdminPagosCrearLiquidaciondeCaja.frx":179E
         End
         Begin GridEX20.GridEX gridCajaOperaciones 
            Height          =   1455
            Left            =   120
            TabIndex        =   47
            Top             =   480
            Width           =   9210
            _ExtentX        =   16245
            _ExtentY        =   2566
            Version         =   "2.0"
            BoundColumnIndex=   ""
            ReplaceColumnIndex=   ""
            ColumnAutoResize=   -1  'True
            MethodHoldFields=   -1  'True
            ContScroll      =   -1  'True
            AllowDelete     =   -1  'True
            GroupByBoxVisible=   0   'False
            RowHeaders      =   -1  'True
            DataMode        =   99
            AllowAddNew     =   -1  'True
            ColumnHeaderHeight=   285
            IntProp1        =   0
            IntProp2        =   0
            IntProp7        =   0
            ColumnsCount    =   5
            Column(1)       =   "frmAdminPagosCrearLiquidaciondeCaja.frx":1976
            Column(2)       =   "frmAdminPagosCrearLiquidaciondeCaja.frx":1AD6
            Column(3)       =   "frmAdminPagosCrearLiquidaciondeCaja.frx":1C12
            Column(4)       =   "frmAdminPagosCrearLiquidaciondeCaja.frx":1D46
            Column(5)       =   "frmAdminPagosCrearLiquidaciondeCaja.frx":1E7A
            FormatStylesCount=   6
            FormatStyle(1)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":1F7E
            FormatStyle(2)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":20B6
            FormatStyle(3)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":2166
            FormatStyle(4)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":221A
            FormatStyle(5)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":22F2
            FormatStyle(6)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":23AA
            ImageCount      =   0
            PrinterProperties=   "frmAdminPagosCrearLiquidaciondeCaja.frx":248A
         End
      End
   End
   Begin XtremeSuiteControls.GroupBox grpCbtesImpagos 
      Height          =   5655
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   8565
      _Version        =   786432
      _ExtentX        =   15108
      _ExtentY        =   9975
      _StockProps     =   79
      Caption         =   "1- Comprobantes Impagos"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Begin VB.TextBox txtInstruccion 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   43
         Text            =   "frmAdminPagosCrearLiquidaciondeCaja.frx":2662
         Top             =   240
         Width           =   8175
      End
      Begin XtremeSuiteControls.PushButton btnConfirmarSeleccion 
         Height          =   495
         Left            =   6000
         TabIndex        =   32
         Top             =   4920
         Width           =   2295
         _Version        =   786432
         _ExtentX        =   4048
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Confirmar Comprobantes"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnDesseleccionarTodo 
         Height          =   495
         Left            =   2040
         TabIndex        =   31
         Top             =   4920
         Width           =   1695
         _Version        =   786432
         _ExtentX        =   2990
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Deseleccionar todo"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnSeleccionarTodo 
         Height          =   495
         Left            =   120
         TabIndex        =   30
         Top             =   4920
         Width           =   1695
         _Version        =   786432
         _ExtentX        =   2990
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Seleccionar todo"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ListBox lstFacturas 
         Height          =   3015
         Left            =   120
         TabIndex        =   6
         Top             =   1680
         Width           =   8205
         _Version        =   786432
         _ExtentX        =   14473
         _ExtentY        =   5318
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
         Style           =   1
      End
      Begin XtremeSuiteControls.Label lblCantidadCbtesSeleccionados 
         Height          =   255
         Left            =   6000
         TabIndex        =   8
         Top             =   1320
         Width           =   2295
         _Version        =   786432
         _ExtentX        =   4048
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Cbtes. Seleccionado: 0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label lblCantidadComprobantes 
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   2895
         _Version        =   786432
         _ExtentX        =   5106
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Total Comprobantes: 0"
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
   Begin GridEX20.GridEX gridBancos 
      Height          =   1845
      Left            =   1440
      TabIndex        =   9
      Top             =   10560
      Visible         =   0   'False
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   3254
      Version         =   "2.0"
      BoundColumnIndex=   "id"
      ReplaceColumnIndex=   "nombre"
      ActAsDropDown   =   -1  'True
      ColumnAutoResize=   -1  'True
      HideSelection   =   2
      MethodHoldFields=   -1  'True
      ContScroll      =   -1  'True
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      NewRowPos       =   1
      RowHeaders      =   -1  'True
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   2
      Column(1)       =   "frmAdminPagosCrearLiquidaciondeCaja.frx":2671
      Column(2)       =   "frmAdminPagosCrearLiquidaciondeCaja.frx":2771
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":2861
      FormatStyle(2)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":2999
      FormatStyle(3)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":2A49
      FormatStyle(4)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":2AFD
      FormatStyle(5)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":2BD5
      FormatStyle(6)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":2C8D
      ImageCount      =   0
      PrinterProperties=   "frmAdminPagosCrearLiquidaciondeCaja.frx":2D6D
   End
   Begin GridEX20.GridEX gridCuentasBancarias 
      Height          =   1935
      Left            =   10080
      TabIndex        =   10
      Top             =   10560
      Visible         =   0   'False
      Width           =   3825
      _ExtentX        =   6747
      _ExtentY        =   3413
      Version         =   "2.0"
      BoundColumnIndex=   "id"
      ReplaceColumnIndex=   "cuenta"
      ActAsDropDown   =   -1  'True
      HideSelection   =   2
      MethodHoldFields=   -1  'True
      ContScroll      =   -1  'True
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      NewRowPos       =   1
      RowHeaders      =   -1  'True
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   2
      Column(1)       =   "frmAdminPagosCrearLiquidaciondeCaja.frx":2F45
      Column(2)       =   "frmAdminPagosCrearLiquidaciondeCaja.frx":3069
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":315D
      FormatStyle(2)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":3295
      FormatStyle(3)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":3345
      FormatStyle(4)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":33F9
      FormatStyle(5)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":34D1
      FormatStyle(6)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":3589
      ImageCount      =   0
      PrinterProperties=   "frmAdminPagosCrearLiquidaciondeCaja.frx":3669
   End
   Begin GridEX20.GridEX gridMonedas 
      Height          =   1815
      Left            =   120
      TabIndex        =   11
      Top             =   10560
      Visible         =   0   'False
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   3201
      Version         =   "2.0"
      BoundColumnIndex=   "id"
      ReplaceColumnIndex=   "moneda"
      ActAsDropDown   =   -1  'True
      HideSelection   =   2
      MethodHoldFields=   -1  'True
      ContScroll      =   -1  'True
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      NewRowPos       =   1
      RowHeaders      =   -1  'True
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   2
      Column(1)       =   "frmAdminPagosCrearLiquidaciondeCaja.frx":3841
      Column(2)       =   "frmAdminPagosCrearLiquidaciondeCaja.frx":3965
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":3A59
      FormatStyle(2)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":3B91
      FormatStyle(3)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":3C41
      FormatStyle(4)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":3CF5
      FormatStyle(5)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":3DCD
      FormatStyle(6)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":3E85
      ImageCount      =   0
      PrinterProperties=   "frmAdminPagosCrearLiquidaciondeCaja.frx":3F65
   End
   Begin GridEX20.GridEX gridCajas 
      Height          =   1935
      Left            =   16080
      TabIndex        =   12
      Top             =   10560
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   3413
      Version         =   "2.0"
      BoundColumnIndex=   "id"
      ReplaceColumnIndex=   "caja"
      ActAsDropDown   =   -1  'True
      HideSelection   =   2
      MethodHoldFields=   -1  'True
      ContScroll      =   -1  'True
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      ColumnHeaders   =   0   'False
      NewRowPos       =   1
      RowHeaders      =   -1  'True
      DataMode        =   99
      ColumnHeaderHeight=   285
      ColumnsCount    =   2
      Column(1)       =   "frmAdminPagosCrearLiquidaciondeCaja.frx":413D
      Column(2)       =   "frmAdminPagosCrearLiquidaciondeCaja.frx":423D
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":4329
      FormatStyle(2)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":4461
      FormatStyle(3)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":4511
      FormatStyle(4)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":45C5
      FormatStyle(5)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":469D
      FormatStyle(6)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":4755
      ImageCount      =   0
      PrinterProperties=   "frmAdminPagosCrearLiquidaciondeCaja.frx":4835
   End
   Begin GridEX20.GridEX gridChequesDisponibles 
      Height          =   1920
      Left            =   4920
      TabIndex        =   13
      Top             =   10560
      Visible         =   0   'False
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   3387
      Version         =   "2.0"
      BoundColumnIndex=   "id"
      ReplaceColumnIndex=   "numero"
      ActAsDropDown   =   -1  'True
      ColumnAutoResize=   -1  'True
      HideSelection   =   2
      MethodHoldFields=   -1  'True
      ContScroll      =   -1  'True
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      RowHeaders      =   -1  'True
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   7
      Column(1)       =   "frmAdminPagosCrearLiquidaciondeCaja.frx":4A0D
      Column(2)       =   "frmAdminPagosCrearLiquidaciondeCaja.frx":4B8D
      Column(3)       =   "frmAdminPagosCrearLiquidaciondeCaja.frx":4D2D
      Column(4)       =   "frmAdminPagosCrearLiquidaciondeCaja.frx":4E69
      Column(5)       =   "frmAdminPagosCrearLiquidaciondeCaja.frx":4F75
      Column(6)       =   "frmAdminPagosCrearLiquidaciondeCaja.frx":5095
      Column(7)       =   "frmAdminPagosCrearLiquidaciondeCaja.frx":51A1
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":5295
      FormatStyle(2)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":53CD
      FormatStyle(3)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":547D
      FormatStyle(4)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":5531
      FormatStyle(5)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":5609
      FormatStyle(6)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":56C1
      ImageCount      =   0
      PrinterProperties=   "frmAdminPagosCrearLiquidaciondeCaja.frx":57A1
   End
   Begin GridEX20.GridEX gridChequeras 
      Height          =   1935
      Left            =   8640
      TabIndex        =   14
      Top             =   10560
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   3413
      Version         =   "2.0"
      BoundColumnIndex=   "id"
      ReplaceColumnIndex=   "chequera"
      ActAsDropDown   =   -1  'True
      ColumnAutoResize=   -1  'True
      HideSelection   =   2
      MethodHoldFields=   -1  'True
      ContScroll      =   -1  'True
      AllowColumnDrag =   0   'False
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      ColumnHeaders   =   0   'False
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   2
      Column(1)       =   "frmAdminPagosCrearLiquidaciondeCaja.frx":5979
      Column(2)       =   "frmAdminPagosCrearLiquidaciondeCaja.frx":5A99
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":5B99
      FormatStyle(2)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":5CD1
      FormatStyle(3)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":5D81
      FormatStyle(4)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":5E35
      FormatStyle(5)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":5F0D
      FormatStyle(6)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":5FC5
      ImageCount      =   0
      PrinterProperties=   "frmAdminPagosCrearLiquidaciondeCaja.frx":60A5
   End
   Begin GridEX20.GridEX gridChequesChequera 
      Height          =   1935
      Left            =   14040
      TabIndex        =   15
      Top             =   10560
      Width           =   1980
      _ExtentX        =   3493
      _ExtentY        =   3413
      Version         =   "2.0"
      HoldSortSettings=   -1  'True
      BoundColumnIndex=   "id"
      ReplaceColumnIndex=   "nro"
      ActAsDropDown   =   -1  'True
      ColumnAutoResize=   -1  'True
      HideSelection   =   2
      MethodHoldFields=   -1  'True
      ContScroll      =   -1  'True
      AllowColumnDrag =   0   'False
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      ColumnHeaders   =   0   'False
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   2
      Column(1)       =   "frmAdminPagosCrearLiquidaciondeCaja.frx":627D
      Column(2)       =   "frmAdminPagosCrearLiquidaciondeCaja.frx":63AD
      SortKeysCount   =   1
      SortKey(1)      =   "frmAdminPagosCrearLiquidaciondeCaja.frx":64AD
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":6515
      FormatStyle(2)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":664D
      FormatStyle(3)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":66FD
      FormatStyle(4)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":67B1
      FormatStyle(5)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":6889
      FormatStyle(6)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":6941
      ImageCount      =   0
      PrinterProperties=   "frmAdminPagosCrearLiquidaciondeCaja.frx":6A21
   End
   Begin XtremeSuiteControls.RadioButton radioFacturaProveedor 
      Height          =   210
      Left            =   17760
      TabIndex        =   16
      Top             =   10920
      Visible         =   0   'False
      Width           =   2760
      _Version        =   786432
      _ExtentX        =   4868
      _ExtentY        =   370
      _StockProps     =   79
      Caption         =   "Seleccione Proveedor"
      Appearance      =   6
      Value           =   -1  'True
   End
   Begin XtremeSuiteControls.ComboBox cboMonedas 
      Height          =   315
      Left            =   18645
      TabIndex        =   17
      Top             =   11280
      Visible         =   0   'False
      Width           =   1245
      _Version        =   786432
      _ExtentX        =   2196
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Style           =   2
      Text            =   "cboMonedas"
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Moneda"
      Height          =   195
      Left            =   17760
      TabIndex        =   18
      Tag             =   "Total: "
      Top             =   11340
      Visible         =   0   'False
      Width           =   570
   End
End
Attribute VB_Name = "frmAdminPagosLiquidaciondeCajaCrear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements ISuscriber
Private id_susc As String
Dim formLoading As Boolean
Dim formLoaded As Boolean
Dim alicuotas As New Collection

Dim vFactElegida As clsFacturaProveedor
Dim vFacturaProveedor As clsFacturaProveedor
Dim colFacturas As New Collection

Dim prov As clsProveedor
Dim Factura As clsFacturaProveedor

'Dim compe As Compensatorio

Private Banco As Banco
Private caja As caja
Private CuentaBancaria As CuentaBancaria
Private moneda As clsMoneda
Private cuentasBancarias As New Collection
Private Monedas As New Collection
Private Cajas As New Collection
Private bancos As New Collection
Private chequesDisponibles As New Collection
Private chequeras As New Collection

Dim compe As Compensatorio

Private LiquidacionCaja As New clsLiquidacionCaja
Private operacion As operacion
Private cheque As cheque
Private tmpChequera As chequera

Private chequesChequeraSeleccionada As New Collection

Public ReadOnly As Boolean
Public EsNueva As Boolean

'Private Sub Command_Click()
'    Cargar (LiquidacionCaja)
'End Sub

Private Sub Form_Load()

    formLoading = True

    Me.gridCompensatorios.ItemCount = 0
    id_susc = funciones.CreateGUID
    Channel.AgregarSuscriptor Me, PasajeChequePropioCartera
    FormHelper.Customize Me

    GridEXHelper.CustomizeGrid Me.gridCajaOperaciones, False, True
    GridEXHelper.CustomizeGrid Me.gridDepositosOperaciones, False, True
    '    GridEXHelper.CustomizeGrid Me.gridCheques, False, True
    GridEXHelper.CustomizeGrid Me.gridChequesDisponibles, False, False
    GridEXHelper.CustomizeGrid Me.gridBancos, False, False
    GridEXHelper.CustomizeGrid Me.gridCuentasBancarias, False, False
    GridEXHelper.CustomizeGrid Me.gridMonedas, False, False
    GridEXHelper.CustomizeGrid Me.gridCajas, False, False
    GridEXHelper.CustomizeGrid Me.gridChequeras, False, False
    '    GridEXHelper.CustomizeGrid Me.gridChequesPropios, False, True
    GridEXHelper.CustomizeGrid Me.gridCompensatorios, False, True
    GridEXHelper.CustomizeGrid Me.gridChequesChequera
    '    GridEXHelper.CustomizeGrid Me.gridRetenciones, False, True

    Set Cajas = DAOCaja.FindAll()
    Me.gridCajas.ItemCount = Cajas.count

    Set Monedas = DAOMoneda.GetAll()
    Me.gridMonedas.ItemCount = Monedas.count

    Set cuentasBancarias = DAOCuentaBancaria.FindAll()
    Me.gridCuentasBancarias.ItemCount = cuentasBancarias.count

    Set bancos = DAOBancos.GetAll()
    Me.gridBancos.ItemCount = bancos.count

    Set chequeras = DAOChequeras.FindAllWithChequesDisponibles()
    Me.gridChequeras.ItemCount = chequeras.count

    CargarChequesDisponibles

    radioFacturaProveedor_Click

    Me.gridCajaOperaciones.ItemCount = LiquidacionCaja.OperacionesCaja.count

    Me.gridDepositosOperaciones.ItemCount = LiquidacionCaja.OperacionesBanco.count

    Set Me.gridDepositosOperaciones.Columns("moneda").DropDownControl = Me.gridMonedas
    Set Me.gridDepositosOperaciones.Columns("cuenta").DropDownControl = Me.gridCuentasBancarias

    Set Me.gridCajaOperaciones.Columns("moneda").DropDownControl = Me.gridMonedas
    Set Me.gridCajaOperaciones.Columns("caja").DropDownControl = Me.gridCajas

    gridChequesChequera.ItemCount = 0
    GridEXHelper.AutoSizeColumns Me.gridChequeras

    DAOMoneda.llenarComboXtremeSuite Me.cboMonedas

    Me.dtpFecha.value = LiquidacionCaja.FEcha

    'lstFacturas_Click
    Totalizar

    Me.txtInstruccion.text = "Se muestran los Comprobantes Impagos." & vbCrLf & "Para realizar el pago, debe seleccionar el Comprobante clickeando en el Check correspondiente." & vbCrLf & "Puede ir confirmando el Comprobante para que se guarde en la lista de Comprobantes Confirmados clickeando en el Botón: Confirmar Comprobantes."
    Me.txtInstruccionDos.text = "Se muestran los Comprobantes Confirmados." & vbCrLf & "La sumatoria de los importes se verá reflejado en el totalizado superior Total Comprobantes" & vbCrLf & "Puede borrar alguno o todos los Comprobantes que se fueron agregando."

    EsNueva = True

    formLoaded = True
    formLoading = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Channel.RemoverSuscripcionTotal Me
End Sub

Public Sub Cargar(liq As clsLiquidacionCaja)

    Me.txtDifCambioNG1.Enabled = Not ReadOnly
    Me.txtDifCambioTOTAL1.Enabled = Not ReadOnly

    Me.radioFacturaProveedor.Enabled = Not ReadOnly

    Me.gridDepositosOperaciones.AllowEdit = Not ReadOnly
    Me.gridDepositosOperaciones.AllowDelete = Not ReadOnly

    Me.gridBancos.AllowEdit = Not ReadOnly
    Me.gridCajaOperaciones.AllowEdit = Not ReadOnly
    Me.gridCajaOperaciones.AllowDelete = Not ReadOnly

    Me.gridCajas.AllowEdit = Not ReadOnly
    Me.gridChequeras.AllowEdit = Not ReadOnly
    Me.gridChequesChequera.AllowEdit = Not ReadOnly
    Me.gridChequesDisponibles.AllowEdit = Not ReadOnly
    Me.lblNúmero.Enabled = Not ReadOnly
    Me.cboMonedas.Enabled = Not ReadOnly
    Me.dtpFecha.Enabled = Not ReadOnly
    Me.btnGuardar.Enabled = Not ReadOnly
    Me.btnFiltrarResultados.Enabled = Not ReadOnly
    Me.btnConfirmarSeleccion.Enabled = Not ReadOnly
    Me.btnDesseleccionarTodo.Enabled = Not ReadOnly
    Me.btnSacarSeleccionado.Enabled = Not ReadOnly
    Me.btnSacarTodos.Enabled = Not ReadOnly
    Me.btnSeleccionarTodo.Enabled = Not ReadOnly
    Me.txtFiltroNumero.Enabled = Not ReadOnly
    Me.lblNumero.Enabled = Not ReadOnly
    Me.btnLimpiarNúmero.Enabled = Not ReadOnly
    Me.lblCantidadCbtesSeleccionados.Visible = Not ReadOnly
    Me.lblCantidadComprobantes.Visible = Not ReadOnly
    Me.Label1.Enabled = Not ReadOnly
    Me.Label2.Enabled = Not ReadOnly
    Me.grpFiltros.Enabled = Not ReadOnly

    Me.txtInstruccion.Enabled = Not ReadOnly
    Me.txtInstruccionDos.Enabled = Not ReadOnly
    Me.txtNumerodeLiquidacion.Enabled = Not ReadOnly
    Me.grpCbtesImpagos.Enabled = Not ReadOnly
    Me.grpCbtesConfirmados(1).Enabled = Not ReadOnly
    Me.txtOtrosDescuentos.Enabled = Not ReadOnly
    Me.lstFacturas.Enabled = Not ReadOnly
    Me.lstFacturasFiltradas.Enabled = Not ReadOnly
    
    
    If Not IsSomething(liq) Then
        MsgBox "La Liquidación que está intentando visualizar está en estado PENDIENTE. " & vbNewLine & "Por lo tanto no puede ser mostrada porque puede estar siendo editada." & vbNewLine & "Verifiquelo por favor.", vbCritical, "OP Pendiente"
        Unload Me
        Exit Sub

    End If


    Set LiquidacionCaja = DAOLiquidacionCaja.FindById(liq.Id)
    Set LiquidacionCaja.Compensatorios = DAOCompensatorios.FindByOP(LiquidacionCaja.Id)

    Dim i As Long
    Dim j As Long
    With LiquidacionCaja

        If .EsParaFacturaProveedor Then
            radioFacturaProveedor.value = True

            If .FacturasProveedor.count > 0 Then

                MostrarFacturas


                Dim idx As Integer
                idx = -1
                For i = 1 To .FacturasProveedor.count
                    For j = 0 To Me.lstFacturas.ListCount - 1
                        If Me.lstFacturas.ItemData(j) = .FacturasProveedor.item(i).Id Then
                            Me.lstFacturas.Checked(j) = True
                            idx = i
                        End If
                    Next j
                Next i

                'acaa

                If ReadOnly Then
                    For j = Me.lstFacturas.ListCount - 1 To 0 Step -1
                        If Not Me.lstFacturas.Checked(j) Then
                            Me.lstFacturas.RemoveItem j
                        End If
                    Next j

                    'Me.lblCantidadComprobantes.caption = Me.lblCantidadCbtesSeleccionados.caption

                End If
            End If

        Else

        End If

        If idx >= 0 Then
            lstFacturas.ListIndex = lstFacturas.ListCount - 1

        End If

        btnConfirmarSeleccion_Click

        Me.gridCajaOperaciones.ItemCount = .OperacionesCaja.count
        Me.gridDepositosOperaciones.ItemCount = .OperacionesBanco.count

        Me.cboMonedas.ListIndex = funciones.PosIndexCbo(.moneda.Id, Me.cboMonedas)
        Me.dtpFecha.value = .FEcha
        Me.txtOtrosDescuentos.text = .OtrosDescuentos

    End With
    mostrarCompensatorios

    Me.caption = "Liquidación Nº " & LiquidacionCaja.NumeroLiq

    Me.txtNumerodeLiquidacion = LiquidacionCaja.NumeroLiq

    Totalizar

    EsNueva = False

End Sub

Public Property Get FacturaProveedor(nvalue As clsFacturaProveedor)
    Set vFacturaProveedor = nvalue
End Property


Private Sub btnDesseleccionarTodo_Click()
    Dim i As Integer

    For i = 0 To Me.lstFacturas.ListCount - 1
        Me.lstFacturas.Checked(i) = False
    Next i
End Sub

Private Sub btnFiltrarResultados_Click()


    MostrarFacturas


End Sub
Private Sub btnGuardar_Click()

    If Me.gridCajaOperaciones.EditMode = jgexEditModeOn Then
        MsgBox "Todavia esta editando la grilla de caja.", vbExclamation
        Exit Sub
    End If

    If Me.gridDepositosOperaciones.EditMode = jgexEditModeOn Then
        MsgBox "Todavia esta editando la grilla de banco.", vbExclamation
        Exit Sub
    End If

    LiquidacionCaja.FEcha = Me.dtpFecha.value

    If Me.txtNumerodeLiquidacion.text = "" Then
        MsgBox ("El número de Liquidación no puede estar vacío.")
        Exit Sub
    Else
        LiquidacionCaja.NumeroLiq = Me.txtNumerodeLiquidacion.text

    End If

    Set LiquidacionCaja.FacturasProveedor = New Collection

    If Me.radioFacturaProveedor.value Then
        Dim T As Long
        For T = 0 To Me.lstFacturasFiltradas.ListCount - 1
            If Me.lstFacturasFiltradas.Checked(T) Then
                LiquidacionCaja.FacturasProveedor.Add colFacturas.item(CStr(Me.lstFacturasFiltradas.ItemData(T)))
            End If
        Next T
    End If

    If LiquidacionCaja.IsValid Then

        Dim n As Boolean: n = (LiquidacionCaja.Id = 0)

        If DAOLiquidacionCaja.Save(LiquidacionCaja, True) Then

            If n Then
                MsgBox "Liquidación de Caja Nº " & Me.txtNumerodeLiquidacion & " creada con exito.", vbInformation
            Else
                MsgBox "Liquidación de Caja modificada con exito.", vbInformation
            End If

            Dim EVENTO As New clsEventoObserver
            Set EVENTO.Elemento = LiquidacionCaja
            EVENTO.Tipo = LiquidacionCaja_
            Set EVENTO.Originador = Me

            If n Then
                EVENTO.EVENTO = agregar_
            Else
                EVENTO.EVENTO = modificar_
            End If
            Channel.Notificar EVENTO, LiquidacionCaja_

            If n Then
                If MsgBox("¿Desea crear una Liquidación de Caja nueva", vbQuestion + vbYesNo) = vbYes Then
                    Dim f12 As New frmAdminPagosLiquidaciondeCajaCrear
                    f12.Show
                End If
            End If

            Unload Me
        Else
            MsgBox "Hubo un problema al guardar la Liquidación.", vbCritical
        End If
    Else
        MsgBox LiquidacionCaja.ValidationMessages, vbCritical, "Error"
    End If

End Sub
Private Sub btnLimpiarNúmero_Click()
    Me.txtFiltroNumero = ""


End Sub

Private Sub btnLimpiarProveedor_Click()
    Me.cboProveedor.ListIndex = -1
End Sub

Private Sub btnSacarSeleccionado_Click()

    Dim indice As Integer

    If Me.lstFacturasFiltradas.ListIndex <> -1 Then    'Se verifica que se haya seleccionado un elemento
        indice = Me.lstFacturasFiltradas.ListIndex    'Se obtiene el índice del elemento seleccionado
        Me.lstFacturasFiltradas.RemoveItem indice    'Se elimina el elemento seleccionado
    Else
        MsgBox "Debe seleccionar un comprobante para eliminarlo.", vbExclamation, "Atención"
    End If

    Me.lblCantidadComprobantesConfirmados.caption = "Cbtes. Confirmados: " & Me.lstFacturasFiltradas.ListCount

    calcularTotalesCbtesFiltrados



End Sub

Private Sub btnSacarTodos_Click()
    If Me.lstFacturasFiltradas.ListCount > 0 Then
        Me.lstFacturasFiltradas.Clear    'Se eliminan todos los elementos del ListBox
    Else
        MsgBox "No hay combrobantes para eliminar.", vbExclamation, "Atención"
    End If

    Me.lblCantidadComprobantesConfirmados.caption = "Cbtes. Confirmados: " & Me.lstFacturasFiltradas.ListCount

    calcularTotalesCbtesFiltrados

End Sub

Private Sub btnSeleccionarTodo_Click()

    Dim i As Integer

    For i = 0 To Me.lstFacturas.ListCount - 1
        Me.lstFacturas.Checked(i) = True
    Next i

End Sub

Private Sub cboMonedas_Click()
    If Me.cboMonedas.ListIndex = -1 Then
        Set LiquidacionCaja.moneda = Nothing
    Else
        Set LiquidacionCaja.moneda = DAOMoneda.GetById(Me.cboMonedas.ItemData(Me.cboMonedas.ListIndex))
    End If

    Totalizar

End Sub

'Private Sub cmdMostrarDatosProveedor_Click()
'    If Me.cboProveedores.ListIndex <> -1 Then
'
'        Set prov = colProveedores.item(CStr(Me.cboProveedores.ItemData(Me.cboProveedores.ListIndex)))
'
'
'
'        Dim d As clsDTOPadronIIBB
'
'        Set d = DTOPadronIIBB.FindByCUIT(prov.Cuit, TipoPadronRetencion)
'
'        If IsSomething(d) Then
'            Me.txtRetenciones = str(d.alicuota)   ' Val(d.Retencion )
'        Else
'            Me.txtRetenciones = 0
'        End If
'
'        'If IsSomething(prov) Then
'        'Set alicuotas = DAORetenciones.FindAllWithAlicuotas(prov.Cuit)
'
'        ' End If
'    Else
'        Set prov = Nothing
'    End If
'
'
'    MostrarFacturas
'    MostrarDeudaCompensatorios
'    btnCargar_Click
'
'End Sub


Private Sub dtpFecha_Change()
    LiquidacionCaja.FEcha = Me.dtpFecha.value
End Sub

Private Sub CargarChequesDisponibles()
    Set chequesDisponibles = DAOCheques.FindAllEnCarteraDeTerceros
    Me.gridChequesDisponibles.ItemCount = chequesDisponibles.count
End Sub


Private Sub MostrarFacturas()

    Me.lstFacturas.Clear

    '    If IsSomething(prov) Then
    'Set colFacturas = DAOFacturaProveedor.FindAll("AdminComprasFacturasProveedores.estado=" & EstadoFacturaProveedor.Aprobada & " or AdminComprasFacturasProveedores.estado=" & EstadoFacturaProveedor.pagoParcial & ")", False, "", False, True)
    Set colFacturas = DAOFacturaProveedor.FindAll("AdminComprasFacturasProveedores.estado=" & EstadoFacturaProveedor.Aprobada)

    If LiquidacionCaja.Id <> 0 And LiquidacionCaja.EsParaFacturaProveedor Then
        '            If prov.Id = LiquidacionCaja.FacturasProveedor.item(1).Proveedor.Id Then
        For Each Factura In LiquidacionCaja.FacturasProveedor
            If Not funciones.BuscarEnColeccion(colFacturas, CStr(Factura.Id)) Then

                colFacturas.Add DAOFacturaProveedor.FindById(Factura.Id), CStr(Factura.Id)
            End If
        Next
        '            End If
    End If

    Dim T As String

    For Each Factura In colFacturas    'en ese for traigo los pendientes a abonar que estan asociados a ops sin aprobar

        Dim c As Collection
        Set c = DAOOrdenPago.FindAbonadoPendiente(Factura.Id, LiquidacionCaja.Id)

        Factura.TotalAbonadoGlobalPendiente = 0    ' c(1) 'que esta en ops sin aprobar
        Factura.NetoGravadoAbonadoGlobalPendiente = 0    ' c(2)
        Factura.OtrosAbonadoGlobalPendiente = 0    'c(3)

        T = Factura.NumeroFormateado & " (" & Factura.moneda.NombreCorto & " " & Factura.Total & ")" & " (" & Factura.FEcha & ")" & "  | " & UCase(Factura.Proveedor.RazonSocial)

        Me.lstFacturas.AddItem T
        Me.lstFacturas.ItemData(Me.lstFacturas.NewIndex) = Factura.Id


    Next

    ' 22/08/2022
    'AGREGO UN LABEL QUE MUESTRA LA CANTIDAD DE COMPROBANTES MOSTRADOS EN EL LIST

    Me.lblCantidadComprobantes.caption = "Cbtes. Mostrados: " & colFacturas.count

    '    Else
    '
    '        Set colFacturas = New Collection
    '
    '        'MsgBox (colFacturas.count)

    '    End If

End Sub



Private Sub gridBancos_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex <= bancos.count Then
        Set Banco = bancos.item(RowIndex)
        Values(1) = Banco.Id
        Values(2) = Banco.nombre
    End If

End Sub

Private Sub gridCajaOperaciones_BeforeUpdate(ByVal Cancel As GridEX20.JSRetBoolean)
    Dim cond1 As Boolean
    Dim cond2 As Boolean
    Dim cond3 As Boolean
    Dim cond4 As Boolean

    cond1 = Not IsNumeric(Me.gridCajaOperaciones.value(1))
    cond2 = Not IsNumeric(Me.gridCajaOperaciones.value(2)) And LenB(Me.gridCajaOperaciones.value(2)) = 0
    cond3 = Not IsDate(Me.gridCajaOperaciones.value(3))
    cond4 = LenB(Me.gridCajaOperaciones.value(4)) = 0 Or IsEmpty(Me.gridCajaOperaciones.value(4))    'or Not IsNumeric(Me.gridCajaOperaciones.value(4))

    Cancel = cond1 Or cond2 Or cond3 Or cond4

End Sub

Private Sub gridCajas_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex > 0 And Cajas.count > 0 Then
        Set caja = Cajas.item(RowIndex)
        Values(1) = caja.Id
        Values(2) = caja.nombre
    End If
End Sub

Private Sub gridChequeras_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex <= chequeras.count Then
        Set tmpChequera = chequeras.item(RowIndex)
        Values(1) = tmpChequera.Description
        Values(2) = tmpChequera.Id
    End If
End Sub

Private Sub gridChequesChequera_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex > 0 And chequesChequeraSeleccionada.count > 0 Then
        Values(1) = chequesChequeraSeleccionada(RowIndex).numero
        Values(2) = chequesChequeraSeleccionada(RowIndex).Id
    End If
End Sub

Private Sub gridChequesDisponibles_ColumnHeaderClick(ByVal Column As GridEX20.JSColumn)
    GridEXHelper.ColumnHeaderClick Me.gridChequesDisponibles, Column
End Sub

Private Sub gridChequesDisponibles_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex <= chequesDisponibles.count Then
        Set cheque = chequesDisponibles.item(RowIndex)
        Values(1) = cheque.numero
        'FORMATCURRENCY
        Values(2) = FormatCurrency(cheque.Monto)
        If IsSomething(cheque.moneda) Then Values(3) = cheque.moneda.NombreCorto
        If IsSomething(cheque.Banco) Then Values(4) = cheque.Banco.nombre
        Values(5) = cheque.Id
        Values(6) = cheque.OrigenCheque
        Values(7) = cheque.OrigenDestino

    End If

End Sub

Private Sub gridCompensatorios_UnboundDelete(ByVal RowIndex As Long, ByVal Bookmark As Variant)
    LiquidacionCaja.Compensatorios.remove (RowIndex)
End Sub

Private Sub gridCompensatorios_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)

    On Error Resume Next
    Set compe = LiquidacionCaja.Compensatorios.item(RowIndex)
    Values(1) = compe.Comprobante.NumeroFormateado
    Values(2) = TiposCompensatorio.item(CStr(compe.Tipo))
    'FORMATCURRENCY
    Values(3) = FormatCurrency(compe.Monto)
    Values(4) = compe.FechaCancelacion
    Values(5) = compe.Observacion

End Sub

Private Sub gridCuentasBancarias_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If cuentasBancarias.count >= RowIndex Then
        Set CuentaBancaria = cuentasBancarias.item(RowIndex)
        Values(1) = CuentaBancaria.Id
        Values(2) = CuentaBancaria.DescripcionFormateada
    End If
End Sub

Private Sub gridMonedas_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex > 0 And Monedas.count > 0 Then
        Set moneda = Monedas.item(RowIndex)
        Values(1) = moneda.Id
        Values(2) = moneda.NombreCorto
    End If
End Sub

Private Property Get ISuscriber_id() As String
    ISuscriber_id = id_susc
End Property

Private Function ISuscriber_Notificarse(EVENTO As clsEventoObserver) As Variant
    CargarChequesDisponibles
End Function

Private Sub MostrarPosiblesRetenciones(col As Collection, Optional colc As Collection = Nothing)
    Dim d As New Dictionary
    Dim ret As Retencion
    Dim colret As Collection
    Set colret = DAORetenciones.FindAllEsAgente
    Set d = DAOCertificadoRetencion.VerPosibleRetenciones2(col, alicuotas, Val(Me.txtDifCambioNG1), LiquidacionCaja.TotalNGCompensatorios)
    Dim totRet As Double

    totRet = 0

    If IsSomething(prov) Then


        For Each ret In colret
            totRet = totRet + d.item(CStr(ret.Id))
        Next ret

    End If


    totRet = funciones.RedondearDecimales(totRet)
    Dim c As Compensatorio
    Dim F As clsFacturaProveedor
    Dim totFact As Double
    Dim TotNG As Double
    Dim totFactHoy As Double
    '    Dim Cambio As Double
    Dim totCambio As Double
    Dim totCambiong As Double
    Dim totNGHoy As Double
    Dim totDeudaCompe As Double
    totDeudaCompe = 0
    For Each F In col

        totFact = totFact + F.Total

        totFactHoy = totFactHoy + MonedaConverter.ConvertirForzado2(IIf(F.tipoDocumentoContable = tipoDocumentoContable.notaCredito, F.TotalDiaPagoAbonado * -1, F.TotalDiaPagoAbonado), F.moneda.Id, LiquidacionCaja.moneda.Id, F.TipoCambioPago)

        TotNG = TotNG + MonedaConverter.ConvertirForzado2(IIf(F.tipoDocumentoContable = tipoDocumentoContable.notaCredito, F.NetoGravadoAbonado * -1, F.NetoGravadoAbonado), F.moneda.Id, LiquidacionCaja.moneda.Id, F.TipoCambioPago)
        totNGHoy = totNGHoy + MonedaConverter.ConvertirForzado2(IIf(F.tipoDocumentoContable = tipoDocumentoContable.notaCredito, F.NetoGravadoAbonadoDiaPago * -1, F.NetoGravadoAbonadoDiaPago), F.moneda.Id, LiquidacionCaja.moneda.Id, F.TipoCambioPago)
        totCambio = totCambio + MonedaConverter.ConvertirForzado2(IIf(F.tipoDocumentoContable = tipoDocumentoContable.notaCredito, F.DiferenciaPorTipoDeCambionTOTAL * -1, F.DiferenciaPorTipoDeCambionTOTAL), F.moneda.Id, LiquidacionCaja.moneda.Id, F.TipoCambioPago)
        totCambiong = totCambiong + MonedaConverter.ConvertirForzado2(IIf(F.tipoDocumentoContable = tipoDocumentoContable.notaCredito, F.DiferenciaPorTipoDeCambionNG * -1, F.DiferenciaPorTipoDeCambionNG), F.moneda.Id, LiquidacionCaja.moneda.Id, F.TipoCambioPago)

    Next F


    If IsSomething(colc) Then
        For Each c In colc

            Dim ff As clsFacturaProveedor

            Set ff = DAOFacturaProveedor.FindById(c.Comprobante.Id)
            totDeudaCompe = totDeudaCompe + MonedaConverter.ConvertirForzado2(IIf(c.Tipo = TC_Credito, c.Monto * -1, c.Monto), ff.moneda.Id, LiquidacionCaja.moneda.Id, ff.TipoCambioPago)

        Next

    End If

    Me.lblTotalFacturas = "Total Comprobantes: :" & FormatCurrency(funciones.FormatearDecimales(totFact))
    Me.lblTotal.caption = "Total valores cargados: " & FormatCurrency(funciones.FormatearDecimales(LiquidacionCaja.StaticTotalOrigenes + LiquidacionCaja.StaticTotalRetenido))

    LiquidacionCaja.StaticTotalFacturas = funciones.RedondearDecimales(totFact)
    LiquidacionCaja.staticTotalDeudaCompensatorios = funciones.RedondearDecimales(totDeudaCompe)

    LiquidacionCaja.StaticTotalFacturasNG = funciones.RedondearDecimales(TotNG + LiquidacionCaja.DiferenciaCambioEnNG)

    LiquidacionCaja.DiferenciaCambio = totCambio

    verCompensatorios

    LiquidacionCaja.StaticTotalRetenido = funciones.RedondearDecimales(totRet)

End Sub


Private Sub verCompensatorios()
'    Me.lblTotalCompensatorios = "Total compensatorios en " & FormatCurrency(funciones.FormatearDecimales(LiquidacionCaja.TotalCompensatorios))
End Sub

Private Sub MostrarPago(F As clsFacturaProveedor)

    If IsSomething(F) Then


        ' If F.ImporteTotalAbonado = 0 Then F.ImporteTotalAbonado = F.Total
        If F.NetoGravadoAbonado = 0 Then F.NetoGravadoAbonado = F.NetoGravado    '- F.NetoNoGravado  (2do cambio en fix 004)
        If F.OtrosAbonado = 0 Then F.OtrosAbonado = F.Total - F.NetoGravado    '- F.NetoNoGravado  (2do cambio en fix 004)

        '        Me.txtParcialAbonar = F.ImporteNetoGravadoSaldo ' F.NetoGravadoAbonado - F.NetoGravadoAbonadoGlobal
        '        Me.txtTotalParcialAbonar = F.ImporteTotalAbonado
        '        Me.txtOtrosParcialAbonar = F.ImporteOtrosSaldo  'F.OtrosAbonado - F.OtrosAbonadoGlobal

        'RecalcularTotalFacturaElegida

        '     vFactElegida.NetoGravadoAbonado = CDbl(Me.txtParcialAbonar)
        '    vFactElegida.ImporteTotalAbonado =    'vFactElegida.CalcularTotalAbonadoParcial(CDbl(Me.txtParcialAbonar))


        'esto debería calcular el total en base a las alícuotas de la factura


        If F.TotalAbonado + F.TotalAbonadoGlobal + F.TotalAbonadoGlobalPendiente > F.Total Then
            MsgBox "El importe que desea abonar, supera el monto total del comprobante seleccionado"
        End If
        'Me.txtnetogravadoabonado = F.NetoGravadoAbonado - F.NetoGravadoAbonadoGlobal
        ' Me.txtParcialAbonado = F.TotalAbonado - F.TotalAbonadoGlobal
    End If
    Totalizar
End Sub

Private Sub lstFacturas_Click()

'    'debug.print (Me.lstFacturas.ItemData(Me.lstFacturas.ListIndex))

    Set vFactElegida = colFacturas.item(CStr(Me.lstFacturas.ItemData(Me.lstFacturas.ListIndex)))


    '    'debug.print (vFactElegida.Id)


    If IsSomething(vFactElegida) Then

        Dim c As Collection

        'Me.lblCantidadCbtesSeleccionados.caption = "Cbtes. Seleccionados: " & c.count


        If LiquidacionCaja.estado = EstadoLiquidacionCaja_pendiente And vFactElegida.NetoGravadoAbonado = 0 And vFactElegida.OtrosAbonado = 0 Then
            Set c = DAOLiquidacionCaja.FindAbonadoFactura(vFactElegida.Id, LiquidacionCaja.Id)

            'vFactElegida.TotalAbonadoGlobalPendiente = c(1)
            vFactElegida.NetoGravadoAbonado = c(2)
            vFactElegida.OtrosAbonado = c(3)
        End If


        '    If vFactElegida.ImporteTotalAbonado = 0 Then
        '        vFactElegida.ImporteTotalAbonado = vFactElegida.TotalPendiente
        '
        '    End If

        MostrarPago vFactElegida
        '    RecalcularFacturaElegida
    End If
    Totalizar

End Sub

Private Sub lstFacturas_DblClick()
    Dim i As Long
    Dim change As Double
    Dim F As clsFacturaProveedor
    '    Dim col As New Collection
    For i = 0 To Me.lstFacturas.ListCount - 1
        If Me.lstFacturas.Selected(i) Then
            Set F = colFacturas.item(CStr(Me.lstFacturas.ItemData(i)))

            MostrarPago vFactElegida
        End If
    Next

    On Error GoTo err1
    change = InputBox("Establezca el tipo de cambio con el cual se va a abonar la factura", "Tipo de cambio", F.TipoCambioPago)


    If LenB(change) = 0 Then
        change = 1
    Else
        F.TipoCambioPago = change

    End If
    Totalizar
    Exit Sub



err1:
    Totalizar
    change = 1
End Sub

Sub calcularOrigenes()
    Dim i As Long
    Dim col As New Collection
    Dim colc As New Collection


    For i = 0 To Me.lstFacturas.ListCount - 1
        If Me.lstFacturas.Checked(i) Then

            If funciones.BuscarEnColeccion(colFacturas, CStr(Me.lstFacturas.ItemData(i))) Then

                col.Add colFacturas.item(CStr(Me.lstFacturas.ItemData(i)))

                Me.lblCantidadCbtesSeleccionados.caption = "Cbtes. Seleccionados: " & col.count

            End If
        End If
    Next i

    TotalizarDiferenciasCambio

    MostrarPosiblesRetenciones col, colc


End Sub

Sub contadorSeleccionados()
    Dim i As Long
    Dim col As New Collection

    For i = 0 To Me.lstFacturas.ListCount - 1
        If Me.lstFacturas.Checked(i) Then

            If funciones.BuscarEnColeccion(colFacturas, CStr(Me.lstFacturas.ItemData(i))) Then

                col.Add colFacturas.item(CStr(Me.lstFacturas.ItemData(i)))

                Me.lblCantidadCbtesSeleccionados.caption = "Cbtes. Seleccionados: " & col.count

            End If
        End If
    Next i

End Sub


Private Sub btnConfirmarSeleccion_Click()
    Dim i As Integer
    Dim j As Integer
    Dim facturas As String
    Dim FacturaExiste As Boolean

    For i = 0 To Me.lstFacturas.ListCount - 1
        If Me.lstFacturas.Checked(i) Then
            idFactura = Me.lstFacturas.ItemData(i)
            facturas = Me.lstFacturas.list(i)
            FacturaExiste = False    'asumimos que la factura no existe en la lista

            'buscamos si la factura ya existe en la lista
            For j = 0 To Me.lstFacturasFiltradas.ListCount - 1
                If InStr(1, Me.lstFacturasFiltradas.list(j), facturas, vbTextCompare) > 0 Then
                    FacturaExiste = True
                    MsgBox ("El Comprobante " & facturas & " ya existe en la lista confirmada.")


                    'Me.lstFacturas.ItemChecked(i) = False

                    Exit For
                End If
            Next j

            'Agregar la nueva factura solo si no existe en la lista
            If Not FacturaExiste Then
                Me.lstFacturasFiltradas.AddItem facturas
                Me.lstFacturasFiltradas.ItemData(Me.lstFacturasFiltradas.NewIndex) = idFactura
                Me.lstFacturasFiltradas.Checked(Me.lstFacturasFiltradas.ListCount - 1) = True
                calcularTotalesCbtesFiltrados
            End If
        End If
    Next i

    btnDesseleccionarTodo_Click

    lstFacturasFiltradas.ListIndex = lstFacturasFiltradas.ListCount - 1

    '    calcularTotalesCbtesFiltrados
    Me.lblCantidadComprobantesConfirmados.caption = "Cbtes. Confirmados: " & Me.lstFacturasFiltradas.ListCount
    Me.lblCantidadCbtesSeleccionados.caption = "Cbtes. Seleccionados: 0"


    Me.txtFiltroNumero.text = ""
    
    If Me.txtFiltroNumero.Enabled = True Then
        Me.txtFiltroNumero.SetFocus
    End If

End Sub


Sub calcularTotalesCbtesFiltrados()

    Dim i As Long
    Dim col As New Collection
    Dim colCheckeadosDos As New Collection

    For i = 0 To Me.lstFacturasFiltradas.ListCount - 1

        If Me.lstFacturasFiltradas.Checked(i) Then

            If funciones.BuscarEnColeccion(colFacturas, CStr(Me.lstFacturasFiltradas.ItemData(i))) Then

                col.Add colFacturas.item(CStr(Me.lstFacturasFiltradas.ItemData(i)))
                colCheckeadosDos.Add colFacturas.item(CStr(Me.lstFacturasFiltradas.ItemData(i)))

            End If
        End If

    Next i

    Me.lblCantidadCbtesSeleccionados.caption = "Cbtes. Seleccionados: " & colCheckeadosDos.count
    '    Debug.Print (colCheckeadosDos.count)


    MostrarPosiblesRetenciones col

    TotalizarImportesCbtesFiltrados col

End Sub

'Sub calcularTotalesCbtesFiltradosNuevo()
'
'    Dim i As Long
'    Dim col As New Collection
'    Dim colCheckeadosDos As New Collection
'
'    For i = 0 To Me.lstFacturasFiltradas.ListCount - 1
'
'        If Me.lstFacturasFiltradas.Checked(i) Then
'
'            If funciones.BuscarEnColeccion(colFacturas, CStr(Me.lstFacturasFiltradas.ItemData(i))) Then
'
'                col.Add colFacturas.item(CStr(Me.lstFacturasFiltradas.ItemData(i)))
'                colCheckeadosDos.Add colFacturas.item(CStr(Me.lstFacturasFiltradas.ItemData(i)))
''           Else
''                colCheckeadosDos.Add colFacturas.item(CStr(Me.lstFacturasFiltradas.ItemData(i)))
''            End If
'
'        End If
'
'
''    Next i
'
'    Me.lblCantidadCbtesSeleccionados.caption = "Cbtes. Seleccionados: " & colCheckeadosDos.count
'
'    MostrarPosiblesRetenciones col
'
'    TotalizarImportesCbtesFiltrados col
'
'End Sub

Private Sub TotalizarImportesCbtesFiltrados(col As Collection)

    Dim F As clsFacturaProveedor
    Dim totFact As Double
    Dim totFactHoy As Double

    For Each F In col

        totFact = totFact + F.Total

        totFactHoy = totFactHoy + MonedaConverter.ConvertirForzado2(IIf(F.tipoDocumentoContable = tipoDocumentoContable.notaCredito, F.TotalDiaPagoAbonado * -1, F.TotalDiaPagoAbonado), F.moneda.Id, LiquidacionCaja.moneda.Id, F.TipoCambioPago)

    Next F

    Me.lblTotalFacturas = "Total Comprobantes: " & FormatCurrency(funciones.FormatearDecimales(totFact))


End Sub


Sub limpiarParciales()
    Me.lblCantidadCbtesSeleccionados.caption = "Cbtes. Seleccionados: 0"
End Sub


Private Sub lstFacturas_ItemCheck(ByVal item As Long)

    If item < -1 Then
        Dim f1
        Set f1 = DAOFacturaProveedor.FindById(CStr(Me.lstFacturas.ItemData(item)))

    End If

    contadorSeleccionados

End Sub

Private Sub lstFacturasFiltradas_AddItem(ByVal ItemData As String)
' Tu código aquí para ejecutar la función
' Puedes acceder al valor agregado utilizando el parámetro ItemData
    MsgBox "Se agregó el valor: " & ItemData
End Sub

Private Sub lstFacturasFiltradas_Click()
' Tu código aquí para ejecutar la función
' Puedes acceder al valor agregado utilizando el parámetro ItemData
'    TotalizarImportesCbtesFiltrados col
End Sub

Private Sub lstFacturasFiltradas_ItemCheck(ByVal item As Long)


    If item < -1 Then
        Dim f1
        Set f1 = DAOFacturaProveedor.FindById(CStr(Me.lstFacturas.ItemData(item)))

    End If


    Me.lblCantidadCbtesSeleccionados.caption = "Cbtes. Seleccionados: 0"
    'calcularOrigenes


    Dim i As Long
    Dim col As New Collection
    Dim colFiltradas As New Collection

    For i = 0 To Me.lstFacturasFiltradas.ListCount - 1

        If Me.lstFacturasFiltradas.Checked(i) Then

            If funciones.BuscarEnColeccion(colFacturas, CStr(Me.lstFacturasFiltradas.ItemData(i))) Then

                col.Add colFacturas.item(CStr(Me.lstFacturasFiltradas.ItemData(i)))

                colFiltradas.Add colFacturas.item(CStr(Me.lstFacturasFiltradas.ItemData(i)))

                '                Me.lblCantidadCbtesSeleccionados.caption = "Cbtes. Seleccionados: " & col.count



            End If
        End If
    Next i

    TotalizarImportesCbtesFiltrados colFiltradas

End Sub

Private Sub mostrarCompensatorios()
    Me.gridCompensatorios.ItemCount = LiquidacionCaja.Compensatorios.count
    verCompensatorios
End Sub

Private Sub LimpiarFacturasYValores()
    Set colFacturas = New Collection
End Sub

Private Sub ActivarControles()
'    Me.cboProveedores.Enabled = Me.radioFacturaProveedor.value
'    Me.lstFacturas.Enabled = Me.radioFacturaProveedor.value

'    Me.cboCuentas.Enabled = Me.radioConcepto.value
'    Me.txtDetalle.Enabled = Me.radioConcepto.value

'    Me.txtRetenciones.text = 0

'    If Not Me.cboProveedores.Enabled Then Me.cboProveedores.ListIndex = -1
    If Not Me.lstFacturas.Enabled Then Me.lstFacturas.Clear
    '
    '    If Not Me.cboCuentas.Enabled Then Me.cboCuentas.ListIndex = -1
    '    If Not Me.txtDetalle.Enabled Then Me.txtDetalle.text = vbNullString


End Sub

'Private Sub PushButton_Click()
'    calcularTotalesCbtesFiltradosNuevo
'End Sub

'Private Sub PushButton_Click()
'    calcularOrigenes
'End Sub

Private Sub radioFacturaProveedor_Click()
    If formLoaded Then
        LimpiarFacturasYValores
        MostrarPosiblesRetenciones New Collection
        Totalizar
    End If
    ActivarControles
End Sub

Private Sub gridCajaOperaciones_UnboundAddNew(ByVal NewRowBookmark As GridEX20.JSRetVariant, ByVal Values As GridEX20.JSRowData)
    Set operacion = New operacion
    'operacion.IdPertenencia = recibo.Id
    operacion.Pertenencia = OrigenOperacion.caja
    operacion.Monto = Values(1)
    operacion.Comprobante = Values(5)
    If IsNumeric(Values(2)) Then
        Set operacion.moneda = DAOMoneda.GetById(Values(2))
    End If
    operacion.FechaOperacion = Values(3)
    If IsNumeric(Values(4)) Then
        Set operacion.caja = DAOCaja.FindById(Values(4))
    End If
    operacion.EntradaSalida = OPSalida
    LiquidacionCaja.OperacionesCaja.Add operacion
    Totalizar
End Sub

Private Sub gridCajaOperaciones_UnboundDelete(ByVal RowIndex As Long, ByVal Bookmark As Variant)
    If RowIndex > 0 And LiquidacionCaja.OperacionesCaja.count >= RowIndex Then
        LiquidacionCaja.OperacionesCaja.remove RowIndex
        Totalizar
    End If
End Sub

Private Sub Totalizar()


    LiquidacionCaja.StaticTotalOrigenes = LiquidacionCaja.TotalOrigenes
    'Me.lblTotalFacturas = "Total Comprobantes: 0"
    Me.lblTotal.caption = "Total valores cargados: " & FormatCurrency(funciones.FormatearDecimales(LiquidacionCaja.StaticTotalOrigenes + LiquidacionCaja.StaticTotalRetenido))
    GridEXHelper.AutoSizeColumns Me.gridCajaOperaciones
    GridEXHelper.AutoSizeColumns Me.gridDepositosOperaciones
    '    GridEXHelper.AutoSizeColumns Me.gridCheques
    'GridEXHelper.AutoSizeColumns Me.gridChequesPropios
    lstFacturas_ItemCheck -1
    'Me.lblCantidadCbtesSeleccionados.caption = "Cbtes. Seleecionados: 0"

    TotalizarDiferenciasCambio



End Sub

Private Function TotalizarDiferenciasCambio()
    Dim F As clsFacturaProveedor
    Dim col As New Collection
    Dim i As Long
    Dim T As Double
    Dim TIVA As Double
    Dim TTOTAL As Double
    For i = 0 To Me.lstFacturas.ListCount - 1
        If Me.lstFacturas.Checked(i) Then

            If funciones.BuscarEnColeccion(colFacturas, CStr(Me.lstFacturas.ItemData(i))) Then
                col.Add colFacturas.item(CStr(Me.lstFacturas.ItemData(i)))
            End If
        End If
    Next



    For Each F In col
        T = T + F.DiferenciaPorTipoDeCambionNG
        TIVA = TIVA + F.DiferenciaPorTipoDeCambionIVA
        TTOTAL = TTOTAL + F.DiferenciaPorTipoDeCambionTOTAL
    Next

    If ReadOnly Then
        Dim s As String
        s = LiquidacionCaja.DiferenciaCambioEnNG
        '        Me.txtDifCambioNG1.text = s
        s = LiquidacionCaja.DiferenciaCambioEnTOTAL
        '        Me.txtDifCambioTOTAL1.text = s
    End If

End Function

Private Sub gridCajaOperaciones_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex <= LiquidacionCaja.OperacionesCaja.count Then
        Set operacion = LiquidacionCaja.OperacionesCaja.item(RowIndex)
        'FORMATCURRENCY
        Values(1) = FormatCurrency(funciones.FormatearDecimales(operacion.Monto))
        If IsSomething(operacion.moneda) Then
            Values(2) = operacion.moneda.NombreCorto
        End If
        Values(3) = operacion.FechaOperacion
        If IsSomething(operacion.caja) Then
            Values(4) = operacion.caja.nombre
        End If
        If IsSomething(operacion) Then
            Values(5) = operacion.Comprobante
        End If
    End If
End Sub

Private Sub gridCajaOperaciones_UnboundUpdate(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex > 0 And LiquidacionCaja.OperacionesCaja.count > 0 Then
        Set operacion = LiquidacionCaja.OperacionesCaja.item(RowIndex)
        'operacion.IdPertenencia = recibo.id
        'operacion.Pertenencia = Banco
        operacion.Monto = Values(1)
        operacion.Comprobante = Values(5)
        If IsNumeric(Values(2)) Then
            Set operacion.moneda = DAOMoneda.GetById(Values(2))
        End If
        operacion.FechaOperacion = Values(3)
        If IsNumeric(Values(4)) Then
            Set operacion.caja = DAOCaja.FindById(Values(4))
        End If
        operacion.EntradaSalida = OPSalida
        Totalizar
    End If
End Sub

Private Sub gridDepositosOperaciones_BeforeUpdate(ByVal Cancel As GridEX20.JSRetBoolean)

    Dim cond1 As Boolean
    Dim cond2 As Boolean
    Dim cond3 As Boolean
    Dim cond4 As Boolean


    cond1 = Not IsNumeric(Me.gridDepositosOperaciones.value(1))
    cond2 = Not IsNumeric(Me.gridDepositosOperaciones.value(2)) And LenB(Me.gridDepositosOperaciones.value(2)) = 0
    cond3 = Not IsDate(Me.gridDepositosOperaciones.value(3))
    cond4 = Not IsNumeric(Me.gridDepositosOperaciones.value(4)) And LenB(Me.gridDepositosOperaciones.value(4)) = 0

    Cancel = cond1 Or cond2 Or cond3 Or cond4
End Sub

Private Sub gridDepositosOperaciones_UnboundAddNew(ByVal NewRowBookmark As GridEX20.JSRetVariant, ByVal Values As GridEX20.JSRowData)
    Set operacion = New operacion
    'operacion.IdPertenencia = recibo.Id
    operacion.Pertenencia = OrigenOperacion.Banco
    operacion.Monto = Values(1)
    operacion.Comprobante = Values(5)
    If IsNumeric(Values(2)) Then
        Set operacion.moneda = DAOMoneda.GetById(Values(2))
    End If
    operacion.FechaOperacion = Values(3)
    If IsNumeric(Values(4)) Then
        Set operacion.CuentaBancaria = DAOCuentaBancaria.FindById(Values(4))
    End If
    operacion.EntradaSalida = OPSalida
    LiquidacionCaja.OperacionesBanco.Add operacion
    Totalizar
End Sub

Private Sub gridDepositosOperaciones_UnboundDelete(ByVal RowIndex As Long, ByVal Bookmark As Variant)
    If RowIndex > 0 And LiquidacionCaja.OperacionesBanco.count >= RowIndex Then
        LiquidacionCaja.OperacionesBanco.remove RowIndex
        Totalizar
    End If
End Sub

Private Sub gridDepositosOperaciones_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex <= LiquidacionCaja.OperacionesBanco.count Then
        Set operacion = LiquidacionCaja.OperacionesBanco.item(RowIndex)
        'FORMATCURRENCY
        Values(1) = FormatCurrency(funciones.FormatearDecimales(operacion.Monto))
        If IsSomething(operacion.moneda) Then
            Values(2) = operacion.moneda.NombreCorto
        End If
        Values(3) = operacion.FechaOperacion
        If IsSomething(operacion.CuentaBancaria) Then
            Values(4) = operacion.CuentaBancaria.DescripcionFormateada
        End If
        If IsSomething(operacion) Then
            Values(5) = operacion.Comprobante
        End If
    End If
End Sub

Private Sub gridDepositosOperaciones_UnboundUpdate(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex > 0 And LiquidacionCaja.OperacionesBanco.count > 0 Then
        Set operacion = LiquidacionCaja.OperacionesBanco.item(RowIndex)
        'operacion.IdPertenencia = recibo.id
        'operacion.Pertenencia = Banco
        operacion.Monto = Values(1)
        operacion.Comprobante = Values(5)
        If IsNumeric(Values(2)) Then
            Set operacion.moneda = DAOMoneda.GetById(Values(2))
        End If
        operacion.FechaOperacion = Values(3)
        If IsNumeric(Values(4)) Then
            Set operacion.CuentaBancaria = DAOCuentaBancaria.FindById(Values(4))
        End If
        operacion.EntradaSalida = OPSalida
        Totalizar
    End If
End Sub

Private Sub TabControl_SelectedChanged(ByVal item As Xtremesuitecontrols.ITabControlItem)
    Me.TabControl.TabIndex = 0
End Sub

Private Sub txtDifCambioNG1_Change()
    LiquidacionCaja.DiferenciaCambioEnNG = Val(Me.txtDifCambioNG1)
    Totalizar
End Sub

Private Sub txtDifCambioTOTAL1_Change()
    LiquidacionCaja.DiferenciaCambioEnTOTAL = Val(Me.txtDifCambioTOTAL1)
    Totalizar
End Sub


Private Sub txtFiltroNumero_Change()
    Dim filterText As String
    filterText = Trim(Me.txtFiltroNumero.text)

    Me.lstFacturas.Clear

    For Each Factura In colFacturas
        ' Aplica el filtro visualmente en base al texto ingresado en el TextBox
        If InStr(1, Factura.NumeroFormateado, filterText, vbTextCompare) > 0 Then
            Dim displayText As String
            displayText = Factura.NumeroFormateado & " (" & Factura.moneda.NombreCorto & " " & Factura.Total & ") " & "(" & Factura.FEcha & ") | " & UCase(Factura.Proveedor.RazonSocial)

            If Factura.TotalAbonadoGlobal + Factura.TotalAbonadoGlobalPendiente > 0 Then
                displayText = Factura.NumeroFormateado & " (" & Factura.moneda.NombreCorto & " " & Factura.Total & " - Abonado: " & Factura.TotalAbonadoGlobal + Factura.TotalAbonadoGlobalPendiente & ") " & "(" & Factura.FEcha & ")"
            End If

            Me.lstFacturas.AddItem displayText
            Me.lstFacturas.ItemData(Me.lstFacturas.NewIndex) = Factura.Id
        End If
    Next

    Me.lblCantidadComprobantes.caption = "Cbtes. Mostrados: " & Me.lstFacturas.ListCount
End Sub


'''Private Sub txtFiltroNumero_Change()
'''
'''    Me.lstFacturas.Clear
'''    Dim condition As String
'''    condition = " 1 = 1 "
'''
'''    If LenB(Me.txtFiltroNumero) > 0 Then
'''        condition = condition & " AND AdminComprasFacturasProveedores.numero_factura like '%" & Trim(Me.txtFiltroNumero.text) & "%'"
'''    End If
'''
'''    If cboProveedor.ListIndex > -1 Then
'''        condition = condition & " AND AdminComprasFacturasProveedores.id_proveedor = " & cboProveedor.ItemData(Me.cboProveedor.ListIndex)
'''    End If
'''
'''
''''    Dim estado As String
'''
'''
'''
'''    If 1 = 1 Then
'''        Set colFacturas = DAOFacturaProveedor.FindAll(condition & " AND AdminComprasFacturasProveedores.estado = " & EstadoFacturaProveedor.Aprobada & "", False, "proveedores.razon ASC", False, True)
'''
'''        Dim T As String
'''
'''        For Each Factura In colFacturas    'en ese for traigo los pendientes a abonar que estan asociados a ops sin aprobar
'''
'''            T = Factura.NumeroFormateado & " (" & Factura.moneda.NombreCorto & " " & Factura.Total & ")" & " (" & Factura.FEcha & ") | " & UCase(Factura.Proveedor.RazonSocial)
'''
'''            If Factura.TotalAbonadoGlobal + Factura.TotalAbonadoGlobalPendiente > 0 Then
'''                T = Factura.NumeroFormateado & " (" & Factura.moneda.NombreCorto & " " & Factura.Total & " - Abonado: " & Factura.TotalAbonadoGlobal + Factura.TotalAbonadoGlobalPendiente & ")" & " (" & Factura.FEcha & ")"
'''            End If
'''
'''            Me.lstFacturas.AddItem T
'''            Me.lstFacturas.ItemData(Me.lstFacturas.NewIndex) = Factura.Id
'''        Next
'''
'''        Me.lblCantidadComprobantes.caption = "Cbtes. Mostrados: " & colFacturas.count
'''
'''
'''
'''    Else
'''        Set colFacturas = New Collection
'''            TotalizarImportesCbtesFiltrados colFacturas
'''    End If
'''End Sub

Private Sub txtOtrosDescuentos_LostFocus()
    LiquidacionCaja.OtrosDescuentos = Val(Me.txtOtrosDescuentos.text)
    Totalizar
End Sub


