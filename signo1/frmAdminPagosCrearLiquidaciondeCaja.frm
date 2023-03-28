VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmAdminPagosLiquidaciondeCajaCrear 
   Caption         =   "Crear LiquidaciÃ³n de Caja"
   ClientHeight    =   11070
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   23010
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11070
   ScaleWidth      =   23010
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtOtrosDescuentos 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   120
      TabIndex        =   30
      Top             =   8520
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.TextBox txtDifCambioNG1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1200
      TabIndex        =   29
      Top             =   8520
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.TextBox txtDifCambioTOTAL1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2400
      TabIndex        =   28
      Top             =   8520
      Visible         =   0   'False
      Width           =   960
   End
   Begin XtremeSuiteControls.GroupBox grpOrigen 
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   4320
      Width           =   11100
      _Version        =   786432
      _ExtentX        =   19579
      _ExtentY        =   5741
      _StockProps     =   79
      Caption         =   "Valores"
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
         Height          =   2820
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   10740
         _Version        =   786432
         _ExtentX        =   18944
         _ExtentY        =   4974
         _StockProps     =   68
         Appearance      =   10
         Color           =   32
         PaintManager.ShowIcons=   -1  'True
         ItemCount       =   2
         Item(0).Caption =   "Banco"
         Item(0).ControlCount=   2
         Item(0).Control(0)=   "gridDepositosOperaciones"
         Item(0).Control(1)=   "gridCompensatorios"
         Item(1).Caption =   "Caja"
         Item(1).ControlCount=   1
         Item(1).Control(0)=   "gridCajaOperaciones"
         Begin GridEX20.GridEX gridDepositosOperaciones 
            Height          =   2070
            Left            =   120
            TabIndex        =   2
            Top             =   600
            Width           =   10410
            _ExtentX        =   18362
            _ExtentY        =   3651
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
            Column(1)       =   "frmAdminPagosCrearLiquidaciondeCaja.frx":0000
            Column(2)       =   "frmAdminPagosCrearLiquidaciondeCaja.frx":0160
            Column(3)       =   "frmAdminPagosCrearLiquidaciondeCaja.frx":029C
            Column(4)       =   "frmAdminPagosCrearLiquidaciondeCaja.frx":03D0
            Column(5)       =   "frmAdminPagosCrearLiquidaciondeCaja.frx":0514
            FormatStylesCount=   6
            FormatStyle(1)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":0618
            FormatStyle(2)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":0750
            FormatStyle(3)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":0800
            FormatStyle(4)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":08B4
            FormatStyle(5)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":098C
            FormatStyle(6)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":0A44
            ImageCount      =   0
            PrinterProperties=   "frmAdminPagosCrearLiquidaciondeCaja.frx":0B24
         End
         Begin GridEX20.GridEX gridCajaOperaciones 
            Height          =   2055
            Left            =   -69880
            TabIndex        =   3
            Top             =   600
            Visible         =   0   'False
            Width           =   10410
            _ExtentX        =   18362
            _ExtentY        =   3625
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
            Column(1)       =   "frmAdminPagosCrearLiquidaciondeCaja.frx":0CFC
            Column(2)       =   "frmAdminPagosCrearLiquidaciondeCaja.frx":0E5C
            Column(3)       =   "frmAdminPagosCrearLiquidaciondeCaja.frx":0F98
            Column(4)       =   "frmAdminPagosCrearLiquidaciondeCaja.frx":10CC
            Column(5)       =   "frmAdminPagosCrearLiquidaciondeCaja.frx":1200
            FormatStylesCount=   6
            FormatStyle(1)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":1304
            FormatStyle(2)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":143C
            FormatStyle(3)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":14EC
            FormatStyle(4)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":15A0
            FormatStyle(5)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":1678
            FormatStyle(6)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":1730
            ImageCount      =   0
            PrinterProperties=   "frmAdminPagosCrearLiquidaciondeCaja.frx":1810
         End
         Begin GridEX20.GridEX gridCompensatorios 
            Height          =   4710
            Left            =   -69895
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
            Column(1)       =   "frmAdminPagosCrearLiquidaciondeCaja.frx":19E8
            Column(2)       =   "frmAdminPagosCrearLiquidaciondeCaja.frx":1B30
            Column(3)       =   "frmAdminPagosCrearLiquidaciondeCaja.frx":1C3C
            Column(4)       =   "frmAdminPagosCrearLiquidaciondeCaja.frx":1D28
            Column(5)       =   "frmAdminPagosCrearLiquidaciondeCaja.frx":1E2C
            FormatStylesCount=   6
            FormatStyle(1)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":1F6C
            FormatStyle(2)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":20A4
            FormatStyle(3)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":2154
            FormatStyle(4)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":2208
            FormatStyle(5)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":22E0
            FormatStyle(6)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":2398
            ImageCount      =   0
            PrinterProperties=   "frmAdminPagosCrearLiquidaciondeCaja.frx":2478
         End
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox2 
      Height          =   2895
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   11085
      _Version        =   786432
      _ExtentX        =   19553
      _ExtentY        =   5106
      _StockProps     =   79
      Caption         =   "Mostrar Facturas"
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
      Begin VB.TextBox txtBuscarFactura 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   5010
      End
      Begin XtremeSuiteControls.ListBox lstFacturas 
         Height          =   1575
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   10770
         _Version        =   786432
         _ExtentX        =   18997
         _ExtentY        =   2778
         _StockProps     =   77
         BackColor       =   -2147483643
         Appearance      =   4
         Style           =   1
      End
      Begin XtremeSuiteControls.Label lblCantidadCbtesSeleccionados 
         Height          =   135
         Left            =   8640
         TabIndex        =   10
         Top             =   960
         Width           =   2175
         _Version        =   786432
         _ExtentX        =   3836
         _ExtentY        =   238
         _StockProps     =   79
         Caption         =   "Seleccionados"
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
         Height          =   135
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   2175
         _Version        =   786432
         _ExtentX        =   3836
         _ExtentY        =   238
         _StockProps     =   79
         Caption         =   "Total Comprobantes"
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Buscar factura en la lista:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1830
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox3 
      Height          =   1335
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   11100
      _Version        =   786432
      _ExtentX        =   19579
      _ExtentY        =   2355
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.PushButton btnMostrarFacturas 
         Height          =   615
         Left            =   720
         TabIndex        =   24
         Top             =   360
         Width           =   1815
         _Version        =   786432
         _ExtentX        =   3201
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "Mostrar Comprobantes"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.DateTimePicker dtpFecha 
         Height          =   330
         Left            =   9720
         TabIndex        =   12
         Top             =   240
         Width           =   1245
         _Version        =   786432
         _ExtentX        =   2196
         _ExtentY        =   582
         _StockProps     =   68
         Format          =   1
         CurrentDate     =   40183.7263657407
      End
      Begin XtremeSuiteControls.PushButton btnGuardar 
         Height          =   495
         Left            =   9240
         TabIndex        =   23
         Top             =   720
         Width           =   1695
         _Version        =   786432
         _ExtentX        =   2990
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Guardar"
         UseVisualStyle  =   -1  'True
      End
      Begin VB.Label lblTotalFacturas 
         AutoSize        =   -1  'True
         Caption         =   "Total facturas: "
         Height          =   195
         Left            =   5040
         TabIndex        =   22
         Top             =   840
         Width           =   1110
      End
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         Caption         =   "Total Pagos:"
         Height          =   195
         Left            =   5040
         TabIndex        =   21
         Tag             =   "Total: "
         Top             =   240
         Width           =   900
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         Height          =   195
         Left            =   9135
         TabIndex        =   13
         Tag             =   "Total: "
         Top             =   330
         Width           =   435
      End
   End
   Begin GridEX20.GridEX gridBancos 
      Height          =   1845
      Left            =   1560
      TabIndex        =   14
      Top             =   11040
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
      Column(1)       =   "frmAdminPagosCrearLiquidaciondeCaja.frx":2650
      Column(2)       =   "frmAdminPagosCrearLiquidaciondeCaja.frx":2750
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":2840
      FormatStyle(2)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":2978
      FormatStyle(3)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":2A28
      FormatStyle(4)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":2ADC
      FormatStyle(5)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":2BB4
      FormatStyle(6)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":2C6C
      ImageCount      =   0
      PrinterProperties=   "frmAdminPagosCrearLiquidaciondeCaja.frx":2D4C
   End
   Begin GridEX20.GridEX gridCuentasBancarias 
      Height          =   1935
      Left            =   10560
      TabIndex        =   15
      Top             =   11040
      Visible         =   0   'False
      Width           =   1665
      _ExtentX        =   2937
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
      Column(1)       =   "frmAdminPagosCrearLiquidaciondeCaja.frx":2F24
      Column(2)       =   "frmAdminPagosCrearLiquidaciondeCaja.frx":3048
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":313C
      FormatStyle(2)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":3274
      FormatStyle(3)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":3324
      FormatStyle(4)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":33D8
      FormatStyle(5)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":34B0
      FormatStyle(6)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":3568
      ImageCount      =   0
      PrinterProperties=   "frmAdminPagosCrearLiquidaciondeCaja.frx":3648
   End
   Begin GridEX20.GridEX gridMonedas 
      Height          =   1815
      Left            =   120
      TabIndex        =   16
      Top             =   11040
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
      Column(1)       =   "frmAdminPagosCrearLiquidaciondeCaja.frx":3820
      Column(2)       =   "frmAdminPagosCrearLiquidaciondeCaja.frx":3944
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":3A38
      FormatStyle(2)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":3B70
      FormatStyle(3)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":3C20
      FormatStyle(4)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":3CD4
      FormatStyle(5)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":3DAC
      FormatStyle(6)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":3E64
      ImageCount      =   0
      PrinterProperties=   "frmAdminPagosCrearLiquidaciondeCaja.frx":3F44
   End
   Begin GridEX20.GridEX gridCajas 
      Height          =   1935
      Left            =   14520
      TabIndex        =   17
      Top             =   11040
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
      Column(1)       =   "frmAdminPagosCrearLiquidaciondeCaja.frx":411C
      Column(2)       =   "frmAdminPagosCrearLiquidaciondeCaja.frx":421C
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":4308
      FormatStyle(2)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":4440
      FormatStyle(3)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":44F0
      FormatStyle(4)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":45A4
      FormatStyle(5)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":467C
      FormatStyle(6)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":4734
      ImageCount      =   0
      PrinterProperties=   "frmAdminPagosCrearLiquidaciondeCaja.frx":4814
   End
   Begin GridEX20.GridEX gridChequesDisponibles 
      Height          =   1920
      Left            =   5160
      TabIndex        =   18
      Top             =   11040
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
      Column(1)       =   "frmAdminPagosCrearLiquidaciondeCaja.frx":49EC
      Column(2)       =   "frmAdminPagosCrearLiquidaciondeCaja.frx":4B6C
      Column(3)       =   "frmAdminPagosCrearLiquidaciondeCaja.frx":4D0C
      Column(4)       =   "frmAdminPagosCrearLiquidaciondeCaja.frx":4E48
      Column(5)       =   "frmAdminPagosCrearLiquidaciondeCaja.frx":4F54
      Column(6)       =   "frmAdminPagosCrearLiquidaciondeCaja.frx":5074
      Column(7)       =   "frmAdminPagosCrearLiquidaciondeCaja.frx":5180
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":5274
      FormatStyle(2)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":53AC
      FormatStyle(3)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":545C
      FormatStyle(4)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":5510
      FormatStyle(5)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":55E8
      FormatStyle(6)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":56A0
      ImageCount      =   0
      PrinterProperties=   "frmAdminPagosCrearLiquidaciondeCaja.frx":5780
   End
   Begin GridEX20.GridEX gridChequeras 
      Height          =   1935
      Left            =   9000
      TabIndex        =   19
      Top             =   11040
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
      Column(1)       =   "frmAdminPagosCrearLiquidaciondeCaja.frx":5958
      Column(2)       =   "frmAdminPagosCrearLiquidaciondeCaja.frx":5A78
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":5B78
      FormatStyle(2)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":5CB0
      FormatStyle(3)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":5D60
      FormatStyle(4)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":5E14
      FormatStyle(5)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":5EEC
      FormatStyle(6)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":5FA4
      ImageCount      =   0
      PrinterProperties=   "frmAdminPagosCrearLiquidaciondeCaja.frx":6084
   End
   Begin GridEX20.GridEX gridChequesChequera 
      Height          =   1935
      Left            =   12360
      TabIndex        =   20
      Top             =   11040
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
      Column(1)       =   "frmAdminPagosCrearLiquidaciondeCaja.frx":625C
      Column(2)       =   "frmAdminPagosCrearLiquidaciondeCaja.frx":638C
      SortKeysCount   =   1
      SortKey(1)      =   "frmAdminPagosCrearLiquidaciondeCaja.frx":648C
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":64F4
      FormatStyle(2)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":662C
      FormatStyle(3)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":66DC
      FormatStyle(4)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":6790
      FormatStyle(5)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":6868
      FormatStyle(6)  =   "frmAdminPagosCrearLiquidaciondeCaja.frx":6920
      ImageCount      =   0
      PrinterProperties=   "frmAdminPagosCrearLiquidaciondeCaja.frx":6A00
   End
   Begin XtremeSuiteControls.RadioButton radioFacturaProveedor 
      Height          =   210
      Left            =   120
      TabIndex        =   25
      Top             =   7680
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
      Left            =   765
      TabIndex        =   26
      Top             =   8040
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
      Left            =   120
      TabIndex        =   27
      Tag             =   "Total: "
      Top             =   8100
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

Dim total_por_factura As New Dictionary
Dim vFactElegida As clsFacturaProveedor
Dim vCompeElegido As Compensatorio
Dim vFacturaProveedor As clsFacturaProveedor
Dim colProveedores As New Collection
Dim colFacturas As New Collection
Dim colDeudaCompensatorios As New Collection
Dim prov As clsProveedor
Dim Factura As clsFacturaProveedor

Private Banco As Banco
Private caja As caja
Private CuentaBancaria As CuentaBancaria
Private moneda As clsMoneda
Private alicuotaRetencion As DTORetencionAlicuota
Private cuentasBancarias As New Collection
Private retenciones As New Collection
Private Monedas As New Collection
Private Cajas As New Collection
Private bancos As New Collection
Private chequesDisponibles As New Collection
Private chequeras As New Collection

Dim compe As Compensatorio

Private LiquidacionCaja As New clsLiquidacionCaja
Private OrdenPago As New OrdenPago

Private operacion As operacion
Private cheque As cheque
Private tmpChequera As chequera

Private chequesChequeraSeleccionada As New Collection

Public ReadOnly As Boolean

Public Sub Cargar(op As OrdenPago)


If Not IsSomething(op) Then
    MsgBox "La OP que está intentando visualizar está en estado PENDIENTE. " & vbNewLine & "Por lo tanto no puede ser mostrada porque puede estar siendo editada." & vbNewLine & "Verifiquelo por favor.", vbCritical, "OP Pendiente"
    Unload Me
    Exit Sub

End If


    Set OrdenPago = DAOLiquidacionCaja.FindById(op.Id)
    Set LiquidacionCaja.Compensatorios = DAOCompensatorios.FindByOP(LiquidacionCaja.Id)

    Dim i As Long
    Dim j As Long
    With OrdenPago

        If .EsParaFacturaProveedor Then
            radioFacturaProveedor.value = True

            If .FacturasProveedor.count > 0 Then

'                Me.cboProveedores.ListIndex = funciones.PosIndexCbo(.FacturasProveedor.item(1).Proveedor.Id, Me.cboProveedores)

'                If Me.cboProveedores.ListIndex = -1 Then    'el proveedor no esta en la lista porque no tiene mas facturas sin saldar
'                    Me.cboProveedores.AddItem .FacturasProveedor.item(1).Proveedor.RazonSocial
'                    Me.cboProveedores.ItemData(Me.cboProveedores.NewIndex) = .FacturasProveedor.item(1).Proveedor.Id
'                    colProveedores.Add .FacturasProveedor.item(1).Proveedor, CStr(.FacturasProveedor.item(1).Proveedor.Id)
'                    Me.cboProveedores.ListIndex = funciones.PosIndexCbo(.FacturasProveedor.item(1).Proveedor.Id, Me.cboProveedores)
'                End If

         cmdMostrarDatosProveedor_Click


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
'            Me.txtRetenciones.text = .alicuota

        Else
'            Me.radioConcepto.value = True

'            If IsSomething(.CuentaContable) Then
'                Me.cboCuentas.ListIndex = funciones.PosIndexCbo(.CuentaContable.Id, Me.cboCuentas)
'                Me.txtDetalle.text = .CuentaContableDescripcion
'            Else
'                Me.cboCuentas.ListIndex = -1
'                Me.txtDetalle.text = vbNullString
'            End If

        End If


        If idx >= 0 Then
             lstFacturas.ListIndex = lstFacturas.ListCount - 1

         End If




'        Me.gridCajaOperaciones.ItemCount = .OperacionesCaja.count
'        Me.gridDepositosOperaciones.ItemCount = .OperacionesBanco.count
'        Me.gridCheques.ItemCount = .ChequesTerceros.count
'        Me.gridChequesPropios.ItemCount = .ChequesPropios.count
'
'        Me.gridRetenciones.ItemCount = .RetencionesAlicuota.count
'        Set alicuotas = .RetencionesAlicuota
'
'
'        Me.cboMonedas.ListIndex = funciones.PosIndexCbo(.moneda.Id, Me.cboMonedas)
'        Me.dtpFecha.value = .FEcha
'        Me.txtDifCambio.text = .DiferenciaCambio
'        Me.txtOtrosDescuentos.text = .OtrosDescuentos

    End With
    mostrarCompensatorios





'    Me.caption = "Orden de Pago Nº " & LiquidacionCaja.Id
'
'    'Me.grpDestino.Enabled = Not ReadOnly
'    Me.txtDifCambioNG1.Enabled = Not ReadOnly
'    Me.txtDifCambioTOTAL1.Enabled = Not ReadOnly
'    Me.cmdMostrarDatosProveedor.Enabled = Not ReadOnly
'    Me.btnPadronAnt.Enabled = Not ReadOnly
'    Me.btnCargar.Enabled = Not ReadOnly
'
'    Me.gridRetenciones.AllowEdit = Not ReadOnly

'    GroupBox2.Enabled = Not ReadOnly
'
'    GroupBox1.Enabled = Not ReadOnly


    Me.radioConcepto.Enabled = Not ReadOnly
    Me.radioFacturaProveedor.Enabled = Not ReadOnly
    Me.cboCuentas.Enabled = Not ReadOnly
    Me.cboProveedores.Enabled = Not ReadOnly
    Me.txtDetalle.Enabled = Not ReadOnly
    Me.btnClearProveedor.Enabled = Not ReadOnly

    'Me.grpOrigen.Enabled = Not ReadOnly



    Me.gridDepositosOperaciones.AllowEdit = Not ReadOnly
    Me.gridDepositosOperaciones.AllowDelete = Not ReadOnly

    Me.gridBancos.AllowEdit = Not ReadOnly
    'Me.gridBancos.AllowDelete = Not ReadOnly

    Me.gridCajaOperaciones.AllowEdit = Not ReadOnly
    Me.gridCajaOperaciones.AllowDelete = Not ReadOnly

    Me.gridCajas.AllowEdit = Not ReadOnly
    'Me.gridCajas.AllowDelete = Not ReadOnly

    Me.gridChequeras.AllowEdit = Not ReadOnly
    'Me.gridChequeras.AllowDelete = Not ReadOnly

    Me.gridCheques.AllowEdit = Not ReadOnly
    Me.gridCheques.AllowDelete = Not ReadOnly

    Me.gridChequesChequera.AllowEdit = Not ReadOnly
    'Me.gridChequesChequera.AllowDelete = Not ReadOnly

    Me.gridChequesDisponibles.AllowEdit = Not ReadOnly
    'Me.gridChequesDisponibles.AllowDelete = Not ReadOnly

    Me.gridChequesPropios.AllowEdit = Not ReadOnly
    Me.gridChequesPropios.AllowDelete = Not ReadOnly

    Me.cboMonedas.Enabled = Not ReadOnly
    Me.dtpFecha.Enabled = Not ReadOnly
    Me.btnGuardar.Enabled = Not ReadOnly
    Me.txtDifCambio.Enabled = Not ReadOnly
    Me.txtOtrosDescuentos.Enabled = Not ReadOnly

    Totalizar

End Sub


Public Property Get FacturaProveedor(nvalue As clsFacturaProveedor)
    Set vFacturaProveedor = nvalue
End Property


Private Sub btnBorrar_Click()

    cboProveedores.ListIndex = -1
    Me.gridRetenciones.ItemCount = 0
    Me.txtRetenciones.text = 0
    Me.lstFacturas.Clear
    Set prov = Nothing


End Sub

Private Sub ActualizarAlicuotas()

  Dim A As DTORetencionAlicuota
                    Dim B As DTORetencionAlicuota
                       For Each A In alicuotas

                       For Each B In LiquidacionCaja.RetencionesAlicuota
                                If A.Retencion.Id = B.Retencion.Id Then
                                  If B.importe > 0 Then
                                    A.importe = B.importe
                                  End If

                                End If

                    Next

                    Next

End Sub


Private Sub btnCargar_Click()

    If Me.cboProveedores.ListIndex <> -1 Then
        Set prov = colProveedores.item(CStr(Me.cboProveedores.ItemData(Me.cboProveedores.ListIndex)))

                If IsSomething(prov) Then

                 ' #fix 180
                If LiquidacionCaja.estado = EstadoOrdenPago_pendiente Then
                    Set alicuotas = DAORetenciones.FindAllWithAlicuotas(prov.Cuit)
                    ActualizarAlicuotas
                End If


                End If
    Else
        Set prov = Nothing

    End If

    Me.gridRetenciones.ItemCount = 0
    Me.gridRetenciones.ItemCount = alicuotas.count
    Me.gridRetenciones.Refresh

'MostrarFacturas
    Totalizar

End Sub

Private Sub btnClearProveedor_Click()
    cboProveedores.ListIndex = -1
    Me.gridRetenciones.ItemCount = 0
    Me.txtRetenciones.text = 0
    Me.lstFacturas.Clear
    Set prov = Nothing
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


'    Set LiquidacionCaja.CuentaContable = Nothing
'    LiquidacionCaja.CuentaContableDescripcion = vbNullString
    Set LiquidacionCaja.FacturasProveedor = New Collection
'    Set LiquidacionCaja.RetencionesAlicuota = alicuotas




    If Me.radioFacturaProveedor.value Then
        Dim T As Long
        For T = 0 To Me.lstFacturas.ListCount - 1
            If Me.lstFacturas.Checked(T) Then
                LiquidacionCaja.FacturasProveedor.Add colFacturas.item(CStr(Me.lstFacturas.ItemData(T)))
            End If
        Next T

    Else

'        If Me.cboCuentas.ListIndex > -1 Then
'            Set LiquidacionCaja.CuentaContable = DAOCuentaContable.GetById(Me.cboCuentas.ItemData(Me.cboCuentas.ListIndex))
'        End If
'        LiquidacionCaja.CuentaContableDescripcion = Me.txtDetalle.text

    End If

'     For i = 0 To Me.lstDeudaCompensatorios.ListCount - 1
'            If Me.lstDeudaCompensatorios.Checked(i) Then
'                LiquidacionCaja.DeudaCompensatorios.Add colDeudaCompensatorios.item(CStr(Me.lstDeudaCompensatorios.ItemData(i)))
'            End If
'        Next i


'    If IsNumeric(Me.txtRetenciones) Then LiquidacionCaja.alicuota = Val(Me.txtRetenciones)

    If LiquidacionCaja.IsValid Then

        Dim n As Boolean: n = (LiquidacionCaja.Id = 0)

        If DAOLiquidacionCaja.Save(LiquidacionCaja, True) Then

            If n Then
                MsgBox "Liquidación de Caja Nº " & LiquidacionCaja.Id & " creada con exito.", vbInformation
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

Private Sub btnPadronAnt_Click()

    If Me.cboProveedores.ListIndex <> -1 Then
        Set prov = colProveedores.item(CStr(Me.cboProveedores.ItemData(Me.cboProveedores.ListIndex)))

                If IsSomething(prov) Then
                    Set alicuotas = DAORetenciones.FindAllWithAlicuotasAnt(prov.Cuit)
        ActualizarAlicuotas

                End If
    Else
        Set prov = Nothing

    End If

    Me.gridRetenciones.ItemCount = 0
    Me.gridRetenciones.ItemCount = alicuotas.count
    Me.gridRetenciones.Refresh

'MostrarFacturas
    Totalizar

End Sub

Private Sub btnMostrarFacturas_Click()
MostrarFacturas
End Sub

Private Sub cboMonedas_Click()
    If Me.cboMonedas.ListIndex = -1 Then
        Set LiquidacionCaja.moneda = Nothing
    Else
        Set LiquidacionCaja.moneda = DAOMoneda.GetById(Me.cboMonedas.ItemData(Me.cboMonedas.ListIndex))
    End If
    Totalizar
End Sub



Private Sub cboProveedores_Click()

Me.gridRetenciones.ItemCount = 0
Me.lstFacturas.Clear

Me.txtBuscarFactura = ""
Me.txtParcialAbonar = ""

End Sub


Private Sub cmdMostrarDatosProveedor_Click()
  If Me.cboProveedores.ListIndex <> -1 Then

        Set prov = colProveedores.item(CStr(Me.cboProveedores.ItemData(Me.cboProveedores.ListIndex)))



        Dim d As clsDTOPadronIIBB

            Set d = DTOPadronIIBB.FindByCUIT(prov.Cuit, TipoPadronRetencion)

            If IsSomething(d) Then
              Me.txtRetenciones = str(d.alicuota)   ' Val(d.Retencion )
            Else
                Me.txtRetenciones = 0
            End If


        'If IsSomething(prov) Then
         'Set alicuotas = DAORetenciones.FindAllWithAlicuotas(prov.Cuit)

       ' End If
    Else
        Set prov = Nothing
    End If


MostrarFacturas
MostrarDeudaCompensatorios
btnCargar_Click

End Sub

Private Sub Command1_Click()


    If Me.cboProveedores.ListIndex <> -1 Then

        Set prov = colProveedores.item(CStr(Me.cboProveedores.ItemData(Me.cboProveedores.ListIndex)))
        If IsSomething(prov) Then

         Set alicuotas = DAORetenciones.FindAllWithAlicuotas(prov.Cuit)

        ActualizarAlicuotas

        End If
    Else
        Set prov = Nothing
    End If
    Me.gridRetenciones.ItemCount = 0

    Me.gridRetenciones.ItemCount = alicuotas.count

    Me.gridRetenciones.Refresh

    MostrarFacturas

End Sub

Private Sub dtpFecha_Change()
    LiquidacionCaja.FEcha = Me.dtpFecha.value
End Sub

Private Sub Form_Load()
    formLoading = True
'    Me.gridChequeras.Visible = False
'    Me.gridChequesChequera.Visible = False
    Me.gridCompensatorios.ItemCount = 0
    id_susc = funciones.CreateGUID
    Channel.AgregarSuscriptor Me, PasajeChequePropioCartera
    FormHelper.Customize Me
    GridEXHelper.CustomizeGrid Me.gridCajaOperaciones, False, True
    GridEXHelper.CustomizeGrid Me.gridDepositosOperaciones, False, True
'    GridEXHelper.CustomizeGrid Me.gridCheques, False, True
'    GridEXHelper.CustomizeGrid Me.gridChequesDisponibles, False, False
    GridEXHelper.CustomizeGrid Me.gridBancos, False, False
'    GridEXHelper.CustomizeGrid Me.gridCuentasBancarias, False, False
    GridEXHelper.CustomizeGrid Me.gridMonedas, False, False
    GridEXHelper.CustomizeGrid Me.gridCajas, False, False
    GridEXHelper.CustomizeGrid Me.gridChequeras, False, False
'    GridEXHelper.CustomizeGrid Me.gridChequesPropios, False, True
    GridEXHelper.CustomizeGrid Me.gridCompensatorios, False, True
'    GridEXHelper.CustomizeGrid Me.gridChequesChequera
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


'    Set colProveedores = DAOProveedor.FindAllProveedoresWithFacturasImpagas
'    For Each prov In colProveedores
'        cboProveedores.AddItem prov.RazonSocial
'        cboProveedores.ItemData(cboProveedores.NewIndex) = prov.Id
'    Next

'    Dim cuentasContables As Collection
'    Set cuentasContables = DAOCuentaContable.GetAll()
'    Dim cc As clsCuentaContable
'    Me.cboCuentas.Clear
'    For Each cc In cuentasContables
'        cboCuentas.AddItem cc.nombre & " - " & cc.codigo
'        cboCuentas.ItemData(cboCuentas.NewIndex) = cc.Id
'    Next cc


    radioFacturaProveedor_Click

    Me.gridCajaOperaciones.ItemCount = LiquidacionCaja.OperacionesCaja.count
    Me.gridDepositosOperaciones.ItemCount = LiquidacionCaja.OperacionesBanco.count
'    Me.gridCheques.ItemCount = LiquidacionCaja.ChequesTerceros.count
'    Me.gridChequesPropios.ItemCount = LiquidacionCaja.ChequesPropios.count



'    Set Me.gridCheques.Columns("numero").DropDownControl = Me.gridChequesDisponibles

    Set Me.gridDepositosOperaciones.Columns("moneda").DropDownControl = Me.gridMonedas
    Set Me.gridDepositosOperaciones.Columns("cuenta").DropDownControl = Me.gridCuentasBancarias

    Set Me.gridCajaOperaciones.Columns("moneda").DropDownControl = Me.gridMonedas
    Set Me.gridCajaOperaciones.Columns("caja").DropDownControl = Me.gridCajas

'    Set Me.gridChequesPropios.Columns("chequera").DropDownControl = Me.gridChequeras
'    Set Me.gridChequesPropios.Columns("numero").DropDownControl = Me.gridChequesChequera
    gridChequesChequera.ItemCount = 0
    GridEXHelper.AutoSizeColumns Me.gridChequeras


    DAOMoneda.llenarComboXtremeSuite Me.cboMonedas

    Me.dtpFecha.value = LiquidacionCaja.FEcha

'lstFacturas_Click
    Totalizar

    formLoaded = True
    formLoading = False
End Sub

Private Sub CargarChequesDisponibles()
    Set chequesDisponibles = DAOCheques.FindAllEnCarteraDeTerceros
    Me.gridChequesDisponibles.ItemCount = chequesDisponibles.count
End Sub

Private Sub MostrarDeudaCompensatorios()
    Me.lstDeudaCompensatorios.Clear
    If IsSomething(prov) Then
        Set colDeudaCompensatorios = DAOCompensatorios.FindAllPendientesByProveedor(prov.Id)  'DAOFacturaProveedor.FindAll("AdminComprasFacturasProveedores.id_proveedor=" & prov.id & " and (AdminComprasFacturasProveedores.estado=" & EstadoFacturaProveedor.pagoParcial & " or  AdminComprasFacturasProveedores.estado=" & EstadoFacturaProveedor.Aprobada & ")", False, "", False, True)



        Dim c As Compensatorio


'        If LiquidacionCaja.id <> 0 And LiquidacionCaja.EsParaFacturaProveedor Then
'            If prov.id = LiquidacionCaja.FacturasProveedor.item(1).Proveedor.id Then
'                For Each Factura In LiquidacionCaja.FacturasProveedor
'                    If Not funciones.BuscarEnColeccion(colFacturas, CStr(Factura.id)) Then
'                        colFacturas.Add DAOFacturaProveedor.FindById(Factura.id), CStr(Factura.id)
'                    End If
'                Next
'            End If
'        End If

        For Each c In colDeudaCompensatorios
            Me.lstDeudaCompensatorios.AddItem "Cód: " & c.Id & " (OP: " & c.IdOrdenPago & ", Cbte: " & c.Comprobante.NumeroFormateado & ", Importe: " & c.Monto & ")"
            Me.lstDeudaCompensatorios.ItemData(Me.lstDeudaCompensatorios.NewIndex) = c.Id
        Next




    Else
        Set colFacturas = New Collection
    End If
End Sub



Private Sub MostrarFacturas()

    Me.lstFacturas.Clear

    If 1 = 1 Then
                'Set colFacturas = DAOFacturaProveedor.FindAll("AdminComprasFacturasProveedores.id_proveedor=" & prov.Id & " and (AdminComprasFacturasProveedores.estado=" & EstadoFacturaProveedor.Aprobada & " or AdminComprasFacturasProveedores.estado=" & EstadoFacturaProveedor.pagoParcial & ")", False, "", False, True)

                Set colFacturas = DAOFacturaProveedor.FindAll("AdminComprasFacturasProveedores.id_proveedor=" & 6003 & " and (AdminComprasFacturasProveedores.estado=" & EstadoFacturaProveedor.Aprobada & ")", False, "", False, True)

                Set colFacturas = DAOFacturaProveedor.FindAll("AdminComprasFacturasProveedores.estado = " & EstadoFacturaProveedor.Aprobada & "", False, "", False, True)


            If LiquidacionCaja.Id <> 0 And LiquidacionCaja.EsParaFacturaProveedor Then
                If prov.Id = LiquidacionCaja.FacturasProveedor.item(1).Proveedor.Id Then
                    For Each Factura In LiquidacionCaja.FacturasProveedor
                        If Not funciones.BuscarEnColeccion(colFacturas, CStr(Factura.Id)) Then

                            colFacturas.Add DAOFacturaProveedor.FindById(Factura.Id), CStr(Factura.Id)
                        End If
                    Next
                End If
            End If

        Dim T As String

            For Each Factura In colFacturas 'en ese for traigo los pendientes a abonar que estan asociados a ops sin aprobar

               Dim c As Collection
                  Set c = DAOLiquidacionCaja.FindAbonadoPendiente(Factura.Id, LiquidacionCaja.Id)

                   Factura.TotalAbonadoGlobalPendiente = 0 ' c(1) 'que esta en ops sin aprobar
                       Factura.NetoGravadoAbonadoGlobalPendiente = 0 ' c(2)
                         Factura.OtrosAbonadoGlobalPendiente = 0 'c(3)

                      T = Factura.NumeroFormateado & " (" & Factura.moneda.NombreCorto & " " & Factura.Total & ")" & " (" & Factura.FEcha & ")" 'TipoCambio: (" & Factura.TipoCambioPago & ")"
                         If Factura.TotalAbonadoGlobal + Factura.TotalAbonadoGlobalPendiente > 0 Then
                          T = Factura.NumeroFormateado & " (" & Factura.moneda.NombreCorto & " " & Factura.Total & " - Abonado: " & Factura.TotalAbonadoGlobal + Factura.TotalAbonadoGlobalPendiente & ")" & " (" & Factura.FEcha & ")"

                 'MsgBox (c.count)

                 End If

                  Me.lstFacturas.AddItem T
                  Me.lstFacturas.ItemData(Me.lstFacturas.NewIndex) = Factura.Id


            Next

' 22/08/2022
'AGREGO UN LABEL QUE MUESTRA LA CANTIDAD DE COMPROBANTES MOSTRADOS EN EL LIST

            Me.lblCantidadComprobantes.caption = "Cbtes. Mostrados: " & colFacturas.count

        Else

            Set colFacturas = New Collection

            'MsgBox (colFacturas.count)

    End If

End Sub


Private Sub Form_Unload(Cancel As Integer)
    Channel.RemoverSuscripcionTotal Me
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



Private Sub gridChequesPropios_BeforeUpdate(ByVal Cancel As GridEX20.JSRetBoolean)
    Dim msg As New Collection

    If LenB(Me.gridChequesPropios.value(1)) = 0 Then
        msg.Add "Debe especificar una chequera."
    End If

    If LenB(Me.gridChequesPropios.value(2)) = 0 Then
        msg.Add "Debe especificar un cheque."
    End If

' REVISA QUE EN LA COLECCION DE CHEQUES PROPIOS QUE SE ESTAN CARGANDO NO ESTÉ INGRESADO EL MISMO CHEQUE, SI LO DETECTA GENERA MSG DE ERROR
    If funciones.BuscarEnColeccion(LiquidacionCaja.ChequesPropios, CStr(Me.gridChequesPropios.value(2))) Then
        msg.Add "El cheque seleccionado ya fue ingresado anteriormente."
    End If

    If Not IsNumeric(Me.gridChequesPropios.value(3)) Then
        msg.Add "Debe especificar un monto válido."
    End If
' REVISA QUE SE HAYA CARGADO UN MONTO DEL CHEQUE INGRESADO, SI NO SE CARGA GENERA MSG DE ERROR

    If LenB(Me.gridChequesPropios.value(3)) = 0 Then
        msg.Add "Debe especificar un monto mayor a 0."
    End If

    If Not IsDate(Me.gridChequesPropios.value(4)) Then
        msg.Add "Debe especificar una fecha valida."
    End If

    Cancel = (msg.count > 0)
    If Cancel Then MsgBox funciones.JoinCollectionValues(msg, vbNewLine), vbExclamation

End Sub





Private Sub gridChequesPropios_ListSelected(ByVal ColIndex As Integer, ByVal ValueListIndex As Long, ByVal value As Variant)
    If ColIndex = 1 Then
        'If Not IsNumeric(Me.gridChequesPropios.Value(1)) Or LenB(Me.gridChequesPropios.Value(1)) = 0 Then
        If Not IsNumeric(value) Or LenB(value) = 0 Then
            Set chequesChequeraSeleccionada = New Collection
        Else
            Set chequesChequeraSeleccionada = DAOCheques.FindAllDisponiblesByChequera(Val(value))  ' Me.gridChequesPropios.Value(1))
        End If

        Me.gridChequesChequera.ItemCount = chequesChequeraSeleccionada.count
    End If
End Sub

Private Sub gridChequesPropios_UnboundAddNew(ByVal NewRowBookmark As GridEX20.JSRetVariant, ByVal Values As GridEX20.JSRowData)
    Set cheque = Nothing
    If IsNumeric(Values(2)) Then Set cheque = DAOCheques.FindById(Values(2))
    If IsSomething(cheque) Then
        cheque.Monto = Values(3)
        cheque.FechaVencimiento = Values(4)

        LiquidacionCaja.ChequesPropios.Add cheque, CStr(cheque.Id)


    End If
    Totalizar
End Sub

Private Sub gridChequesPropios_UnboundDelete(ByVal RowIndex As Long, ByVal Bookmark As Variant)
    If RowIndex > 0 Then
        LiquidacionCaja.ChequesPropios.remove RowIndex
        Totalizar
    End If
End Sub

Private Sub gridChequesPropios_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If LiquidacionCaja.ChequesPropios.count >= RowIndex Then
        Set cheque = LiquidacionCaja.ChequesPropios.item(RowIndex)
        Values(1) = cheque.chequera.Description
        Values(2) = vbNullString
        'FORMATCURRENCY
        Values(3) = FormatCurrency(cheque.Monto)
        Values(4) = cheque.FechaVencimiento
        Values(5) = cheque.numero


        Totalizar
    End If
End Sub

Private Sub gridChequesPropios_UnboundUpdate(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If LiquidacionCaja.ChequesPropios.count >= RowIndex Then
        Set cheque = LiquidacionCaja.ChequesPropios.item(RowIndex)

        '        If Values(2) <> Cheque.Id Then
        '            LiquidacionCaja.ChequesPropios.remove CStr(Cheque.Id)
        '            Set Cheque = DAOCheques.FindById(Values(2))
        '            LiquidacionCaja.ChequesPropios.Add Cheque, CStr(Cheque.Id)
        '        End If

        cheque.Monto = Values(3)
        cheque.FechaVencimiento = Values(4)
    End If

    Totalizar
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


Private Sub gridRetenciones_RowFormat(RowBuffer As GridEX20.JSRowData)

 On Error GoTo err1

        Set alicuotaRetencion = alicuotas.item(RowBuffer.RowIndex)

            If alicuotaRetencion.importe > 0 Then '.Retencion.id <> 2 Then
            RowBuffer.RowStyle = "padronganancias"
            Else
            RowBuffer.RowStyle = "padroningresos"

        End If

      Exit Sub

err1:

End Sub

Private Sub gridRetenciones_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If alicuotas.count >= RowIndex Then
        Set alicuotaRetencion = alicuotas.item(RowIndex)
        Values(2) = alicuotaRetencion.alicuotaRetencion
        Values(1) = alicuotaRetencion.Retencion.nombre
        Values(3) = alicuotaRetencion.importe
    End If
End Sub

Private Sub gridRetenciones_UnboundUpdate(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
 If alicuotas.count >= RowIndex Then
        Set alicuotaRetencion = alicuotas.item(RowIndex)
       alicuotaRetencion.alicuotaRetencion = Values(2)
       If Not IsNumeric(Values(3)) Then
       alicuotaRetencion.importe = 0
       Else
       alicuotaRetencion.importe = Values(3)
       End If
       Totalizar

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
    Dim Cambio As Double
    Dim totCambio As Double
    Dim totCambiong As Double
    Dim totNGHoy As Double
    Dim totDeudaCompe As Double
    totDeudaCompe = 0
    For Each F In col


        'totNGHoy = totNGHoy + MonedaConverter.ConvertirForzado2(IIf(f.tipoDocumentoContable = tipoDocumentoContable.notaCredito, f.NetoGravadoDiaPago * -1, f.NetoGravadoDiaPago), f.Moneda.Id, LiquidacionCaja.Moneda.Id, f.TipoCambioPago)
        ' totFact = totFact + MonedaConverter.ConvertirForzado2(IIf(f.tipoDocumentoContable = tipoDocumentoContable.notaCredito, f.total * -1, f.total), f.Moneda.Id, LiquidacionCaja.Moneda.Id, f.TipoCambioPago) cambiado el 22-9-14 por tema de pagos parciales
        'totFactHoy = totFactHoy + MonedaConverter.ConvertirForzado2(IIf(f.tipoDocumentoContable = tipoDocumentoContable.notaCredito, f.TotalDiaPago * -1, f.TotalDiaPago), f.Moneda.Id, LiquidacionCaja.Moneda.Id, f.TipoCambioPago)
        'totNG = TotNG + MonedaConverter.ConvertirForzado2(IIf(f.tipoDocumentoContable = tipoDocumentoContable.notaCredito, f.NetoGravado * -1, f.NetoGravado), f.Moneda.Id, LiquidacionCaja.Moneda.Id, f.TipoCambioPago)
        'totFact = totFact + MonedaConverter.ConvertirForzado2(IIf(F.tipoDocumentoContable = tipoDocumentoContable.notaCredito, F.ImporteTotalAbonado * -1, F.ImporteTotalAbonado), F.moneda.id, LiquidacionCaja.moneda.id, F.TipoCambioPago)
        'fix 004


        'ORIGINAL- totFact = totFact + MonedaConverter.ConvertirForzado2(IIf(F.tipoDocumentoContable = tipoDocumentoContable.notaCredito, F.TotalAbonado * -1, F.TotalAbonado), F.moneda.Id, LiquidacionCaja.moneda.Id, F.TipoCambioPago)

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

    'FORMATCURRENCY
'    Me.lblNgAbonar = "Total NG a Abonar en " & FormatCurrency(funciones.FormatearDecimales(LiquidacionCaja.DiferenciaCambioEnNG + totNGHoy))

    'FORMATCURRENCY
'    MsgBox ("Total Facturas en " & FormatCurrency(funciones.FormatearDecimales(totFact)))
    Me.lblTotalFacturas = "Total Facturas en " & FormatCurrency(funciones.FormatearDecimales(totFact))

    'FORMATCURRENCY
'    Me.lblDeudaCompensatorios = "Total deuda compensatorios en " & FormatCurrency(funciones.FormatearDecimales(totDeudaCompe))

    LiquidacionCaja.StaticTotalFacturas = funciones.RedondearDecimales(totFact)
    LiquidacionCaja.staticTotalDeudaCompensatorios = funciones.RedondearDecimales(totDeudaCompe)

    'FORMATCURRENCY
'    Me.lblTotalFacturasNG = "Total NG Facturas en " & FormatCurrency(funciones.FormatearDecimales(TotNG + LiquidacionCaja.DiferenciaCambioEnNG))

    LiquidacionCaja.StaticTotalFacturasNG = funciones.RedondearDecimales(TotNG + LiquidacionCaja.DiferenciaCambioEnNG)

    'FORMATCURRENCY
'    Me.lblDiferenciaCambio = "Diferencia Cambio en " & FormatCurrency(totCambiong)
    'Me.lblDiferenciaCambio = "Diferencia Cambio en " & LiquidacionCaja.moneda.NombreCorto & " " & totCambiong

    LiquidacionCaja.DiferenciaCambio = totCambio

    verCompensatorios

    'FORMATCURRENCY
'    Me.lblTotalARetener = "Total a retener en " & FormatCurrency(funciones.FormatearDecimales(totRet))
    'Me.lblTotalARetener = "Total a retener en " & LiquidacionCaja.moneda.NombreCorto & " " & funciones.FormatearDecimales(totRet)

    LiquidacionCaja.StaticTotalRetenido = funciones.RedondearDecimales(totRet)

    'FORMATCURRENCY
'    Me.lblTotalOrdenPago = "Total a abonar en " & FormatCurrency(funciones.FormatearDecimales(LiquidacionCaja.DiferenciaCambioEnTOTAL + totFactHoy - totRet - LiquidacionCaja.OtrosDescuentos + LiquidacionCaja.TotalCompensatorios + totDeudaCompe))
'    'Me.lblTotalOP = "Total OP: " & LiquidacionCaja.moneda.NombreCorto & " " & LiquidacionCaja.StaticTotal
End Sub

Private Sub verCompensatorios()
'    Me.lblTotalCompensatorios = "Total compensatorios en " & FormatCurrency(funciones.FormatearDecimales(LiquidacionCaja.TotalCompensatorios))
End Sub



Private Sub MostrarPago(F As clsFacturaProveedor)

    If IsSomething(F) Then

'        Me.txtTotalParcialAbonado = F.TotalAbonadoGlobal
'        Me.txtOtrosParcialAbonado = F.OtrosAbonadoGlobal + F.OtrosAbonadoGlobalPendiente
'        Me.txtParcialAbonado = F.NetoGravadoAbonadoGlobal + F.NetoGravadoAbonadoGlobalPendiente


       ' If F.ImporteTotalAbonado = 0 Then F.ImporteTotalAbonado = F.Total
        If F.NetoGravadoAbonado = 0 Then F.NetoGravadoAbonado = F.NetoGravado '- F.NetoNoGravado  (2do cambio en fix 004)
          If F.OtrosAbonado = 0 Then F.OtrosAbonado = F.Total - F.NetoGravado '- F.NetoNoGravado  (2do cambio en fix 004)

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


Private Sub Label13_Click()

End Sub

Private Sub lstDeudaCompensatorios_Click()


Set vCompeElegido = colDeudaCompensatorios.item(CStr(Me.lstDeudaCompensatorios.ItemData(Me.lstDeudaCompensatorios.ListIndex)))
If IsSomething(vCompeElegido) Then


'    MostrarPago vFactElegida
End If

End Sub

Private Sub lstDeudaCompensatorios_ItemCheck(ByVal item As Long)
    calcularOrigenes
End Sub

Private Sub lstFacturas_Click()

'    'debug.print (Me.lstFacturas.ItemData(Me.lstFacturas.ListIndex))

    Set vFactElegida = colFacturas.item(CStr(Me.lstFacturas.ItemData(Me.lstFacturas.ListIndex)))


'    'debug.print (vFactElegida.Id)


If IsSomething(vFactElegida) Then

   Dim c As Collection

'Me.lblCantidadCbtesSeleccionados.caption = "Cbtes. Seleccionados: " & c.count


    If LiquidacionCaja.estado = EstadoOrdenPago_pendiente And vFactElegida.NetoGravadoAbonado = 0 And vFactElegida.OtrosAbonado = 0 Then
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
    Dim col As New Collection
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

        Else


            'si destildo tengo q ver q no existan compensatorios. Si existen debería primero eliminarlos.
            Dim ff As clsFacturaProveedor
            Dim c As Compensatorio
            For Each c In LiquidacionCaja.Compensatorios
                Set ff = colFacturas.item(CStr(Me.lstFacturas.ItemData(i)))
                If c.Comprobante.Id = ff.Id Then
                    MsgBox "Existen compensatorios para este comprobante. Eliminelos primero!", vbCritical, "Error"
                    Me.lstFacturas.Checked(i) = True
                End If
            Next


        End If
    Next i


'    For i = 0 To Me.lstDeudaCompensatorios.ListCount - 1
'        If Me.lstDeudaCompensatorios.Checked(i) Then
'
'            If funciones.BuscarEnColeccion(colDeudaCompensatorios, CStr(Me.lstDeudaCompensatorios.ItemData(i))) Then
'                colc.Add colDeudaCompensatorios.item(CStr(Me.lstDeudaCompensatorios.ItemData(i)))
'
'
'            End If
'
'        End If
'    Next i

    TotalizarDiferenciasCambio
    MostrarPosiblesRetenciones col, colc
End Sub


Sub limpiarParciales()
    Me.txtParcialAbonado = 0
    Me.txtParcialAbonar = 0
    Me.txtOtrosParcialAbonado = 0
    Me.txtOtrosParcialAbonar = 0
    Me.txtTotalParcialAbonado = 0
    Me.txtTotalParcialAbonar = 0

    Me.lblCantidadCbtesSeleccionados.caption = "Cbtes. Seleccionados: 0"
End Sub

Private Sub lstFacturas_ItemCheck(ByVal item As Long)

    If item < -1 Then
                 Dim f1
                         Set f1 = DAOFacturaProveedor.FindById(CStr(Me.lstFacturas.ItemData(item)))

    End If

Me.lblCantidadCbtesSeleccionados.caption = "Cbtes. Seleccionados: 0"
    calcularOrigenes


    If lstFacturas.ListCount > 0 And item > -1 Then

        Dim x As Integer

''            Me.txtParcialAbonado.Enabled = lstFacturas.Checked(item)
'            Me.txtParcialAbonar.Enabled = lstFacturas.Checked(item)
'            Me.txtOtrosParcialAbonado.Enabled = lstFacturas.Checked(item)
'            Me.txtOtrosParcialAbonar.Enabled = lstFacturas.Checked(item)
'            Me.txtTotalParcialAbonado.Enabled = lstFacturas.Checked(item)
'            Me.txtTotalParcialAbonar.Enabled = lstFacturas.Checked(item)



    End If

End Sub

'Private Sub lstFacturas_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'    Dim i As Integer
'    If Button = 2 Then
'
'        For i = 0 To Me.lstFacturas.ListCount - 1
'
'            If Me.lstFacturas.Selected(i) Then
'                Me.mnuCrearCompensatorio.Enabled = Me.lstFacturas.Checked(i)
'                PopupMenu Me.emergente
'            End If
'        Next
'
'
'    End If
'
'End Sub

Private Sub mnuCrearCompensatorio_Click()

    Dim d As New frmCrearCompensatorio
    Dim i As Long
    Dim ivamax As Boolean

    For i = 0 To Me.lstFacturas.ListCount - 1
        If Me.lstFacturas.Selected(i) Then
            Set Factura = colFacturas(CStr(Me.lstFacturas.ItemData(i)))

            If Factura.IvaAplicado.count > 1 Then ivamax = True


            'chequeo que no exista un compensatorio para esa factura.

            Dim c As Compensatorio
            Dim hay As Boolean
            hay = False
            For Each c In LiquidacionCaja.Compensatorios
                If c.Comprobante.Id = Factura.Id Then
                    hay = True
                    Exit For
                End If

            Next c

            Dim Cant As Long

            If DAOCompensatorios.FindAll("id_orden_pago= " & LiquidacionCaja.Id & " and  id_comprobante=" & Factura.Id).count > 0 Then hay = True

            If hay Then
                MsgBox "Ya existe un compensatorio para el comprobante indicado!", vbInformation, "Error"
            Else
                If ivamax Then
                    MsgBox "No puede crear un compensatorio cuando hay multiples alícuotas!", vbInformation, "Error"
                Else
                    d.Cargar Factura, OrdenPago
                    d.Show 1
                    mostrarCompensatorios
                    lstFacturas_ItemCheck 1
                End If
            End If
        End If
    Next i
End Sub

Private Sub mostrarCompensatorios()
    Me.gridCompensatorios.ItemCount = LiquidacionCaja.Compensatorios.count
    verCompensatorios
End Sub



Private Sub PushButton1_Click()

    If Me.cboProveedores.ListIndex <> -1 Then
        Set prov = colProveedores.item(CStr(Me.cboProveedores.ItemData(Me.cboProveedores.ListIndex)))

                If IsSomething(prov) Then
                    Dim Nueva As New Collection
                Set Nueva = DAORetenciones.FindAllWithAlicuotas(prov.Cuit) '


                   Set alicuotas = DAORetenciones.FindAllWithAlicuotas(prov.Cuit) '
        ActualizarAlicuotas
                End If
    Else
        Set prov = Nothing

    End If

    MostrarFacturas
End Sub

Private Sub radioConcepto_Click()
    If formLoaded Then
        LimpiarFacturasYValores
        MostrarPosiblesRetenciones New Collection
        Totalizar
    End If
    ActivarControles
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

    Me.lblTotal.caption = "Total orden de pago en " & FormatCurrency(funciones.FormatearDecimales(LiquidacionCaja.StaticTotalOrigenes + LiquidacionCaja.StaticTotalRetenido))
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

'    Me.txtDiferenciaCambioPago.text = T
''    Me.txtDifTipoCambioIVA.text = TIVA
'    Me.txtDifCambio = TTOTAL


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


Private Sub gridCheques_BeforeUpdate(ByVal Cancel As GridEX20.JSRetBoolean)
    Dim msg As New Collection

' REVISA QUE EN LA COLECCION DE CHEQUES DE TERCEROS QUE SE ESTAN CARGANDO NO ESTÉ INGRESADO EL MISMO CHEQUE, SI LO DETECTA GENERA MSG DE ERROR
    If funciones.BuscarEnColeccion(LiquidacionCaja.ChequesTerceros, CStr(Me.gridCheques.value(1))) Then
        msg.Add "El cheque seleccionado ya fue ingresado anteriormente."
    End If

    Cancel = (msg.count > 0)
    If Cancel Then MsgBox funciones.JoinCollectionValues(msg, vbNewLine), vbExclamation

End Sub



Private Sub gridCheques_UnboundAddNew(ByVal NewRowBookmark As GridEX20.JSRetVariant, ByVal Values As GridEX20.JSRowData)
    Set cheque = Nothing
    If IsNumeric(Values(1)) Then Set cheque = DAOCheques.FindById(Values(1))
    If IsSomething(cheque) Then
        LiquidacionCaja.ChequesTerceros.Add cheque, CStr(cheque.Id)

    End If
    Totalizar


End Sub

Private Sub gridCheques_UnboundDelete(ByVal RowIndex As Long, ByVal Bookmark As Variant)
    If RowIndex > 0 Then
        LiquidacionCaja.ChequesTerceros.remove RowIndex
        Totalizar
    End If
End Sub


Private Sub gridCheques_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex <= LiquidacionCaja.ChequesTerceros.count Then
        Set cheque = LiquidacionCaja.ChequesTerceros.item(RowIndex)


        'Values(1) = "ID: " & cheque.Id & "N " & cheque.numero
        'Values(1) = "ID: " & cheque.numero & "N " & cheque.numero
        Values(1) = cheque.numero & " "
        'Values(1) = cheque.numero

        'If IsNumeric(Values(1)) Then Values(1) = cheque.numero



        'FORMATCURRENCY
        Values(2) = FormatCurrency(cheque.Monto)
        If IsSomething(cheque.moneda) Then Values(3) = cheque.moneda.NombreCorto
        If IsSomething(cheque.Banco) Then Values(4) = cheque.Banco.nombre
        Values(5) = cheque.OrigenDestino
        Values(6) = cheque.OrigenCheque
 '       Totalizar
    End If
End Sub


Private Sub gridCheques_UnboundUpdate(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex > 0 And LiquidacionCaja.ChequesTerceros.count >= RowIndex Then
        Set cheque = Nothing
        If IsNumeric(Values(1)) Then Set cheque = DAOCheques.FindById(Values(1))
        If IsSomething(cheque) Then
            LiquidacionCaja.ChequesTerceros.Add cheque, , , RowIndex
            LiquidacionCaja.ChequesTerceros.remove RowIndex
        End If
        Totalizar
    End If
End Sub

Private Sub Text1_Change()

End Sub

Private Sub MostrarFacturas_Click()
MostrarFacturas
End Sub

Private Sub txtBuscarFactura_GotFocus()
    Me.txtBuscarFactura.SelStart = 0
    Me.txtBuscarFactura.SelLength = Len(Me.txtBuscarFactura.text)
End Sub

'Private Sub txtBuscarFactura_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = 13 Then
'        'buscar en facturas y tildar
'
'        If LenB(Me.txtBuscarFactura.text) > 0 Then
'            Dim cont As Long
'
'            If colFacturas.count > 0 Then
'                Dim i As Long
'                For Each vFacturaProveedor In colFacturas
'                    If InStr(1, vFacturaProveedor.numero, Me.txtBuscarFactura.text) > 0 Then    'aplica
'                        For i = 0 To Me.lstFacturas.ListCount - 1
'                            If Me.lstFacturas.ItemData(i) = vFacturaProveedor.Id Then
'                                Me.lstFacturas.Checked(i) = True
'                                cont = cont + 1
'                                Exit For
'                            End If
'                        Next i
'                    End If
'                Next vFacturaProveedor
'
'                If cont = 0 Then
'                    MsgBox "No se encontraron facturas con ese número en la lista.", vbOKOnly + vbExclamation
'                Else
'                    lstFacturas_ItemCheck -1
'                    MsgBox "Se encontró " & cont & " factura/s.", vbOKOnly + vbInformation
'                    Me.txtBuscarFactura.text = vbNullString
'                    Me.txtBuscarFactura.SetFocus
'                End If
'            End If
'        End If
'    End If
'End Sub

Private Sub txtBuscarFactura_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        'buscar en facturas y tildar

        If LenB(Me.txtBuscarFactura.text) > 0 Then
            Dim cont As Long
            Dim i As Long
            
            ' Elimina todas las facturas existentes en la lista
            Me.lstFacturas.Clear
            
            If colFacturas.count > 0 Then
                For Each vFacturaProveedor In colFacturas
                    If InStr(1, vFacturaProveedor.numero, Me.txtBuscarFactura.text) > 0 Then    'aplica
                        ' Agrega solo las facturas que coinciden con el texto del cuadro de búsqueda
                        Me.lstFacturas.AddItem vFacturaProveedor.numero
                        Me.lstFacturas.ItemData(Me.lstFacturas.ListCount - 1) = vFacturaProveedor.Id
                        
                        cont = cont + 1
                    End If
                Next vFacturaProveedor
                
                If cont = 0 Then
                    MsgBox "No se encontraron facturas con ese número en la lista.", vbOKOnly + vbExclamation
                Else
                    lstFacturas_ItemCheck -1
                    MsgBox "Se encontró " & cont & " factura/s.", vbOKOnly + vbInformation
                    Me.txtBuscarFactura.text = vbNullString
                    Me.txtBuscarFactura.SetFocus
                End If
            End If
        End If
    End If
End Sub



Private Sub txtDifCambio_GotFocus()
    foco Me.txtDifCambio
End Sub


Private Sub txtDifCambioNG1_Change()
    LiquidacionCaja.DiferenciaCambioEnNG = Val(Me.txtDifCambioNG1)
    Totalizar
End Sub

Private Sub txtDifCambioTOTAL1_Change()
    LiquidacionCaja.DiferenciaCambioEnTOTAL = Val(Me.txtDifCambioTOTAL1)
    Totalizar
End Sub

Private Sub txtnetogravadoabonado_Change()
    If LenB(Me.txtnetogravadoabonado) > 0 Then
        vFactElegida.NetoGravadoAbonado = CDbl(Me.txtnetogravadoabonado)
    Else
        vFactElegida.ImporteTotalAbonado = 0
    End If

    Totalizar
End Sub

Private Sub txtOtrosDescuentos_LostFocus()
    LiquidacionCaja.OtrosDescuentos = Val(Me.txtOtrosDescuentos.text)
    Totalizar
End Sub


Public Sub RecalcularOtrosFacturaelegida()
 If LenB(Me.txtOtrosParcialAbonar) > 0 And IsNumeric(Me.txtOtrosParcialAbonar) Then

        vFactElegida.OtrosAbonado = CDbl(Me.txtOtrosParcialAbonar)
        RecalcularTotalFacturaElegida


    End If

End Sub

Private Sub txtOtrosParcialAbonar_KeyUp(KeyCode As Integer, Shift As Integer)
 RecalcularOtrosFacturaelegida

    Totalizar
End Sub

''''''Private Sub RecalcularTotalFacturaElegida()
''''''    Me.txtTotalParcialAbonar = (CDbl(txtParcialAbonar)) + (CDbl(Me.txtOtrosParcialAbonar))
''''''
''''''    If Me.txtTotalParcialAbonar = "0" Then Me.txtTotalParcialAbonar = "0.00"
''''''
''''''
''''''       vFactElegida.TotalAbonado = CDbl(txtTotalParcialAbonar)
''''''
''''''End Sub


Private Sub txtOtrosParcialAbonar_LostFocus()
'  If LenB(Me.txtOtrosParcialAbonar) > 0 Then
'
'        vFactElegida.OtrosAbonado = CDbl(Me.txtOtrosParcialAbonar)
'        recalcularTotalFacturaelegida
'
'
'    End If
'
'    Totalizar
End Sub

Private Sub txtOtrosParcialAbonar_Validate(Cancel As Boolean)
If Not IsNumeric(Me.txtOtrosParcialAbonar) Then
 Cancel = True
Else
'COMENTO ESTA LINEA PORQUE ESTA COMPROBACIÓN HACE QUE EL FORM SE CONGELE Y NO SE PUEDA AVANZAR CON LA CARGA.
'QUEDA PARA VER CON NICOLAS

'Cancel = CDbl(Me.txtOtrosParcialAbonar) > vFactElegida.ImporteOtrosSaldo Or Not IsNumeric(Me.txtOtrosParcialAbonar) Or CDbl(Me.txtOtrosParcialAbonar) < 0
End If
If Cancel Then
    Me.txtOtrosParcialAbonar.backColor = vbRed
    Me.txtOtrosParcialAbonar.ForeColor = vbWhite
Else
    Me.txtOtrosParcialAbonar.backColor = vbWhite
    Me.txtOtrosParcialAbonar.ForeColor = vbBlack
End If
End Sub


'Private Sub RecalcularFacturaElegida()
'RecalcularNetoGravadoFacturaElegida
'RecalcularOtrosFacturaelegida
'End Sub

'''''Private Sub RecalcularNetoGravadoFacturaElegida()
''''' If LenB(txtParcialAbonar) > 0 And IsNumeric(txtParcialAbonar) Then
'''''
'''''
'''''       vFactElegida.NetoGravadoAbonado = CDbl(txtParcialAbonar)
'''''        RecalcularTotalFacturaElegida
'''''    End If
'''''End Sub

Private Sub txtParcialAbonar_KeyUp(KeyCode As Integer, Shift As Integer)
  RecalcularNetoGravadoFacturaElegida

'

    Totalizar
End Sub

Private Sub txtParcialAbonar_Validate(Cancel As Boolean)
If Not IsNumeric(Me.txtParcialAbonar) Then
Cancel = True
Else
'Cancel = CDbl(Me.txtParcialAbonar) > vFactElegida.ImporteNetoGravadoSaldo Or Not IsNumeric(Me.txtParcialAbonar) Or CDbl(Me.txtParcialAbonar) < 0
End If
If Cancel Then
    Me.txtParcialAbonar.backColor = vbRed
    Me.txtParcialAbonar.ForeColor = vbWhite
Else
    Me.txtParcialAbonar.backColor = vbWhite
    Me.txtParcialAbonar.ForeColor = vbBlack
End If
End Sub

Private Sub txtRetenciones_GotFocus()
    foco Me.txtRetenciones
End Sub

Private Sub txtRetenciones_LostFocus()
    Totalizar
End Sub

Private Sub txtRetenciones_Validate(Cancel As Boolean)
    funciones.ValidarTextBox Me.txtRetenciones, Cancel
End Sub


Private Sub txtTotalParcialAbonar_Change()
 If IsSomething(vFactElegida) Then
     If CDbl(Me.txtTotalParcialAbonar) > vFactElegida.ImporteTotalSaldo Or CDbl(Me.txtParcialAbonar) < 0 Then
                  Me.txtTotalParcialAbonar.backColor = vbRed
        Me.txtTotalParcialAbonar.ForeColor = vbWhite
    Else
        Me.txtTotalParcialAbonar.backColor = vbWhite
        Me.txtTotalParcialAbonar.ForeColor = vbBlack
    End If
End If
End Sub

