VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmAdminPagosCrearOrdenPago 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Orden de Pago"
   ClientHeight    =   12780
   ClientLeft      =   2340
   ClientTop       =   3105
   ClientWidth     =   17580
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAdminPagosCrearOrdenPago.frx":0000
   LinkTopic       =   "Orden de Pago"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   12780
   ScaleWidth      =   17580
   Begin XtremeSuiteControls.GroupBox GroupBox5 
      Height          =   1815
      Left            =   120
      TabIndex        =   82
      Top             =   9600
      Width           =   6885
      _Version        =   786432
      _ExtentX        =   12144
      _ExtentY        =   3201
      _StockProps     =   79
      Caption         =   "Detalle de comprobante"
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
      Begin GridEX20.GridEX gridDetalleComprobante 
         Height          =   1455
         Left            =   120
         TabIndex        =   83
         Top             =   240
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   2566
         Version         =   "2.0"
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         ColumnAutoResize=   -1  'True
         MethodHoldFields=   -1  'True
         GroupByBoxVisible=   0   'False
         DataMode        =   99
         ColumnHeaderHeight=   285
         IntProp1        =   0
         ColumnsCount    =   3
         Column(1)       =   "frmAdminPagosCrearOrdenPago.frx":000C
         Column(2)       =   "frmAdminPagosCrearOrdenPago.frx":0154
         Column(3)       =   "frmAdminPagosCrearOrdenPago.frx":0294
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmAdminPagosCrearOrdenPago.frx":03CC
         FormatStyle(2)  =   "frmAdminPagosCrearOrdenPago.frx":04F4
         FormatStyle(3)  =   "frmAdminPagosCrearOrdenPago.frx":05A4
         FormatStyle(4)  =   "frmAdminPagosCrearOrdenPago.frx":0658
         FormatStyle(5)  =   "frmAdminPagosCrearOrdenPago.frx":0730
         FormatStyle(6)  =   "frmAdminPagosCrearOrdenPago.frx":07E8
         ImageCount      =   0
         PrinterProperties=   "frmAdminPagosCrearOrdenPago.frx":08C8
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox7 
      Height          =   5535
      Left            =   13800
      TabIndex        =   65
      Top             =   120
      Width           =   3660
      _Version        =   786432
      _ExtentX        =   6456
      _ExtentY        =   9763
      _StockProps     =   79
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
      Begin XtremeSuiteControls.PushButton btnMoneda 
         Height          =   495
         Left            =   3240
         TabIndex        =   81
         Top             =   9840
         Width           =   495
         _Version        =   786432
         _ExtentX        =   873
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "PushButton1"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnExportarDatos 
         Height          =   495
         Left            =   840
         TabIndex        =   76
         Top             =   4320
         Width           =   1935
         _Version        =   786432
         _ExtentX        =   3413
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Exportar"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnGuardar 
         Height          =   525
         Left            =   840
         TabIndex        =   77
         Top             =   4920
         Width           =   1950
         _Version        =   786432
         _ExtentX        =   3440
         _ExtentY        =   926
         _StockProps     =   79
         Caption         =   "Guardar"
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
      End
      Begin XtremeSuiteControls.Label lblTotalPercepciones 
         Height          =   255
         Left            =   120
         TabIndex        =   85
         Top             =   2760
         Width           =   3375
         _Version        =   786432
         _ExtentX        =   5953
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Total Percepciones:"
      End
      Begin XtremeSuiteControls.Label lblFacturasTotal 
         Height          =   375
         Left            =   120
         TabIndex        =   80
         Top             =   3720
         Visible         =   0   'False
         Width           =   2535
         _Version        =   786432
         _ExtentX        =   4471
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "lblFacturasTotal"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XtremeSuiteControls.Label lblTotalPagoACuenta 
         Height          =   255
         Left            =   120
         TabIndex        =   78
         Top             =   3120
         Width           =   3375
         _Version        =   786432
         _ExtentX        =   5953
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "lblTotalPagoACuenta"
      End
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         Caption         =   "Total Pagos:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   74
         Tag             =   "Total: "
         Top             =   360
         Width           =   1035
      End
      Begin VB.Label lblNgAbonar 
         AutoSize        =   -1  'True
         Caption         =   "Neto gravado a abonar:"
         Height          =   195
         Left            =   120
         TabIndex        =   73
         Top             =   1800
         Width           =   1740
      End
      Begin VB.Label lblTotalCompensatorios 
         AutoSize        =   -1  'True
         Caption         =   "Total compensatorios: "
         Height          =   195
         Left            =   120
         TabIndex        =   72
         Tag             =   "Total: "
         Top             =   2280
         Width           =   1635
      End
      Begin VB.Label lblTotalARetener 
         AutoSize        =   -1  'True
         Caption         =   "Total a retener:"
         Height          =   195
         Left            =   120
         TabIndex        =   71
         Top             =   600
         Width           =   1140
      End
      Begin VB.Label lblTotalFacturas 
         AutoSize        =   -1  'True
         Caption         =   "Total facturas: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   70
         Top             =   1080
         Width           =   1275
      End
      Begin VB.Label lblTotalOrdenPago 
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total a pagar:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   69
         Tag             =   "tot fac - tot ret"
         Top             =   840
         Width           =   1230
      End
      Begin VB.Label lblTotalFacturasNG 
         AutoSize        =   -1  'True
         Caption         =   "Total NG Facturas: "
         Height          =   195
         Left            =   120
         TabIndex        =   68
         Top             =   1320
         Width           =   1395
      End
      Begin VB.Label lblDiferenciaCambio 
         AutoSize        =   -1  'True
         Caption         =   "Diferencia Cambio:"
         Height          =   195
         Left            =   120
         TabIndex        =   67
         Top             =   2040
         Width           =   1350
      End
      Begin VB.Label lblDeudaCompensatorios 
         AutoSize        =   -1  'True
         Caption         =   "Total compensatorios pendientes:"
         Height          =   195
         Left            =   120
         TabIndex        =   66
         Top             =   1560
         Width           =   2430
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox6 
      Height          =   2055
      Left            =   7080
      TabIndex        =   64
      Top             =   3600
      Width           =   6660
      _Version        =   786432
      _ExtentX        =   11747
      _ExtentY        =   3625
      _StockProps     =   79
      Caption         =   "Pagos a Cuenta"
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
      Begin XtremeSuiteControls.ListBox ListPagosACuenta 
         Height          =   1575
         Left            =   120
         TabIndex        =   75
         Top             =   240
         Width           =   6375
         _Version        =   786432
         _ExtentX        =   11245
         _ExtentY        =   2778
         _StockProps     =   77
         BackColor       =   -2147483643
         Style           =   1
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox3 
      Height          =   1335
      Left            =   120
      TabIndex        =   39
      Top             =   0
      Width           =   6900
      _Version        =   786432
      _ExtentX        =   12171
      _ExtentY        =   2355
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin VB.TextBox txtOtrosDescuentos 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5880
         TabIndex        =   42
         Top             =   225
         Width           =   960
      End
      Begin VB.TextBox txtDifCambioNG1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5880
         TabIndex        =   41
         Top             =   600
         Width           =   960
      End
      Begin VB.TextBox txtDifCambioTOTAL1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5880
         TabIndex        =   40
         Top             =   945
         Width           =   960
      End
      Begin XtremeSuiteControls.ComboBox cboMonedas 
         Height          =   315
         Left            =   885
         TabIndex        =   46
         Top             =   240
         Width           =   1245
         _Version        =   786432
         _ExtentX        =   2196
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Style           =   2
         Text            =   "cboMonedas"
      End
      Begin XtremeSuiteControls.DateTimePicker dtpFecha 
         Height          =   330
         Left            =   885
         TabIndex        =   47
         Top             =   735
         Width           =   1245
         _Version        =   786432
         _ExtentX        =   2196
         _ExtentY        =   582
         _StockProps     =   68
         Format          =   1
         CurrentDate     =   40183.7263657407
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         Height          =   195
         Left            =   375
         TabIndex        =   49
         Tag             =   "Total: "
         Top             =   810
         Width           =   435
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Moneda"
         Height          =   195
         Left            =   240
         TabIndex        =   48
         Tag             =   "Total: "
         Top             =   300
         Width           =   570
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Otros Descuentos"
         Height          =   195
         Left            =   4440
         TabIndex        =   45
         Top             =   270
         Width           =   1275
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Dif. Cambio manual NG "
         Height          =   195
         Left            =   4080
         TabIndex        =   44
         Top             =   645
         Width           =   1680
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Dif. Cambio manual TOTAL"
         Height          =   195
         Left            =   3840
         TabIndex        =   43
         Top             =   990
         Width           =   1905
      End
   End
   Begin VB.TextBox txtnetogravadoabonado 
      Height          =   315
      Left            =   3600
      TabIndex        =   24
      Top             =   240
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.TextBox txtDifTipoCambioIVA 
      Height          =   285
      Left            =   3000
      TabIndex        =   23
      Top             =   840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtDiferenciaCambioPago 
      Height          =   285
      Left            =   4680
      TabIndex        =   22
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtDifCambio 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2880
      TabIndex        =   13
      Top             =   840
      Visible         =   0   'False
      Width           =   1200
   End
   Begin XtremeSuiteControls.GroupBox grpOrigen 
      Height          =   5655
      Left            =   7080
      TabIndex        =   0
      Top             =   5760
      Width           =   10380
      _Version        =   786432
      _ExtentX        =   18309
      _ExtentY        =   9975
      _StockProps     =   79
      Caption         =   "Valores de pago"
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
         Height          =   5265
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   10140
         _Version        =   786432
         _ExtentX        =   17886
         _ExtentY        =   9287
         _StockProps     =   68
         Appearance      =   10
         Color           =   32
         PaintManager.ShowIcons=   -1  'True
         ItemCount       =   6
         Item(0).Caption =   "Cheques Propios"
         Item(0).ControlCount=   2
         Item(0).Control(0)=   "gridChequesPropios"
         Item(0).Control(1)=   "txtTotalizadorCHEQUESPROPIOS"
         Item(1).Caption =   "Banco"
         Item(1).ControlCount=   2
         Item(1).Control(0)=   "gridDepositosOperaciones"
         Item(1).Control(1)=   "txtTotalizadorBANCO"
         Item(2).Caption =   "Cheques 3ros"
         Item(2).ControlCount=   2
         Item(2).Control(0)=   "gridCheques"
         Item(2).Control(1)=   "txtTotalizadorCHEQUES3ROS"
         Item(3).Caption =   "Caja"
         Item(3).ControlCount=   2
         Item(3).Control(0)=   "gridCajaOperaciones"
         Item(3).Control(1)=   "txtTotalizadorCAJA"
         Item(4).Caption =   "Percepciones"
         Item(4).ControlCount=   2
         Item(4).Control(0)=   "gridPercepciones"
         Item(4).Control(1)=   "txtTotalizadorPERCEPCIONES"
         Item(5).Caption =   "Compensatorios"
         Item(5).ControlCount=   2
         Item(5).Control(0)=   "gridCompensatorios"
         Item(5).Control(1)=   "txtTotalizadorCOMPENSATORIOS"
         Begin GridEX20.GridEX gridPercepciones 
            Height          =   4335
            Left            =   -69880
            TabIndex        =   84
            Top             =   435
            Visible         =   0   'False
            Width           =   9810
            _ExtentX        =   17304
            _ExtentY        =   7646
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
            Column(1)       =   "frmAdminPagosCrearOrdenPago.frx":0A98
            Column(2)       =   "frmAdminPagosCrearOrdenPago.frx":0BD0
            Column(3)       =   "frmAdminPagosCrearOrdenPago.frx":0D0C
            Column(4)       =   "frmAdminPagosCrearOrdenPago.frx":0E40
            Column(5)       =   "frmAdminPagosCrearOrdenPago.frx":0F60
            FormatStylesCount=   6
            FormatStyle(1)  =   "frmAdminPagosCrearOrdenPago.frx":1074
            FormatStyle(2)  =   "frmAdminPagosCrearOrdenPago.frx":119C
            FormatStyle(3)  =   "frmAdminPagosCrearOrdenPago.frx":124C
            FormatStyle(4)  =   "frmAdminPagosCrearOrdenPago.frx":1300
            FormatStyle(5)  =   "frmAdminPagosCrearOrdenPago.frx":13D8
            FormatStyle(6)  =   "frmAdminPagosCrearOrdenPago.frx":1490
            ImageCount      =   0
            PrinterProperties=   "frmAdminPagosCrearOrdenPago.frx":1570
         End
         Begin GridEX20.GridEX gridDepositosOperaciones 
            Height          =   4335
            Left            =   -69880
            TabIndex        =   2
            Top             =   435
            Visible         =   0   'False
            Width           =   9810
            _ExtentX        =   17304
            _ExtentY        =   7646
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
            Column(1)       =   "frmAdminPagosCrearOrdenPago.frx":1740
            Column(2)       =   "frmAdminPagosCrearOrdenPago.frx":18A0
            Column(3)       =   "frmAdminPagosCrearOrdenPago.frx":19DC
            Column(4)       =   "frmAdminPagosCrearOrdenPago.frx":1B10
            Column(5)       =   "frmAdminPagosCrearOrdenPago.frx":1C54
            FormatStylesCount=   6
            FormatStyle(1)  =   "frmAdminPagosCrearOrdenPago.frx":1D58
            FormatStyle(2)  =   "frmAdminPagosCrearOrdenPago.frx":1E90
            FormatStyle(3)  =   "frmAdminPagosCrearOrdenPago.frx":1F40
            FormatStyle(4)  =   "frmAdminPagosCrearOrdenPago.frx":1FF4
            FormatStyle(5)  =   "frmAdminPagosCrearOrdenPago.frx":20CC
            FormatStyle(6)  =   "frmAdminPagosCrearOrdenPago.frx":2184
            ImageCount      =   0
            PrinterProperties=   "frmAdminPagosCrearOrdenPago.frx":2264
         End
         Begin GridEX20.GridEX gridCajaOperaciones 
            Height          =   4335
            Left            =   -69880
            TabIndex        =   10
            Top             =   435
            Visible         =   0   'False
            Width           =   9810
            _ExtentX        =   17304
            _ExtentY        =   7646
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
            Column(1)       =   "frmAdminPagosCrearOrdenPago.frx":243C
            Column(2)       =   "frmAdminPagosCrearOrdenPago.frx":259C
            Column(3)       =   "frmAdminPagosCrearOrdenPago.frx":26D8
            Column(4)       =   "frmAdminPagosCrearOrdenPago.frx":280C
            Column(5)       =   "frmAdminPagosCrearOrdenPago.frx":2940
            FormatStylesCount=   6
            FormatStyle(1)  =   "frmAdminPagosCrearOrdenPago.frx":2A44
            FormatStyle(2)  =   "frmAdminPagosCrearOrdenPago.frx":2B7C
            FormatStyle(3)  =   "frmAdminPagosCrearOrdenPago.frx":2C2C
            FormatStyle(4)  =   "frmAdminPagosCrearOrdenPago.frx":2CE0
            FormatStyle(5)  =   "frmAdminPagosCrearOrdenPago.frx":2DB8
            FormatStyle(6)  =   "frmAdminPagosCrearOrdenPago.frx":2E70
            ImageCount      =   0
            PrinterProperties=   "frmAdminPagosCrearOrdenPago.frx":2F50
         End
         Begin GridEX20.GridEX gridChequesPropios 
            Height          =   4335
            Left            =   120
            TabIndex        =   9
            Top             =   435
            Width           =   9810
            _ExtentX        =   17304
            _ExtentY        =   7646
            Version         =   "2.0"
            BoundColumnIndex=   ""
            ReplaceColumnIndex=   ""
            ColumnAutoResize=   -1  'True
            MethodHoldFields=   -1  'True
            ContScroll      =   -1  'True
            AllowColumnDrag =   0   'False
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
            Column(1)       =   "frmAdminPagosCrearOrdenPago.frx":3128
            Column(2)       =   "frmAdminPagosCrearOrdenPago.frx":3290
            Column(3)       =   "frmAdminPagosCrearOrdenPago.frx":33C4
            Column(4)       =   "frmAdminPagosCrearOrdenPago.frx":3500
            Column(5)       =   "frmAdminPagosCrearOrdenPago.frx":3668
            FormatStylesCount=   6
            FormatStyle(1)  =   "frmAdminPagosCrearOrdenPago.frx":3760
            FormatStyle(2)  =   "frmAdminPagosCrearOrdenPago.frx":3898
            FormatStyle(3)  =   "frmAdminPagosCrearOrdenPago.frx":3948
            FormatStyle(4)  =   "frmAdminPagosCrearOrdenPago.frx":39FC
            FormatStyle(5)  =   "frmAdminPagosCrearOrdenPago.frx":3AD4
            FormatStyle(6)  =   "frmAdminPagosCrearOrdenPago.frx":3B8C
            ImageCount      =   0
            PrinterProperties=   "frmAdminPagosCrearOrdenPago.frx":3C6C
         End
         Begin GridEX20.GridEX gridCheques 
            Height          =   4335
            Left            =   -69880
            TabIndex        =   8
            Top             =   435
            Visible         =   0   'False
            Width           =   9810
            _ExtentX        =   17304
            _ExtentY        =   7646
            Version         =   "2.0"
            BoundColumnIndex=   ""
            ReplaceColumnIndex=   ""
            ColumnAutoResize=   -1  'True
            MethodHoldFields=   -1  'True
            ContScroll      =   -1  'True
            AllowColumnDrag =   0   'False
            AllowDelete     =   -1  'True
            GroupByBoxVisible=   0   'False
            RowHeaders      =   -1  'True
            DataMode        =   99
            AllowAddNew     =   -1  'True
            ColumnHeaderHeight=   285
            IntProp1        =   0
            IntProp2        =   0
            IntProp7        =   0
            ColumnsCount    =   7
            Column(1)       =   "frmAdminPagosCrearOrdenPago.frx":3E44
            Column(2)       =   "frmAdminPagosCrearOrdenPago.frx":3FC4
            Column(3)       =   "frmAdminPagosCrearOrdenPago.frx":4164
            Column(4)       =   "frmAdminPagosCrearOrdenPago.frx":425C
            Column(5)       =   "frmAdminPagosCrearOrdenPago.frx":4398
            Column(6)       =   "frmAdminPagosCrearOrdenPago.frx":44A4
            Column(7)       =   "frmAdminPagosCrearOrdenPago.frx":4574
            FormatStylesCount=   6
            FormatStyle(1)  =   "frmAdminPagosCrearOrdenPago.frx":4660
            FormatStyle(2)  =   "frmAdminPagosCrearOrdenPago.frx":4798
            FormatStyle(3)  =   "frmAdminPagosCrearOrdenPago.frx":4848
            FormatStyle(4)  =   "frmAdminPagosCrearOrdenPago.frx":48FC
            FormatStyle(5)  =   "frmAdminPagosCrearOrdenPago.frx":49D4
            FormatStyle(6)  =   "frmAdminPagosCrearOrdenPago.frx":4A8C
            ImageCount      =   0
            PrinterProperties=   "frmAdminPagosCrearOrdenPago.frx":4B6C
         End
         Begin GridEX20.GridEX gridCompensatorios 
            Height          =   4335
            Left            =   -69880
            TabIndex        =   14
            Top             =   435
            Visible         =   0   'False
            Width           =   9810
            _ExtentX        =   17304
            _ExtentY        =   7646
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
            Column(1)       =   "frmAdminPagosCrearOrdenPago.frx":4D44
            Column(2)       =   "frmAdminPagosCrearOrdenPago.frx":4E8C
            Column(3)       =   "frmAdminPagosCrearOrdenPago.frx":4F98
            Column(4)       =   "frmAdminPagosCrearOrdenPago.frx":5084
            Column(5)       =   "frmAdminPagosCrearOrdenPago.frx":5188
            FormatStylesCount=   6
            FormatStyle(1)  =   "frmAdminPagosCrearOrdenPago.frx":52C8
            FormatStyle(2)  =   "frmAdminPagosCrearOrdenPago.frx":5400
            FormatStyle(3)  =   "frmAdminPagosCrearOrdenPago.frx":54B0
            FormatStyle(4)  =   "frmAdminPagosCrearOrdenPago.frx":5564
            FormatStyle(5)  =   "frmAdminPagosCrearOrdenPago.frx":563C
            FormatStyle(6)  =   "frmAdminPagosCrearOrdenPago.frx":56F4
            ImageCount      =   0
            PrinterProperties=   "frmAdminPagosCrearOrdenPago.frx":57D4
         End
         Begin XtremeSuiteControls.Label txtTotalizadorCOMPENSATORIOS 
            Height          =   375
            Left            =   -69880
            TabIndex        =   91
            Top             =   4800
            Visible         =   0   'False
            Width           =   9855
            _Version        =   786432
            _ExtentX        =   17383
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Label13"
         End
         Begin XtremeSuiteControls.Label txtTotalizadorPERCEPCIONES 
            Height          =   375
            Left            =   -69880
            TabIndex        =   90
            Top             =   4800
            Visible         =   0   'False
            Width           =   9855
            _Version        =   786432
            _ExtentX        =   17383
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Label13"
         End
         Begin XtremeSuiteControls.Label txtTotalizadorCAJA 
            Height          =   375
            Left            =   -69880
            TabIndex        =   89
            Top             =   4800
            Visible         =   0   'False
            Width           =   9855
            _Version        =   786432
            _ExtentX        =   17383
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Label13"
         End
         Begin XtremeSuiteControls.Label txtTotalizadorCHEQUES3ROS 
            Height          =   375
            Left            =   -69880
            TabIndex        =   88
            Top             =   4800
            Visible         =   0   'False
            Width           =   9855
            _Version        =   786432
            _ExtentX        =   17383
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Label13"
         End
         Begin XtremeSuiteControls.Label txtTotalizadorBANCO 
            Height          =   375
            Left            =   -69880
            TabIndex        =   87
            Top             =   4800
            Visible         =   0   'False
            Width           =   9855
            _Version        =   786432
            _ExtentX        =   17383
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Label13"
         End
         Begin XtremeSuiteControls.Label txtTotalizadorCHEQUESPROPIOS 
            Height          =   375
            Left            =   120
            TabIndex        =   86
            Top             =   4800
            Width           =   9855
            _Version        =   786432
            _ExtentX        =   17383
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Label13"
         End
      End
   End
   Begin GridEX20.GridEX gridBancos 
      Height          =   1845
      Left            =   5160
      TabIndex        =   3
      Top             =   12000
      Visible         =   0   'False
      Width           =   5745
      _ExtentX        =   10134
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
      Column(1)       =   "frmAdminPagosCrearOrdenPago.frx":59AC
      Column(2)       =   "frmAdminPagosCrearOrdenPago.frx":5AAC
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminPagosCrearOrdenPago.frx":5B9C
      FormatStyle(2)  =   "frmAdminPagosCrearOrdenPago.frx":5CD4
      FormatStyle(3)  =   "frmAdminPagosCrearOrdenPago.frx":5D84
      FormatStyle(4)  =   "frmAdminPagosCrearOrdenPago.frx":5E38
      FormatStyle(5)  =   "frmAdminPagosCrearOrdenPago.frx":5F10
      FormatStyle(6)  =   "frmAdminPagosCrearOrdenPago.frx":5FC8
      ImageCount      =   0
      PrinterProperties=   "frmAdminPagosCrearOrdenPago.frx":60A8
   End
   Begin GridEX20.GridEX gridCuentasBancarias 
      Height          =   1695
      Left            =   5880
      TabIndex        =   4
      Top             =   12000
      Visible         =   0   'False
      Width           =   6345
      _ExtentX        =   11192
      _ExtentY        =   2990
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
      Column(1)       =   "frmAdminPagosCrearOrdenPago.frx":6280
      Column(2)       =   "frmAdminPagosCrearOrdenPago.frx":63A4
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminPagosCrearOrdenPago.frx":6498
      FormatStyle(2)  =   "frmAdminPagosCrearOrdenPago.frx":65D0
      FormatStyle(3)  =   "frmAdminPagosCrearOrdenPago.frx":6680
      FormatStyle(4)  =   "frmAdminPagosCrearOrdenPago.frx":6734
      FormatStyle(5)  =   "frmAdminPagosCrearOrdenPago.frx":680C
      FormatStyle(6)  =   "frmAdminPagosCrearOrdenPago.frx":68C4
      ImageCount      =   0
      PrinterProperties=   "frmAdminPagosCrearOrdenPago.frx":69A4
   End
   Begin GridEX20.GridEX gridMonedas 
      Height          =   1815
      Left            =   1680
      TabIndex        =   5
      Top             =   12120
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
      Column(1)       =   "frmAdminPagosCrearOrdenPago.frx":6B7C
      Column(2)       =   "frmAdminPagosCrearOrdenPago.frx":6CA0
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminPagosCrearOrdenPago.frx":6D94
      FormatStyle(2)  =   "frmAdminPagosCrearOrdenPago.frx":6ECC
      FormatStyle(3)  =   "frmAdminPagosCrearOrdenPago.frx":6F7C
      FormatStyle(4)  =   "frmAdminPagosCrearOrdenPago.frx":7030
      FormatStyle(5)  =   "frmAdminPagosCrearOrdenPago.frx":7108
      FormatStyle(6)  =   "frmAdminPagosCrearOrdenPago.frx":71C0
      ImageCount      =   0
      PrinterProperties=   "frmAdminPagosCrearOrdenPago.frx":72A0
   End
   Begin GridEX20.GridEX gridCajas 
      Height          =   1695
      Left            =   240
      TabIndex        =   6
      Top             =   12120
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   2990
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
      Column(1)       =   "frmAdminPagosCrearOrdenPago.frx":7478
      Column(2)       =   "frmAdminPagosCrearOrdenPago.frx":7578
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminPagosCrearOrdenPago.frx":7664
      FormatStyle(2)  =   "frmAdminPagosCrearOrdenPago.frx":779C
      FormatStyle(3)  =   "frmAdminPagosCrearOrdenPago.frx":784C
      FormatStyle(4)  =   "frmAdminPagosCrearOrdenPago.frx":7900
      FormatStyle(5)  =   "frmAdminPagosCrearOrdenPago.frx":79D8
      FormatStyle(6)  =   "frmAdminPagosCrearOrdenPago.frx":7A90
      ImageCount      =   0
      PrinterProperties=   "frmAdminPagosCrearOrdenPago.frx":7B70
   End
   Begin GridEX20.GridEX gridChequesDisponibles 
      Height          =   1905
      Left            =   3120
      TabIndex        =   7
      Top             =   12120
      Visible         =   0   'False
      Width           =   8235
      _ExtentX        =   14526
      _ExtentY        =   3360
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
      ColumnsCount    =   8
      Column(1)       =   "frmAdminPagosCrearOrdenPago.frx":7D48
      Column(2)       =   "frmAdminPagosCrearOrdenPago.frx":7EC8
      Column(3)       =   "frmAdminPagosCrearOrdenPago.frx":8068
      Column(4)       =   "frmAdminPagosCrearOrdenPago.frx":8160
      Column(5)       =   "frmAdminPagosCrearOrdenPago.frx":829C
      Column(6)       =   "frmAdminPagosCrearOrdenPago.frx":83A8
      Column(7)       =   "frmAdminPagosCrearOrdenPago.frx":84C8
      Column(8)       =   "frmAdminPagosCrearOrdenPago.frx":85D4
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminPagosCrearOrdenPago.frx":86C8
      FormatStyle(2)  =   "frmAdminPagosCrearOrdenPago.frx":8800
      FormatStyle(3)  =   "frmAdminPagosCrearOrdenPago.frx":88B0
      FormatStyle(4)  =   "frmAdminPagosCrearOrdenPago.frx":8964
      FormatStyle(5)  =   "frmAdminPagosCrearOrdenPago.frx":8A3C
      FormatStyle(6)  =   "frmAdminPagosCrearOrdenPago.frx":8AF4
      ImageCount      =   0
      PrinterProperties=   "frmAdminPagosCrearOrdenPago.frx":8BD4
   End
   Begin GridEX20.GridEX gridChequeras 
      Height          =   1815
      Left            =   9000
      TabIndex        =   11
      Top             =   12000
      Visible         =   0   'False
      Width           =   8235
      _ExtentX        =   14526
      _ExtentY        =   3201
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
      Column(1)       =   "frmAdminPagosCrearOrdenPago.frx":8DAC
      Column(2)       =   "frmAdminPagosCrearOrdenPago.frx":8ECC
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminPagosCrearOrdenPago.frx":8FB0
      FormatStyle(2)  =   "frmAdminPagosCrearOrdenPago.frx":90E8
      FormatStyle(3)  =   "frmAdminPagosCrearOrdenPago.frx":9198
      FormatStyle(4)  =   "frmAdminPagosCrearOrdenPago.frx":924C
      FormatStyle(5)  =   "frmAdminPagosCrearOrdenPago.frx":9324
      FormatStyle(6)  =   "frmAdminPagosCrearOrdenPago.frx":93DC
      ImageCount      =   0
      PrinterProperties=   "frmAdminPagosCrearOrdenPago.frx":94BC
   End
   Begin GridEX20.GridEX gridChequesChequera 
      Height          =   1710
      Left            =   10560
      TabIndex        =   12
      Top             =   12120
      Visible         =   0   'False
      Width           =   3420
      _ExtentX        =   6033
      _ExtentY        =   3016
      Version         =   "2.0"
      HoldSortSettings=   -1  'True
      BoundColumnIndex=   "id"
      ReplaceColumnIndex=   "nro"
      ActAsDropDown   =   -1  'True
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
      Column(1)       =   "frmAdminPagosCrearOrdenPago.frx":9694
      Column(2)       =   "frmAdminPagosCrearOrdenPago.frx":97C4
      SortKeysCount   =   1
      SortKey(1)      =   "frmAdminPagosCrearOrdenPago.frx":98C4
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminPagosCrearOrdenPago.frx":992C
      FormatStyle(2)  =   "frmAdminPagosCrearOrdenPago.frx":9A64
      FormatStyle(3)  =   "frmAdminPagosCrearOrdenPago.frx":9B14
      FormatStyle(4)  =   "frmAdminPagosCrearOrdenPago.frx":9BC8
      FormatStyle(5)  =   "frmAdminPagosCrearOrdenPago.frx":9CA0
      FormatStyle(6)  =   "frmAdminPagosCrearOrdenPago.frx":9D58
      ImageCount      =   0
      PrinterProperties=   "frmAdminPagosCrearOrdenPago.frx":9E38
   End
   Begin XtremeSuiteControls.GroupBox grpDestino 
      Height          =   2295
      Left            =   120
      TabIndex        =   15
      Top             =   1320
      Width           =   6885
      _Version        =   786432
      _ExtentX        =   12144
      _ExtentY        =   4048
      _StockProps     =   79
      Caption         =   "Destino"
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
      Begin XtremeSuiteControls.PushButton cmdMostrarDatosProveedor 
         Height          =   345
         Left            =   3870
         TabIndex        =   26
         Top             =   480
         Width           =   1095
         _Version        =   786432
         _ExtentX        =   1931
         _ExtentY        =   617
         _StockProps     =   79
         Caption         =   "Seleccionar"
         UseVisualStyle  =   -1  'True
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   255
         Left            =   9960
         TabIndex        =   25
         Top             =   6840
         Width           =   1335
      End
      Begin XtremeSuiteControls.RadioButton radioFacturaProveedor 
         Height          =   210
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   2760
         _Version        =   786432
         _ExtentX        =   4868
         _ExtentY        =   370
         _StockProps     =   79
         Caption         =   "Seleccione Proveedor"
         Appearance      =   6
         Value           =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton radioConcepto 
         Height          =   210
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   1500
         _Version        =   786432
         _ExtentX        =   2646
         _ExtentY        =   370
         _StockProps     =   79
         Caption         =   "Cuenta Contable"
         Appearance      =   6
      End
      Begin XtremeSuiteControls.ComboBox cboProveedores 
         Height          =   315
         Left            =   120
         TabIndex        =   18
         Top             =   498
         Width           =   3690
         _Version        =   786432
         _ExtentX        =   6509
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Sorted          =   -1  'True
         Style           =   2
         Appearance      =   6
         UseVisualStyle  =   -1  'True
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.PushButton btnClearProveedor 
         Height          =   345
         Left            =   5040
         TabIndex        =   21
         Top             =   480
         Width           =   375
         _Version        =   786432
         _ExtentX        =   661
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "X"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtDetalle 
         Height          =   480
         Left            =   120
         TabIndex        =   20
         Top             =   1680
         Width           =   5295
         _Version        =   786432
         _ExtentX        =   9340
         _ExtentY        =   847
         _StockProps     =   77
         BackColor       =   -2147483643
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   6
      End
      Begin XtremeSuiteControls.ComboBox cboCuentas 
         Height          =   315
         Left            =   120
         TabIndex        =   19
         Top             =   1200
         Width           =   3690
         _Version        =   786432
         _ExtentX        =   6509
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Sorted          =   -1  'True
         Style           =   2
         Appearance      =   6
         UseVisualStyle  =   -1  'True
         Text            =   "ComboBox1"
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   2655
      Left            =   7080
      TabIndex        =   27
      Top             =   1320
      Width           =   6660
      _Version        =   786432
      _ExtentX        =   11747
      _ExtentY        =   4683
      _StockProps     =   79
      Caption         =   "Retenciones"
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
      Begin VB.TextBox txtRetenciones 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   195
         Left            =   3600
         TabIndex        =   28
         Top             =   600
         Width           =   585
      End
      Begin GridEX20.GridEX gridRetenciones 
         Height          =   1215
         Left            =   120
         TabIndex        =   29
         Top             =   960
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   2143
         Version         =   "2.0"
         AllowRowSizing  =   -1  'True
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         MethodHoldFields=   -1  'True
         ContScroll      =   -1  'True
         SelectionStyle  =   1
         AllowColumnDrag =   0   'False
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         RowHeaders      =   -1  'True
         DataMode        =   99
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   4
         Column(1)       =   "frmAdminPagosCrearOrdenPago.frx":A010
         Column(2)       =   "frmAdminPagosCrearOrdenPago.frx":A14C
         Column(3)       =   "frmAdminPagosCrearOrdenPago.frx":A24C
         Column(4)       =   "frmAdminPagosCrearOrdenPago.frx":A350
         FormatStylesCount=   8
         FormatStyle(1)  =   "frmAdminPagosCrearOrdenPago.frx":A458
         FormatStyle(2)  =   "frmAdminPagosCrearOrdenPago.frx":A580
         FormatStyle(3)  =   "frmAdminPagosCrearOrdenPago.frx":A630
         FormatStyle(4)  =   "frmAdminPagosCrearOrdenPago.frx":A6E4
         FormatStyle(5)  =   "frmAdminPagosCrearOrdenPago.frx":A7BC
         FormatStyle(6)  =   "frmAdminPagosCrearOrdenPago.frx":A874
         FormatStyle(7)  =   "frmAdminPagosCrearOrdenPago.frx":A954
         FormatStyle(8)  =   "frmAdminPagosCrearOrdenPago.frx":A9F0
         ImageCount      =   0
         PrinterProperties=   "frmAdminPagosCrearOrdenPago.frx":AA90
      End
      Begin XtremeSuiteControls.PushButton btnCargar 
         Height          =   345
         Left            =   4200
         TabIndex        =   30
         Top             =   240
         Width           =   2175
         _Version        =   786432
         _ExtentX        =   3836
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Traer Alicuotas Actuales"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnPadronAnt 
         Height          =   345
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   2175
         _Version        =   786432
         _ExtentX        =   3836
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Traer Alicuotas Anteriores"
         UseVisualStyle  =   -1  'True
      End
      Begin VB.Label lblRetenciones 
         AutoSize        =   -1  'True
         Caption         =   "Retenciones previamente aplicadas IIBB BSAS"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   32
         Top             =   600
         Width           =   3300
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox2 
      Height          =   6135
      Left            =   120
      TabIndex        =   33
      Top             =   3600
      Width           =   6885
      _Version        =   786432
      _ExtentX        =   12144
      _ExtentY        =   10821
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
      Begin XtremeSuiteControls.PushButton btnExportarPagos 
         Height          =   375
         Left            =   4860
         TabIndex        =   92
         Top             =   5640
         Width           =   1815
         _Version        =   786432
         _ExtentX        =   3201
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Exportar pagos"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnExportarCbtes 
         Height          =   375
         Left            =   120
         TabIndex        =   79
         Top             =   5640
         Width           =   2055
         _Version        =   786432
         _ExtentX        =   3625
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Exportar cbtes"
         UseVisualStyle  =   -1  'True
      End
      Begin VB.TextBox txtOtrosParcialAbonar 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1800
         TabIndex        =   59
         Top             =   1680
         Width           =   1545
      End
      Begin VB.TextBox txtOtrosParcialAbonado 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   58
         Top             =   1080
         Width           =   1545
      End
      Begin VB.TextBox txtTotalParcialAbonado 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   56
         Top             =   1080
         Width           =   1545
      End
      Begin VB.TextBox txtTotalParcialAbonar 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   54
         Top             =   1680
         Width           =   1545
      End
      Begin VB.TextBox txtParcialAbonado 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   52
         Top             =   1080
         Width           =   1425
      End
      Begin VB.TextBox txtBuscarFactura 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         TabIndex        =   35
         Top             =   480
         Width           =   5010
      End
      Begin VB.TextBox txtParcialAbonar 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         TabIndex        =   34
         Top             =   1680
         Width           =   1425
      End
      Begin XtremeSuiteControls.ListBox lstFacturas 
         Height          =   3135
         Left            =   120
         TabIndex        =   36
         Top             =   2400
         Width           =   6570
         _Version        =   786432
         _ExtentX        =   11589
         _ExtentY        =   5530
         _StockProps     =   77
         BackColor       =   -2147483643
         Appearance      =   6
         Style           =   1
      End
      Begin XtremeSuiteControls.Label lblCantidadCbtesSeleccionados 
         Height          =   255
         Left            =   4500
         TabIndex        =   63
         Top             =   2100
         Width           =   2175
         _Version        =   786432
         _ExtentX        =   3836
         _ExtentY        =   450
         _StockProps     =   79
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label lblCantidadComprobantes 
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Top             =   2100
         Width           =   2175
         _Version        =   786432
         _ExtentX        =   3836
         _ExtentY        =   450
         _StockProps     =   79
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Otros Parcial a abonar:"
         Height          =   195
         Left            =   1800
         TabIndex        =   61
         Top             =   1440
         Width           =   1665
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Otros Parcial abonado:"
         Height          =   195
         Left            =   1800
         TabIndex        =   60
         Top             =   840
         Width           =   1650
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Total Parcial abonado:"
         Height          =   195
         Left            =   3600
         TabIndex        =   57
         Top             =   840
         Width           =   1605
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Total Parcial a abonar:"
         Height          =   195
         Left            =   3600
         TabIndex        =   55
         Top             =   1440
         Width           =   1620
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "NG Parcial abonado:"
         Height          =   195
         Left            =   120
         TabIndex        =   53
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Buscar factura en la lista:"
         Height          =   195
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   1830
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "NG Parcial a abonar:"
         Height          =   195
         Left            =   120
         TabIndex        =   37
         Top             =   1440
         Width           =   1470
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox4 
      Height          =   1095
      Left            =   12960
      TabIndex        =   50
      Top             =   12480
      Width           =   4125
      _Version        =   786432
      _ExtentX        =   7276
      _ExtentY        =   1931
      _StockProps     =   79
      Caption         =   "Mostrar Compensatorios Pendientes"
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
      Begin XtremeSuiteControls.ListBox lstDeudaCompensatorios 
         Height          =   495
         Left            =   14280
         TabIndex        =   51
         Top             =   -6480
         Width           =   5250
         _Version        =   786432
         _ExtentX        =   9260
         _ExtentY        =   873
         _StockProps     =   77
         BackColor       =   -2147483643
         Appearance      =   6
         Style           =   1
      End
   End
   Begin VB.Menu emergente 
      Caption         =   "emergente"
      Visible         =   0   'False
      Begin VB.Menu mnuCrearCompensatorio 
         Caption         =   "Crear Compensatorio"
      End
   End
End
Attribute VB_Name = "frmAdminPagosCrearOrdenPago"
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
Dim colPagosACuenta As New Collection
Dim colMonedas As New Collection
Dim colDeudaCompensatorios As New Collection
Dim prov As clsProveedor
Dim Factura As clsFacturaProveedor
Private Banco As Banco
Private caja As caja
Private CuentaBancaria As CuentaBancaria
Private moneda As clsMoneda
Private alicuotaRetencion As DTORetencionAlicuota
Private CuentasBancarias As New Collection
Private retenciones As New Collection
Private Monedas As New Collection
Private Cajas As New Collection
Private bancos As New Collection
Private chequesDisponibles As New Collection
Private chequeras As New Collection
Dim compe As Compensatorio
Private OrdenPago As New OrdenPago
Private operacion As operacion
Private cheque As cheque
Private tmpChequera As chequera
Private chequesChequeraSeleccionada As New Collection
Public ReadOnly As Boolean
Dim PagoACta As clsPagoACta
Public monedaplicada As clsMonedaAplicada
Dim monedaDefault As clsMoneda
Public DetalleComprobante As clsDetalleComprobante
Public colDetalles As New Collection
Public colDetallesOP As New Collection

Private Percepcion As clsPercepcionesOrdenPago


Public Sub Cargar(ByVal op As OrdenPago)
    Dim i As Long
    Dim j As Long

    On Error GoTo err1

    formLoading = True
    formLoaded = False

    If op Is Nothing Then
        MsgBox "No se recibió una orden de pago válida.", vbExclamation
        Unload Me
        Exit Sub
    End If

    Set OrdenPago = DAOOrdenPago.FindById(op.Id)

    If OrdenPago Is Nothing Then
        MsgBox "No se pudo cargar la orden de pago.", vbCritical
        Unload Me
        Exit Sub
    End If

    Set OrdenPago.Compensatorios = DAOCompensatorios.FindByOP(OrdenPago.Id)

    Me.caption = "Orden de Pago Nro " & OrdenPago.Id

    '==========================
    ' DATOS GENERALES
    '==========================
    If Not OrdenPago.moneda Is Nothing Then
        Me.cboMonedas.ListIndex = funciones.PosIndexCbo(OrdenPago.moneda.Id, Me.cboMonedas)
    Else
        Me.cboMonedas.ListIndex = -1
    End If

    Me.dtpFecha.value = OrdenPago.FEcha
    Me.txtDifCambioNG1.Text = OrdenPago.DiferenciaCambioEnNG
    Me.txtDifCambioTOTAL1.Text = OrdenPago.DiferenciaCambioEnTOTAL
    Me.txtOtrosDescuentos.Text = OrdenPago.OtrosDescuentos

    '==========================
    ' DESTINO
    '==========================
    If OrdenPago.EsParaFacturaProveedor Then
        Me.radioFacturaProveedor.value = True
        Me.radioConcepto.value = False

        If OrdenPago.FacturasProveedor.count > 0 Then
            Dim idProv As Long
            idProv = OrdenPago.FacturasProveedor.item(1).Proveedor.Id

            Me.cboProveedores.ListIndex = funciones.PosIndexCbo(idProv, Me.cboProveedores)

            If Me.cboProveedores.ListIndex = -1 Then
                Me.cboProveedores.AddItem OrdenPago.FacturasProveedor.item(1).Proveedor.RazonSocial
                Me.cboProveedores.ItemData(Me.cboProveedores.NewIndex) = idProv
                colProveedores.Add OrdenPago.FacturasProveedor.item(1).Proveedor, CStr(idProv)
                Me.cboProveedores.ListIndex = funciones.PosIndexCbo(idProv, Me.cboProveedores)
            End If

            Set prov = colProveedores.item(CStr(idProv))

            ' Cargar listas SIN disparar cmdMostrarDatosProveedor_Click
            MostrarFacturas
            MostrarDeudaCompensatorios
            MostrarPagosACuenta

            For i = 1 To OrdenPago.FacturasProveedor.count
            
                For j = 0 To Me.lstFacturas.ListCount - 1
            
                    If Me.lstFacturas.ItemData(j) = _
                            OrdenPago.FacturasProveedor.item(i).Id Then
            
                        Me.lstFacturas.Checked(j) = True
                        InicializarFacturaSeleccionada j
                        Exit For
            
                    End If
            
                Next j
            
            Next i

            For i = 1 To OrdenPago.pagosacuenta.count
                For j = 0 To Me.ListPagosACuenta.ListCount - 1
                    If Me.ListPagosACuenta.ItemData(j) = OrdenPago.pagosacuenta.item(i).Id Then
                        Me.ListPagosACuenta.Checked(j) = True
                        Exit For
                    End If
                Next j
            Next i
        End If

        Me.txtRetenciones.Text = OrdenPago.alicuota
    Else
        Me.radioFacturaProveedor.value = False
        Me.radioConcepto.value = True

        If Not OrdenPago.CuentaContable Is Nothing Then
            Me.cboCuentas.ListIndex = funciones.PosIndexCbo(OrdenPago.CuentaContable.Id, Me.cboCuentas)
        Else
            Me.cboCuentas.ListIndex = -1
        End If

        Me.txtDetalle.Text = OrdenPago.CuentaContableDescripcion
    End If

    '==========================
    ' GRILLAS
    '==========================
    Me.gridCajaOperaciones.ItemCount = OrdenPago.OperacionesCaja.count
    Me.gridDepositosOperaciones.ItemCount = OrdenPago.operacionesBanco.count
    Me.gridCheques.ItemCount = OrdenPago.ChequesTerceros.count
    Me.gridChequesPropios.ItemCount = OrdenPago.ChequesPropios.count
    Me.gridPercepciones.ItemCount = OrdenPago.percepciones.count
    Me.gridRetenciones.ItemCount = OrdenPago.RetencionesAlicuota.count
    Me.gridCompensatorios.ItemCount = OrdenPago.Compensatorios.count

    Set alicuotas = OrdenPago.RetencionesAlicuota

    ActivarControles

    Me.gridCajaOperaciones.AllowEdit = Not ReadOnly
    Me.gridDepositosOperaciones.AllowEdit = Not ReadOnly
    Me.gridCheques.AllowEdit = Not ReadOnly
    Me.gridCheques.AllowDelete = Not ReadOnly
    Me.gridChequesPropios.AllowEdit = Not ReadOnly
    Me.gridChequesPropios.AllowDelete = Not ReadOnly

    Me.cboMonedas.Enabled = Not ReadOnly
    Me.dtpFecha.Enabled = Not ReadOnly
    Me.btnGuardar.Enabled = Not ReadOnly
    Me.txtDifCambio.Enabled = Not ReadOnly
    Me.txtOtrosDescuentos.Enabled = Not ReadOnly

    formLoaded = True
    formLoading = False

    Totalizar
    Exit Sub

err1:
    formLoading = False
    formLoaded = True
    MsgBox "Error al cargar la orden de pago: " & Err.Description, vbCritical
End Sub


Public Property Set FacturaProveedor(ByVal nValue As clsFacturaProveedor)
    Set vFacturaProveedor = nValue
End Property

Public Property Get FacturaProveedor() As clsFacturaProveedor
    Set FacturaProveedor = vFacturaProveedor
End Property


Private Sub btnBorrar_Click()


    Me.cboProveedores.ListIndex = -1
    Me.gridRetenciones.ItemCount = 0
    Me.txtRetenciones.Text = "0"

    Me.lstFacturas.Clear
    Me.ListPagosACuenta.Clear
    Me.lstDeudaCompensatorios.Clear

    Set prov = Nothing
    Set vFactElegida = Nothing
    Set vCompeElegido = Nothing

    Set colFacturas = New Collection
    Set colPagosACuenta = New Collection
    Set colDeudaCompensatorios = New Collection

    limpiarParciales

    Me.lblCantidadComprobantes.caption = _
        "Cbtes. Mostrados: 0"

    calcularOrigenes

End Sub


Private Sub ActualizarAlicuotas()

    Dim A As DTORetencionAlicuota
    Dim B As DTORetencionAlicuota
    For Each A In alicuotas

        For Each B In OrdenPago.RetencionesAlicuota
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
            If OrdenPago.estado = EstadoOrdenPago_pendiente Then
                Set alicuotas = DAORetenciones.FindAllWithAlicuotas(prov.cuit)
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

    Me.cboProveedores.ListIndex = -1
    Me.gridRetenciones.ItemCount = 0
    Me.txtRetenciones.Text = "0"

    Me.lstFacturas.Clear
    Me.ListPagosACuenta.Clear
    Me.lstDeudaCompensatorios.Clear

    Set prov = Nothing
    Set vFactElegida = Nothing
    Set vCompeElegido = Nothing

    Set colFacturas = New Collection
    Set colPagosACuenta = New Collection
    Set colDeudaCompensatorios = New Collection

    limpiarParciales

    Me.lblCantidadComprobantes.caption = _
        "Cbtes. Mostrados: 0"

    calcularOrigenes

End Sub


Private Sub btnExportarCbtes_Click()
    ExportarListBoxAExcel

End Sub


Private Sub ExportarListBoxAExcel()
    Dim xlApp As Object
    Dim xlWorkbook As Object
    Dim xlWorksheet As Object
    Dim i As Integer
    Dim datos() As String
    Dim item As String
    Dim totalAbonado As Double
    Dim totalTotal As Double
    Dim LastRow As Integer
    Dim tipoComprobante As String
    Dim valorTotal As Double
    Dim valorAbonado As Double
    
    ' Crear una nueva instancia de Excel
    Set xlApp = CreateObject("Excel.Application")
    Set xlWorkbook = xlApp.Workbooks.Add
    Set xlWorksheet = xlWorkbook.Sheets(1)
    
    ' Escribir los encabezados de las columnas en negrita
    With xlWorksheet
        .Cells(1, 1).value = "Tipo"
        .Cells(1, 2).value = "Numero"
        .Cells(1, 3).value = "Total"
        .Cells(1, 4).value = "Abonado"
        .Cells(1, 5).value = "Fecha"
        .Cells(1, 6).value = "TC"
        
        ' Poner los encabezados en negrita
        .rows(1).Font.Bold = True
    End With
    
    ' Inicializar totales
    totalAbonado = 0
    totalTotal = 0
    
    ' Recorrer los elementos del ListBox (Me.lstFacturas)
    For i = 0 To Me.lstFacturas.ListCount - 1
        ' Obtener el elemento del ListBox
        item = Me.lstFacturas.list(i)
        
        ' Eliminar los textos "Abonado :" y "TC:"
        item = Replace(item, "Abonado: ", "")
        item = Replace(item, "TC: ", "")
        
        ' Dividir el texto por el carácter "|"
        datos = Split(item, "|")
        
        ' Escribir los datos en Excel
        If UBound(datos) >= 5 Then ' Asegurarse de que hay suficientes datos
            tipoComprobante = Trim(datos(0)) ' Tipo de comprobante (primera columna)
            
            ' Procesar valorTotal
            Dim valorTextoTotal As String
            valorTextoTotal = Trim(datos(2))
            valorTextoTotal = Replace(valorTextoTotal, ".", "")
            valorTextoTotal = Replace(valorTextoTotal, ",", ".")
            valorTotal = CDbl(valorTextoTotal)
            
            ' Procesar valorAbonado
            Dim valorTextoAbonado As String
            valorTextoAbonado = Trim(datos(3))
            valorTextoAbonado = Replace(valorTextoAbonado, ".", "")
            valorTextoAbonado = Replace(valorTextoAbonado, ",", ".")
            valorAbonado = CDbl(valorTextoAbonado)
            
            ' Resto del código permanece igual...
            If UCase$(Left$(Trim$(tipoComprobante), 2)) = "NC" Then
                valorTotal = valorTotal * -1
                valorAbonado = valorAbonado * -1
            End If
            
            ' Escribir los datos en Excel
            xlWorksheet.Cells(i + 2, 1).value = tipoComprobante ' Tipo
            xlWorksheet.Cells(i + 2, 2).value = Trim(datos(1)) ' Numero
            xlWorksheet.Cells(i + 2, 3).value = valorTotal ' Total (puede ser negativo si es NC)
            xlWorksheet.Cells(i + 2, 4).value = valorAbonado ' Abonado (puede ser negativo si es NC)
            xlWorksheet.Cells(i + 2, 5).value = Trim(datos(4)) ' Fecha
            xlWorksheet.Cells(i + 2, 6).value = Trim(datos(5)) ' TC
            
            ' Sumar las columnas 3 (Total) y 4 (Abonado)
            totalTotal = totalTotal + valorTotal
            totalAbonado = totalAbonado + valorAbonado
        End If
    Next i
    
    ' Calcular la última fila con datos
    LastRow = Me.lstFacturas.ListCount + 2
    
    ' Escribir los totales en la última fila
    With xlWorksheet
        .Cells(LastRow, 1).value = "Totales"
        .Cells(LastRow, 3).value = totalTotal
        .Cells(LastRow, 4).value = totalAbonado
        
        ' Poner los totales en negrita
        .rows(LastRow).Font.Bold = True
    End With
    
    ' Ajustar el ancho de las columnas en Excel
    xlWorksheet.Columns("A:F").AutoFit
    
    ' Mostrar Excel
    xlApp.Visible = True
    
    ' Liberar objetos
    Set xlWorksheet = Nothing
    Set xlWorkbook = Nothing
    Set xlApp = Nothing
End Sub


Private Sub btnExportarDatos_Click()
    Dim i As Long
    
    Dim OrdenPago As New OrdenPago
    
    For i = 0 To Me.lstFacturas.ListCount - 1
        If Me.lstFacturas.Checked(i) Then
            OrdenPago.FacturasProveedor.Add colFacturas.item(CStr(Me.lstFacturas.ItemData(i)))
        End If
    Next i
        
    If IsSomething(OrdenPago) Then
        If Not DAOOrdenPago.ExportarOrdenPago(OrdenPago) Then GoTo err1
    End If

    Exit Sub
err1:
    MsgBox "Se produjo un error al exportar!", vbCritical, "Error"
    
End Sub


Private Sub btnExportarPagos_Click()

    ExportarPagosComprobanteAExcel

End Sub


Private Sub ExportarPagosComprobanteAExcel()

    On Error GoTo ControlarError

    Dim xlApp As Object
    Dim xlLibro As Object
    Dim xlHoja As Object

    Dim detalle As clsDetalleComprobante
    Dim fila As Long
    Dim totalPagado As Double
    Dim importe As Double

    '==================================================
    ' VALIDACIONES
    '==================================================
    If vFactElegida Is Nothing Then

        MsgBox "Primero seleccione un comprobante de la lista.", _
               vbExclamation, _
               "Exportar pagos"

        Exit Sub

    End If

    If colDetalles Is Nothing Then

        MsgBox "El comprobante seleccionado no tiene pagos efectuados.", _
               vbInformation, _
               "Exportar pagos"

        Exit Sub

    End If

    If colDetalles.count = 0 Then

        MsgBox "El comprobante seleccionado no tiene pagos efectuados.", _
               vbInformation, _
               "Exportar pagos"

        Exit Sub

    End If

    '==================================================
    ' CREAR EXCEL
    '==================================================
    Set xlApp = CreateObject("Excel.Application")
    Set xlLibro = xlApp.Workbooks.Add
    Set xlHoja = xlLibro.Worksheets(1)

    xlHoja.Name = "Pagos efectuados"

    '==================================================
    ' DATOS DEL COMPROBANTE
    '==================================================
    With xlHoja

        .Cells(1, 1).value = "PAGOS EFECTUADOS DEL COMPROBANTE"
        .Range("A1:C1").Merge
        .Range("A1:C1").Font.Bold = True
        .Range("A1:C1").Font.Size = 14

        .Cells(2, 1).value = "Comprobante:"
        .Cells(2, 2).value = vFactElegida.NumeroFormateado

        .Cells(3, 1).value = "Proveedor:"

        If Not vFactElegida.Proveedor Is Nothing Then
            .Cells(3, 2).value = _
                vFactElegida.Proveedor.RazonSocial
        End If

        .Cells(4, 1).value = "Fecha del comprobante:"
        .Cells(4, 2).value = vFactElegida.FEcha
        .Cells(4, 2).NumberFormat = "dd/mm/yyyy"

        'Encabezados
        .Cells(6, 1).value = "Importe abonado"
        .Cells(6, 2).value = "Fecha"
        .Cells(6, 3).value = "Orden de pago"

        .Range("A6:C6").Font.Bold = True


    End With

    '==================================================
    ' EXPORTAR PAGOS
    '==================================================
    fila = 7
    totalPagado = 0

    For Each detalle In colDetalles

        importe = detalle.NetoGravado + detalle.Otros

        xlHoja.Cells(fila, 1).value = importe
        xlHoja.Cells(fila, 2).value = detalle.FechaEmision
        xlHoja.Cells(fila, 3).value = detalle.IdOrdenPago

        totalPagado = totalPagado + importe
        fila = fila + 1

    Next detalle

    '==================================================
    ' TOTAL
    '==================================================
    xlHoja.Cells(fila + 1, 1).value = "TOTAL PAGADO"
    xlHoja.Cells(fila + 1, 2).value = totalPagado

    xlHoja.Range( _
        xlHoja.Cells(fila + 1, 1), _
        xlHoja.Cells(fila + 1, 2) _
    ).Font.Bold = True

    '==================================================
    ' FORMATO
    '==================================================
    xlHoja.Range("A7:A" & CStr(fila - 1)).NumberFormat = _
        "#,##0.00"

    xlHoja.Cells(fila + 1, 2).NumberFormat = "#,##0.00"

    xlHoja.Range("B7:B" & CStr(fila - 1)).NumberFormat = _
        "dd/mm/yyyy"

    xlHoja.Columns("A:C").AutoFit
    xlHoja.Range("A6:C" & CStr(fila - 1)).Borders.LineStyle = 1

    xlApp.Visible = True

    Set xlHoja = Nothing
    Set xlLibro = Nothing
    Set xlApp = Nothing

    Exit Sub

ControlarError:

    MsgBox "No se pudieron exportar los pagos." & _
           vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, _
           "Exportar pagos"

    Set xlHoja = Nothing
    Set xlLibro = Nothing
    Set xlApp = Nothing

End Sub


Private Sub btnGuardar_Click()
    If Me.gridChequesPropios.EditMode = jgexEditModeOn Then
        MsgBox "Todavia esta editando la tabla de cheques propios.", vbExclamation
        Exit Sub
    End If

    If Me.gridCheques.EditMode = jgexEditModeOn Then
        MsgBox "Todavia esta editando la tabla de cheques de 3ros.", vbExclamation
        Exit Sub
    End If

    If Me.gridCajaOperaciones.EditMode = jgexEditModeOn Then
        MsgBox "Todavia esta editando la tabla de caja.", vbExclamation
        Exit Sub
    End If

    If Me.gridDepositosOperaciones.EditMode = jgexEditModeOn Then
        MsgBox "Todavia esta editando la tabla de banco.", vbExclamation
        Exit Sub
    End If
    
    If Me.gridPercepciones.EditMode = jgexEditModeOn Then
        MsgBox "Todavia esta editando la tabla de percepciones.", vbExclamation
        Exit Sub
    End If

    Set OrdenPago.CuentaContable = Nothing
    OrdenPago.CuentaContableDescripcion = vbNullString
    Set OrdenPago.FacturasProveedor = New Collection
    Set OrdenPago.RetencionesAlicuota = alicuotas

    If Me.radioFacturaProveedor.value Then
        Dim i As Long
        For i = 0 To Me.lstFacturas.ListCount - 1
            If Me.lstFacturas.Checked(i) Then
                OrdenPago.FacturasProveedor.Add colFacturas.item(CStr(Me.lstFacturas.ItemData(i)))
            End If
        Next i
    Else
        If Me.cboCuentas.ListIndex > -1 Then
            Set OrdenPago.CuentaContable = DAOCuentaContable.GetById(Me.cboCuentas.ItemData(Me.cboCuentas.ListIndex))
        End If
        OrdenPago.CuentaContableDescripcion = Me.txtDetalle.Text

    End If

    For i = 0 To Me.lstDeudaCompensatorios.ListCount - 1
        If Me.lstDeudaCompensatorios.Checked(i) Then
            OrdenPago.DeudaCompensatorios.Add colDeudaCompensatorios.item(CStr(Me.lstDeudaCompensatorios.ItemData(i)))
        End If
    Next i

    Set OrdenPago.pagosacuenta = New Collection
    
    For i = 0 To Me.ListPagosACuenta.ListCount - 1
        If Me.ListPagosACuenta.Checked(i) Then
            OrdenPago.pagosacuenta.Add colPagosACuenta.item(CStr(Me.ListPagosACuenta.ItemData(i)))
            End If
    Next i

    If IsNumeric(Me.txtRetenciones) Then OrdenPago.alicuota = val(Me.txtRetenciones)


    If OrdenPago.IsValid Then

        Dim n As Boolean: n = (OrdenPago.Id = 0)

        If DAOOrdenPago.Save(OrdenPago, True) Then

            'Me.btnGuardar.Enabled = False

            If n Then
                MsgBox "Orden de pago Nro " & OrdenPago.Id & " creada con éxito.", vbInformation
            Else

                MsgBox "Orden de pago modificada con exito.", vbInformation
            End If

            Dim EVENTO As New clsEventoObserver
            Set EVENTO.Elemento = OrdenPago
            EVENTO.Tipo = OrdenesPago_
            Set EVENTO.Originador = Me

            If n Then
                EVENTO.EVENTO = agregar_
            Else
                EVENTO.EVENTO = modificar_
            End If
            Channel.Notificar EVENTO, OrdenesPago_

            If n Then
                If MsgBox("Desea crear una nueva orden de pago?", vbQuestion + vbYesNo) = vbYes Then
                    Dim f12 As New frmAdminPagosCrearOrdenPago
                    f12.Show
                End If
            End If

            Unload Me
        Else
            MsgBox "Hubo un problema al guardar la orden de pago.", vbCritical
        End If
    Else
        MsgBox OrdenPago.ValidationMessages, vbCritical, "Error"
    End If

End Sub


Private Sub btnPadronAnt_Click()
    If Me.cboProveedores.ListIndex <> -1 Then
        Set prov = colProveedores.item(CStr(Me.cboProveedores.ItemData(Me.cboProveedores.ListIndex)))

        If IsSomething(prov) Then
            Set alicuotas = DAORetenciones.FindAllWithAlicuotasAnt(prov.cuit)
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


Private Sub cboMonedas_Click()
'''    If formLoading Then Exit Sub
    If Me.cboMonedas.ListIndex = -1 Then
        Set OrdenPago.moneda = Nothing
    Else
        Set OrdenPago.moneda = DAOMoneda.GetById(Me.cboMonedas.ItemData(Me.cboMonedas.ListIndex))
    End If
    Totalizar
End Sub


Private Sub cboProveedores_Click()
    If formLoading Then Exit Sub
    Me.gridRetenciones.ItemCount = 0

    Me.txtBuscarFactura = ""
    Me.txtParcialAbonar = ""
    
    Me.ListPagosACuenta.Clear
    Me.lstFacturas.Clear
    
    Me.lblCantidadComprobantes.caption = "Cbtes. Mostrados: 0"
    
    Me.GroupBox5.caption = "Detalle de comprobante: "
    
    Me.txtTotalParcialAbonado = ""
    Me.txtOtrosParcialAbonado = ""
    Me.txtParcialAbonado = ""
    Me.txtTotalParcialAbonar = ""
    Me.txtOtrosParcialAbonar = ""
    Me.txtParcialAbonar = ""
    
    Me.gridDetalleComprobante.ItemCount = 0
    Me.gridDetalleComprobante.Refresh
    
End Sub


Private Sub cmdMostrarDatosProveedor_Click()
    If Me.cboProveedores.ListIndex <> -1 Then
        Set prov = colProveedores.item(CStr(Me.cboProveedores.ItemData(Me.cboProveedores.ListIndex)))

        Dim d As clsDTOPadronIIBB

        Set d = DTOPadronIIBB.FindByCUIT(prov.cuit, TipoPadronRetencion)

        If IsSomething(d) Then
            Me.txtRetenciones = str(d.alicuota)   ' Val(d.Retencion )
        Else
            Me.txtRetenciones = 0
        End If
    
    Else
        Set prov = Nothing
    End If

    MostrarFacturas
    MostrarDeudaCompensatorios
    MostrarPagosACuenta
     
    btnCargar_Click

End Sub


Private Sub Command1_Click()
    If Me.cboProveedores.ListIndex <> -1 Then

        Set prov = colProveedores.item(CStr(Me.cboProveedores.ItemData(Me.cboProveedores.ListIndex)))
        If IsSomething(prov) Then

            Set alicuotas = DAORetenciones.FindAllWithAlicuotas(prov.cuit)

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
    If formLoading Then Exit Sub
    OrdenPago.FEcha = Me.dtpFecha.value

End Sub


Private Sub Exportarpagos_Click()

End Sub

Private Sub Form_Load()
    formLoading = True
    
    Me.Left = frmPrincipal.ScaleWidth / 6
    Me.Top = frmPrincipal.ScaleHeight / 22
    
    Me.gridChequeras.Visible = False
    Me.gridChequesChequera.Visible = False
    Me.gridCompensatorios.ItemCount = 0
    id_susc = funciones.CreateGUID
    Channel.AgregarSuscriptor Me, PasajeChequePropioCartera
    FormHelper.Customize Me
    
    GridEXHelper.CustomizeGrid Me.gridCajaOperaciones, False, True
    GridEXHelper.CustomizeGrid Me.gridDepositosOperaciones, False, True
    GridEXHelper.CustomizeGrid Me.gridCheques, False, True
    GridEXHelper.CustomizeGrid Me.gridChequesDisponibles, False, False
    GridEXHelper.CustomizeGrid Me.gridBancos, False, False
    GridEXHelper.CustomizeGrid Me.gridCuentasBancarias, False, False
    GridEXHelper.CustomizeGrid Me.gridMonedas, False, False
    GridEXHelper.CustomizeGrid Me.gridCajas, False, False
    GridEXHelper.CustomizeGrid Me.gridChequeras, False, False
    GridEXHelper.CustomizeGrid Me.gridChequesPropios, False, True
    GridEXHelper.CustomizeGrid Me.gridCompensatorios, False, True
    GridEXHelper.CustomizeGrid Me.gridChequesChequera
    GridEXHelper.CustomizeGrid Me.gridRetenciones, False, True
    GridEXHelper.CustomizeGrid Me.gridDetalleComprobante, False, False
    GridEXHelper.CustomizeGrid Me.gridPercepciones, False, True
    
    Set Cajas = DAOCaja.FindAll()
    Me.gridCajas.ItemCount = Cajas.count

    Set Monedas = DAOMoneda.GetAll()
    Me.gridMonedas.ItemCount = Monedas.count

    Set CuentasBancarias = DAOCuentaBancaria.FindAll()
    Me.gridCuentasBancarias.ItemCount = CuentasBancarias.count

    Set bancos = DAOBancos.GetAll()
    Me.gridBancos.ItemCount = bancos.count

    Set chequeras = DAOChequeras.FindAllWithChequesDisponibles()
    Me.gridChequeras.ItemCount = chequeras.count


    CargarChequesDisponibles


    Set colProveedores = DAOProveedor.FindAllProveedoresWithFacturasImpagas
    For Each prov In colProveedores
        cboProveedores.AddItem prov.RazonSocial
        cboProveedores.ItemData(cboProveedores.NewIndex) = prov.Id
    Next

    Dim cuentasContables As Collection
    Set cuentasContables = DAOCuentaContable.GetAll()
    Dim cc As clsCuentaContable
    Me.cboCuentas.Clear
    For Each cc In cuentasContables
        cboCuentas.AddItem cc.nombre & " - " & cc.codigo
        cboCuentas.ItemData(cboCuentas.NewIndex) = cc.Id
    Next cc

    radioFacturaProveedor_Click

    Me.gridCajaOperaciones.ItemCount = OrdenPago.OperacionesCaja.count
    Me.gridPercepciones.ItemCount = OrdenPago.percepciones.count
    Me.gridDepositosOperaciones.ItemCount = OrdenPago.operacionesBanco.count
    Me.gridCheques.ItemCount = OrdenPago.ChequesTerceros.count
    Me.gridChequesPropios.ItemCount = OrdenPago.ChequesPropios.count


    Set Me.gridCheques.Columns("numero").DropDownControl = Me.gridChequesDisponibles

    Set Me.gridDepositosOperaciones.Columns("moneda").DropDownControl = Me.gridMonedas
   
    Set Me.gridDepositosOperaciones.Columns("cuenta").DropDownControl = Me.gridCuentasBancarias

    Set Me.gridCajaOperaciones.Columns("monedas").DropDownControl = Me.gridMonedas
    
    Set Me.gridPercepciones.Columns("moneda").DropDownControl = Me.gridMonedas
    
    Set Me.gridCajaOperaciones.Columns("caja").DropDownControl = Me.gridCajas

    Set Me.gridChequesPropios.Columns("chequera").DropDownControl = Me.gridChequeras
    
    Set Me.gridChequesPropios.Columns("numero").DropDownControl = Me.gridChequesChequera
    
'''    cargarCamposPredefinidos

    gridChequesChequera.ItemCount = 0
    
    GridEXHelper.AutoSizeColumns Me.gridChequeras

    DAOMoneda.llenarComboXtremeSuite Me.cboMonedas

    Me.dtpFecha.value = OrdenPago.FEcha
    
    If OrdenPago.estado = EstadoOrdenPago_pendiente Then
        btnExportarDatos.Enabled = True
    Else
            btnExportarDatos.Enabled = False
    End If

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

        For Each c In colDeudaCompensatorios
            Me.lstDeudaCompensatorios.AddItem "Cód: " & c.Id & " (OP: " & c.IdOrdenPago & ", Cbte: " & c.Comprobante.NumeroFormateado & ", Importe: " & c.Monto & ")"
            Me.lstDeudaCompensatorios.ItemData(Me.lstDeudaCompensatorios.NewIndex) = c.Id
        Next
    Else
        Set colDeudaCompensatorios = New Collection
    End If
    
End Sub


Private Sub MostrarFacturas()
    Me.lstFacturas.Clear

    If IsSomething(prov) Then
        Set colFacturas = DAOFacturaProveedor.FindAll("AdminComprasFacturasProveedores.id_proveedor=" & prov.Id & " and (AdminComprasFacturasProveedores.estado=" & EstadoFacturaProveedor.Aprobada & " or AdminComprasFacturasProveedores.estado=" & EstadoFacturaProveedor.pagoParcial & ")", False, "", False, True)

        If OrdenPago.Id <> 0 And OrdenPago.EsParaFacturaProveedor Then
            If prov.Id = OrdenPago.FacturasProveedor.item(1).Proveedor.Id Then
                For Each Factura In OrdenPago.FacturasProveedor
                    If Not funciones.BuscarEnColeccion(colFacturas, CStr(Factura.Id)) Then
                        colFacturas.Add DAOFacturaProveedor.FindById(Factura.Id), CStr(Factura.Id)
                    End If
                Next
            End If
        End If

        Dim T As String

        For Each Factura In colFacturas    'en ese for traigo los pendientes a abonar que estan asociados a ops sin aprobar

            Dim c As Collection
            Set c = DAOOrdenPago.FindAbonadoPendiente(Factura.Id, OrdenPago.Id)

            Factura.TotalAbonadoGlobalPendiente = 0    ' c(1) 'que esta en ops sin aprobar
            Factura.NetoGravadoAbonadoGlobalPendiente = 0    ' c(2)
            Factura.OtrosAbonadoGlobalPendiente = 0    'c(3)

'''                T = Factura.NumeroFormateado & " (" & Factura.moneda.NombreCorto & " " & Factura.total & ")" & " (" & Factura.FEcha & ")"    'TipoCambio: (" & Factura.TipoCambioPago & ")"
'''            If Factura.TotalAbonadoGlobal + Factura.TotalAbonadoGlobalPendiente > 0 Then
'''                T = Factura.NumeroFormateado & " (" & Factura.moneda.NombreCorto & " " & Factura.total & " - Abonado: " & Factura.TotalAbonadoGlobal + Factura.TotalAbonadoGlobalPendiente & ")" & " (" & Factura.FEcha & ")"

                T = Factura.NumeroFormateadoCorto & " | " & Factura.numero & " | " & Replace(FormatCurrency(funciones.FormatearDecimales(Factura.total)), "$", "") & " | Abonado: 0  " & " | " & Factura.FEcha & " | TC: " & Factura.TipoCambioPago & " | "
            If Factura.TotalAbonadoGlobal + Factura.TotalAbonadoGlobalPendiente > 0 Then
                T = Factura.NumeroFormateadoCorto & " | " & Factura.numero & " | " & Replace(FormatCurrency(funciones.FormatearDecimales(Factura.total)), "$", "") & " | Abonado: " & Replace(FormatCurrency(funciones.FormatearDecimales(Factura.TotalAbonadoGlobal + Factura.TotalAbonadoGlobalPendiente)), "$", "") & " | " & Factura.FEcha & " | TC: " & Factura.TipoCambioPago & " | "

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


Private Sub MostrarPagosACuenta()
    Me.ListPagosACuenta.Clear
    Set colPagosACuenta = New Collection

    If Not IsSomething(prov) Then Exit Sub

    Dim filtro As String

    If OrdenPago.Id <> 0 And (ReadOnly Or OrdenPago.estado <> EstadoOrdenPago_pendiente) Then

        ' OP aprobada / cerrada / solo lectura:
        ' mostrar solamente los pagos a cuenta usados en esta OP
        filtro = "pagos_a_cuenta.id_proveedor = " & prov.Id & _
                 " AND pagos_a_cuenta.id IN (" & _
                 " SELECT opac.id_pago_a_cuenta " & _
                 " FROM ordenes_pago_pagos_a_cuenta opac " & _
                 " WHERE opac.id_orden_pago = " & OrdenPago.Id & _
                 " )"

    ElseIf OrdenPago.Id <> 0 Then

        ' OP pendiente existente:
        ' mostrar los disponibles + los que ya estaban usados en esta misma OP
        filtro = "pagos_a_cuenta.id_proveedor = " & prov.Id & _
                 " AND (" & _
                 " pagos_a_cuenta.estado = 0 " & _
                 " OR pagos_a_cuenta.id IN (" & _
                 "     SELECT opac.id_pago_a_cuenta " & _
                 "     FROM ordenes_pago_pagos_a_cuenta opac " & _
                 "     WHERE opac.id_orden_pago = " & OrdenPago.Id & _
                 " )" & _
                 " )"

    Else

        ' OP nueva:
        ' mostrar solamente pagos a cuenta disponibles
        filtro = "pagos_a_cuenta.estado = 0 AND pagos_a_cuenta.id_proveedor = " & prov.Id

    End If

    Set colPagosACuenta = DAOPagoACta.FindAll(filtro)

    Dim T As String

    For Each PagoACta In colPagosACuenta

        T = "N°: " & PagoACta.Id & _
            " ( " & PagoACta.moneda.NombreCorto & " " & _
            Replace(FormatCurrency(funciones.FormatearDecimales(PagoACta.StaticTotalOrigenes)), "$", "") & ")"

        Me.ListPagosACuenta.AddItem T
        Me.ListPagosACuenta.ItemData(Me.ListPagosACuenta.NewIndex) = PagoACta.Id

    Next

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
        Values(3) = cheque.FechaVencimiento
        If IsSomething(cheque.moneda) Then Values(4) = cheque.moneda.NombreCorto
        If IsSomething(cheque.Banco) Then Values(5) = cheque.Banco.nombre
        Values(6) = cheque.Id
        Values(7) = cheque.OrigenCheque
        Values(8) = cheque.OrigenDestino
    End If
End Sub


Private Sub gridChequesPropios_BeforeUpdate(ByVal Cancel As GridEX20.JSRetBoolean)

    Dim msg As New Collection

    Dim valorChequera As String
    Dim idChequeSeleccionado As String
    Dim numeroChequeCargado As String
    Dim fechaTexto As String

    '==================================================
    ' Obtener los valores evitando errores por Null
    '==================================================
    If Not IsNull(Me.gridChequesPropios.value(1)) Then
        valorChequera = Trim$(CStr(Me.gridChequesPropios.value(1)))
    End If

    If Not IsNull(Me.gridChequesPropios.value(2)) Then
        idChequeSeleccionado = Trim$(CStr(Me.gridChequesPropios.value(2)))
    End If

    If Not IsNull(Me.gridChequesPropios.value(5)) Then
        numeroChequeCargado = Trim$(CStr(Me.gridChequesPropios.value(5)))
    End If

    If Not IsNull(Me.gridChequesPropios.value(4)) Then
        fechaTexto = Trim$(CStr(Me.gridChequesPropios.value(4)))
    End If

    '==================================================
    ' Chequera
    '==================================================
    If LenB(valorChequera) = 0 Then
        msg.Add "Debe especificar una chequera."
    End If

    '==================================================
    ' Cheque
    '
    ' En una fila nueva, el ID se encuentra en Value(2).
    ' En una fila existente, Value(2) queda vacío pero
    ' el número se muestra en Value(5).
    '==================================================
    If LenB(idChequeSeleccionado) = 0 And _
       LenB(numeroChequeCargado) = 0 Then

        msg.Add "Debe especificar un cheque."

    End If

    ' Revisar duplicado solamente cuando se está
    ' seleccionando un cheque nuevo.
    If LenB(idChequeSeleccionado) > 0 Then

        If IsNumeric(idChequeSeleccionado) Then

            If funciones.BuscarEnColeccion( _
                    OrdenPago.ChequesPropios, _
                    CStr(CLng(idChequeSeleccionado))) Then

                msg.Add "El cheque seleccionado ya fue ingresado anteriormente."

            End If

        End If

    End If

    '==================================================
    ' Monto
    '==================================================
    If Not EsImporteValido(Me.gridChequesPropios.value(3)) Then
        msg.Add "Debe especificar un monto válido."
    End If

    '==================================================
    ' Fecha
    '==================================================
    If Not EsFechaValida(fechaTexto) Then
        msg.Add "Debe especificar una fecha válida con formato dd/mm/aaaa."
    End If

    Cancel = (msg.count > 0)

    If Cancel Then
        MsgBox funciones.JoinCollectionValues(msg, vbNewLine), _
               vbExclamation, _
               "No se puede actualizar el cheque"
    End If

End Sub

Private Sub gridChequesPropios_ListSelected(ByVal ColIndex As Integer, ByVal ValueListIndex As Long, ByVal value As Variant)
    If ColIndex = 1 Then
        'If Not IsNumeric(Me.gridChequesPropios.Value(1)) Or LenB(Me.gridChequesPropios.Value(1)) = 0 Then
        If Not IsNumeric(value) Or LenB(value) = 0 Then
            Set chequesChequeraSeleccionada = New Collection
        Else
            Set chequesChequeraSeleccionada = DAOCheques.FindAllDisponiblesByChequera(val(value))  ' Me.gridChequesPropios.Value(1))
        End If

        Me.gridChequesChequera.ItemCount = chequesChequeraSeleccionada.count
    End If
End Sub


Private Sub gridChequesPropios_UnboundAddNew(ByVal NewRowBookmark As GridEX20.JSRetVariant, ByVal Values As GridEX20.JSRowData)
    Set cheque = Nothing
    If IsNumeric(Values(2)) Then Set cheque = DAOCheques.FindById(Values(2))
    If IsSomething(cheque) Then
        cheque.Monto = ImporteDesdeTexto(Values(3))
        cheque.FechaVencimiento = FechaDesdeTexto(CStr(Values(4)))
        
        OrdenPago.ChequesPropios.Add cheque, CStr(cheque.Id)

    End If
    Totalizar
End Sub


Private Sub gridChequesPropios_UnboundDelete(ByVal RowIndex As Long, ByVal Bookmark As Variant)
    If RowIndex > 0 Then
        OrdenPago.ChequesPropios.remove RowIndex
        Totalizar
    End If
End Sub


Private Sub gridChequesPropios_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If OrdenPago.ChequesPropios.count >= RowIndex Then
        Set cheque = OrdenPago.ChequesPropios.item(RowIndex)
        Values(1) = cheque.chequera.Description
        Values(2) = vbNullString
        'FORMATCURRENCY
        Values(3) = FormatCurrency(cheque.Monto)
        Values(4) = Format$(cheque.FechaVencimiento, "dd/mm/yyyy")
        Values(5) = cheque.numero

    End If
End Sub


Private Sub gridChequesPropios_UnboundUpdate( _
        ByVal RowIndex As Long, _
        ByVal Bookmark As Variant, _
        ByVal Values As GridEX20.JSRowData)

    Dim fechaTexto As String

    If RowIndex <= 0 Then Exit Sub
    If OrdenPago.ChequesPropios.count < RowIndex Then Exit Sub

    Set cheque = OrdenPago.ChequesPropios.item(RowIndex)

    If Not IsNull(Values(4)) Then
        fechaTexto = Trim$(CStr(Values(4)))
    End If

    If Not EsImporteValido(Values(3)) Then
        MsgBox "El monto ingresado no es válido.", vbExclamation
        Exit Sub
    End If

    If Not EsFechaValida(fechaTexto) Then
        MsgBox "La fecha debe tener formato dd/mm/aaaa.", vbExclamation
        Exit Sub
    End If

    cheque.Monto = ImporteDesdeTexto(Values(3))
    cheque.FechaVencimiento = FechaDesdeTexto(fechaTexto)

    Totalizar

    Me.gridChequesPropios.Refresh

End Sub


Private Sub gridCompensatorios_UnboundDelete(ByVal RowIndex As Long, ByVal Bookmark As Variant)
    OrdenPago.Compensatorios.remove (RowIndex)
End Sub


Private Sub gridCompensatorios_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    On Error Resume Next
    Set compe = OrdenPago.Compensatorios.item(RowIndex)
    Values(1) = compe.Comprobante.NumeroFormateado
    Values(2) = TiposCompensatorio.item(CStr(compe.Tipo))
    'FORMATCURRENCY
    Values(3) = FormatCurrency(compe.Monto)
    Values(4) = compe.FechaCancelacion
    Values(5) = compe.Observacion

End Sub


Private Sub gridCuentasBancarias_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If CuentasBancarias.count >= RowIndex Then
        Set CuentaBancaria = CuentasBancarias.item(RowIndex)
        Values(1) = CuentaBancaria.Id
        Values(2) = CuentaBancaria.DescripcionFormateada
    End If
End Sub



Private Sub gridDetalleComprobante_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If colDetalles.count >= RowIndex Then
        Set DetalleComprobante = colDetalles.item(RowIndex)
        Values(1) = Replace(FormatCurrency(funciones.FormatearDecimales(DetalleComprobante.NetoGravado + DetalleComprobante.Otros)), "$", "")
        Values(2) = DetalleComprobante.FechaEmision
        Values(3) = DetalleComprobante.IdOrdenPago
        
    End If
End Sub

Private Sub gridMonedas_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex > 0 And Monedas.count > 0 Then
        Set moneda = Monedas.item(RowIndex)
        Values(1) = moneda.Id
        Values(2) = moneda.NombreCorto
    End If
End Sub


Private Sub gridPercepciones_BeforeUpdate(ByVal Cancel As GridEX20.JSRetBoolean)
    Dim cond1 As Boolean
    Dim cond2 As Boolean
    Dim cond3 As Boolean
    Dim cond4 As Boolean

    cond1 = Not IsNumeric(Me.gridPercepciones.value(1))
    cond2 = Not IsNumeric(Me.gridPercepciones.value(2)) And LenB(Me.gridPercepciones.value(2)) = 0
    cond3 = Not IsDate(Me.gridPercepciones.value(3))
    cond4 = LenB(Me.gridPercepciones.value(4)) = 0 Or IsEmpty(Me.gridPercepciones.value(4))
    Cancel = cond1 Or cond2 Or cond3 Or cond4
    
End Sub


Private Sub gridPercepciones_UnboundAddNew(ByVal NewRowBookmark As GridEX20.JSRetVariant, ByVal Values As GridEX20.JSRowData)
    
    Set Percepcion = New clsPercepcionesOrdenPago
    
    Percepcion.Monto = Values(1)
    
    If IsNumeric(Values(2)) Then
        Set Percepcion.moneda = DAOMoneda.GetById(Values(2))
    End If
    
    Percepcion.FEcha = Values(3)

    Percepcion.Comprobante = Values(4)
    
    Percepcion.Tipo = Values(5)
    
    OrdenPago.percepciones.Add Percepcion
    
    Totalizar

End Sub


Private Sub gridPercepciones_UnboundDelete(ByVal RowIndex As Long, ByVal Bookmark As Variant)
    
    If RowIndex > 0 And OrdenPago.percepciones.count >= RowIndex Then
        OrdenPago.percepciones.remove RowIndex
        
        Totalizar
    End If
    
End Sub

Private Sub gridPercepciones_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    
    If RowIndex <= OrdenPago.percepciones.count Then
        Set Percepcion = OrdenPago.percepciones.item(RowIndex)
        'FORMATCURRENCY
        Values(1) = FormatCurrency(funciones.FormatearDecimales(Percepcion.Monto))
        If IsSomething(Percepcion.moneda) Then
            Values(2) = Percepcion.moneda.NombreCorto
        End If
        Values(3) = Percepcion.FEcha
        
        If IsSomething(Percepcion) Then
            Values(4) = Percepcion.Comprobante
        End If
        If IsSomething(Percepcion) Then
            Values(5) = Percepcion.Tipo
        End If
    End If
        
End Sub

Private Sub gridPercepciones_UnboundUpdate(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex > 0 And OrdenPago.percepciones.count > 0 Then
        Set Percepcion = OrdenPago.percepciones.item(RowIndex)
        
        Percepcion.Monto = Values(1)

        If IsNumeric(Values(2)) Then
            Set Percepcion.moneda = DAOMoneda.GetById(Values(2))
        End If
        
        Percepcion.FEcha = Values(3)
        Percepcion.Comprobante = Values(4)
        Percepcion.Tipo = Values(5)
        
        Totalizar
    End If
End Sub

Private Sub gridRetenciones_RowFormat(RowBuffer As GridEX20.JSRowData)

    On Error GoTo err1

    Set alicuotaRetencion = alicuotas.item(RowBuffer.RowIndex)

    If alicuotaRetencion.importe > 0 Then    '.Retencion.id <> 2 Then
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
        Values(4) = alicuotaRetencion.certificados
    End If
End Sub


Private Sub gridRetenciones_UnboundUpdate(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If alicuotas.count >= RowIndex Then
        Set alicuotaRetencion = alicuotas.item(RowIndex)
        alicuotaRetencion.alicuotaRetencion = Values(2)
        If Not IsNumeric(Values(3)) Then
            alicuotaRetencion.importe = 0
            alicuotaRetencion.certificados = "-"
        Else
            alicuotaRetencion.importe = Values(3)
            alicuotaRetencion.certificados = Values(4)
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


Private Sub MostrarPosiblesRetenciones(col As Collection, Optional colc As Collection = Nothing, Optional colpcta As Collection = Nothing, Optional colPercepciones As Collection = Nothing)
    Dim d As New Dictionary
    
    Dim ret As Retencion
    Dim colret As Collection
    Set colret = DAORetenciones.FindAllEsAgente
    Set d = DAOCertificadoRetencion.VerPosibleRetenciones2(col, alicuotas, val(Me.txtDifCambioNG1), OrdenPago.TotalNGCompensatorios)
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
    Dim P As clsPagoACta
    Dim totFact As Double
    Dim TotNG As Double
    Dim totFactHoy As Double
    Dim Cambio As Double
    Dim totCambio As Double
    Dim totCambiong As Double
    Dim totNGHoy As Double
    Dim totDeudaCompe As Double
    Dim totPagoACuenta As Double
    Dim totFactNuevo As Double
    
    Dim totPercepciones As Double
    
    totDeudaCompe = 0
    totFactNuevo = 0
    totPercepciones = 0
    
    Dim totalComprobantes As Double
    totalComprobantes = 0
         
    For Each F In col
        
    ' Inicializar la variable
    
    ' Recorrer la lista de facturas

        ' Verificar si es Nota de Crédito
       
        Dim importeComprobante As Double
        
        importeComprobante = 0
        
        If OrdenPago.estado = EstadoOrdenPago_pendiente Then
        
            importeComprobante = F.totalAbonado
        
        ElseIf OrdenPago.estado = EstadoOrdenPago_Aprobada Then
        
            Set colDetallesOP = _
                DAOOrdenPago.FindAllDetallesAbonadoOP(F.Id, OrdenPago.Id)
        
            For Each DetalleComprobante In colDetallesOP
        
                importeComprobante = importeComprobante + _
                                     DetalleComprobante.Otros + _
                                     DetalleComprobante.NetoGravado
        
            Next DetalleComprobante
        
        End If
        
        If F.tipoDocumentoContable = _
                tipoDocumentoContable.notaCredito Then
        
            totalComprobantes = totalComprobantes - importeComprobante
        
        Else
        
            totalComprobantes = totalComprobantes + importeComprobante
        
        End If
    
        totFact = totFact + MonedaConverter.ConvertirForzado2(IIf(F.tipoDocumentoContable = tipoDocumentoContable.notaCredito, F.totalAbonado * -1, F.totalAbonado), F.moneda.Id, OrdenPago.moneda.Id, F.TipoCambioPago)
        totFactHoy = totFactHoy + MonedaConverter.ConvertirForzado2(IIf(F.tipoDocumentoContable = tipoDocumentoContable.notaCredito, F.TotalDiaPagoAbonado * -1, F.TotalDiaPagoAbonado), F.moneda.Id, OrdenPago.moneda.Id, F.TipoCambioPago)
        TotNG = TotNG + MonedaConverter.ConvertirForzado2(IIf(F.tipoDocumentoContable = tipoDocumentoContable.notaCredito, F.NetoGravadoAbonado * -1, F.NetoGravadoAbonado), F.moneda.Id, OrdenPago.moneda.Id, F.TipoCambioPago)
        totNGHoy = totNGHoy + MonedaConverter.ConvertirForzado2(IIf(F.tipoDocumentoContable = tipoDocumentoContable.notaCredito, F.NetoGravadoAbonadoDiaPago * -1, F.NetoGravadoAbonadoDiaPago), F.moneda.Id, OrdenPago.moneda.Id, F.TipoCambioPago)
        totCambio = totCambio + MonedaConverter.ConvertirForzado2(IIf(F.tipoDocumentoContable = tipoDocumentoContable.notaCredito, F.DiferenciaPorTipoDeCambionTOTAL * -1, F.DiferenciaPorTipoDeCambionTOTAL), F.moneda.Id, OrdenPago.moneda.Id, F.TipoCambioPago)
        totCambiong = totCambiong + MonedaConverter.ConvertirForzado2(IIf(F.tipoDocumentoContable = tipoDocumentoContable.notaCredito, F.DiferenciaPorTipoDeCambionNG * -1, F.DiferenciaPorTipoDeCambionNG), F.moneda.Id, OrdenPago.moneda.Id, F.TipoCambioPago)
 
        Next F
        
        Me.lblTotalFacturas.caption = _
            "Total Facturas en " & _
            FormatCurrency(funciones.FormatearDecimales(totalComprobantes))


    If IsSomething(colc) Then
        For Each c In colc

            Dim ff As clsFacturaProveedor

            Set ff = DAOFacturaProveedor.FindById(c.Comprobante.Id)
            totDeudaCompe = totDeudaCompe + MonedaConverter.ConvertirForzado2(IIf(c.Tipo = TC_Credito, c.Monto * -1, c.Monto), ff.moneda.Id, OrdenPago.moneda.Id, ff.TipoCambioPago)

        Next
    End If
    
    
   If IsSomething(colpcta) Then
        For Each P In colpcta
            totPagoACuenta = totPagoACuenta + P.StaticTotalOrigenes
        Next
    End If


   Dim per As clsPercepcionesOrdenPago

    For Each per In OrdenPago.percepciones
        If IsSomething(per) Then
            If IsSomething(per.moneda) And IsSomething(OrdenPago.moneda) Then
                totPercepciones = totPercepciones + MonedaConverter.ConvertirForzado2(per.Monto, per.moneda.Id, OrdenPago.moneda.Id, 1)
            Else
                totPercepciones = totPercepciones + per.Monto
            End If
        End If
    Next
    

        

    Me.lblNgAbonar = "Total NG a Abonar en " & FormatCurrency(funciones.FormatearDecimales(OrdenPago.DiferenciaCambioEnNG + totNGHoy))
  
    Me.lblDeudaCompensatorios = "Total deuda compensatorios en " & FormatCurrency(funciones.FormatearDecimales(totDeudaCompe))
    
    OrdenPago.StaticTotalFacturas = funciones.RedondearDecimales(totalComprobantes)
    
    OrdenPago.staticTotalDeudaCompensatorios = funciones.RedondearDecimales(totDeudaCompe)

    Me.lblTotalFacturasNG = "Total NG Facturas en " & FormatCurrency(funciones.FormatearDecimales(TotNG + OrdenPago.DiferenciaCambioEnNG))

    OrdenPago.StaticTotalFacturasNG = funciones.RedondearDecimales(TotNG + OrdenPago.DiferenciaCambioEnNG)

    Me.lblDiferenciaCambio = "Diferencia Cambio en " & FormatCurrency(totCambiong)

    OrdenPago.DiferenciaCambio = totCambio

    verCompensatorios

    Me.lblTotalARetener = "Total a retener en " & FormatCurrency(funciones.FormatearDecimales(totRet))

    OrdenPago.StaticTotalRetenido = funciones.RedondearDecimales(totRet)

    Me.lblTotalOrdenPago = "Total a abonar en " & FormatCurrency(funciones.FormatearDecimales(totalComprobantes - (totRet + totPagoACuenta + totPercepciones)))
    
    Me.lblTotalPagoACuenta.caption = "Total Pago a Cuenta en " & FormatCurrency(funciones.FormatearDecimales(totPagoACuenta))
    
    OrdenPago.StaticTotalPercepciones = totPercepciones
    
    Me.lblTotalPercepciones.caption = "Total Percepciones en " & FormatCurrency(funciones.FormatearDecimales(totPercepciones))

    
End Sub


Private Sub verCompensatorios()
    Me.lblTotalCompensatorios = "Total compensatorios en " & FormatCurrency(funciones.FormatearDecimales(OrdenPago.TotalCompensatorios))

End Sub


Private Sub MostrarPago(F As clsFacturaProveedor)

    If F Is Nothing Then Exit Sub

    Me.txtTotalParcialAbonado.Text = _
        CStr(F.TotalAbonadoGlobal)

    Me.txtOtrosParcialAbonado.Text = _
        CStr(F.OtrosAbonadoGlobal + _
             F.OtrosAbonadoGlobalPendiente)

    Me.txtParcialAbonado.Text = _
        CStr(F.NetoGravadoAbonadoGlobal + _
             F.NetoGravadoAbonadoGlobalPendiente)

    Me.txtParcialAbonar.Text = _
        CStr(F.NetoGravadoAbonado)

    Me.txtOtrosParcialAbonar.Text = _
        CStr(F.OtrosAbonado)

    Me.txtTotalParcialAbonar.Text = _
        CStr(F.totalAbonado)

    RecalcularTotalFacturaElegida

    If F.totalAbonado + _
       F.TotalAbonadoGlobal + _
       F.TotalAbonadoGlobalPendiente > F.total Then

        MsgBox "El importe que desea abonar supera el monto total del comprobante seleccionado.", _
               vbExclamation, _
               "Importe inválido"
    End If

End Sub


Private Sub MotrarHistorialPagos(F As clsFacturaProveedor)

    If IsSomething(F) Then

    Me.GroupBox5.caption = "Detalle de comprobante: " & F.tipoDocumentoContable & " " & F.NumeroFormateado & " (ID: " & F.Id & ") "
    
    Set colDetalles = DAOOrdenPago.FindAllDetallesAbonado(F.Id)
    
    Me.gridDetalleComprobante.ItemCount = 0
    
    Me.gridDetalleComprobante.ItemCount = colDetalles.count
    
    End If

End Sub


Private Sub ListPagosACuenta_ItemCheck(ByVal item As Long)
    calcularOrigenes
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

    Dim item As Long

    item = Me.lstFacturas.ListIndex

    If item < 0 Or item >= Me.lstFacturas.ListCount Then Exit Sub

    If Not funciones.BuscarEnColeccion( _
            colFacturas, _
            CStr(Me.lstFacturas.ItemData(item))) Then
        Exit Sub
    End If

    Set vFactElegida = colFacturas.item( _
                            CStr(Me.lstFacturas.ItemData(item)))

    If Me.lstFacturas.Checked(item) Then
        InicializarFacturaSeleccionada item
    End If

    MostrarPago vFactElegida
    MotrarHistorialPagos vFactElegida
    RecalcularFacturaElegida
    Totalizar

End Sub


Private Sub lstFacturas_DblClick()

    Dim item As Long
    Dim F As clsFacturaProveedor
    Dim respuesta As String
    Dim nuevoCambio As Double

    item = Me.lstFacturas.ListIndex

    If item < 0 Or item >= Me.lstFacturas.ListCount Then Exit Sub

    If Not funciones.BuscarEnColeccion( _
            colFacturas, _
            CStr(Me.lstFacturas.ItemData(item))) Then
        Exit Sub
    End If

    Set F = colFacturas.item( _
                CStr(Me.lstFacturas.ItemData(item)))

    Set vFactElegida = F

    respuesta = InputBox( _
                    "Establezca el tipo de cambio con el cual se va a abonar la factura.", _
                    "Tipo de cambio", _
                    CStr(F.TipoCambioPago))

    If LenB(Trim$(respuesta)) = 0 Then Exit Sub

    If Not IsNumeric(respuesta) Then
        MsgBox "El tipo de cambio ingresado no es válido.", _
               vbExclamation, _
               "Tipo de cambio"
        Exit Sub
    End If

    nuevoCambio = CDbl(respuesta)

    If nuevoCambio <= 0 Then
        MsgBox "El tipo de cambio debe ser mayor que cero.", _
               vbExclamation, _
               "Tipo de cambio"
        Exit Sub
    End If

    F.TipoCambioPago = nuevoCambio

    MostrarPago F
    Totalizar

End Sub


Private Sub calcularOrigenes()

    Dim i As Long
    Dim col As New Collection
    Dim colc As New Collection
    Dim colpcta As New Collection
    Dim colPercepciones As New Collection

    Dim ff As clsFacturaProveedor
    Dim c As Compensatorio

    For i = 0 To Me.lstFacturas.ListCount - 1

        If Me.lstFacturas.Checked(i) Then

            If funciones.BuscarEnColeccion( _
                    colFacturas, _
                    CStr(Me.lstFacturas.ItemData(i))) Then

                col.Add colFacturas.item( _
                            CStr(Me.lstFacturas.ItemData(i)))

            End If

        Else

            If funciones.BuscarEnColeccion( _
                    colFacturas, _
                    CStr(Me.lstFacturas.ItemData(i))) Then

                Set ff = colFacturas.item( _
                            CStr(Me.lstFacturas.ItemData(i)))

                For Each c In OrdenPago.Compensatorios

                    If Not c.Comprobante Is Nothing Then

                        If c.Comprobante.Id = ff.Id Then

                            MsgBox _
                                "Existen compensatorios para este comprobante. Elimínelos primero.", _
                                vbCritical, _
                                "Comprobante con compensatorio"

                            Me.lstFacturas.Checked(i) = True
                            col.Add ff
                            Exit For

                        End If

                    End If

                Next c

            End If

        End If

    Next i

    Me.lblCantidadCbtesSeleccionados.caption = _
        "Cbtes. Seleccionados [ " & col.count & " ]"

    For i = 0 To Me.lstDeudaCompensatorios.ListCount - 1

        If Me.lstDeudaCompensatorios.Checked(i) Then

            If funciones.BuscarEnColeccion( _
                    colDeudaCompensatorios, _
                    CStr(Me.lstDeudaCompensatorios.ItemData(i))) Then

                colc.Add colDeudaCompensatorios.item( _
                            CStr(Me.lstDeudaCompensatorios.ItemData(i)))

            End If

        End If

    Next i

    For i = 0 To Me.ListPagosACuenta.ListCount - 1

        If Me.ListPagosACuenta.Checked(i) Then

            If funciones.BuscarEnColeccion( _
                    colPagosACuenta, _
                    CStr(Me.ListPagosACuenta.ItemData(i))) Then

                colpcta.Add colPagosACuenta.item( _
                                CStr(Me.ListPagosACuenta.ItemData(i)))

            End If

        End If

    Next i

    TotalizarDiferenciasCambio

    MostrarPosiblesRetenciones _
        col, _
        colc, _
        colpcta, _
        colPercepciones

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

    If formLoading Then Exit Sub

    If item >= 0 And item < Me.lstFacturas.ListCount Then

        If Me.lstFacturas.Checked(item) Then
            InicializarFacturaSeleccionada item
        End If

        Me.txtParcialAbonado.Enabled = Me.lstFacturas.Checked(item)
        Me.txtParcialAbonar.Enabled = Me.lstFacturas.Checked(item)
        Me.txtOtrosParcialAbonado.Enabled = Me.lstFacturas.Checked(item)
        Me.txtOtrosParcialAbonar.Enabled = Me.lstFacturas.Checked(item)
        Me.txtTotalParcialAbonado.Enabled = Me.lstFacturas.Checked(item)
        Me.txtTotalParcialAbonar.Enabled = Me.lstFacturas.Checked(item)

    End If

    calcularOrigenes

End Sub


Private Sub lstFacturas_MouseDown(Button As Integer, _
                                  Shift As Integer, _
                                  x As Single, _
                                  y As Single)

    Dim i As Long

    If Button <> vbRightButton Then Exit Sub

    For i = 0 To Me.lstFacturas.ListCount - 1

        If Me.lstFacturas.Selected(i) Then

            Me.mnuCrearCompensatorio.Enabled = _
                Me.lstFacturas.Checked(i)

            PopupMenu Me.emergente
            Exit For

        End If

    Next i

End Sub


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
            
            For Each c In OrdenPago.Compensatorios
            
                If Not c.Comprobante Is Nothing Then
            
                    If c.Comprobante.Id = Factura.Id Then
                        hay = True
                        Exit For
                    End If
            
                End If
            
            Next c

            Dim Cant As Long

            If DAOCompensatorios.FindAll("id_orden_pago= " & OrdenPago.Id & " and  id_comprobante=" & Factura.Id).count > 0 Then hay = True

            If hay Then
                MsgBox "Ya existe un compensatorio para el comprobante indicado!", vbInformation, "Error"
            Else
                If ivamax Then
                    MsgBox "No puede crear un compensatorio cuando hay multiples alícuotas!", vbInformation, "Error"
                Else
                
                d.Cargar Factura, OrdenPago
                d.Show 1

                mostrarCompensatorios
                Totalizar
                Exit For

            End If
        End If

    End If

Next i
End Sub

Private Sub mostrarCompensatorios()
    Me.gridCompensatorios.ItemCount = OrdenPago.Compensatorios.count
    verCompensatorios
End Sub


Private Sub PushButton1_Click()

    If Me.cboProveedores.ListIndex <> -1 Then
        Set prov = colProveedores.item(CStr(Me.cboProveedores.ItemData(Me.cboProveedores.ListIndex)))

        If IsSomething(prov) Then
            Dim Nueva As New Collection
            Set Nueva = DAORetenciones.FindAllWithAlicuotas(prov.cuit)    '
            
            Set alicuotas = DAORetenciones.FindAllWithAlicuotas(prov.cuit)    '
            ActualizarAlicuotas
        End If
    Else
        Set prov = Nothing

    End If

    MostrarFacturas
End Sub


Private Sub radioConcepto_Click()
    If formLoading Then Exit Sub
    If formLoaded Then
        LimpiarFacturasYValores
        MostrarPosiblesRetenciones New Collection
        Totalizar
    End If
    ActivarControles
End Sub


Private Sub LimpiarFacturasYValores()

    Set colFacturas = New Collection
    Set colPagosACuenta = New Collection
    Set colDeudaCompensatorios = New Collection

    Set vFactElegida = Nothing
    Set vCompeElegido = Nothing
    Set prov = Nothing

    Me.lstFacturas.Clear
    Me.ListPagosACuenta.Clear
    Me.lstDeudaCompensatorios.Clear

    Me.lblCantidadComprobantes.caption = _
        "Cbtes. Mostrados: 0"

    limpiarParciales

End Sub


Private Sub ActivarControles()

    Me.cboProveedores.Enabled = Me.radioFacturaProveedor.value
    Me.lstFacturas.Enabled = Me.radioFacturaProveedor.value

    Me.cboCuentas.Enabled = Me.radioConcepto.value
    Me.txtDetalle.Enabled = Me.radioConcepto.value

    If Not Me.cboProveedores.Enabled Then
        Me.cboProveedores.ListIndex = -1
    End If

    If Not Me.lstFacturas.Enabled Then
        Me.lstFacturas.Clear
    End If

End Sub


Private Sub radioFacturaProveedor_Click()
    If formLoading Then Exit Sub
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
    
    
    If Not EsFechaValida(Values(3)) Then
        MsgBox "Fecha inválida. Use formato dd/mm/yyyy", vbExclamation
        Exit Sub
    End If
    
    operacion.FechaOperacion = CDate(Values(3))
    
    If IsNumeric(Values(4)) Then
        Set operacion.caja = DAOCaja.FindById(Values(4))
    End If
    operacion.EntradaSalida = OPSalida
    OrdenPago.OperacionesCaja.Add operacion
    Totalizar
End Sub


Private Sub gridCajaOperaciones_UnboundDelete(ByVal RowIndex As Long, ByVal Bookmark As Variant)
    If RowIndex > 0 And OrdenPago.OperacionesCaja.count >= RowIndex Then
        OrdenPago.OperacionesCaja.remove RowIndex
        Totalizar
    End If
End Sub


Private Sub Totalizar()

    OrdenPago.StaticTotalOrigenes = OrdenPago.TotalOrigenes

    Me.lblTotal.caption = _
        "Total orden de pago en " & _
        FormatCurrency(funciones.FormatearDecimales( _
            OrdenPago.StaticTotalOrigenes + _
            OrdenPago.StaticTotalRetenido))

    GridEXHelper.AutoSizeColumns Me.gridCajaOperaciones
    GridEXHelper.AutoSizeColumns Me.gridDepositosOperaciones
    GridEXHelper.AutoSizeColumns Me.gridCheques
    GridEXHelper.AutoSizeColumns Me.gridPercepciones

    TotalizarSolapas
    calcularOrigenes

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

    Me.txtDiferenciaCambioPago.Text = T
    Me.txtDifTipoCambioIVA.Text = TIVA
    Me.txtDifCambio = TTOTAL

    If ReadOnly Then
        Dim s As String
        s = OrdenPago.DiferenciaCambioEnNG
        Me.txtDifCambioNG1.Text = s
        s = OrdenPago.DiferenciaCambioEnTOTAL
        Me.txtDifCambioTOTAL1.Text = s
    End If

End Function


Private Sub gridCajaOperaciones_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex <= OrdenPago.OperacionesCaja.count Then
        Set operacion = OrdenPago.OperacionesCaja.item(RowIndex)
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
    If RowIndex > 0 And OrdenPago.OperacionesCaja.count > 0 Then
        Set operacion = OrdenPago.OperacionesCaja.item(RowIndex)
        operacion.Monto = Values(1)
        operacion.Comprobante = Values(5)
        If IsNumeric(Values(2)) Then
            Set operacion.moneda = DAOMoneda.GetById(Values(2))
        End If
        
        If Not EsFechaValida(Values(3)) Then
            MsgBox "Fecha inválida. Use formato dd/mm/yyyy", vbExclamation
            Exit Sub
        End If
        
        operacion.FechaOperacion = CDate(Values(3))
        
        If IsNumeric(Values(4)) Then
            Set operacion.caja = DAOCaja.FindById(Values(4))
        End If
        operacion.EntradaSalida = OPSalida
        Totalizar
    End If
End Sub


Private Sub gridDepositosOperaciones_BeforeUpdate(ByVal Cancel As GridEX20.JSRetBoolean)

    Dim msg As String
    msg = ""

    If Not EsImporteValido(Me.gridDepositosOperaciones.value(1)) Then
        msg = msg & "Monto inválido." & vbNewLine
    End If

    If LenB(CStr(Me.gridDepositosOperaciones.value(2))) = 0 Then
        msg = msg & "Debe indicar moneda." & vbNewLine
    End If

    If Not EsFechaValida(CStr(Me.gridDepositosOperaciones.value(3))) Then
        msg = msg & "Fecha inválida. Use formato dd/mm/aaaa." & vbNewLine
    End If

    If LenB(CStr(Me.gridDepositosOperaciones.value(4))) = 0 Then
        msg = msg & "Debe indicar cuenta bancaria." & vbNewLine
    End If

    Cancel = LenB(msg) > 0

    If Cancel Then
        MsgBox msg, vbExclamation, "No se puede actualizar la fila de banco"
    End If

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
    
    If Not EsFechaValida(Values(3)) Then
        MsgBox "Fecha inválida. Use formato dd/mm/aaaa", vbExclamation
        Exit Sub
    End If
    
    operacion.FechaOperacion = CDate(Values(3))
    
    If IsNumeric(Values(4)) Then
        Set operacion.CuentaBancaria = DAOCuentaBancaria.FindById(Values(4))
    End If
    operacion.EntradaSalida = OPSalida
    OrdenPago.operacionesBanco.Add operacion
    
    Totalizar
End Sub


Private Sub gridDepositosOperaciones_UnboundDelete(ByVal RowIndex As Long, ByVal Bookmark As Variant)
    If RowIndex > 0 And OrdenPago.operacionesBanco.count >= RowIndex Then
        OrdenPago.operacionesBanco.remove RowIndex
        Totalizar
        
    End If
End Sub


Private Sub gridDepositosOperaciones_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex <= OrdenPago.operacionesBanco.count Then
        Set operacion = OrdenPago.operacionesBanco.item(RowIndex)
        'FORMATCURRENCY
        Values(1) = FormatCurrency(funciones.FormatearDecimales(operacion.Monto))
        If IsSomething(operacion.moneda) Then
            Values(2) = operacion.moneda.NombreCorto
        End If
        Values(3) = Format$(operacion.FechaOperacion, "dd/mm/yyyy")
        If IsSomething(operacion.CuentaBancaria) Then
            Values(4) = operacion.CuentaBancaria.DescripcionFormateada
        End If
        If IsSomething(operacion) Then
            Values(5) = operacion.Comprobante
        End If
    End If
End Sub


Private Sub gridDepositosOperaciones_UnboundUpdate(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)

    If RowIndex > 0 And OrdenPago.operacionesBanco.count >= RowIndex Then

        Set operacion = OrdenPago.operacionesBanco.item(RowIndex)

        operacion.Monto = ImporteDesdeTexto(Values(1))
        operacion.Comprobante = Values(5)

        If IsNumeric(Values(2)) Then
            Set operacion.moneda = DAOMoneda.GetById(CLng(Values(2)))
        End If

        If Not EsFechaValida(CStr(Values(3))) Then
            MsgBox "Fecha inválida. Use formato dd/mm/aaaa.", vbExclamation
            Exit Sub
        End If

        operacion.FechaOperacion = FechaDesdeTexto(CStr(Values(3)))

        If IsNumeric(Values(4)) Then
            Set operacion.CuentaBancaria = DAOCuentaBancaria.FindById(CLng(Values(4)))
        End If

        operacion.EntradaSalida = OPSalida

        Totalizar
        Me.gridDepositosOperaciones.Refresh

    End If

End Sub


Private Sub gridCheques_BeforeUpdate(ByVal Cancel As GridEX20.JSRetBoolean)
    Dim msg As New Collection

    ' REVISA QUE EN LA COLECCION DE CHEQUES DE TERCEROS QUE SE ESTAN CARGANDO NO EST? INGRESADO EL MISMO CHEQUE, SI LO DETECTA GENERA MSG DE ERROR
    If funciones.BuscarEnColeccion(OrdenPago.ChequesTerceros, CStr(Me.gridCheques.value(1))) Then
        msg.Add "El cheque seleccionado ya fue ingresado anteriormente."
    End If

    Cancel = (msg.count > 0)
    If Cancel Then MsgBox funciones.JoinCollectionValues(msg, vbNewLine), vbExclamation

End Sub


Private Sub gridCheques_UnboundAddNew(ByVal NewRowBookmark As GridEX20.JSRetVariant, ByVal Values As GridEX20.JSRowData)
    Set cheque = Nothing
    If IsNumeric(Values(1)) Then Set cheque = DAOCheques.FindById(Values(1))
    If IsSomething(cheque) Then
        OrdenPago.ChequesTerceros.Add cheque, CStr(cheque.Id)
    End If
    
    Totalizar

End Sub


Private Sub gridCheques_UnboundDelete(ByVal RowIndex As Long, ByVal Bookmark As Variant)
    If RowIndex > 0 Then
        OrdenPago.ChequesTerceros.remove RowIndex
        Totalizar
    End If
End Sub


Private Sub gridCheques_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex <= OrdenPago.ChequesTerceros.count Then
        Set cheque = OrdenPago.ChequesTerceros.item(RowIndex)

        Values(1) = cheque.numero & " "

        'FORMATCURRENCY
        Values(2) = FormatCurrency(cheque.Monto)
        Values(3) = cheque.FechaVencimiento
        If IsSomething(cheque.moneda) Then Values(4) = cheque.moneda.NombreCorto
        If IsSomething(cheque.Banco) Then Values(5) = cheque.Banco.nombre
        Values(6) = cheque.OrigenDestino
        Values(7) = cheque.OrigenCheque
    
    End If
End Sub


Private Sub gridCheques_UnboundUpdate(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex > 0 And OrdenPago.ChequesTerceros.count >= RowIndex Then
        Set cheque = Nothing
        If IsNumeric(Values(1)) Then Set cheque = DAOCheques.FindById(Values(1))
        If IsSomething(cheque) Then
            OrdenPago.ChequesTerceros.Add cheque, , , RowIndex
            OrdenPago.ChequesTerceros.remove RowIndex
        End If
        Totalizar
    End If
End Sub


Private Sub txtBuscarFactura_GotFocus()
    Me.txtBuscarFactura.SelStart = 0
    Me.txtBuscarFactura.SelLength = Len(Me.txtBuscarFactura.Text)
End Sub


Private Sub txtBuscarFactura_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim cont As Long
    Dim i As Long

    If KeyCode <> vbKeyReturn Then Exit Sub
    If LenB(Trim$(Me.txtBuscarFactura.Text)) = 0 Then Exit Sub

    If colFacturas.count > 0 Then

        For Each vFacturaProveedor In colFacturas

            If InStr(1, _
                     vFacturaProveedor.numero, _
                     Me.txtBuscarFactura.Text, _
                     vbTextCompare) > 0 Then

                For i = 0 To Me.lstFacturas.ListCount - 1

                    If Me.lstFacturas.ItemData(i) = _
                            vFacturaProveedor.Id Then

                        Me.lstFacturas.Checked(i) = True
                        InicializarFacturaSeleccionada i

                        cont = cont + 1
                        Exit For

                    End If

                Next i

            End If

        Next vFacturaProveedor

    End If

    If cont = 0 Then

        MsgBox "No se encontraron facturas con ese número en la lista.", _
               vbExclamation, _
               "Buscar comprobante"

    Else

        calcularOrigenes

        MsgBox "Se encontró " & cont & " factura/s.", _
               vbInformation, _
               "Buscar comprobante"

        Me.txtBuscarFactura.Text = vbNullString
        Me.txtBuscarFactura.SetFocus

    End If

End Sub


Private Sub txtDifCambio_GotFocus()
    foco Me.txtDifCambio
End Sub


Private Sub txtDifCambioNG1_Change()
    If formLoading Then Exit Sub
    OrdenPago.DiferenciaCambioEnNG = val(Me.txtDifCambioNG1)
    Totalizar
End Sub


Private Sub txtDifCambioTOTAL1_Change()
    If formLoading Then Exit Sub
    OrdenPago.DiferenciaCambioEnTOTAL = val(Me.txtDifCambioTOTAL1)
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
    If formLoading Then Exit Sub
    OrdenPago.OtrosDescuentos = val(Me.txtOtrosDescuentos.Text)
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

Private Sub RecalcularTotalFacturaElegida()
    Me.txtTotalParcialAbonar = (CDbl(txtParcialAbonar)) + (CDbl(Me.txtOtrosParcialAbonar))

    If Me.txtTotalParcialAbonar = "0" Then Me.txtTotalParcialAbonar = "0.00"


    vFactElegida.totalAbonado = CDbl(txtTotalParcialAbonar)

End Sub


Private Sub txtOtrosParcialAbonar_LostFocus()
'  If LenB(Me.txtOtrosParcialAbonar) > 0 Then
'
'        vFactElegida.OtrosAbonado = CDbl(Me.txtOtrosParcialAbonar)
'        RecalcularTotalFacturaElegida
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
        'COMENTO ESTA LINEA PORQUE ESTA COMPROBACI?N HACE QUE EL FORM SE CONGELE Y NO SE PUEDA AVANZAR CON LA CARGA.
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


Private Sub RecalcularFacturaElegida()
    RecalcularNetoGravadoFacturaElegida
    RecalcularOtrosFacturaelegida
End Sub

Private Sub RecalcularNetoGravadoFacturaElegida()
    If LenB(txtParcialAbonar) > 0 And IsNumeric(txtParcialAbonar) Then

        vFactElegida.NetoGravadoAbonado = CDbl(Me.txtParcialAbonar)
        RecalcularTotalFacturaElegida
    End If
End Sub

Private Sub txtParcialAbonar_KeyUp(KeyCode As Integer, Shift As Integer)
    RecalcularNetoGravadoFacturaElegida

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
    If Me.txtTotalParcialAbonar = "" Then Me.txtTotalParcialAbonar = 0
    If Me.txtParcialAbonar = "" Then Me.txtParcialAbonar = 0
    
        If CDbl(Me.txtTotalParcialAbonar) > vFactElegida.ImporteTotalSaldo Or CDbl(Me.txtParcialAbonar) < 0 Then
            Me.txtTotalParcialAbonar.backColor = vbRed
            Me.txtTotalParcialAbonar.ForeColor = vbWhite
        Else
            Me.txtTotalParcialAbonar.backColor = vbWhite
            Me.txtTotalParcialAbonar.ForeColor = vbBlack
        End If
    End If
End Sub


Private Function EsFechaValida(ByVal txt As String) As Boolean
    Dim partes() As String
    Dim d As Integer, m As Integer, y As Integer
    
    EsFechaValida = False
    
    'Debe tener formato con /
    If InStr(txt, "/") = 0 Then Exit Function
    
    partes = Split(txt, "/")
    
    'Debe tener 3 partes
    If UBound(partes) <> 2 Then Exit Function
    
    'Día, mes y año numéricos
    If Not IsNumeric(partes(0)) Or Not IsNumeric(partes(1)) Or Not IsNumeric(partes(2)) Then Exit Function
    
    d = CInt(partes(0))
    m = CInt(partes(1))
    y = CInt(partes(2))
    
    'Validar largo del año (clave)
    If Len(partes(2)) <> 4 Then Exit Function
    
    'Validaciones básicas
    If d < 1 Or d > 31 Then Exit Function
    If m < 1 Or m > 12 Then Exit Function
    If y < 1900 Or y > 2100 Then Exit Function
    
    'Validar fecha real (ej: 31/02 no pasa)
    On Error GoTo invalida
    Dim fechaTest As Date
    fechaTest = DateSerial(y, m, d)
    
    'Chequeo cruzado
    If Day(fechaTest) <> d Or Month(fechaTest) <> m Or Year(fechaTest) <> y Then Exit Function
    
    EsFechaValida = True
    Exit Function
    
invalida:
    EsFechaValida = False
End Function


Private Sub TotalizarSolapas()
    Dim totalCaja As Double
    Dim totalChequesPropios As Double
    Dim totalBanco As Double
    Dim totalChequesTerceros As Double
    Dim totalPercepciones As Double
    Dim TotalCompensatorios As Double

    Dim OpCaja As operacion
    Dim opBanco As operacion
    Dim chPropio As cheque
    Dim chTercero As cheque
    Dim per As clsPercepcionesOrdenPago
    Dim comp As Compensatorio

    '=========================
    ' CAJA
    '=========================
    totalCaja = 0

    For Each OpCaja In OrdenPago.OperacionesCaja
        If IsSomething(OpCaja) Then
            totalCaja = totalCaja + OpCaja.Monto
        End If
    Next OpCaja

    '=========================
    ' CHEQUES PROPIOS
    '=========================
    totalChequesPropios = 0

    For Each chPropio In OrdenPago.ChequesPropios
        If IsSomething(chPropio) Then
            totalChequesPropios = totalChequesPropios + chPropio.Monto
        End If
    Next chPropio

    '=========================
    ' BANCO
    '=========================
    totalBanco = 0

    For Each opBanco In OrdenPago.operacionesBanco
        If IsSomething(opBanco) Then
            totalBanco = totalBanco + opBanco.Monto
        End If
    Next opBanco

    '=========================
    ' CHEQUES DE TERCEROS
    '=========================
    totalChequesTerceros = 0

    For Each chTercero In OrdenPago.ChequesTerceros
        If IsSomething(chTercero) Then
            totalChequesTerceros = totalChequesTerceros + chTercero.Monto
        End If
    Next chTercero

    '=========================
    ' PERCEPCIONES
    '=========================
    totalPercepciones = 0

    For Each per In OrdenPago.percepciones
        If IsSomething(per) Then
            totalPercepciones = totalPercepciones + per.Monto
        End If
    Next per

    '=========================
    ' COMPENSATORIOS
    '=========================
    TotalCompensatorios = 0

    For Each comp In OrdenPago.Compensatorios
        If IsSomething(comp) Then
            TotalCompensatorios = TotalCompensatorios + comp.Monto
        End If
    Next comp

    '=========================
    ' MOSTRAR RESULTADOS
    '=========================
    Me.txtTotalizadorCAJA.caption = _
        FormatCurrency(funciones.FormatearDecimales(totalCaja))

    Me.txtTotalizadorCHEQUESPROPIOS.caption = _
        FormatCurrency(funciones.FormatearDecimales(totalChequesPropios))

    Me.txtTotalizadorBANCO.caption = _
        FormatCurrency(funciones.FormatearDecimales(totalBanco))

    Me.txtTotalizadorCHEQUES3ROS.caption = _
        FormatCurrency(funciones.FormatearDecimales(totalChequesTerceros))

    Me.txtTotalizadorPERCEPCIONES.caption = _
        FormatCurrency(funciones.FormatearDecimales(totalPercepciones))

    Me.txtTotalizadorCOMPENSATORIOS.caption = _
        FormatCurrency(funciones.FormatearDecimales(TotalCompensatorios))
End Sub


Private Function EsImporteValido(ByVal valor As Variant) As Boolean

    Dim importe As Double

    EsImporteValido = TryImporteDesdeTexto(valor, importe)

End Function


Private Function ImporteDesdeTexto(ByVal valor As Variant) As Double

    Dim importe As Double

    If Not TryImporteDesdeTexto(valor, importe) Then
        Err.Raise 13, _
                  "ImporteDesdeTexto", _
                  "El importe ingresado no es válido."
    End If

    ImporteDesdeTexto = importe

End Function


Private Function FechaDesdeTexto(ByVal txt As String) As Date
    Dim partes() As String

    txt = Trim$(txt)
    partes = Split(txt, "/")

    FechaDesdeTexto = DateSerial(CInt(partes(2)), CInt(partes(1)), CInt(partes(0)))
End Function


Private Sub InicializarFacturaSeleccionada(ByVal item As Long)

    Dim F As clsFacturaProveedor
    Dim c As Collection

    If item < 0 Or item >= Me.lstFacturas.ListCount Then Exit Sub

    If Not funciones.BuscarEnColeccion( _
            colFacturas, _
            CStr(Me.lstFacturas.ItemData(item))) Then
        Exit Sub
    End If

    Set F = colFacturas.item( _
                CStr(Me.lstFacturas.ItemData(item)))

    If OrdenPago.estado <> EstadoOrdenPago_pendiente Then Exit Sub

    'Solamente inicializar cuando los dos componentes están en cero.
    If F.NetoGravadoAbonado = 0 And _
       F.OtrosAbonado = 0 Then

        Set c = DAOOrdenPago.FindAbonadoFactura( _
                    F.Id, _
                    OrdenPago.Id)

        If Not c Is Nothing Then
            If c.count >= 3 Then
                F.NetoGravadoAbonado = CDbl(c(2))
                F.OtrosAbonado = CDbl(c(3))
            End If
        End If

        'Si el DAO no encontró un pago previo,
        'proponer el saldo pendiente completo.
        If F.NetoGravadoAbonado = 0 And _
           F.OtrosAbonado = 0 Then

            F.NetoGravadoAbonado = _
                F.ImporteNetoGravadoSaldo

            F.OtrosAbonado = _
                F.ImporteOtrosSaldo
        End If

    End If

    F.totalAbonado = _
        F.NetoGravadoAbonado + _
        F.OtrosAbonado

End Sub


Private Function TryImporteDesdeTexto( _
        ByVal valor As Variant, _
        ByRef resultado As Double) As Boolean

    On Error GoTo ImporteInvalido

    Dim s As String
    Dim posPunto As Long
    Dim posComa As Long
    Dim cantidadDecimales As Long
    Dim separadorDecimalSistema As String

    If IsNull(valor) Or IsEmpty(valor) Then GoTo ImporteInvalido

    s = Trim$(CStr(valor))

    s = Replace(s, "AR$", "", 1, -1, vbTextCompare)
    s = Replace(s, "$", "")
    s = Replace(s, " ", "")

    If LenB(s) = 0 Then GoTo ImporteInvalido

    posPunto = InStrRev(s, ".")
    posComa = InStrRev(s, ",")

    If posPunto > 0 And posComa > 0 Then

        'Si contiene ambos separadores, el último es el decimal.
        'Ejemplos:
        '1.010,50  -> 1010,50
        '1,010.50  -> 1010.50
        If posComa > posPunto Then
            s = Replace(s, ".", "")
            s = Replace(s, ",", "|")
        Else
            s = Replace(s, ",", "")
            s = Replace(s, ".", "|")
        End If

    ElseIf posPunto > 0 Then

        cantidadDecimales = Len(s) - posPunto

        If cantidadDecimales = 1 Or cantidadDecimales = 2 Then
            '10.10 se interpreta como diez con diez centavos.
            s = Replace(s, ".", "|")
        Else
            '1.010 se interpreta como mil diez.
            s = Replace(s, ".", "")
        End If

    ElseIf posComa > 0 Then

        cantidadDecimales = Len(s) - posComa

        If cantidadDecimales = 1 Or cantidadDecimales = 2 Then
            '10,10 se interpreta como diez con diez centavos.
            s = Replace(s, ",", "|")
        Else
            '1,010 se interpreta como mil diez.
            s = Replace(s, ",", "")
        End If

    End If

    'Utilizar el separador decimal configurado en Windows.
    separadorDecimalSistema = Mid$(Format$(0.5, "0.0"), 2, 1)
    s = Replace(s, "|", separadorDecimalSistema)

    If Not IsNumeric(s) Then GoTo ImporteInvalido

    resultado = CDbl(s)
    TryImporteDesdeTexto = True
    Exit Function

ImporteInvalido:
    resultado = 0
    TryImporteDesdeTexto = False

End Function
