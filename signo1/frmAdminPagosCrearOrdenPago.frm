VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmAdminPagosCrearOrdenPago 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Orden de Pago"
   ClientHeight    =   11595
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
   ScaleHeight     =   11595
   ScaleWidth      =   17580
   Begin XtremeSuiteControls.GroupBox GroupBox5 
      Height          =   1575
      Left            =   120
      TabIndex        =   82
      Top             =   9840
      Width           =   6855
      _Version        =   786432
      _ExtentX        =   12091
      _ExtentY        =   2778
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
      Begin GridEX20.GridEX GridEX1 
         Height          =   1215
         Left            =   120
         TabIndex        =   83
         Top             =   240
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   2143
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
         Column(2)       =   "frmAdminPagosCrearOrdenPago.frx":0100
         Column(3)       =   "frmAdminPagosCrearOrdenPago.frx":01EC
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmAdminPagosCrearOrdenPago.frx":02D0
         FormatStyle(2)  =   "frmAdminPagosCrearOrdenPago.frx":03F8
         FormatStyle(3)  =   "frmAdminPagosCrearOrdenPago.frx":04A8
         FormatStyle(4)  =   "frmAdminPagosCrearOrdenPago.frx":055C
         FormatStyle(5)  =   "frmAdminPagosCrearOrdenPago.frx":0634
         FormatStyle(6)  =   "frmAdminPagosCrearOrdenPago.frx":06EC
         ImageCount      =   0
         PrinterProperties=   "frmAdminPagosCrearOrdenPago.frx":07CC
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
         Top             =   2760
         Width           =   3375
         _Version        =   786432
         _ExtentX        =   5953
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Label13"
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
         Height          =   195
         Left            =   120
         TabIndex        =   69
         Tag             =   "tot fac - tot ret"
         Top             =   840
         Width           =   1170
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
         Height          =   5145
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   10140
         _Version        =   786432
         _ExtentX        =   17886
         _ExtentY        =   9075
         _StockProps     =   68
         Appearance      =   10
         Color           =   32
         PaintManager.ShowIcons=   -1  'True
         ItemCount       =   5
         SelectedItem    =   4
         Item(0).Caption =   "Cheques Propios"
         Item(0).ControlCount=   1
         Item(0).Control(0)=   "gridChequesPropios"
         Item(1).Caption =   "Banco"
         Item(1).ControlCount=   1
         Item(1).Control(0)=   "gridDepositosOperaciones"
         Item(2).Caption =   "Cheques 3ros"
         Item(2).ControlCount=   1
         Item(2).Control(0)=   "gridCheques"
         Item(3).Caption =   "Caja"
         Item(3).ControlCount=   1
         Item(3).Control(0)=   "gridCajaOperaciones"
         Item(4).Caption =   "Compensatorios"
         Item(4).ControlCount=   1
         Item(4).Control(0)=   "gridCompensatorios"
         Begin GridEX20.GridEX gridDepositosOperaciones 
            Height          =   4575
            Left            =   -69880
            TabIndex        =   2
            Top             =   435
            Visible         =   0   'False
            Width           =   9810
            _ExtentX        =   17304
            _ExtentY        =   8070
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
            Column(1)       =   "frmAdminPagosCrearOrdenPago.frx":099C
            Column(2)       =   "frmAdminPagosCrearOrdenPago.frx":0AFC
            Column(3)       =   "frmAdminPagosCrearOrdenPago.frx":0C38
            Column(4)       =   "frmAdminPagosCrearOrdenPago.frx":0D6C
            Column(5)       =   "frmAdminPagosCrearOrdenPago.frx":0EB0
            FormatStylesCount=   6
            FormatStyle(1)  =   "frmAdminPagosCrearOrdenPago.frx":0FB4
            FormatStyle(2)  =   "frmAdminPagosCrearOrdenPago.frx":10EC
            FormatStyle(3)  =   "frmAdminPagosCrearOrdenPago.frx":119C
            FormatStyle(4)  =   "frmAdminPagosCrearOrdenPago.frx":1250
            FormatStyle(5)  =   "frmAdminPagosCrearOrdenPago.frx":1328
            FormatStyle(6)  =   "frmAdminPagosCrearOrdenPago.frx":13E0
            ImageCount      =   0
            PrinterProperties=   "frmAdminPagosCrearOrdenPago.frx":14C0
         End
         Begin GridEX20.GridEX gridCajaOperaciones 
            Height          =   4575
            Left            =   -69895
            TabIndex        =   10
            Top             =   435
            Visible         =   0   'False
            Width           =   9810
            _ExtentX        =   17304
            _ExtentY        =   8070
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
            Column(1)       =   "frmAdminPagosCrearOrdenPago.frx":1698
            Column(2)       =   "frmAdminPagosCrearOrdenPago.frx":17F8
            Column(3)       =   "frmAdminPagosCrearOrdenPago.frx":1934
            Column(4)       =   "frmAdminPagosCrearOrdenPago.frx":1A68
            Column(5)       =   "frmAdminPagosCrearOrdenPago.frx":1B9C
            FormatStylesCount=   6
            FormatStyle(1)  =   "frmAdminPagosCrearOrdenPago.frx":1CA0
            FormatStyle(2)  =   "frmAdminPagosCrearOrdenPago.frx":1DD8
            FormatStyle(3)  =   "frmAdminPagosCrearOrdenPago.frx":1E88
            FormatStyle(4)  =   "frmAdminPagosCrearOrdenPago.frx":1F3C
            FormatStyle(5)  =   "frmAdminPagosCrearOrdenPago.frx":2014
            FormatStyle(6)  =   "frmAdminPagosCrearOrdenPago.frx":20CC
            ImageCount      =   0
            PrinterProperties=   "frmAdminPagosCrearOrdenPago.frx":21AC
         End
         Begin GridEX20.GridEX gridChequesPropios 
            Height          =   4575
            Left            =   -69895
            TabIndex        =   9
            Top             =   435
            Visible         =   0   'False
            Width           =   9810
            _ExtentX        =   17304
            _ExtentY        =   8070
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
            Column(1)       =   "frmAdminPagosCrearOrdenPago.frx":2384
            Column(2)       =   "frmAdminPagosCrearOrdenPago.frx":24EC
            Column(3)       =   "frmAdminPagosCrearOrdenPago.frx":2620
            Column(4)       =   "frmAdminPagosCrearOrdenPago.frx":275C
            Column(5)       =   "frmAdminPagosCrearOrdenPago.frx":28C4
            FormatStylesCount=   6
            FormatStyle(1)  =   "frmAdminPagosCrearOrdenPago.frx":29BC
            FormatStyle(2)  =   "frmAdminPagosCrearOrdenPago.frx":2AF4
            FormatStyle(3)  =   "frmAdminPagosCrearOrdenPago.frx":2BA4
            FormatStyle(4)  =   "frmAdminPagosCrearOrdenPago.frx":2C58
            FormatStyle(5)  =   "frmAdminPagosCrearOrdenPago.frx":2D30
            FormatStyle(6)  =   "frmAdminPagosCrearOrdenPago.frx":2DE8
            ImageCount      =   0
            PrinterProperties=   "frmAdminPagosCrearOrdenPago.frx":2EC8
         End
         Begin GridEX20.GridEX gridCheques 
            Height          =   4575
            Left            =   -69895
            TabIndex        =   8
            Top             =   435
            Visible         =   0   'False
            Width           =   9810
            _ExtentX        =   17304
            _ExtentY        =   8070
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
            Column(1)       =   "frmAdminPagosCrearOrdenPago.frx":30A0
            Column(2)       =   "frmAdminPagosCrearOrdenPago.frx":3220
            Column(3)       =   "frmAdminPagosCrearOrdenPago.frx":33C0
            Column(4)       =   "frmAdminPagosCrearOrdenPago.frx":34B8
            Column(5)       =   "frmAdminPagosCrearOrdenPago.frx":35F4
            Column(6)       =   "frmAdminPagosCrearOrdenPago.frx":3700
            Column(7)       =   "frmAdminPagosCrearOrdenPago.frx":37D0
            FormatStylesCount=   6
            FormatStyle(1)  =   "frmAdminPagosCrearOrdenPago.frx":38BC
            FormatStyle(2)  =   "frmAdminPagosCrearOrdenPago.frx":39F4
            FormatStyle(3)  =   "frmAdminPagosCrearOrdenPago.frx":3AA4
            FormatStyle(4)  =   "frmAdminPagosCrearOrdenPago.frx":3B58
            FormatStyle(5)  =   "frmAdminPagosCrearOrdenPago.frx":3C30
            FormatStyle(6)  =   "frmAdminPagosCrearOrdenPago.frx":3CE8
            ImageCount      =   0
            PrinterProperties=   "frmAdminPagosCrearOrdenPago.frx":3DC8
         End
         Begin GridEX20.GridEX gridCompensatorios 
            Height          =   4575
            Left            =   105
            TabIndex        =   14
            Top             =   435
            Width           =   9810
            _ExtentX        =   17304
            _ExtentY        =   8070
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
            Column(1)       =   "frmAdminPagosCrearOrdenPago.frx":3FA0
            Column(2)       =   "frmAdminPagosCrearOrdenPago.frx":40E8
            Column(3)       =   "frmAdminPagosCrearOrdenPago.frx":41F4
            Column(4)       =   "frmAdminPagosCrearOrdenPago.frx":42E0
            Column(5)       =   "frmAdminPagosCrearOrdenPago.frx":43E4
            FormatStylesCount=   6
            FormatStyle(1)  =   "frmAdminPagosCrearOrdenPago.frx":4524
            FormatStyle(2)  =   "frmAdminPagosCrearOrdenPago.frx":465C
            FormatStyle(3)  =   "frmAdminPagosCrearOrdenPago.frx":470C
            FormatStyle(4)  =   "frmAdminPagosCrearOrdenPago.frx":47C0
            FormatStyle(5)  =   "frmAdminPagosCrearOrdenPago.frx":4898
            FormatStyle(6)  =   "frmAdminPagosCrearOrdenPago.frx":4950
            ImageCount      =   0
            PrinterProperties=   "frmAdminPagosCrearOrdenPago.frx":4A30
         End
      End
   End
   Begin GridEX20.GridEX gridBancos 
      Height          =   1845
      Left            =   1800
      TabIndex        =   3
      Top             =   12240
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
      Column(1)       =   "frmAdminPagosCrearOrdenPago.frx":4C08
      Column(2)       =   "frmAdminPagosCrearOrdenPago.frx":4D08
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminPagosCrearOrdenPago.frx":4DF8
      FormatStyle(2)  =   "frmAdminPagosCrearOrdenPago.frx":4F30
      FormatStyle(3)  =   "frmAdminPagosCrearOrdenPago.frx":4FE0
      FormatStyle(4)  =   "frmAdminPagosCrearOrdenPago.frx":5094
      FormatStyle(5)  =   "frmAdminPagosCrearOrdenPago.frx":516C
      FormatStyle(6)  =   "frmAdminPagosCrearOrdenPago.frx":5224
      ImageCount      =   0
      PrinterProperties=   "frmAdminPagosCrearOrdenPago.frx":5304
   End
   Begin GridEX20.GridEX gridCuentasBancarias 
      Height          =   1695
      Left            =   7800
      TabIndex        =   4
      Top             =   12360
      Visible         =   0   'False
      Width           =   4785
      _ExtentX        =   8440
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
      Column(1)       =   "frmAdminPagosCrearOrdenPago.frx":54DC
      Column(2)       =   "frmAdminPagosCrearOrdenPago.frx":5600
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminPagosCrearOrdenPago.frx":56F4
      FormatStyle(2)  =   "frmAdminPagosCrearOrdenPago.frx":582C
      FormatStyle(3)  =   "frmAdminPagosCrearOrdenPago.frx":58DC
      FormatStyle(4)  =   "frmAdminPagosCrearOrdenPago.frx":5990
      FormatStyle(5)  =   "frmAdminPagosCrearOrdenPago.frx":5A68
      FormatStyle(6)  =   "frmAdminPagosCrearOrdenPago.frx":5B20
      ImageCount      =   0
      PrinterProperties=   "frmAdminPagosCrearOrdenPago.frx":5C00
   End
   Begin GridEX20.GridEX gridMonedas 
      Height          =   1815
      Left            =   240
      TabIndex        =   5
      Top             =   12000
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
      Column(1)       =   "frmAdminPagosCrearOrdenPago.frx":5DD8
      Column(2)       =   "frmAdminPagosCrearOrdenPago.frx":5EFC
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminPagosCrearOrdenPago.frx":5FF0
      FormatStyle(2)  =   "frmAdminPagosCrearOrdenPago.frx":6128
      FormatStyle(3)  =   "frmAdminPagosCrearOrdenPago.frx":61D8
      FormatStyle(4)  =   "frmAdminPagosCrearOrdenPago.frx":628C
      FormatStyle(5)  =   "frmAdminPagosCrearOrdenPago.frx":6364
      FormatStyle(6)  =   "frmAdminPagosCrearOrdenPago.frx":641C
      ImageCount      =   0
      PrinterProperties=   "frmAdminPagosCrearOrdenPago.frx":64FC
   End
   Begin GridEX20.GridEX gridCajas 
      Height          =   1695
      Left            =   120
      TabIndex        =   6
      Top             =   12240
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
      Column(1)       =   "frmAdminPagosCrearOrdenPago.frx":66D4
      Column(2)       =   "frmAdminPagosCrearOrdenPago.frx":67D4
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminPagosCrearOrdenPago.frx":68C0
      FormatStyle(2)  =   "frmAdminPagosCrearOrdenPago.frx":69F8
      FormatStyle(3)  =   "frmAdminPagosCrearOrdenPago.frx":6AA8
      FormatStyle(4)  =   "frmAdminPagosCrearOrdenPago.frx":6B5C
      FormatStyle(5)  =   "frmAdminPagosCrearOrdenPago.frx":6C34
      FormatStyle(6)  =   "frmAdminPagosCrearOrdenPago.frx":6CEC
      ImageCount      =   0
      PrinterProperties=   "frmAdminPagosCrearOrdenPago.frx":6DCC
   End
   Begin GridEX20.GridEX gridChequesDisponibles 
      Height          =   1905
      Left            =   1800
      TabIndex        =   7
      Top             =   12000
      Visible         =   0   'False
      Width           =   8355
      _ExtentX        =   14737
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
      Column(1)       =   "frmAdminPagosCrearOrdenPago.frx":6FA4
      Column(2)       =   "frmAdminPagosCrearOrdenPago.frx":7124
      Column(3)       =   "frmAdminPagosCrearOrdenPago.frx":72C4
      Column(4)       =   "frmAdminPagosCrearOrdenPago.frx":73BC
      Column(5)       =   "frmAdminPagosCrearOrdenPago.frx":74F8
      Column(6)       =   "frmAdminPagosCrearOrdenPago.frx":7604
      Column(7)       =   "frmAdminPagosCrearOrdenPago.frx":7724
      Column(8)       =   "frmAdminPagosCrearOrdenPago.frx":7830
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminPagosCrearOrdenPago.frx":7924
      FormatStyle(2)  =   "frmAdminPagosCrearOrdenPago.frx":7A5C
      FormatStyle(3)  =   "frmAdminPagosCrearOrdenPago.frx":7B0C
      FormatStyle(4)  =   "frmAdminPagosCrearOrdenPago.frx":7BC0
      FormatStyle(5)  =   "frmAdminPagosCrearOrdenPago.frx":7C98
      FormatStyle(6)  =   "frmAdminPagosCrearOrdenPago.frx":7D50
      ImageCount      =   0
      PrinterProperties=   "frmAdminPagosCrearOrdenPago.frx":7E30
   End
   Begin GridEX20.GridEX gridChequeras 
      Height          =   1815
      Left            =   10560
      TabIndex        =   11
      Top             =   12000
      Width           =   6795
      _ExtentX        =   11986
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
      Column(1)       =   "frmAdminPagosCrearOrdenPago.frx":8008
      Column(2)       =   "frmAdminPagosCrearOrdenPago.frx":8128
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminPagosCrearOrdenPago.frx":8228
      FormatStyle(2)  =   "frmAdminPagosCrearOrdenPago.frx":8360
      FormatStyle(3)  =   "frmAdminPagosCrearOrdenPago.frx":8410
      FormatStyle(4)  =   "frmAdminPagosCrearOrdenPago.frx":84C4
      FormatStyle(5)  =   "frmAdminPagosCrearOrdenPago.frx":859C
      FormatStyle(6)  =   "frmAdminPagosCrearOrdenPago.frx":8654
      ImageCount      =   0
      PrinterProperties=   "frmAdminPagosCrearOrdenPago.frx":8734
   End
   Begin GridEX20.GridEX gridChequesChequera 
      Height          =   1710
      Left            =   10560
      TabIndex        =   12
      Top             =   11880
      Width           =   3420
      _ExtentX        =   6033
      _ExtentY        =   3016
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
      Column(1)       =   "frmAdminPagosCrearOrdenPago.frx":890C
      Column(2)       =   "frmAdminPagosCrearOrdenPago.frx":8A3C
      SortKeysCount   =   1
      SortKey(1)      =   "frmAdminPagosCrearOrdenPago.frx":8B3C
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminPagosCrearOrdenPago.frx":8BA4
      FormatStyle(2)  =   "frmAdminPagosCrearOrdenPago.frx":8CDC
      FormatStyle(3)  =   "frmAdminPagosCrearOrdenPago.frx":8D8C
      FormatStyle(4)  =   "frmAdminPagosCrearOrdenPago.frx":8E40
      FormatStyle(5)  =   "frmAdminPagosCrearOrdenPago.frx":8F18
      FormatStyle(6)  =   "frmAdminPagosCrearOrdenPago.frx":8FD0
      ImageCount      =   0
      PrinterProperties=   "frmAdminPagosCrearOrdenPago.frx":90B0
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
         Column(1)       =   "frmAdminPagosCrearOrdenPago.frx":9288
         Column(2)       =   "frmAdminPagosCrearOrdenPago.frx":93C4
         Column(3)       =   "frmAdminPagosCrearOrdenPago.frx":94C4
         Column(4)       =   "frmAdminPagosCrearOrdenPago.frx":95C8
         FormatStylesCount=   8
         FormatStyle(1)  =   "frmAdminPagosCrearOrdenPago.frx":96D0
         FormatStyle(2)  =   "frmAdminPagosCrearOrdenPago.frx":97F8
         FormatStyle(3)  =   "frmAdminPagosCrearOrdenPago.frx":98A8
         FormatStyle(4)  =   "frmAdminPagosCrearOrdenPago.frx":995C
         FormatStyle(5)  =   "frmAdminPagosCrearOrdenPago.frx":9A34
         FormatStyle(6)  =   "frmAdminPagosCrearOrdenPago.frx":9AEC
         FormatStyle(7)  =   "frmAdminPagosCrearOrdenPago.frx":9BCC
         FormatStyle(8)  =   "frmAdminPagosCrearOrdenPago.frx":9C68
         ImageCount      =   0
         PrinterProperties=   "frmAdminPagosCrearOrdenPago.frx":9D08
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
      Begin XtremeSuiteControls.PushButton btnExportarCbtes 
         Height          =   375
         Left            =   5280
         TabIndex        =   79
         Top             =   5640
         Width           =   1335
         _Version        =   786432
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Exportar"
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
Private cuentasBancarias As New Collection
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


Public Sub Cargar(op As OrdenPago)

    Me.caption = "Orden de Pago Nro " & OrdenPago.Id

    If Not IsSomething(op) Then
        MsgBox "La OP que está intentando visualizar est? en estado PENDIENTE. " & vbNewLine & "Por lo tanto no puede ser mostrada porque puede estar siendo editada." & vbNewLine & "Verifiquelo por favor.", vbCritical, "OP Pendiente"
        Unload Me
        Exit Sub

    End If

    Set OrdenPago = DAOOrdenPago.FindById(op.Id)
    
    Set OrdenPago.Compensatorios = DAOCompensatorios.FindByOP(OrdenPago.Id)
    
    Me.caption = "Orden de Pago Nro " & OrdenPago.Id

    Dim i As Long
    Dim j As Long
    With OrdenPago

        If .EsParaFacturaProveedor Then
            radioFacturaProveedor.value = True

            If .FacturasProveedor.count > 0 Then

                Me.cboProveedores.ListIndex = funciones.PosIndexCbo(.FacturasProveedor.item(1).Proveedor.Id, Me.cboProveedores)

                If Me.cboProveedores.ListIndex = -1 Then    'el proveedor no esta en la lista porque no tiene mas facturas sin saldar
                    Me.cboProveedores.AddItem .FacturasProveedor.item(1).Proveedor.RazonSocial
                    Me.cboProveedores.ItemData(Me.cboProveedores.NewIndex) = .FacturasProveedor.item(1).Proveedor.Id
                    colProveedores.Add .FacturasProveedor.item(1).Proveedor, CStr(.FacturasProveedor.item(1).Proveedor.Id)
                    Me.cboProveedores.ListIndex = funciones.PosIndexCbo(.FacturasProveedor.item(1).Proveedor.Id, Me.cboProveedores)
                End If

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
            Me.txtRetenciones.Text = .alicuota

        Else
            Me.radioConcepto.value = True

            If IsSomething(.CuentaContable) Then
                Me.cboCuentas.ListIndex = funciones.PosIndexCbo(.CuentaContable.Id, Me.cboCuentas)
                Me.txtDetalle.Text = .CuentaContableDescripcion
            Else
                Me.cboCuentas.ListIndex = -1
                Me.txtDetalle.Text = vbNullString
            End If

        End If


        If idx >= 0 Then
            lstFacturas.ListIndex = lstFacturas.ListCount - 1

        End If

        Me.gridCajaOperaciones.ItemCount = .operacionesCaja.count
        Me.gridDepositosOperaciones.ItemCount = .operacionesBanco.count
        Me.gridCheques.ItemCount = .ChequesTerceros.count
        Me.gridChequesPropios.ItemCount = .ChequesPropios.count

        Me.gridRetenciones.ItemCount = .RetencionesAlicuota.count
        Set alicuotas = .RetencionesAlicuota


        Me.cboMonedas.ListIndex = funciones.PosIndexCbo(.moneda.Id, Me.cboMonedas)
        Me.dtpFecha.value = .FEcha
        Me.txtDifCambio.Text = .DiferenciaCambio
        Me.txtOtrosDescuentos.Text = .OtrosDescuentos

    End With
    mostrarCompensatorios

    'Me.grpDestino.Enabled = Not ReadOnly
    Me.txtDifCambioNG1.Enabled = Not ReadOnly
    Me.txtDifCambioTOTAL1.Enabled = Not ReadOnly
    Me.cmdMostrarDatosProveedor.Enabled = Not ReadOnly
    Me.btnPadronAnt.Enabled = Not ReadOnly
    Me.btnCargar.Enabled = Not ReadOnly

    Me.gridRetenciones.AllowEdit = Not ReadOnly

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
    Me.txtRetenciones.Text = 0
    Me.lstFacturas.Clear
    Set prov = Nothing

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
    Me.txtRetenciones.Text = 0
    Me.lstFacturas.Clear
    Set prov = Nothing
    
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
            If Left(tipoComprobante, 2) = "NC" Then
                valorTotal = valorTotal * -1
                valorAbonado = valorAbonado * -1
            End If
            ' Si el comprobante comienza con "NC", convertir los valores a negativos
            If Left(tipoComprobante, 2) = "NC" Then
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
    If IsSomething(OrdenPago) Then
        If Not DAOOrdenPago.ExportarOrdenPago(OrdenPago) Then GoTo err1
    End If

    Exit Sub
err1:
    MsgBox "Se produjo un error al exportar!", vbCritical, "Error"
    
End Sub


Private Sub btnGuardar_Click()
    If Me.gridChequesPropios.EditMode = jgexEditModeOn Then
        MsgBox "Todavia esta editando la grilla de cheques propios.", vbExclamation
        Exit Sub
    End If

    If Me.gridCheques.EditMode = jgexEditModeOn Then
        MsgBox "Todavia esta editando la grilla de cheques de 3ros.", vbExclamation
        Exit Sub
    End If

    If Me.gridCajaOperaciones.EditMode = jgexEditModeOn Then
        MsgBox "Todavia esta editando la grilla de caja.", vbExclamation
        Exit Sub
    End If

    If Me.gridDepositosOperaciones.EditMode = jgexEditModeOn Then
        MsgBox "Todavia esta editando la grilla de banco.", vbExclamation
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

    For i = 0 To Me.ListPagosACuenta.ListCount - 1
        If Me.ListPagosACuenta.Checked(i) Then
            OrdenPago.pagosacuenta.Add colPagosACuenta.item(CStr(Me.ListPagosACuenta.ItemData(i)))
            End If
    Next i

    If IsNumeric(Me.txtRetenciones) Then OrdenPago.alicuota = Val(Me.txtRetenciones)


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


Private Sub btnMoneda_Click()

''' If colAlicuotas.count > 0 Then
'''    vFactura.IvaAplicado = Nothing
'''    Me.grilla_alicuotas.ItemCount = 0
'''    Me.grilla_alicuotas.Refresh
'''    AddDefaultAlicuota colAlicuotas(1).Id
'''End If

'''Private Sub AddDefaultAlicuota(id_alicuota As Long)
'''    Set aliaplicada = New clsAlicuotaAplicada
'''    aliaplicada.Monto = 0
'''    aliaplicada.alicuota = DAOAlicuotas.GetById(id_alicuota)
'''    vFactura.IvaAplicado.Add aliaplicada
'''    mostrarALicuotas
'''End Sub


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


Private Sub cboMonedas_Click()
    If Me.cboMonedas.ListIndex = -1 Then
        Set OrdenPago.moneda = Nothing
    Else
        Set OrdenPago.moneda = DAOMoneda.GetById(Me.cboMonedas.ItemData(Me.cboMonedas.ListIndex))
    End If
    Totalizar
End Sub


Private Sub cboProveedores_Click()
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
    OrdenPago.FEcha = Me.dtpFecha.value
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

    Me.gridCajaOperaciones.ItemCount = OrdenPago.operacionesCaja.count
    Me.gridDepositosOperaciones.ItemCount = OrdenPago.operacionesBanco.count
    Me.gridCheques.ItemCount = OrdenPago.ChequesTerceros.count
    Me.gridChequesPropios.ItemCount = OrdenPago.ChequesPropios.count


    Set Me.gridCheques.Columns("numero").DropDownControl = Me.gridChequesDisponibles

    Set Me.gridDepositosOperaciones.Columns("moneda").DropDownControl = Me.gridMonedas
   
    Set Me.gridDepositosOperaciones.Columns("cuenta").DropDownControl = Me.gridCuentasBancarias

    Set Me.gridCajaOperaciones.Columns("monedas").DropDownControl = Me.gridMonedas
    
    Set Me.gridCajaOperaciones.Columns("caja").DropDownControl = Me.gridCajas

    Set Me.gridChequesPropios.Columns("chequera").DropDownControl = Me.gridChequeras
    
    Set Me.gridChequesPropios.Columns("numero").DropDownControl = Me.gridChequesChequera
    
'''    cargarCamposPredefinidos

    gridChequesChequera.ItemCount = 0
    
    GridEXHelper.AutoSizeColumns Me.gridChequeras

    DAOMoneda.llenarComboXtremeSuite Me.cboMonedas

    Me.dtpFecha.value = OrdenPago.FEcha

    'lstFacturas_Click
    Totalizar

    formLoaded = True
    formLoading = False

End Sub


'''Private Sub cargarCamposPredefinidos()
'''
'''   Set monedaplicada = New clsMonedaAplicada
'''   Set colMonedas = New Collection
'''
'''   If Monedas.count > 0 Then
'''
'''       Me.gridCajaOperaciones.ItemCount = 0
'''       Me.gridCajaOperaciones.Refresh
'''
'''       monedaplicada.moneda = DAOMoneda.GetById(Monedas(1).Id)
'''
'''       colMonedas.Add monedaplicada.moneda
'''
'''       Me.gridCajaOperaciones.ItemCount = 0
'''       Me.gridCajaOperaciones.ItemCount = colMonedas.count
'''
'''    End If
'''
'''End Sub


Private Sub CargarChequesDisponibles()
    Set chequesDisponibles = DAOCheques.FindAllEnCarteraDeTerceros
    Me.gridChequesDisponibles.ItemCount = chequesDisponibles.count
End Sub


Private Sub MostrarDeudaCompensatorios()
    Me.lstDeudaCompensatorios.Clear
    If IsSomething(prov) Then
        Set colDeudaCompensatorios = DAOCompensatorios.FindAllPendientesByProveedor(prov.Id)  'DAOFacturaProveedor.FindAll("AdminComprasFacturasProveedores.id_proveedor=" & prov.id & " and (AdminComprasFacturasProveedores.estado=" & EstadoFacturaProveedor.pagoParcial & " or  AdminComprasFacturasProveedores.estado=" & EstadoFacturaProveedor.Aprobada & ")", False, "", False, True)

        Dim C As Compensatorio

        For Each C In colDeudaCompensatorios
            Me.lstDeudaCompensatorios.AddItem "Cód: " & C.Id & " (OP: " & C.IdOrdenPago & ", Cbte: " & C.Comprobante.NumeroFormateado & ", Importe: " & C.Monto & ")"
            Me.lstDeudaCompensatorios.ItemData(Me.lstDeudaCompensatorios.NewIndex) = C.Id
        Next
    Else
        Set colFacturas = New Collection
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

            Dim C As Collection
            Set C = DAOOrdenPago.FindAbonadoPendiente(Factura.Id, OrdenPago.Id)

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

    If IsSomething(prov) Then
        Dim filtro As String
        
        filtro = "pagos_a_cuenta.estado = 0 AND pagos_a_cuenta.id_proveedor=" & prov.Id
        
        Set colPagosACuenta = DAOPagoACta.FindAll(filtro)


        Dim T As String

        For Each PagoACta In colPagosACuenta
            Dim C As Collection

            T = "N°: " & PagoACta.Id & " ( " & PagoACta.moneda.NombreCorto & " " & Replace(FormatCurrency(funciones.FormatearDecimales(PagoACta.StaticTotalOrigenes)), "$", "") & ")"
            
            Me.ListPagosACuenta.AddItem T
            Me.ListPagosACuenta.ItemData(Me.ListPagosACuenta.NewIndex) = PagoACta.Id


        Next

    Else

        Set colPagosACuenta = New Collection

        'MsgBox (colFacturas.count)

    End If

End Sub


Private Sub Form_Unload(Cancel As Integer)
    Channel.RemoverSuscripcionTotal Me
    
End Sub


Private Sub gridBancos_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If rowIndex <= bancos.count Then
        Set Banco = bancos.item(rowIndex)
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


Private Sub gridCajas_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If rowIndex > 0 And Cajas.count > 0 Then
        Set caja = Cajas.item(rowIndex)
        Values(1) = caja.Id
        Values(2) = caja.nombre
    End If
End Sub


Private Sub gridChequeras_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If rowIndex <= chequeras.count Then
        Set tmpChequera = chequeras.item(rowIndex)
        Values(1) = tmpChequera.Description
        Values(2) = tmpChequera.Id
    End If
End Sub


Private Sub gridChequesChequera_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If rowIndex > 0 And chequesChequeraSeleccionada.count > 0 Then
        Values(1) = chequesChequeraSeleccionada(rowIndex).numero
        Values(2) = chequesChequeraSeleccionada(rowIndex).Id
    End If
End Sub


Private Sub gridChequesDisponibles_ColumnHeaderClick(ByVal Column As GridEX20.JSColumn)
    GridEXHelper.ColumnHeaderClick Me.gridChequesDisponibles, Column
End Sub


Private Sub gridChequesDisponibles_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If rowIndex <= chequesDisponibles.count Then
        Set cheque = chequesDisponibles.item(rowIndex)
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

    If LenB(Me.gridChequesPropios.value(1)) = 0 Then
        msg.Add "Debe especificar una chequera."
    End If

    If LenB(Me.gridChequesPropios.value(2)) = 0 Then
        msg.Add "Debe especificar un cheque."
    End If

    ' REVISA QUE EN LA COLECCION DE CHEQUES PROPIOS QUE SE ESTAN CARGANDO NO EST? INGRESADO EL MISMO CHEQUE, SI LO DETECTA GENERA MSG DE ERROR
    If funciones.BuscarEnColeccion(OrdenPago.ChequesPropios, CStr(Me.gridChequesPropios.value(2))) Then
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

        OrdenPago.ChequesPropios.Add cheque, CStr(cheque.Id)

    End If
    Totalizar
End Sub


Private Sub gridChequesPropios_UnboundDelete(ByVal rowIndex As Long, ByVal Bookmark As Variant)
    If rowIndex > 0 Then
        OrdenPago.ChequesPropios.remove rowIndex
        Totalizar
    End If
End Sub


Private Sub gridChequesPropios_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If OrdenPago.ChequesPropios.count >= rowIndex Then
        Set cheque = OrdenPago.ChequesPropios.item(rowIndex)
        Values(1) = cheque.chequera.Description
        Values(2) = vbNullString
        'FORMATCURRENCY
        Values(3) = FormatCurrency(cheque.Monto)
        Values(4) = cheque.FechaVencimiento
        Values(5) = cheque.numero


        Totalizar
    End If
End Sub


Private Sub gridChequesPropios_UnboundUpdate(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If OrdenPago.ChequesPropios.count >= rowIndex Then
        Set cheque = OrdenPago.ChequesPropios.item(rowIndex)

        '        If Values(2) <> Cheque.Id Then
        '            ordenPago.ChequesPropios.remove CStr(Cheque.Id)
        '            Set Cheque = DAOCheques.FindById(Values(2))
        '            ordenPago.ChequesPropios.Add Cheque, CStr(Cheque.Id)
        '        End If

        cheque.Monto = Values(3)
        cheque.FechaVencimiento = Values(4)
    End If

    Totalizar
    
End Sub


Private Sub gridCompensatorios_UnboundDelete(ByVal rowIndex As Long, ByVal Bookmark As Variant)
    OrdenPago.Compensatorios.remove (rowIndex)
End Sub


Private Sub gridCompensatorios_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    On Error Resume Next
    Set compe = OrdenPago.Compensatorios.item(rowIndex)
    Values(1) = compe.Comprobante.NumeroFormateado
    Values(2) = TiposCompensatorio.item(CStr(compe.Tipo))
    'FORMATCURRENCY
    Values(3) = FormatCurrency(compe.Monto)
    Values(4) = compe.FechaCancelacion
    Values(5) = compe.Observacion

End Sub


Private Sub gridCuentasBancarias_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If cuentasBancarias.count >= rowIndex Then
        Set CuentaBancaria = cuentasBancarias.item(rowIndex)
        Values(1) = CuentaBancaria.Id
        Values(2) = CuentaBancaria.DescripcionFormateada
    End If
End Sub


Private Sub gridMonedas_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If rowIndex > 0 And Monedas.count > 0 Then
        Set moneda = Monedas.item(rowIndex)
        Values(1) = moneda.Id
        Values(2) = moneda.NombreCorto
    End If
End Sub


Private Sub gridRetenciones_RowFormat(RowBuffer As GridEX20.JSRowData)

    On Error GoTo err1

    Set alicuotaRetencion = alicuotas.item(RowBuffer.rowIndex)

    If alicuotaRetencion.importe > 0 Then    '.Retencion.id <> 2 Then
        RowBuffer.RowStyle = "padronganancias"
    Else
        RowBuffer.RowStyle = "padroningresos"

    End If

    Exit Sub

err1:

End Sub


Private Sub gridRetenciones_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If alicuotas.count >= rowIndex Then
        Set alicuotaRetencion = alicuotas.item(rowIndex)
        Values(2) = alicuotaRetencion.alicuotaRetencion
        Values(1) = alicuotaRetencion.Retencion.nombre
        Values(3) = alicuotaRetencion.importe
        Values(4) = alicuotaRetencion.certificados
    End If
End Sub


Private Sub gridRetenciones_UnboundUpdate(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If alicuotas.count >= rowIndex Then
        Set alicuotaRetencion = alicuotas.item(rowIndex)
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


Private Sub MostrarPosiblesRetenciones(col As Collection, Optional colc As Collection = Nothing, Optional colpcta As Collection = Nothing)
    Dim d As New Dictionary
    Dim ret As Retencion
    Dim colret As Collection
    Set colret = DAORetenciones.FindAllEsAgente
    Set d = DAOCertificadoRetencion.VerPosibleRetenciones2(col, alicuotas, Val(Me.txtDifCambioNG1), OrdenPago.TotalNGCompensatorios)
    Dim totRet As Double

    totRet = 0

    If IsSomething(prov) Then
        
        For Each ret In colret
            totRet = totRet + d.item(CStr(ret.Id))
        Next ret
    End If

    totRet = funciones.RedondearDecimales(totRet)
    Dim C As Compensatorio
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
    
    totDeudaCompe = 0
    totFactNuevo = 0
    
    For Each F In col
        
    ' Inicializar la variable
    
    ' Recorrer la lista de facturas

        ' Verificar si es Nota de Crédito
        
    If OrdenPago.estado = EstadoOrdenPago_pendiente Then
            If F.tipoDocumentoContable = tipoDocumentoContable.notaCredito Then
                ' Restar el total (convertir a negativo)
                totFactNuevo = totFactNuevo - (F.total - (F.TotalAbonadoGlobal + F.TotalAbonadoGlobalPendiente))
            Else
                ' Sumar el total normalmente
                totFactNuevo = totFactNuevo + (F.total - (F.TotalAbonadoGlobal + F.TotalAbonadoGlobalPendiente))
            End If
    Else
            If F.tipoDocumentoContable = tipoDocumentoContable.notaCredito Then
            ' Restar el total (convertir a negativo)
            totFactNuevo = totFactNuevo - (F.total)
        Else
            ' Sumar el total normalmente
            totFactNuevo = totFactNuevo + (F.total)
        End If
    End If
    

        'totNGHoy = totNGHoy + MonedaConverter.ConvertirForzado2(IIf(f.tipoDocumentoContable = tipoDocumentoContable.notaCredito, f.NetoGravadoDiaPago * -1, f.NetoGravadoDiaPago), f.Moneda.Id, OrdenPago.Moneda.Id, f.TipoCambioPago)
        ' totFact = totFact + MonedaConverter.ConvertirForzado2(IIf(f.tipoDocumentoContable = tipoDocumentoContable.notaCredito, f.total * -1, f.total), f.Moneda.Id, OrdenPago.Moneda.Id, f.TipoCambioPago) cambiado el 22-9-14 por tema de pagos parciales
        'totFactHoy = totFactHoy + MonedaConverter.ConvertirForzado2(IIf(f.tipoDocumentoContable = tipoDocumentoContable.notaCredito, f.TotalDiaPago * -1, f.TotalDiaPago), f.Moneda.Id, OrdenPago.Moneda.Id, f.TipoCambioPago)
        'totNG = TotNG + MonedaConverter.ConvertirForzado2(IIf(f.tipoDocumentoContable = tipoDocumentoContable.notaCredito, f.NetoGravado * -1, f.NetoGravado), f.Moneda.Id, OrdenPago.Moneda.Id, f.TipoCambioPago)
        'totFact = totFact + MonedaConverter.ConvertirForzado2(IIf(F.tipoDocumentoContable = tipoDocumentoContable.notaCredito, F.ImporteTotalAbonado * -1, F.ImporteTotalAbonado), F.moneda.id, OrdenPago.moneda.id, F.TipoCambioPago)
        'fix 004
        
        totFact = totFact + MonedaConverter.ConvertirForzado2(IIf(F.tipoDocumentoContable = tipoDocumentoContable.notaCredito, F.totalAbonado * -1, F.totalAbonado), F.moneda.Id, OrdenPago.moneda.Id, F.TipoCambioPago)

        totFactHoy = totFactHoy + MonedaConverter.ConvertirForzado2(IIf(F.tipoDocumentoContable = tipoDocumentoContable.notaCredito, F.TotalDiaPagoAbonado * -1, F.TotalDiaPagoAbonado), F.moneda.Id, OrdenPago.moneda.Id, F.TipoCambioPago)

        TotNG = TotNG + MonedaConverter.ConvertirForzado2(IIf(F.tipoDocumentoContable = tipoDocumentoContable.notaCredito, F.NetoGravadoAbonado * -1, F.NetoGravadoAbonado), F.moneda.Id, OrdenPago.moneda.Id, F.TipoCambioPago)
        totNGHoy = totNGHoy + MonedaConverter.ConvertirForzado2(IIf(F.tipoDocumentoContable = tipoDocumentoContable.notaCredito, F.NetoGravadoAbonadoDiaPago * -1, F.NetoGravadoAbonadoDiaPago), F.moneda.Id, OrdenPago.moneda.Id, F.TipoCambioPago)
        totCambio = totCambio + MonedaConverter.ConvertirForzado2(IIf(F.tipoDocumentoContable = tipoDocumentoContable.notaCredito, F.DiferenciaPorTipoDeCambionTOTAL * -1, F.DiferenciaPorTipoDeCambionTOTAL), F.moneda.Id, OrdenPago.moneda.Id, F.TipoCambioPago)
        totCambiong = totCambiong + MonedaConverter.ConvertirForzado2(IIf(F.tipoDocumentoContable = tipoDocumentoContable.notaCredito, F.DiferenciaPorTipoDeCambionNG * -1, F.DiferenciaPorTipoDeCambionNG), F.moneda.Id, OrdenPago.moneda.Id, F.TipoCambioPago)
    
    Next F


    If IsSomething(colc) Then
        For Each C In colc

            Dim ff As clsFacturaProveedor

            Set ff = DAOFacturaProveedor.FindById(C.Comprobante.Id)
            totDeudaCompe = totDeudaCompe + MonedaConverter.ConvertirForzado2(IIf(C.Tipo = TC_Credito, C.Monto * -1, C.Monto), ff.moneda.Id, OrdenPago.moneda.Id, ff.TipoCambioPago)

        Next
    End If
    
    
   If IsSomething(colpcta) Then
        For Each P In colpcta

            totPagoACuenta = totPagoACuenta + P.StaticTotalOrigenes

        Next
    End If

    Me.lblNgAbonar = "Total NG a Abonar en " & FormatCurrency(funciones.FormatearDecimales(OrdenPago.DiferenciaCambioEnNG + totNGHoy))

'''    Me.lblTotalFacturas = "Total Facturas en " & FormatCurrency(funciones.FormatearDecimales(totFact))

    Me.lblTotalFacturas = "Total Facturas en " & FormatCurrency(funciones.FormatearDecimales(totFactNuevo))
    
    Me.lblDeudaCompensatorios = "Total deuda compensatorios en " & FormatCurrency(funciones.FormatearDecimales(totDeudaCompe))

    OrdenPago.StaticTotalFacturas = funciones.RedondearDecimales(totFact)
    
    OrdenPago.staticTotalDeudaCompensatorios = funciones.RedondearDecimales(totDeudaCompe)

    Me.lblTotalFacturasNG = "Total NG Facturas en " & FormatCurrency(funciones.FormatearDecimales(TotNG + OrdenPago.DiferenciaCambioEnNG))

    OrdenPago.StaticTotalFacturasNG = funciones.RedondearDecimales(TotNG + OrdenPago.DiferenciaCambioEnNG)

    Me.lblDiferenciaCambio = "Diferencia Cambio en " & FormatCurrency(totCambiong)

    OrdenPago.DiferenciaCambio = totCambio

    verCompensatorios

    Me.lblTotalARetener = "Total a retener en " & FormatCurrency(funciones.FormatearDecimales(totRet))

    OrdenPago.StaticTotalRetenido = funciones.RedondearDecimales(totRet)

'''    Me.lblTotalOrdenPago = "Total a abonar en " & FormatCurrency(funciones.FormatearDecimales((OrdenPago.DiferenciaCambioEnTOTAL + totFactHoy - (totRet - OrdenPago.OtrosDescuentos) + OrdenPago.TotalCompensatorios + totDeudaCompe)) - totPagoACuenta)
    Me.lblTotalOrdenPago = "Total a abonar en " & FormatCurrency(funciones.FormatearDecimales(totFactNuevo - totRet))
        
    Me.lblTotalPagoACuenta.caption = "Total Pago a Cuenta en " & FormatCurrency(funciones.FormatearDecimales(totPagoACuenta))
    
    Me.lblFacturasTotal.caption = FormatCurrency(funciones.FormatearDecimales(totFactNuevo))
    
End Sub


Private Sub verCompensatorios()
    Me.lblTotalCompensatorios = "Total compensatorios en " & FormatCurrency(funciones.FormatearDecimales(OrdenPago.TotalCompensatorios))

End Sub


Private Sub MostrarPago(F As clsFacturaProveedor)

    If IsSomething(F) Then

        Me.txtTotalParcialAbonado = F.TotalAbonadoGlobal
        Me.txtOtrosParcialAbonado = F.OtrosAbonadoGlobal + F.OtrosAbonadoGlobalPendiente
        Me.txtParcialAbonado = F.NetoGravadoAbonadoGlobal + F.NetoGravadoAbonadoGlobalPendiente


        ' If F.ImporteTotalAbonado = 0 Then F.ImporteTotalAbonado = F.Total
        If F.NetoGravadoAbonado = 0 Then F.NetoGravadoAbonado = F.NetoGravado    '- F.NetoNoGravado  (2do cambio en fix 004)
        If F.OtrosAbonado = 0 Then F.OtrosAbonado = F.total - F.NetoGravado    '- F.NetoNoGravado  (2do cambio en fix 004)

        Me.txtParcialAbonar = F.ImporteNetoGravadoSaldo    ' F.NetoGravadoAbonado - F.NetoGravadoAbonadoGlobal
        Me.txtTotalParcialAbonar = F.ImporteTotalAbonado
        Me.txtOtrosParcialAbonar = F.ImporteOtrosSaldo  'F.OtrosAbonado - F.OtrosAbonadoGlobal

        RecalcularTotalFacturaElegida
        
        'esto deber?a calcular el total en base a las al?cuotas de la factura

        If F.totalAbonado + F.TotalAbonadoGlobal + F.TotalAbonadoGlobalPendiente > F.total Then
            MsgBox "El importe que desea abonar, supera el monto total del comprobante seleccionado"
        End If
        'Me.txtnetogravadoabonado = F.NetoGravadoAbonado - F.NetoGravadoAbonadoGlobal
        ' Me.txtParcialAbonado = F.TotalAbonado - F.TotalAbonadoGlobal
    End If
    Totalizar
End Sub


Private Sub MotrarHistorialPagos(F As clsFacturaProveedor)

    If IsSomething(F) Then

    Me.GroupBox5.caption = "Detalle de comprobante: " & F.tipoDocumentoContable & " " & F.NumeroFormateado & " (ID: " & F.Id & ") "
    
    
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

'    'debug.print (Me.lstFacturas.ItemData(Me.lstFacturas.ListIndex))

    Set vFactElegida = colFacturas.item(CStr(Me.lstFacturas.ItemData(Me.lstFacturas.ListIndex)))

    If IsSomething(vFactElegida) Then

        Dim C As Collection

        If OrdenPago.estado = EstadoOrdenPago_pendiente And vFactElegida.NetoGravadoAbonado = 0 And vFactElegida.OtrosAbonado = 0 Then
            Set C = DAOOrdenPago.FindAbonadoFactura(vFactElegida.Id, OrdenPago.Id)

            vFactElegida.NetoGravadoAbonado = C(2)
            vFactElegida.OtrosAbonado = C(3)
        End If

        MostrarPago vFactElegida
        
        MotrarHistorialPagos vFactElegida
        
        RecalcularFacturaElegida
        
        
        
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
    Dim colpcta As New Collection

    For i = 0 To Me.lstFacturas.ListCount - 1
        If Me.lstFacturas.Checked(i) Then
            If funciones.BuscarEnColeccion(colFacturas, CStr(Me.lstFacturas.ItemData(i))) Then
                
                col.Add colFacturas.item(CStr(Me.lstFacturas.ItemData(i)))
                
                Me.lblCantidadCbtesSeleccionados.caption = "Cbtes. Seleccionados: " & col.count

            End If
        Else

            'si destildo tengo q ver q no existan compensatorios. Si existen deber?a primero eliminarlos.
            Dim ff As clsFacturaProveedor
            Dim C As Compensatorio
            For Each C In OrdenPago.Compensatorios
                Set ff = colFacturas.item(CStr(Me.lstFacturas.ItemData(i)))
                If C.Comprobante.Id = ff.Id Then
                    MsgBox "Existen compensatorios para este comprobante. Eliminelos primero!", vbCritical, "Error"
                    Me.lstFacturas.Checked(i) = True
                End If
            Next
        
        End If
    Next i
    
 
    For i = 0 To Me.lstDeudaCompensatorios.ListCount - 1
        If Me.lstDeudaCompensatorios.Checked(i) Then

            If funciones.BuscarEnColeccion(colDeudaCompensatorios, CStr(Me.lstDeudaCompensatorios.ItemData(i))) Then
                colc.Add colDeudaCompensatorios.item(CStr(Me.lstDeudaCompensatorios.ItemData(i)))

            End If

        End If
    Next i

    
   For i = 0 To Me.ListPagosACuenta.ListCount - 1
        If Me.ListPagosACuenta.Checked(i) Then

            If funciones.BuscarEnColeccion(colPagosACuenta, CStr(Me.ListPagosACuenta.ItemData(i))) Then
                colpcta.Add colPagosACuenta.item(CStr(Me.ListPagosACuenta.ItemData(i)))

            End If

        End If
    Next i


    TotalizarDiferenciasCambio
    MostrarPosiblesRetenciones col, colc, colpcta
    
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

        Me.txtParcialAbonado.Enabled = lstFacturas.Checked(item)
        Me.txtParcialAbonar.Enabled = lstFacturas.Checked(item)
        Me.txtOtrosParcialAbonado.Enabled = lstFacturas.Checked(item)
        Me.txtOtrosParcialAbonar.Enabled = lstFacturas.Checked(item)
        Me.txtTotalParcialAbonado.Enabled = lstFacturas.Checked(item)
        Me.txtTotalParcialAbonar.Enabled = lstFacturas.Checked(item)
    End If
End Sub


Private Sub lstFacturas_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Integer
    If Button = 2 Then

        For i = 0 To Me.lstFacturas.ListCount - 1

            If Me.lstFacturas.Selected(i) Then
                Me.mnuCrearCompensatorio.Enabled = Me.lstFacturas.Checked(i)
                PopupMenu Me.emergente
            End If
        Next
    End If
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

            Dim C As Compensatorio
            Dim hay As Boolean
            hay = False
            For Each C In OrdenPago.Compensatorios
                If C.Comprobante.Id = Factura.Id Then
                    hay = True
                    Exit For
                End If

            Next C

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
                    lstFacturas_ItemCheck 1
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
            Set Nueva = DAORetenciones.FindAllWithAlicuotas(prov.Cuit)    '
            
            Set alicuotas = DAORetenciones.FindAllWithAlicuotas(prov.Cuit)    '
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
    Me.cboProveedores.Enabled = Me.radioFacturaProveedor.value
    Me.lstFacturas.Enabled = Me.radioFacturaProveedor.value

    Me.cboCuentas.Enabled = Me.radioConcepto.value
    Me.txtDetalle.Enabled = Me.radioConcepto.value

    Me.txtRetenciones.Text = 0

    If Not Me.cboProveedores.Enabled Then Me.cboProveedores.ListIndex = -1
    If Not Me.lstFacturas.Enabled Then Me.lstFacturas.Clear

    If Not Me.cboCuentas.Enabled Then Me.cboCuentas.ListIndex = -1
    If Not Me.txtDetalle.Enabled Then Me.txtDetalle.Text = vbNullString

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
    OrdenPago.operacionesCaja.Add operacion
    Totalizar
End Sub


Private Sub gridCajaOperaciones_UnboundDelete(ByVal rowIndex As Long, ByVal Bookmark As Variant)
    If rowIndex > 0 And OrdenPago.operacionesCaja.count >= rowIndex Then
        OrdenPago.operacionesCaja.remove rowIndex
        Totalizar
    End If
End Sub


Private Sub Totalizar()
    OrdenPago.StaticTotalOrigenes = OrdenPago.TotalOrigenes

    Me.lblTotal.caption = "Total orden de pago en " & FormatCurrency(funciones.FormatearDecimales(OrdenPago.StaticTotalOrigenes + OrdenPago.StaticTotalRetenido))
    GridEXHelper.AutoSizeColumns Me.gridCajaOperaciones
    GridEXHelper.AutoSizeColumns Me.gridDepositosOperaciones
    GridEXHelper.AutoSizeColumns Me.gridCheques
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


Private Sub gridCajaOperaciones_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If rowIndex <= OrdenPago.operacionesCaja.count Then
        Set operacion = OrdenPago.operacionesCaja.item(rowIndex)
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


Private Sub gridCajaOperaciones_UnboundUpdate(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If rowIndex > 0 And OrdenPago.operacionesCaja.count > 0 Then
        Set operacion = OrdenPago.operacionesCaja.item(rowIndex)
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
    OrdenPago.operacionesBanco.Add operacion
    
    Totalizar
End Sub


Private Sub gridDepositosOperaciones_UnboundDelete(ByVal rowIndex As Long, ByVal Bookmark As Variant)
    If rowIndex > 0 And OrdenPago.operacionesBanco.count >= rowIndex Then
        OrdenPago.operacionesBanco.remove rowIndex
        Totalizar
        
    End If
End Sub


Private Sub gridDepositosOperaciones_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If rowIndex <= OrdenPago.operacionesBanco.count Then
        Set operacion = OrdenPago.operacionesBanco.item(rowIndex)
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


Private Sub gridDepositosOperaciones_UnboundUpdate(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If rowIndex > 0 And OrdenPago.operacionesBanco.count > 0 Then
        Set operacion = OrdenPago.operacionesBanco.item(rowIndex)
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


Private Sub gridCheques_UnboundDelete(ByVal rowIndex As Long, ByVal Bookmark As Variant)
    If rowIndex > 0 Then
        OrdenPago.ChequesTerceros.remove rowIndex
        Totalizar
    End If
End Sub


Private Sub gridCheques_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If rowIndex <= OrdenPago.ChequesTerceros.count Then
        Set cheque = OrdenPago.ChequesTerceros.item(rowIndex)

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


Private Sub gridCheques_UnboundUpdate(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If rowIndex > 0 And OrdenPago.ChequesTerceros.count >= rowIndex Then
        Set cheque = Nothing
        If IsNumeric(Values(1)) Then Set cheque = DAOCheques.FindById(Values(1))
        If IsSomething(cheque) Then
            OrdenPago.ChequesTerceros.Add cheque, , , rowIndex
            OrdenPago.ChequesTerceros.remove rowIndex
        End If
        Totalizar
    End If
End Sub


Private Sub txtBuscarFactura_GotFocus()
    Me.txtBuscarFactura.SelStart = 0
    Me.txtBuscarFactura.SelLength = Len(Me.txtBuscarFactura.Text)
End Sub


Private Sub txtBuscarFactura_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        'buscar en facturas y tildar

        If LenB(Me.txtBuscarFactura.Text) > 0 Then
            Dim cont As Long

            If colFacturas.count > 0 Then
                Dim i As Long
                For Each vFacturaProveedor In colFacturas
                    If InStr(1, vFacturaProveedor.numero, Me.txtBuscarFactura.Text) > 0 Then    'aplica
                        For i = 0 To Me.lstFacturas.ListCount - 1
                            If Me.lstFacturas.ItemData(i) = vFacturaProveedor.Id Then
                                Me.lstFacturas.Checked(i) = True
                                cont = cont + 1
                                Exit For
                            End If
                        Next i
                    End If
                Next vFacturaProveedor

                If cont = 0 Then
                    MsgBox "No se encontraron facturas con ese número en la lista.", vbOKOnly + vbExclamation
                Else
                    lstFacturas_ItemCheck -1
                    MsgBox "Se encontró " & cont & " factura/s.", vbOKOnly + vbInformation
                    Me.txtBuscarFactura.Text = vbNullString
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
    OrdenPago.DiferenciaCambioEnNG = Val(Me.txtDifCambioNG1)
    Totalizar
End Sub


Private Sub txtDifCambioTOTAL1_Change()
    OrdenPago.DiferenciaCambioEnTOTAL = Val(Me.txtDifCambioTOTAL1)
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
    OrdenPago.OtrosDescuentos = Val(Me.txtOtrosDescuentos.Text)
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


        vFactElegida.NetoGravadoAbonado = CDbl(txtParcialAbonar)
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

