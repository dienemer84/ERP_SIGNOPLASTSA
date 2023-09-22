VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmAdminPagosCrearOrdenPago 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Orden de Pago"
   ClientHeight    =   9960
   ClientLeft      =   2340
   ClientTop       =   3105
   ClientWidth     =   16980
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
   ScaleHeight     =   9960
   ScaleWidth      =   16980
   Begin XtremeSuiteControls.PushButton btnExportarDatos 
      Height          =   495
      Left            =   14880
      TabIndex        =   74
      Top             =   2760
      Visible         =   0   'False
      Width           =   1935
      _Version        =   786432
      _ExtentX        =   3413
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Exportar"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.GroupBox GroupBox3 
      Height          =   1335
      Left            =   120
      TabIndex        =   48
      Top             =   0
      Width           =   11100
      _Version        =   786432
      _ExtentX        =   19579
      _ExtentY        =   2355
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin VB.TextBox txtOtrosDescuentos 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   9960
         TabIndex        =   51
         Top             =   225
         Width           =   960
      End
      Begin VB.TextBox txtDifCambioNG1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   9960
         TabIndex        =   50
         Top             =   600
         Width           =   960
      End
      Begin VB.TextBox txtDifCambioTOTAL1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   9960
         TabIndex        =   49
         Top             =   945
         Width           =   960
      End
      Begin XtremeSuiteControls.ComboBox cboMonedas 
         Height          =   315
         Left            =   885
         TabIndex        =   55
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
         TabIndex        =   56
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
         TabIndex        =   58
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
         TabIndex        =   57
         Tag             =   "Total: "
         Top             =   300
         Width           =   570
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Otros Descuentos"
         Height          =   195
         Left            =   8520
         TabIndex        =   54
         Top             =   270
         Width           =   1275
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Dif. Cambio manual NG "
         Height          =   195
         Left            =   8160
         TabIndex        =   53
         Top             =   645
         Width           =   1680
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Dif. Cambio manual TOTAL"
         Height          =   195
         Left            =   7920
         TabIndex        =   52
         Top             =   990
         Width           =   1905
      End
   End
   Begin VB.TextBox txtnetogravadoabonado 
      Height          =   315
      Left            =   6240
      TabIndex        =   26
      Top             =   120
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.TextBox txtDifTipoCambioIVA 
      Height          =   285
      Left            =   6240
      TabIndex        =   25
      Top             =   840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtDiferenciaCambioPago 
      Height          =   285
      Left            =   4680
      TabIndex        =   24
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtDifCambio 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6240
      TabIndex        =   13
      Top             =   480
      Visible         =   0   'False
      Width           =   1200
   End
   Begin XtremeSuiteControls.GroupBox grpOrigen 
      Height          =   5775
      Left            =   7080
      TabIndex        =   0
      Top             =   4080
      Width           =   9780
      _Version        =   786432
      _ExtentX        =   17251
      _ExtentY        =   10186
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
         Height          =   5340
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   9540
         _Version        =   786432
         _ExtentX        =   16828
         _ExtentY        =   9419
         _StockProps     =   68
         Appearance      =   10
         Color           =   32
         PaintManager.ShowIcons=   -1  'True
         ItemCount       =   5
         SelectedItem    =   2
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
            Height          =   4710
            Left            =   -69895
            TabIndex        =   2
            Top             =   435
            Visible         =   0   'False
            Width           =   9330
            _ExtentX        =   16457
            _ExtentY        =   8308
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
            Column(1)       =   "frmAdminPagosCrearOrdenPago.frx":000C
            Column(2)       =   "frmAdminPagosCrearOrdenPago.frx":016C
            Column(3)       =   "frmAdminPagosCrearOrdenPago.frx":02A8
            Column(4)       =   "frmAdminPagosCrearOrdenPago.frx":03DC
            Column(5)       =   "frmAdminPagosCrearOrdenPago.frx":0520
            FormatStylesCount=   6
            FormatStyle(1)  =   "frmAdminPagosCrearOrdenPago.frx":0624
            FormatStyle(2)  =   "frmAdminPagosCrearOrdenPago.frx":075C
            FormatStyle(3)  =   "frmAdminPagosCrearOrdenPago.frx":080C
            FormatStyle(4)  =   "frmAdminPagosCrearOrdenPago.frx":08C0
            FormatStyle(5)  =   "frmAdminPagosCrearOrdenPago.frx":0998
            FormatStyle(6)  =   "frmAdminPagosCrearOrdenPago.frx":0A50
            ImageCount      =   0
            PrinterProperties=   "frmAdminPagosCrearOrdenPago.frx":0B30
         End
         Begin GridEX20.GridEX gridCajaOperaciones 
            Height          =   4710
            Left            =   -69895
            TabIndex        =   10
            Top             =   435
            Visible         =   0   'False
            Width           =   9330
            _ExtentX        =   16457
            _ExtentY        =   8308
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
            Column(1)       =   "frmAdminPagosCrearOrdenPago.frx":0D08
            Column(2)       =   "frmAdminPagosCrearOrdenPago.frx":0E68
            Column(3)       =   "frmAdminPagosCrearOrdenPago.frx":0FA4
            Column(4)       =   "frmAdminPagosCrearOrdenPago.frx":10D8
            Column(5)       =   "frmAdminPagosCrearOrdenPago.frx":120C
            FormatStylesCount=   6
            FormatStyle(1)  =   "frmAdminPagosCrearOrdenPago.frx":1310
            FormatStyle(2)  =   "frmAdminPagosCrearOrdenPago.frx":1448
            FormatStyle(3)  =   "frmAdminPagosCrearOrdenPago.frx":14F8
            FormatStyle(4)  =   "frmAdminPagosCrearOrdenPago.frx":15AC
            FormatStyle(5)  =   "frmAdminPagosCrearOrdenPago.frx":1684
            FormatStyle(6)  =   "frmAdminPagosCrearOrdenPago.frx":173C
            ImageCount      =   0
            PrinterProperties=   "frmAdminPagosCrearOrdenPago.frx":181C
         End
         Begin GridEX20.GridEX gridChequesPropios 
            Height          =   4710
            Left            =   -69895
            TabIndex        =   9
            Top             =   435
            Visible         =   0   'False
            Width           =   9330
            _ExtentX        =   16457
            _ExtentY        =   8308
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
            Column(1)       =   "frmAdminPagosCrearOrdenPago.frx":19F4
            Column(2)       =   "frmAdminPagosCrearOrdenPago.frx":1B5C
            Column(3)       =   "frmAdminPagosCrearOrdenPago.frx":1C90
            Column(4)       =   "frmAdminPagosCrearOrdenPago.frx":1DCC
            Column(5)       =   "frmAdminPagosCrearOrdenPago.frx":1F34
            FormatStylesCount=   6
            FormatStyle(1)  =   "frmAdminPagosCrearOrdenPago.frx":202C
            FormatStyle(2)  =   "frmAdminPagosCrearOrdenPago.frx":2164
            FormatStyle(3)  =   "frmAdminPagosCrearOrdenPago.frx":2214
            FormatStyle(4)  =   "frmAdminPagosCrearOrdenPago.frx":22C8
            FormatStyle(5)  =   "frmAdminPagosCrearOrdenPago.frx":23A0
            FormatStyle(6)  =   "frmAdminPagosCrearOrdenPago.frx":2458
            ImageCount      =   0
            PrinterProperties=   "frmAdminPagosCrearOrdenPago.frx":2538
         End
         Begin GridEX20.GridEX gridCheques 
            Height          =   4710
            Left            =   105
            TabIndex        =   8
            Top             =   435
            Width           =   9330
            _ExtentX        =   16457
            _ExtentY        =   8308
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
            Column(1)       =   "frmAdminPagosCrearOrdenPago.frx":2710
            Column(2)       =   "frmAdminPagosCrearOrdenPago.frx":2890
            Column(3)       =   "frmAdminPagosCrearOrdenPago.frx":2A30
            Column(4)       =   "frmAdminPagosCrearOrdenPago.frx":2B28
            Column(5)       =   "frmAdminPagosCrearOrdenPago.frx":2C64
            Column(6)       =   "frmAdminPagosCrearOrdenPago.frx":2D70
            Column(7)       =   "frmAdminPagosCrearOrdenPago.frx":2E40
            FormatStylesCount=   6
            FormatStyle(1)  =   "frmAdminPagosCrearOrdenPago.frx":2F2C
            FormatStyle(2)  =   "frmAdminPagosCrearOrdenPago.frx":3064
            FormatStyle(3)  =   "frmAdminPagosCrearOrdenPago.frx":3114
            FormatStyle(4)  =   "frmAdminPagosCrearOrdenPago.frx":31C8
            FormatStyle(5)  =   "frmAdminPagosCrearOrdenPago.frx":32A0
            FormatStyle(6)  =   "frmAdminPagosCrearOrdenPago.frx":3358
            ImageCount      =   0
            PrinterProperties=   "frmAdminPagosCrearOrdenPago.frx":3438
         End
         Begin GridEX20.GridEX gridCompensatorios 
            Height          =   4710
            Left            =   -69895
            TabIndex        =   16
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
            Column(1)       =   "frmAdminPagosCrearOrdenPago.frx":3610
            Column(2)       =   "frmAdminPagosCrearOrdenPago.frx":3758
            Column(3)       =   "frmAdminPagosCrearOrdenPago.frx":3864
            Column(4)       =   "frmAdminPagosCrearOrdenPago.frx":3950
            Column(5)       =   "frmAdminPagosCrearOrdenPago.frx":3A54
            FormatStylesCount=   6
            FormatStyle(1)  =   "frmAdminPagosCrearOrdenPago.frx":3B94
            FormatStyle(2)  =   "frmAdminPagosCrearOrdenPago.frx":3CCC
            FormatStyle(3)  =   "frmAdminPagosCrearOrdenPago.frx":3D7C
            FormatStyle(4)  =   "frmAdminPagosCrearOrdenPago.frx":3E30
            FormatStyle(5)  =   "frmAdminPagosCrearOrdenPago.frx":3F08
            FormatStyle(6)  =   "frmAdminPagosCrearOrdenPago.frx":3FC0
            ImageCount      =   0
            PrinterProperties=   "frmAdminPagosCrearOrdenPago.frx":40A0
         End
      End
   End
   Begin GridEX20.GridEX gridBancos 
      Height          =   1845
      Left            =   2520
      TabIndex        =   3
      Top             =   10080
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
      Column(1)       =   "frmAdminPagosCrearOrdenPago.frx":4278
      Column(2)       =   "frmAdminPagosCrearOrdenPago.frx":4378
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminPagosCrearOrdenPago.frx":4468
      FormatStyle(2)  =   "frmAdminPagosCrearOrdenPago.frx":45A0
      FormatStyle(3)  =   "frmAdminPagosCrearOrdenPago.frx":4650
      FormatStyle(4)  =   "frmAdminPagosCrearOrdenPago.frx":4704
      FormatStyle(5)  =   "frmAdminPagosCrearOrdenPago.frx":47DC
      FormatStyle(6)  =   "frmAdminPagosCrearOrdenPago.frx":4894
      ImageCount      =   0
      PrinterProperties=   "frmAdminPagosCrearOrdenPago.frx":4974
   End
   Begin GridEX20.GridEX gridCuentasBancarias 
      Height          =   1695
      Left            =   11400
      TabIndex        =   4
      Top             =   10680
      Visible         =   0   'False
      Width           =   4185
      _ExtentX        =   7382
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
      Column(1)       =   "frmAdminPagosCrearOrdenPago.frx":4B4C
      Column(2)       =   "frmAdminPagosCrearOrdenPago.frx":4C70
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminPagosCrearOrdenPago.frx":4D64
      FormatStyle(2)  =   "frmAdminPagosCrearOrdenPago.frx":4E9C
      FormatStyle(3)  =   "frmAdminPagosCrearOrdenPago.frx":4F4C
      FormatStyle(4)  =   "frmAdminPagosCrearOrdenPago.frx":5000
      FormatStyle(5)  =   "frmAdminPagosCrearOrdenPago.frx":50D8
      FormatStyle(6)  =   "frmAdminPagosCrearOrdenPago.frx":5190
      ImageCount      =   0
      PrinterProperties=   "frmAdminPagosCrearOrdenPago.frx":5270
   End
   Begin GridEX20.GridEX gridMonedas 
      Height          =   1695
      Left            =   240
      TabIndex        =   5
      Top             =   10200
      Visible         =   0   'False
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   2990
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
      Column(1)       =   "frmAdminPagosCrearOrdenPago.frx":5448
      Column(2)       =   "frmAdminPagosCrearOrdenPago.frx":556C
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminPagosCrearOrdenPago.frx":5660
      FormatStyle(2)  =   "frmAdminPagosCrearOrdenPago.frx":5798
      FormatStyle(3)  =   "frmAdminPagosCrearOrdenPago.frx":5848
      FormatStyle(4)  =   "frmAdminPagosCrearOrdenPago.frx":58FC
      FormatStyle(5)  =   "frmAdminPagosCrearOrdenPago.frx":59D4
      FormatStyle(6)  =   "frmAdminPagosCrearOrdenPago.frx":5A8C
      ImageCount      =   0
      PrinterProperties=   "frmAdminPagosCrearOrdenPago.frx":5B6C
   End
   Begin GridEX20.GridEX gridCajas 
      Height          =   1695
      Left            =   12960
      TabIndex        =   6
      Top             =   10320
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
      Column(1)       =   "frmAdminPagosCrearOrdenPago.frx":5D44
      Column(2)       =   "frmAdminPagosCrearOrdenPago.frx":5E44
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminPagosCrearOrdenPago.frx":5F30
      FormatStyle(2)  =   "frmAdminPagosCrearOrdenPago.frx":6068
      FormatStyle(3)  =   "frmAdminPagosCrearOrdenPago.frx":6118
      FormatStyle(4)  =   "frmAdminPagosCrearOrdenPago.frx":61CC
      FormatStyle(5)  =   "frmAdminPagosCrearOrdenPago.frx":62A4
      FormatStyle(6)  =   "frmAdminPagosCrearOrdenPago.frx":635C
      ImageCount      =   0
      PrinterProperties=   "frmAdminPagosCrearOrdenPago.frx":643C
   End
   Begin GridEX20.GridEX gridChequesDisponibles 
      Height          =   2640
      Left            =   6240
      TabIndex        =   7
      Top             =   9960
      Visible         =   0   'False
      Width           =   9555
      _ExtentX        =   16854
      _ExtentY        =   4657
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
      Column(1)       =   "frmAdminPagosCrearOrdenPago.frx":6614
      Column(2)       =   "frmAdminPagosCrearOrdenPago.frx":6794
      Column(3)       =   "frmAdminPagosCrearOrdenPago.frx":6934
      Column(4)       =   "frmAdminPagosCrearOrdenPago.frx":6A2C
      Column(5)       =   "frmAdminPagosCrearOrdenPago.frx":6B68
      Column(6)       =   "frmAdminPagosCrearOrdenPago.frx":6C74
      Column(7)       =   "frmAdminPagosCrearOrdenPago.frx":6D94
      Column(8)       =   "frmAdminPagosCrearOrdenPago.frx":6EA0
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminPagosCrearOrdenPago.frx":6F94
      FormatStyle(2)  =   "frmAdminPagosCrearOrdenPago.frx":70CC
      FormatStyle(3)  =   "frmAdminPagosCrearOrdenPago.frx":717C
      FormatStyle(4)  =   "frmAdminPagosCrearOrdenPago.frx":7230
      FormatStyle(5)  =   "frmAdminPagosCrearOrdenPago.frx":7308
      FormatStyle(6)  =   "frmAdminPagosCrearOrdenPago.frx":73C0
      ImageCount      =   0
      PrinterProperties=   "frmAdminPagosCrearOrdenPago.frx":74A0
   End
   Begin GridEX20.GridEX gridChequeras 
      Height          =   1815
      Left            =   7200
      TabIndex        =   11
      Top             =   10560
      Width           =   6435
      _ExtentX        =   11351
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
      Column(1)       =   "frmAdminPagosCrearOrdenPago.frx":7678
      Column(2)       =   "frmAdminPagosCrearOrdenPago.frx":7798
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminPagosCrearOrdenPago.frx":7898
      FormatStyle(2)  =   "frmAdminPagosCrearOrdenPago.frx":79D0
      FormatStyle(3)  =   "frmAdminPagosCrearOrdenPago.frx":7A80
      FormatStyle(4)  =   "frmAdminPagosCrearOrdenPago.frx":7B34
      FormatStyle(5)  =   "frmAdminPagosCrearOrdenPago.frx":7C0C
      FormatStyle(6)  =   "frmAdminPagosCrearOrdenPago.frx":7CC4
      ImageCount      =   0
      PrinterProperties=   "frmAdminPagosCrearOrdenPago.frx":7DA4
   End
   Begin GridEX20.GridEX gridChequesChequera 
      Height          =   1710
      Left            =   12240
      TabIndex        =   12
      Top             =   10200
      Width           =   1980
      _ExtentX        =   3493
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
      Column(1)       =   "frmAdminPagosCrearOrdenPago.frx":7F7C
      Column(2)       =   "frmAdminPagosCrearOrdenPago.frx":80AC
      SortKeysCount   =   1
      SortKey(1)      =   "frmAdminPagosCrearOrdenPago.frx":81AC
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminPagosCrearOrdenPago.frx":8214
      FormatStyle(2)  =   "frmAdminPagosCrearOrdenPago.frx":834C
      FormatStyle(3)  =   "frmAdminPagosCrearOrdenPago.frx":83FC
      FormatStyle(4)  =   "frmAdminPagosCrearOrdenPago.frx":84B0
      FormatStyle(5)  =   "frmAdminPagosCrearOrdenPago.frx":8588
      FormatStyle(6)  =   "frmAdminPagosCrearOrdenPago.frx":8640
      ImageCount      =   0
      PrinterProperties=   "frmAdminPagosCrearOrdenPago.frx":8720
   End
   Begin XtremeSuiteControls.PushButton btnGuardar 
      Height          =   525
      Left            =   14880
      TabIndex        =   14
      Top             =   3480
      Width           =   1950
      _Version        =   786432
      _ExtentX        =   3440
      _ExtentY        =   926
      _StockProps     =   79
      Caption         =   "Guardar"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.GroupBox grpDestino 
      Height          =   2655
      Left            =   120
      TabIndex        =   17
      Top             =   1320
      Width           =   5565
      _Version        =   786432
      _ExtentX        =   9816
      _ExtentY        =   4683
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
         TabIndex        =   28
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
         TabIndex        =   27
         Top             =   6840
         Width           =   1335
      End
      Begin XtremeSuiteControls.RadioButton radioFacturaProveedor 
         Height          =   210
         Left            =   120
         TabIndex        =   18
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
         TabIndex        =   19
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
         TabIndex        =   20
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
         TabIndex        =   23
         Top             =   480
         Width           =   270
         _Version        =   786432
         _ExtentX        =   476
         _ExtentY        =   617
         _StockProps     =   79
         Caption         =   "X"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtDetalle 
         Height          =   840
         Left            =   120
         TabIndex        =   22
         Top             =   1680
         Width           =   5175
         _Version        =   786432
         _ExtentX        =   9128
         _ExtentY        =   1482
         _StockProps     =   77
         BackColor       =   -2147483643
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   6
      End
      Begin XtremeSuiteControls.ComboBox cboCuentas 
         Height          =   315
         Left            =   120
         TabIndex        =   21
         Top             =   1200
         Width           =   3735
         _Version        =   786432
         _ExtentX        =   6588
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
      Left            =   5880
      TabIndex        =   29
      Top             =   1320
      Width           =   5340
      _Version        =   786432
      _ExtentX        =   9419
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
         Left            =   4320
         TabIndex        =   30
         Top             =   960
         Width           =   585
      End
      Begin GridEX20.GridEX gridRetenciones 
         Height          =   1335
         Left            =   240
         TabIndex        =   31
         Top             =   1200
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   2355
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
         ColumnsCount    =   3
         Column(1)       =   "frmAdminPagosCrearOrdenPago.frx":88F8
         Column(2)       =   "frmAdminPagosCrearOrdenPago.frx":8A34
         Column(3)       =   "frmAdminPagosCrearOrdenPago.frx":8B34
         FormatStylesCount=   8
         FormatStyle(1)  =   "frmAdminPagosCrearOrdenPago.frx":8C38
         FormatStyle(2)  =   "frmAdminPagosCrearOrdenPago.frx":8D60
         FormatStyle(3)  =   "frmAdminPagosCrearOrdenPago.frx":8E10
         FormatStyle(4)  =   "frmAdminPagosCrearOrdenPago.frx":8EC4
         FormatStyle(5)  =   "frmAdminPagosCrearOrdenPago.frx":8F9C
         FormatStyle(6)  =   "frmAdminPagosCrearOrdenPago.frx":9054
         FormatStyle(7)  =   "frmAdminPagosCrearOrdenPago.frx":9134
         FormatStyle(8)  =   "frmAdminPagosCrearOrdenPago.frx":91D0
         ImageCount      =   0
         PrinterProperties=   "frmAdminPagosCrearOrdenPago.frx":9270
      End
      Begin XtremeSuiteControls.PushButton btnCargar 
         Height          =   345
         Left            =   3000
         TabIndex        =   32
         Top             =   480
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
         Left            =   240
         TabIndex        =   33
         Top             =   480
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
         Left            =   840
         TabIndex        =   34
         Top             =   960
         Width           =   3300
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox2 
      Height          =   5775
      Left            =   120
      TabIndex        =   42
      Top             =   4080
      Width           =   6885
      _Version        =   786432
      _ExtentX        =   12144
      _ExtentY        =   10186
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
      Begin VB.TextBox txtOtrosParcialAbonar 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1800
         TabIndex        =   69
         Top             =   1680
         Width           =   1545
      End
      Begin VB.TextBox txtOtrosParcialAbonado 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   68
         Top             =   1080
         Width           =   1545
      End
      Begin VB.TextBox txtTotalParcialAbonado 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   66
         Top             =   1080
         Width           =   1545
      End
      Begin VB.TextBox txtTotalParcialAbonar 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   64
         Top             =   1680
         Width           =   1545
      End
      Begin VB.TextBox txtParcialAbonado 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   62
         Top             =   1080
         Width           =   1425
      End
      Begin VB.TextBox txtBuscarFactura 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         TabIndex        =   44
         Top             =   480
         Width           =   5010
      End
      Begin VB.TextBox txtParcialAbonar 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         TabIndex        =   43
         Top             =   1680
         Width           =   1425
      End
      Begin XtremeSuiteControls.ListBox lstFacturas 
         Height          =   3135
         Left            =   120
         TabIndex        =   45
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
         TabIndex        =   73
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
         TabIndex        =   72
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
         TabIndex        =   71
         Top             =   1440
         Width           =   1665
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Otros Parcial abonado:"
         Height          =   195
         Left            =   1800
         TabIndex        =   70
         Top             =   840
         Width           =   1650
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Total Parcial abonado:"
         Height          =   195
         Left            =   3600
         TabIndex        =   67
         Top             =   840
         Width           =   1605
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Total Parcial a abonar:"
         Height          =   195
         Left            =   3600
         TabIndex        =   65
         Top             =   1440
         Width           =   1620
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "NG Parcial abonado:"
         Height          =   195
         Left            =   120
         TabIndex        =   63
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Buscar factura en la lista:"
         Height          =   195
         Left            =   120
         TabIndex        =   47
         Top             =   240
         Width           =   1830
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "NG Parcial a abonar:"
         Height          =   195
         Left            =   120
         TabIndex        =   46
         Top             =   1440
         Width           =   1470
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox4 
      Height          =   1095
      Left            =   17400
      TabIndex        =   59
      Top             =   5640
      Width           =   5565
      _Version        =   786432
      _ExtentX        =   9816
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
         TabIndex        =   60
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
   Begin VB.Label lblDeudaCompensatorios 
      AutoSize        =   -1  'True
      Caption         =   "Total compensatorios pendientes:"
      Height          =   195
      Left            =   11400
      TabIndex        =   61
      Top             =   1920
      Width           =   2430
   End
   Begin VB.Label lblDiferenciaCambio 
      AutoSize        =   -1  'True
      Caption         =   "Diferencia Cambio:"
      Height          =   195
      Left            =   11400
      TabIndex        =   41
      Top             =   2880
      Width           =   1350
   End
   Begin VB.Label lblTotalFacturasNG 
      AutoSize        =   -1  'True
      Caption         =   "Total NG Facturas: "
      Height          =   195
      Left            =   11400
      TabIndex        =   40
      Top             =   1680
      Width           =   1395
   End
   Begin VB.Label lblTotalOrdenPago 
      AutoSize        =   -1  'True
      Caption         =   "Total a pagar:"
      Height          =   195
      Left            =   11400
      TabIndex        =   39
      Tag             =   "tot fac - tot ret"
      Top             =   720
      Width           =   1020
   End
   Begin VB.Label lblTotalFacturas 
      AutoSize        =   -1  'True
      Caption         =   "Total facturas: "
      Height          =   195
      Left            =   11400
      TabIndex        =   38
      Top             =   1440
      Width           =   1110
   End
   Begin VB.Label lblTotalARetener 
      AutoSize        =   -1  'True
      Caption         =   "Total a retener:"
      Height          =   195
      Left            =   11400
      TabIndex        =   37
      Top             =   480
      Width           =   1140
   End
   Begin VB.Label lblTotalCompensatorios 
      AutoSize        =   -1  'True
      Caption         =   "Total compensatorios: "
      Height          =   195
      Left            =   11400
      TabIndex        =   36
      Tag             =   "Total: "
      Top             =   3120
      Width           =   1635
   End
   Begin VB.Label lblNgAbonar 
      AutoSize        =   -1  'True
      Caption         =   "Neto gravado a abonar:"
      Height          =   195
      Left            =   11400
      TabIndex        =   35
      Top             =   2640
      Width           =   1740
   End
   Begin VB.Label lblTotal 
      AutoSize        =   -1  'True
      Caption         =   "Total Pagos:"
      Height          =   195
      Left            =   11400
      TabIndex        =   15
      Tag             =   "Total: "
      Top             =   120
      Width           =   900
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
            Me.txtRetenciones.text = .alicuota

        Else
            Me.radioConcepto.value = True

            If IsSomething(.CuentaContable) Then
                Me.cboCuentas.ListIndex = funciones.PosIndexCbo(.CuentaContable.Id, Me.cboCuentas)
                Me.txtDetalle.text = .CuentaContableDescripcion
            Else
                Me.cboCuentas.ListIndex = -1
                Me.txtDetalle.text = vbNullString
            End If

        End If


        If idx >= 0 Then
            lstFacturas.ListIndex = lstFacturas.ListCount - 1

        End If




        Me.gridCajaOperaciones.ItemCount = .OperacionesCaja.count
'                MsgBox (.OperacionesCaja.count)
        Me.gridDepositosOperaciones.ItemCount = .OperacionesBanco.count
'                MsgBox (.OperacionesBanco.count)
        Me.gridCheques.ItemCount = .ChequesTerceros.count
        Me.gridChequesPropios.ItemCount = .ChequesPropios.count

        Me.gridRetenciones.ItemCount = .RetencionesAlicuota.count
        Set alicuotas = .RetencionesAlicuota


        Me.cboMonedas.ListIndex = funciones.PosIndexCbo(.moneda.Id, Me.cboMonedas)
        Me.dtpFecha.value = .FEcha
        Me.txtDifCambio.text = .DiferenciaCambio
        Me.txtOtrosDescuentos.text = .OtrosDescuentos

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
    Me.txtRetenciones.text = 0
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
    Me.txtRetenciones.text = 0
    Me.lstFacturas.Clear
    Set prov = Nothing
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
        OrdenPago.CuentaContableDescripcion = Me.txtDetalle.text

    End If

    For i = 0 To Me.lstDeudaCompensatorios.ListCount - 1
        If Me.lstDeudaCompensatorios.Checked(i) Then
            OrdenPago.DeudaCompensatorios.Add colDeudaCompensatorios.item(CStr(Me.lstDeudaCompensatorios.ItemData(i)))
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

    Me.gridCajaOperaciones.ItemCount = OrdenPago.OperacionesCaja.count
'        MsgBox (OrdenPago.OperacionesCaja.count)
    Me.gridDepositosOperaciones.ItemCount = OrdenPago.OperacionesBanco.count
'        MsgBox (OrdenPago.OperacionesBanco.count)
    Me.gridCheques.ItemCount = OrdenPago.ChequesTerceros.count
    Me.gridChequesPropios.ItemCount = OrdenPago.ChequesPropios.count



    Set Me.gridCheques.Columns("numero").DropDownControl = Me.gridChequesDisponibles

    Set Me.gridDepositosOperaciones.Columns("moneda").DropDownControl = Me.gridMonedas
    Set Me.gridDepositosOperaciones.Columns("cuenta").DropDownControl = Me.gridCuentasBancarias

    Set Me.gridCajaOperaciones.Columns("moneda").DropDownControl = Me.gridMonedas
    Set Me.gridCajaOperaciones.Columns("caja").DropDownControl = Me.gridCajas

    Set Me.gridChequesPropios.Columns("chequera").DropDownControl = Me.gridChequeras
    Set Me.gridChequesPropios.Columns("numero").DropDownControl = Me.gridChequesChequera
    gridChequesChequera.ItemCount = 0
    GridEXHelper.AutoSizeColumns Me.gridChequeras


    DAOMoneda.llenarComboXtremeSuite Me.cboMonedas

    Me.dtpFecha.value = OrdenPago.FEcha

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

            Dim c As Collection
            Set c = DAOOrdenPago.FindAbonadoPendiente(Factura.Id, OrdenPago.Id)

            Factura.TotalAbonadoGlobalPendiente = 0    ' c(1) 'que esta en ops sin aprobar
            Factura.NetoGravadoAbonadoGlobalPendiente = 0    ' c(2)
            Factura.OtrosAbonadoGlobalPendiente = 0    'c(3)

            T = Factura.NumeroFormateado & " (" & Factura.moneda.NombreCorto & " " & Factura.total & ")" & " (" & Factura.FEcha & ")" & " TC: (" & Factura.TipoCambioPago & ")"
            If Factura.TotalAbonadoGlobal + Factura.TotalAbonadoGlobalPendiente > 0 Then
                T = Factura.NumeroFormateado & " (" & Factura.moneda.NombreCorto & " " & Factura.total & " - Abonado: " & Factura.TotalAbonadoGlobal + Factura.TotalAbonadoGlobalPendiente & ")" & " (" & Factura.FEcha & ")" & " TC: (" & Factura.TipoCambioPago & ")"

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
    End If
End Sub

Private Sub gridRetenciones_UnboundUpdate(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If alicuotas.count >= rowIndex Then
        Set alicuotaRetencion = alicuotas.item(rowIndex)
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
    Set d = DAOCertificadoRetencion.VerPosibleRetenciones2(col, alicuotas, Val(Me.txtDifCambioNG1), OrdenPago.TotalNGCompensatorios)
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


        'totNGHoy = totNGHoy + MonedaConverter.ConvertirForzado2(IIf(f.tipoDocumentoContable = tipoDocumentoContable.notaCredito, f.NetoGravadoDiaPago * -1, f.NetoGravadoDiaPago), f.Moneda.Id, OrdenPago.Moneda.Id, f.TipoCambioPago)
        ' totFact = totFact + MonedaConverter.ConvertirForzado2(IIf(f.tipoDocumentoContable = tipoDocumentoContable.notaCredito, f.total * -1, f.total), f.Moneda.Id, OrdenPago.Moneda.Id, f.TipoCambioPago) cambiado el 22-9-14 por tema de pagos parciales
        'totFactHoy = totFactHoy + MonedaConverter.ConvertirForzado2(IIf(f.tipoDocumentoContable = tipoDocumentoContable.notaCredito, f.TotalDiaPago * -1, f.TotalDiaPago), f.Moneda.Id, OrdenPago.Moneda.Id, f.TipoCambioPago)
        'totNG = TotNG + MonedaConverter.ConvertirForzado2(IIf(f.tipoDocumentoContable = tipoDocumentoContable.notaCredito, f.NetoGravado * -1, f.NetoGravado), f.Moneda.Id, OrdenPago.Moneda.Id, f.TipoCambioPago)
        'totFact = totFact + MonedaConverter.ConvertirForzado2(IIf(F.tipoDocumentoContable = tipoDocumentoContable.notaCredito, F.ImporteTotalAbonado * -1, F.ImporteTotalAbonado), F.moneda.id, OrdenPago.moneda.id, F.TipoCambioPago)
        'fix 004
        totFact = totFact + MonedaConverter.ConvertirForzado2(IIf(F.tipoDocumentoContable = tipoDocumentoContable.NotaCredito, F.TotalAbonado * -1, F.TotalAbonado), F.moneda.Id, OrdenPago.moneda.Id, F.TipoCambioPago)

        totFactHoy = totFactHoy + MonedaConverter.ConvertirForzado2(IIf(F.tipoDocumentoContable = tipoDocumentoContable.NotaCredito, F.TotalDiaPagoAbonado * -1, F.TotalDiaPagoAbonado), F.moneda.Id, OrdenPago.moneda.Id, F.TipoCambioPago)

        TotNG = TotNG + MonedaConverter.ConvertirForzado2(IIf(F.tipoDocumentoContable = tipoDocumentoContable.NotaCredito, F.NetoGravadoAbonado * -1, F.NetoGravadoAbonado), F.moneda.Id, OrdenPago.moneda.Id, F.TipoCambioPago)
        totNGHoy = totNGHoy + MonedaConverter.ConvertirForzado2(IIf(F.tipoDocumentoContable = tipoDocumentoContable.NotaCredito, F.NetoGravadoAbonadoDiaPago * -1, F.NetoGravadoAbonadoDiaPago), F.moneda.Id, OrdenPago.moneda.Id, F.TipoCambioPago)
        totCambio = totCambio + MonedaConverter.ConvertirForzado2(IIf(F.tipoDocumentoContable = tipoDocumentoContable.NotaCredito, F.DiferenciaPorTipoDeCambionTOTAL * -1, F.DiferenciaPorTipoDeCambionTOTAL), F.moneda.Id, OrdenPago.moneda.Id, F.TipoCambioPago)
        totCambiong = totCambiong + MonedaConverter.ConvertirForzado2(IIf(F.tipoDocumentoContable = tipoDocumentoContable.NotaCredito, F.DiferenciaPorTipoDeCambionNG * -1, F.DiferenciaPorTipoDeCambionNG), F.moneda.Id, OrdenPago.moneda.Id, F.TipoCambioPago)

    Next F


    If IsSomething(colc) Then
        For Each c In colc

            Dim ff As clsFacturaProveedor

            Set ff = DAOFacturaProveedor.FindById(c.Comprobante.Id)
            totDeudaCompe = totDeudaCompe + MonedaConverter.ConvertirForzado2(IIf(c.Tipo = TC_Credito, c.Monto * -1, c.Monto), ff.moneda.Id, OrdenPago.moneda.Id, ff.TipoCambioPago)

        Next
    End If

    'FORMATCURRENCY
    Me.lblNgAbonar = "Total NG a Abonar en " & FormatCurrency(funciones.FormatearDecimales(OrdenPago.DiferenciaCambioEnNG + totNGHoy))

    'FORMATCURRENCY
    'MsgBox ("Total Facturas en " & FormatCurrency(funciones.FormatearDecimales(totFact)))
    Me.lblTotalFacturas = "Total Facturas en " & FormatCurrency(funciones.FormatearDecimales(totFact))

    'FORMATCURRENCY
    Me.lblDeudaCompensatorios = "Total deuda compensatorios en " & FormatCurrency(funciones.FormatearDecimales(totDeudaCompe))

    OrdenPago.StaticTotalFacturas = funciones.RedondearDecimales(totFact)
    OrdenPago.staticTotalDeudaCompensatorios = funciones.RedondearDecimales(totDeudaCompe)

    'FORMATCURRENCY
    Me.lblTotalFacturasNG = "Total NG Facturas en " & FormatCurrency(funciones.FormatearDecimales(TotNG + OrdenPago.DiferenciaCambioEnNG))

    OrdenPago.StaticTotalFacturasNG = funciones.RedondearDecimales(TotNG + OrdenPago.DiferenciaCambioEnNG)

    'FORMATCURRENCY
    Me.lblDiferenciaCambio = "Diferencia Cambio en " & FormatCurrency(totCambiong)
    'Me.lblDiferenciaCambio = "Diferencia Cambio en " & OrdenPago.moneda.NombreCorto & " " & totCambiong

    OrdenPago.DiferenciaCambio = totCambio

    verCompensatorios

    'FORMATCURRENCY
    Me.lblTotalARetener = "Total a retener en " & FormatCurrency(funciones.FormatearDecimales(totRet))
    'Me.lblTotalARetener = "Total a retener en " & OrdenPago.moneda.NombreCorto & " " & funciones.FormatearDecimales(totRet)

    OrdenPago.StaticTotalRetenido = funciones.RedondearDecimales(totRet)

    'FORMATCURRENCY
    Me.lblTotalOrdenPago = "Total a abonar en " & FormatCurrency(funciones.FormatearDecimales(OrdenPago.DiferenciaCambioEnTOTAL + totFactHoy - totRet - OrdenPago.OtrosDescuentos + OrdenPago.TotalCompensatorios + totDeudaCompe))
    'Me.lblTotalOP = "Total OP: " & OrdenPago.moneda.NombreCorto & " " & OrdenPago.StaticTotal
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

        If F.TotalAbonado + F.TotalAbonadoGlobal + F.TotalAbonadoGlobalPendiente > F.total Then
            MsgBox "El importe que desea abonar, supera el monto total del comprobante seleccionado"
        End If
        'Me.txtnetogravadoabonado = F.NetoGravadoAbonado - F.NetoGravadoAbonadoGlobal
        ' Me.txtParcialAbonado = F.TotalAbonado - F.TotalAbonadoGlobal
    End If
    Totalizar
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

        Dim c As Collection

        If OrdenPago.estado = EstadoOrdenPago_pendiente And vFactElegida.NetoGravadoAbonado = 0 And vFactElegida.OtrosAbonado = 0 Then
            Set c = DAOOrdenPago.FindAbonadoFactura(vFactElegida.Id, OrdenPago.Id)

            vFactElegida.NetoGravadoAbonado = c(2)
            vFactElegida.OtrosAbonado = c(3)
        End If

       MostrarPago vFactElegida
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




    For i = 0 To Me.lstFacturas.ListCount - 1
        If Me.lstFacturas.Checked(i) Then


            If funciones.BuscarEnColeccion(colFacturas, CStr(Me.lstFacturas.ItemData(i))) Then


                col.Add colFacturas.item(CStr(Me.lstFacturas.ItemData(i)))


                Me.lblCantidadCbtesSeleccionados.caption = "Cbtes. Seleccionados: " & col.count

            End If

        Else


            'si destildo tengo q ver q no existan compensatorios. Si existen deber?a primero eliminarlos.
            Dim ff As clsFacturaProveedor
            Dim c As Compensatorio
            For Each c In OrdenPago.Compensatorios
                Set ff = colFacturas.item(CStr(Me.lstFacturas.ItemData(i)))
                If c.Comprobante.Id = ff.Id Then
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

            Dim c As Compensatorio
            Dim hay As Boolean
            hay = False
            For Each c In OrdenPago.Compensatorios
                If c.Comprobante.Id = Factura.Id Then
                    hay = True
                    Exit For
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

    Me.txtRetenciones.text = 0

    If Not Me.cboProveedores.Enabled Then Me.cboProveedores.ListIndex = -1
    If Not Me.lstFacturas.Enabled Then Me.lstFacturas.Clear

    If Not Me.cboCuentas.Enabled Then Me.cboCuentas.ListIndex = -1
    If Not Me.txtDetalle.Enabled Then Me.txtDetalle.text = vbNullString


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
    OrdenPago.OperacionesCaja.Add operacion
    Totalizar
End Sub

Private Sub gridCajaOperaciones_UnboundDelete(ByVal rowIndex As Long, ByVal Bookmark As Variant)
    If rowIndex > 0 And OrdenPago.OperacionesCaja.count >= rowIndex Then
        OrdenPago.OperacionesCaja.remove rowIndex
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

    Me.txtDiferenciaCambioPago.text = T
    Me.txtDifTipoCambioIVA.text = TIVA
    Me.txtDifCambio = TTOTAL


    If ReadOnly Then
        Dim s As String
        s = OrdenPago.DiferenciaCambioEnNG
        Me.txtDifCambioNG1.text = s
        s = OrdenPago.DiferenciaCambioEnTOTAL
        Me.txtDifCambioTOTAL1.text = s
    End If

End Function
Private Sub gridCajaOperaciones_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If rowIndex <= OrdenPago.OperacionesCaja.count Then
        Set operacion = OrdenPago.OperacionesCaja.item(rowIndex)
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
    If rowIndex > 0 And OrdenPago.OperacionesCaja.count > 0 Then
        Set operacion = OrdenPago.OperacionesCaja.item(rowIndex)
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
    OrdenPago.OperacionesBanco.Add operacion
    Totalizar
End Sub

Private Sub gridDepositosOperaciones_UnboundDelete(ByVal rowIndex As Long, ByVal Bookmark As Variant)
    If rowIndex > 0 And OrdenPago.OperacionesBanco.count >= rowIndex Then
        OrdenPago.OperacionesBanco.remove rowIndex
        Totalizar
    End If
End Sub

Private Sub gridDepositosOperaciones_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If rowIndex <= OrdenPago.OperacionesBanco.count Then
        Set operacion = OrdenPago.OperacionesBanco.item(rowIndex)
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
    If rowIndex > 0 And OrdenPago.OperacionesBanco.count > 0 Then
        Set operacion = OrdenPago.OperacionesBanco.item(rowIndex)
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

Private Sub Text1_Change()

End Sub

Private Sub txtBuscarFactura_GotFocus()
    Me.txtBuscarFactura.SelStart = 0
    Me.txtBuscarFactura.SelLength = Len(Me.txtBuscarFactura.text)
End Sub

Private Sub txtBuscarFactura_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        'buscar en facturas y tildar

        If LenB(Me.txtBuscarFactura.text) > 0 Then
            Dim cont As Long

            If colFacturas.count > 0 Then
                Dim i As Long
                For Each vFacturaProveedor In colFacturas
                    If InStr(1, vFacturaProveedor.numero, Me.txtBuscarFactura.text) > 0 Then    'aplica
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
    OrdenPago.OtrosDescuentos = Val(Me.txtOtrosDescuentos.text)
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


    vFactElegida.TotalAbonado = CDbl(txtTotalParcialAbonar)

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

