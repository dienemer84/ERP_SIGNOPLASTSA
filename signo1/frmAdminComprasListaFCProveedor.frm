VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmAdminComprasListaFCProveedor 
   Caption         =   "Comprobantes de Proveedores"
   ClientHeight    =   9345
   ClientLeft      =   1440
   ClientTop       =   4725
   ClientWidth     =   19080
   Icon            =   "frmAdminComprasListaFCProveedor.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5081.851
   ScaleMode       =   0  'User
   ScaleWidth      =   24955.33
   Begin XtremeSuiteControls.GroupBox GroupBox5 
      Height          =   3360
      Left            =   11160
      TabIndex        =   35
      Top             =   0
      Width           =   7575
      _Version        =   786432
      _ExtentX        =   13361
      _ExtentY        =   5927
      _StockProps     =   79
      BackColor       =   16744576
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.GroupBox GroupBox4 
         Height          =   2415
         Left            =   120
         TabIndex        =   36
         Top             =   120
         Width           =   7335
         _Version        =   786432
         _ExtentX        =   12938
         _ExtentY        =   4260
         _StockProps     =   79
         BackColor       =   16744576
         Appearance      =   4
         Begin XtremeSuiteControls.ProgressBar progreso 
            Height          =   255
            Left            =   120
            TabIndex        =   42
            Top             =   2125
            Visible         =   0   'False
            Width           =   6975
            _Version        =   786432
            _ExtentX        =   12303
            _ExtentY        =   450
            _StockProps     =   93
            Appearance      =   6
         End
         Begin VB.Label lblTotal 
            AutoSize        =   -1  'True
            Caption         =   "Total Filtrado $:"
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
            Left            =   120
            TabIndex        =   52
            Top             =   1890
            Width           =   1365
         End
         Begin VB.Label lblTotalNeto 
            AutoSize        =   -1  'True
            Caption         =   "Total Filtrado $:"
            Height          =   195
            Left            =   120
            TabIndex        =   51
            Top             =   780
            Width           =   1095
         End
         Begin VB.Label lblTotalIVA 
            AutoSize        =   -1  'True
            Caption         =   "Total Filtrado $:"
            Height          =   195
            Left            =   120
            TabIndex        =   50
            Top             =   1050
            Width           =   1095
         End
         Begin VB.Label lblTotalNoGravadoFiltrado 
            AutoSize        =   -1  'True
            Caption         =   "Total Filtrado $:"
            Height          =   195
            Left            =   120
            TabIndex        =   49
            Top             =   510
            Width           =   1095
         End
         Begin VB.Label lblNetoGravadoFiltrado 
            AutoSize        =   -1  'True
            Caption         =   "Total Filtrado $:"
            Height          =   195
            Left            =   120
            TabIndex        =   48
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lblTotalPercepciones 
            AutoSize        =   -1  'True
            Caption         =   "Total Filtrado $:"
            Height          =   195
            Left            =   120
            TabIndex        =   47
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label lblTotalPendiente 
            AutoSize        =   -1  'True
            Caption         =   "Total Filtrado $:"
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
            Left            =   120
            TabIndex        =   46
            Top             =   1605
            Width           =   1365
         End
      End
      Begin XtremeSuiteControls.GroupBox gbBotones 
         Height          =   735
         Left            =   120
         TabIndex        =   37
         Top             =   2520
         Width           =   7335
         _Version        =   786432
         _ExtentX        =   12938
         _ExtentY        =   1296
         _StockProps     =   79
         BackColor       =   16744576
         Appearance      =   4
         Begin XtremeSuiteControls.PushButton btnBuscar 
            Default         =   -1  'True
            Height          =   390
            Left            =   120
            TabIndex        =   38
            Top             =   240
            Width           =   1245
            _Version        =   786432
            _ExtentX        =   2196
            _ExtentY        =   688
            _StockProps     =   79
            Caption         =   "Buscar"
            BackColor       =   16744576
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton btnImprimir 
            Height          =   390
            Left            =   3000
            TabIndex        =   39
            Top             =   240
            Width           =   1245
            _Version        =   786432
            _ExtentX        =   2196
            _ExtentY        =   688
            _StockProps     =   79
            Caption         =   "Imprimir"
            BackColor       =   16744576
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton btnExportar 
            Height          =   390
            Left            =   1560
            TabIndex        =   40
            Top             =   240
            Width           =   1245
            _Version        =   786432
            _ExtentX        =   2196
            _ExtentY        =   688
            _StockProps     =   79
            Caption         =   "Exportar"
            BackColor       =   16744576
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox checkVerIds 
            Height          =   255
            Left            =   4560
            TabIndex        =   41
            Top             =   285
            Width           =   1095
            _Version        =   786432
            _ExtentX        =   1931
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Ver Id's"
            BackColor       =   16744576
            UseVisualStyle  =   -1  'True
         End
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   3360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11130
      _Version        =   786432
      _ExtentX        =   19632
      _ExtentY        =   5927
      _StockProps     =   79
      Caption         =   "Parámetros de búsqueda"
      BackColor       =   16744576
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.PushButton btnClearCtaCble_Click 
         Height          =   255
         Left            =   5520
         TabIndex        =   45
         Top             =   1920
         Width           =   495
         _Version        =   786432
         _ExtentX        =   873
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "X"
         BackColor       =   12632256
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cboOrdenImporte 
         Height          =   315
         Left            =   1440
         TabIndex        =   19
         Top             =   2835
         Width           =   3885
         _Version        =   786432
         _ExtentX        =   6853
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin VB.TextBox txtMontoHasta1 
         Height          =   315
         Left            =   3720
         TabIndex        =   16
         Top             =   2445
         Width           =   1600
      End
      Begin VB.TextBox txtMontoDesde1 
         Height          =   315
         Left            =   1440
         TabIndex        =   15
         Top             =   2445
         Width           =   1600
      End
      Begin XtremeSuiteControls.PushButton btnClearFormaDePago 
         Height          =   255
         Left            =   5520
         TabIndex        =   11
         Top             =   1485
         Width           =   420
         _Version        =   786432
         _ExtentX        =   741
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "X"
         BackColor       =   12632256
         UseVisualStyle  =   -1  'True
      End
      Begin VB.TextBox txtComprobante 
         Height          =   315
         Left            =   1440
         TabIndex        =   6
         Top             =   645
         Width           =   3885
      End
      Begin XtremeSuiteControls.ComboBox cboProveedores 
         Height          =   315
         Left            =   1440
         TabIndex        =   3
         Top             =   240
         Width           =   3885
         _Version        =   786432
         _ExtentX        =   6853
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Text            =   "cboProveedores"
      End
      Begin XtremeSuiteControls.PushButton btnRemoveProveedor 
         Height          =   255
         Left            =   5520
         TabIndex        =   4
         Top             =   285
         Width           =   420
         _Version        =   786432
         _ExtentX        =   741
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "X"
         BackColor       =   12632256
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnRemoveEstado 
         Height          =   255
         Left            =   5520
         TabIndex        =   8
         Top             =   1080
         Width           =   420
         _Version        =   786432
         _ExtentX        =   741
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "X"
         BackColor       =   12632256
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cboEstado 
         Height          =   315
         Left            =   1440
         TabIndex        =   7
         Top             =   1050
         Width           =   3885
         _Version        =   786432
         _ExtentX        =   6853
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Sorted          =   -1  'True
         Style           =   2
         Text            =   "cboProveedores"
      End
      Begin XtremeSuiteControls.ComboBox cboBoxFormaDePago 
         Height          =   315
         Left            =   1440
         TabIndex        =   10
         Top             =   1440
         Width           =   3885
         _Version        =   786432
         _ExtentX        =   6853
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Sorted          =   -1  'True
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.GroupBox GroupBox3 
         Height          =   1215
         Left            =   6240
         TabIndex        =   21
         Top             =   240
         Width           =   4695
         _Version        =   786432
         _ExtentX        =   8281
         _ExtentY        =   2143
         _StockProps     =   79
         Caption         =   "Fecha Carga"
         BackColor       =   16744576
         Appearance      =   4
         Begin XtremeSuiteControls.DateTimePicker dtpDesdeCarga 
            Height          =   315
            Left            =   720
            TabIndex        =   22
            Top             =   720
            Width           =   1470
            _Version        =   786432
            _ExtentX        =   2593
            _ExtentY        =   556
            _StockProps     =   68
            CheckBox        =   -1  'True
            Format          =   1
         End
         Begin XtremeSuiteControls.DateTimePicker dtpHastaCarga 
            Height          =   315
            Left            =   2925
            TabIndex        =   23
            Top             =   720
            Width           =   1470
            _Version        =   786432
            _ExtentX        =   2593
            _ExtentY        =   556
            _StockProps     =   68
            CheckBox        =   -1  'True
            Format          =   1
         End
         Begin XtremeSuiteControls.ComboBox cboRangosCarga 
            Height          =   315
            Left            =   720
            TabIndex        =   24
            Top             =   300
            Width           =   3675
            _Version        =   786432
            _ExtentX        =   6482
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Style           =   2
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.Label Label7 
            Height          =   195
            Left            =   120
            TabIndex        =   27
            Top             =   360
            Width           =   480
            _Version        =   786432
            _ExtentX        =   847
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Rango"
            BackColor       =   12632256
            AutoSize        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label8 
            Height          =   195
            Index           =   0
            Left            =   165
            TabIndex        =   26
            Top             =   780
            Width           =   465
            _Version        =   786432
            _ExtentX        =   820
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Desde"
            BackColor       =   12632256
            AutoSize        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label9 
            Height          =   195
            Index           =   0
            Left            =   2400
            TabIndex        =   25
            Top             =   780
            Width           =   420
            _Version        =   786432
            _ExtentX        =   741
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Hasta"
            BackColor       =   12632256
            AutoSize        =   -1  'True
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox2 
         Height          =   1215
         Left            =   6240
         TabIndex        =   28
         Top             =   1560
         Width           =   4695
         _Version        =   786432
         _ExtentX        =   8281
         _ExtentY        =   2143
         _StockProps     =   79
         Caption         =   "Fecha Comprobante"
         BackColor       =   16744576
         Appearance      =   4
         Begin XtremeSuiteControls.DateTimePicker dtpDesde 
            Height          =   315
            Left            =   720
            TabIndex        =   29
            Top             =   720
            Width           =   1470
            _Version        =   786432
            _ExtentX        =   2593
            _ExtentY        =   556
            _StockProps     =   68
            CheckBox        =   -1  'True
            Format          =   1
         End
         Begin XtremeSuiteControls.DateTimePicker dtpHasta 
            Height          =   315
            Left            =   2925
            TabIndex        =   30
            Top             =   720
            Width           =   1470
            _Version        =   786432
            _ExtentX        =   2593
            _ExtentY        =   556
            _StockProps     =   68
            CheckBox        =   -1  'True
            Format          =   1
         End
         Begin XtremeSuiteControls.ComboBox cboRangos 
            Height          =   315
            Left            =   720
            TabIndex        =   31
            Top             =   300
            Width           =   3675
            _Version        =   786432
            _ExtentX        =   6482
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Style           =   2
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.Label Label6 
            Height          =   195
            Left            =   2400
            TabIndex        =   34
            Top             =   780
            Width           =   420
            _Version        =   786432
            _ExtentX        =   741
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Hasta"
            BackColor       =   12632256
            AutoSize        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label5 
            Height          =   195
            Left            =   165
            TabIndex        =   33
            Top             =   780
            Width           =   465
            _Version        =   786432
            _ExtentX        =   820
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Desde"
            BackColor       =   12632256
            AutoSize        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   195
            Left            =   120
            TabIndex        =   32
            Top             =   360
            Width           =   480
            _Version        =   786432
            _ExtentX        =   847
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Rango"
            BackColor       =   12632256
            AutoSize        =   -1  'True
         End
      End
      Begin XtremeSuiteControls.ComboBox cboCuentasContables 
         Height          =   315
         Left            =   1425
         TabIndex        =   43
         Top             =   1920
         Width           =   3885
         _Version        =   786432
         _ExtentX        =   6853
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Sorted          =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label12 
         Height          =   195
         Left            =   360
         TabIndex        =   44
         Top             =   1950
         Width           =   960
         _Version        =   786432
         _ExtentX        =   1693
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Cta. Contable"
         BackColor       =   12632256
         Alignment       =   1
         AutoSize        =   -1  'True
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ordenar"
         Height          =   195
         Left            =   720
         TabIndex        =   20
         Top             =   2895
         Width           =   615
      End
      Begin XtremeSuiteControls.Label Label8 
         Height          =   195
         Index           =   1
         Left            =   300
         TabIndex        =   18
         Top             =   2520
         Width           =   1035
         _Version        =   786432
         _ExtentX        =   1826
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Importe Desde"
         BackColor       =   12632256
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label9 
         Height          =   195
         Index           =   1
         Left            =   3240
         TabIndex        =   17
         Top             =   2520
         Width           =   420
         _Version        =   786432
         _ExtentX        =   741
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Hasta"
         BackColor       =   12632256
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblFormaDePago 
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   1480
         Width           =   1095
         _Version        =   786432
         _ExtentX        =   1931
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Forma de Pago"
         BackColor       =   12632256
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label Label10 
         Height          =   195
         Left            =   840
         TabIndex        =   9
         Top             =   1120
         Width           =   495
         _Version        =   786432
         _ExtentX        =   873
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Estado"
         BackColor       =   12632256
         Alignment       =   1
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   195
         Left            =   165
         TabIndex        =   5
         Top             =   720
         Width           =   1170
         _Version        =   786432
         _ExtentX        =   2064
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Nº Comprobante"
         BackColor       =   12632256
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   195
         Left            =   600
         TabIndex        =   2
         Top             =   300
         Width           =   735
         _Version        =   786432
         _ExtentX        =   1296
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Proveedor"
         BackColor       =   12632256
         AutoSize        =   -1  'True
      End
   End
   Begin GridEX20.GridEX grilla 
      Height          =   5175
      Left            =   0
      TabIndex        =   0
      Top             =   3480
      Width           =   18690
      _ExtentX        =   32967
      _ExtentY        =   9128
      Version         =   "2.0"
      PreviewRowIndent=   100
      AutomaticSort   =   -1  'True
      DefaultGroupMode=   1
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      PreviewRowLines =   1
      OLEDropMode     =   1
      ColumnAutoResize=   -1  'True
      HeaderStyle     =   2
      MethodHoldFields=   -1  'True
      ContScroll      =   -1  'True
      LockType        =   4
      GroupByBoxInfoText=   ""
      AllowEdit       =   0   'False
      BorderStyle     =   0
      BackColorGBBox  =   16744576
      BackColorHeader =   16761024
      ImageCount      =   1
      ImagePicture1   =   "frmAdminComprasListaFCProveedor.frx":000C
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   22
      Column(1)       =   "frmAdminComprasListaFCProveedor.frx":0326
      Column(2)       =   "frmAdminComprasListaFCProveedor.frx":048A
      Column(3)       =   "frmAdminComprasListaFCProveedor.frx":05B2
      Column(4)       =   "frmAdminComprasListaFCProveedor.frx":06D2
      Column(5)       =   "frmAdminComprasListaFCProveedor.frx":0816
      Column(6)       =   "frmAdminComprasListaFCProveedor.frx":0996
      Column(7)       =   "frmAdminComprasListaFCProveedor.frx":0ADA
      Column(8)       =   "frmAdminComprasListaFCProveedor.frx":0D46
      Column(9)       =   "frmAdminComprasListaFCProveedor.frx":0F42
      Column(10)      =   "frmAdminComprasListaFCProveedor.frx":1172
      Column(11)      =   "frmAdminComprasListaFCProveedor.frx":1362
      Column(12)      =   "frmAdminComprasListaFCProveedor.frx":156E
      Column(13)      =   "frmAdminComprasListaFCProveedor.frx":176E
      Column(14)      =   "frmAdminComprasListaFCProveedor.frx":18DE
      Column(15)      =   "frmAdminComprasListaFCProveedor.frx":1A1E
      Column(16)      =   "frmAdminComprasListaFCProveedor.frx":1B86
      Column(17)      =   "frmAdminComprasListaFCProveedor.frx":1CDE
      Column(18)      =   "frmAdminComprasListaFCProveedor.frx":1E36
      Column(19)      =   "frmAdminComprasListaFCProveedor.frx":1F6E
      Column(20)      =   "frmAdminComprasListaFCProveedor.frx":20C6
      Column(21)      =   "frmAdminComprasListaFCProveedor.frx":226A
      Column(22)      =   "frmAdminComprasListaFCProveedor.frx":23C2
      FormatStylesCount=   9
      FormatStyle(1)  =   "frmAdminComprasListaFCProveedor.frx":24D2
      FormatStyle(2)  =   "frmAdminComprasListaFCProveedor.frx":260A
      FormatStyle(3)  =   "frmAdminComprasListaFCProveedor.frx":26BA
      FormatStyle(4)  =   "frmAdminComprasListaFCProveedor.frx":276E
      FormatStyle(5)  =   "frmAdminComprasListaFCProveedor.frx":2846
      FormatStyle(6)  =   "frmAdminComprasListaFCProveedor.frx":28FE
      FormatStyle(7)  =   "frmAdminComprasListaFCProveedor.frx":29DE
      FormatStyle(8)  =   "frmAdminComprasListaFCProveedor.frx":2A9E
      FormatStyle(9)  =   "frmAdminComprasListaFCProveedor.frx":2B62
      ImageCount      =   1
      ImagePicture(1) =   "frmAdminComprasListaFCProveedor.frx":2C22
      PrinterProperties=   "frmAdminComprasListaFCProveedor.frx":2F3C
   End
   Begin XtremeSuiteControls.ComboBox cboFantasia 
      Height          =   315
      Left            =   1080
      TabIndex        =   13
      Top             =   0
      Width           =   3885
      _Version        =   786432
      _ExtentX        =   6853
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Sorted          =   -1  'True
      Text            =   "cboProveedores"
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   195
      Left            =   0
      TabIndex        =   14
      Top             =   30
      Width           =   975
      _Version        =   786432
      _ExtentX        =   1720
      _ExtentY        =   344
      _StockProps     =   79
      Caption         =   "Nom Fantasia"
      BackColor       =   12632256
      AutoSize        =   -1  'True
   End
   Begin VB.Menu menu 
      Caption         =   "menu"
      Visible         =   0   'False
      Begin VB.Menu VerDetalle 
         Caption         =   "Ver Factura..."
      End
      Begin VB.Menu editar 
         Caption         =   "Editar..."
      End
      Begin VB.Menu finalizar 
         Caption         =   "Aprobar..."
      End
      Begin VB.Menu mnuPagarEnEfectivo 
         Caption         =   "Pagar en Efectivo..."
      End
      Begin VB.Menu mnuArchivos 
         Caption         =   "Archivos Asociados..."
      End
      Begin VB.Menu mnuScan 
         Caption         =   "Adquirir..."
      End
      Begin VB.Menu n1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEliminar 
         Caption         =   "Eliminar"
      End
      Begin VB.Menu n2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuVerLIQ 
         Caption         =   "Ver Liquidacion de Caja..."
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu MnuVerOP 
         Caption         =   "Ver Orden de Pago..."
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu verHistorial 
         Caption         =   "Historial..."
      End
      Begin VB.Menu Imprimir 
         Caption         =   "Imprimir"
      End
   End
End
Attribute VB_Name = "frmAdminComprasListaFCProveedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Implements ISuscriber
Dim vId As String
Private desde
Private Factura As clsFacturaProveedor
Private facturas As Collection
Dim m_Archivos As Dictionary

Private Sub btnClearCtaCble_Click_Click()
    Me.cboCuentasContables.ListIndex = -1
End Sub

Private Sub btnClearFormaDePago_Click()
    Me.cboBoxFormaDePago.ListIndex = -1
End Sub

Private Sub btnRemoveEstado_Click()
    Me.cboEstado.ListIndex = -1
End Sub

Private Sub cboRangos_Click()
    funciones.CalculateDateRange Me.cboRangos, Me.dtpDesde, Me.dtpHasta
End Sub

Private Sub cboRangosCarga_Click()
    funciones.CalculateDateRange Me.cboRangosCarga, Me.dtpDesdeCarga, Me.dtpHastaCarga
End Sub

Private Sub checkVerIds_Click()
    If Me.checkVerIds.value = xtpUnchecked Then
        Me.grilla.Columns(22).Visible = False

    ElseIf Me.checkVerIds.value = xtpChecked Then
        Me.grilla.Columns(22).Visible = True
        Me.grilla.Columns(22).Width = 800
    End If

End Sub

Private Sub btnExportar_Click()

    Me.progreso.Visible = True


    If IsSomething(facturas) Then
        If Not DAOFacturaProveedor.ExportarColeccion(facturas, Me.progreso) Then GoTo err1
    End If

    Me.progreso.Visible = False

    Exit Sub
err1:
    MsgBox "Se produjo un error al exportar!", vbCritical, "Error"
End Sub

Private Sub btnImprimir_Click()
    Dim elegidos As Boolean
    Dim q As String

    Dim rf As String

    rf = Me.lblNetoGravadoFiltrado & Chr(10)
    rf = rf & Me.lblTotalNoGravadoFiltrado & Chr(10)
    rf = rf & Me.lblTotalNeto & Chr(10)
    rf = rf & Me.lblTotalIVA & Chr(10)
    rf = rf & Me.lblTotal & Chr(10)


    If Not IsNull(Me.dtpDesde) Then
        q = "Desde " & Format(Me.dtpDesde, "dd-mm-yyyy") & Chr(10)
    End If
    If Not IsNull(Me.dtpHasta) Then
        q = q & "Hasta " & Format(Me.dtpHasta, "dd-mm-yyyy") & Chr(10)

    End If


    If IsNull(Me.dtpHasta) And IsNull(Me.dtpDesde) Then
        q = "PERIODO SIN ESPECIFICAR" & Chr(10)
    End If

    Dim pro As String
    If Me.cboProveedores.ListIndex > -1 Then
        pro = " Proveedor: " & Me.cboProveedores.text
    End If

    With Me.grilla.PrinterProperties

        .FitColumns = True
        .RepeatHeaders = True
        .Orientation = jgexPPLandscape
        .HeaderString(jgexHFCenter) = "Listado de Comprobantes de Proveedores" & Chr(10) & pro
        .FooterString(jgexHFCenter) = Now
        .FooterString(jgexHFLeft) = rf
        .BottomMargin = 1500
        .FooterDistance = 1400

    End With
    Load frmPrintPreview
    frmPrintPreview.Move Me.Left, Me.Top, Me.Width, Me.Height
    Me.grilla.PrintPreview frmPrintPreview.GEXPreview1, elegidos
    frmPrintPreview.Show 1
End Sub

Private Sub btnRemoveProveedor_Click()
    Me.cboProveedores.ListIndex = -1
End Sub

Private Sub btnBuscar_Click()
    llenarGrilla
End Sub

Private Sub editar_Click()
    Set Factura = facturas.item(grilla.rowIndex(grilla.row))
    Dim frm As frmAdminComprasNuevaFCProveedor
    Set frm = New frmAdminComprasNuevaFCProveedor

    frm.ver = False
    frm.Factura = Factura
    frm.Show
End Sub
Private Sub finalizar_Click()
    If Me.grilla.ItemCount > 0 Then
        SeleccionarFactura
        Dim l As Long
        l = grilla.rowIndex(grilla.row)
        If MsgBox("¿Desea aprobar la factura?", vbQuestion + vbYesNo) = vbYes Then
            If DAOFacturaProveedor.aprobar(Factura) Then
                MsgBox "Factura aprobada con éxito!", vbInformation, "Información"
                '--------------- added 28-1-11
                txtComprobante.SetFocus
                funciones.foco Me.txtComprobante
                '---------------------------------------
                If Not Factura.FormaPagoCuentaCorriente Then MsgBox "El pago de la factura ha sido registrado con la orden de pago Nº " & DAOOrdenPago.FindLast().Id & ".", vbInformation

                '                Dim tmp As clsFacturaProveedor
                facturas.item(grilla.rowIndex(grilla.row)).estado = Factura.estado


                grilla.RefreshRowIndex l
            Else
                'MsgBox "Se produjo algún error, no se aprobó la factura!", vbCritical, "Error"
            End If
        End If
    End If
End Sub
Private Sub Form_Load()
    Set m_Archivos = DAOArchivo.GetCantidadArchivosPorReferencia(OA_FacturaProveedor)
    vId = funciones.CreateGUID
    Channel.AgregarSuscriptor Me, TipoSuscripcion.FacturaProveedor_
    FormHelper.Customize Me
    GridEXHelper.CustomizeGrid Me.grilla, True

    Set colProveedores = DAOProveedor.FindAll
    For Each prov In colProveedores
        cboProveedores.AddItem prov.RazonSocial
        cboProveedores.ItemData(cboProveedores.NewIndex) = prov.Id
    Next

    llenarComboEstado

    llenarComboFormaPago

    llenarComboOrdenImporte

    Dim P As clsProveedor
    For Each P In DAOProveedor.FindAll()
        If LenB(Trim$(P.razonFantasia)) > 0 Then
            Me.cboFantasia.AddItem P.razonFantasia
            Me.cboFantasia.ItemData(Me.cboFantasia.NewIndex) = P.Id
        End If
    Next P
    Me.cboFantasia.ListIndex = -1


    Dim cc As clsCuentaContable
    For Each cc In DAOCuentaContable.GetAll
        If LenB(Trim$(cc.nombre)) > 0 Then
            Me.cboCuentasContables.AddItem cc.codigo & "- " & cc.nombre
            Me.cboCuentasContables.ItemData(Me.cboCuentasContables.NewIndex) = cc.Id
        End If
    Next cc

    Me.cboCuentasContables.ListIndex = -1

    Me.grilla.ItemCount = 0
    btnRemoveProveedor_Click
    desde = DateSerial(Year(Date), Month(Date), 1)   ' CDate(1 & "-" & Month(Now) & "-" & Year(Now))
    funciones.FillComboBoxDateRanges Me.cboRangos

    For i = 0 To Me.cboRangos.ListCount - 1
        If Me.cboRangos.ItemData(i) = DateRangeValue.DRV_YearCurrent Then Exit For
    Next i
    Me.cboRangos.ListIndex = i

    funciones.FillComboBoxDateRanges Me.cboRangosCarga

    Me.dtpDesdeCarga.value = Null
    Me.dtpHastaCarga.value = Null

    Me.grilla.Refresh

End Sub

Private Sub llenarComboEstado()
    Me.cboEstado.Clear
    Me.cboEstado.AddItem enums.enumEstadoFacturaProveedor(1)
    Me.cboEstado.ItemData(Me.cboEstado.NewIndex) = EstadoFacturaProveedor.EnProceso
    Me.cboEstado.AddItem enums.enumEstadoFacturaProveedor(2)
    Me.cboEstado.ItemData(Me.cboEstado.NewIndex) = EstadoFacturaProveedor.Aprobada
    Me.cboEstado.AddItem enums.enumEstadoFacturaProveedor(3)
    Me.cboEstado.ItemData(Me.cboEstado.NewIndex) = EstadoFacturaProveedor.Saldada
    Me.cboEstado.AddItem enums.enumEstadoFacturaProveedor(4)
    Me.cboEstado.ItemData(Me.cboEstado.NewIndex) = EstadoFacturaProveedor.pagoParcial

End Sub


Private Sub llenarComboFormaPago()
    Me.cboBoxFormaDePago.Clear
    Me.cboBoxFormaDePago.AddItem enums.enumFormaDePagoFacturaProveedor(1)
    Me.cboBoxFormaDePago.ItemData(Me.cboBoxFormaDePago.NewIndex) = FormadePagoFacturaProveedor.PagoContado
    Me.cboBoxFormaDePago.AddItem enums.enumFormaDePagoFacturaProveedor(0)
    Me.cboBoxFormaDePago.ItemData(Me.cboBoxFormaDePago.NewIndex) = FormadePagoFacturaProveedor.PagoCuentaCorriente
End Sub


Private Sub llenarComboOrdenImporte()
    Me.cboOrdenImporte.Clear
    cboOrdenImporte.AddItem "Ascendente"
    cboOrdenImporte.ItemData(cboOrdenImporte.NewIndex) = 0
    cboOrdenImporte.AddItem "Descendente"
    cboOrdenImporte.ItemData(cboOrdenImporte.NewIndex) = 1
End Sub

Public Sub llenarGrilla()
'    Dim tot As Double
    grilla.ItemCount = 0
    Dim condition As String
    condition = " 1 = 1 "

    If Not IsNull(Me.dtpDesde.value) Then
        condition = condition & " AND AdminComprasFacturasProveedores.fecha >= " & conectar.Escape(Me.dtpDesde.value)
    End If

    If Not IsNull(Me.dtpHasta.value) Then
        condition = condition & " AND AdminComprasFacturasProveedores.fecha <= " & conectar.Escape(Me.dtpHasta.value)
    End If

    If Not IsNull(Me.dtpDesdeCarga.value) Then
        condition = condition & " AND AdminComprasFacturasProveedores.fecha_carga >= " & conectar.Escape(CDate(Int(CDbl(Me.dtpDesdeCarga.value))))
    End If

    If Not IsNull(Me.dtpHastaCarga.value) Then
        condition = condition & " AND AdminComprasFacturasProveedores.fecha_carga <= " & conectar.Escape(CDate(Int(CDbl(Me.dtpHastaCarga.value))) + TimeSerial(23, 59, 59))
    End If

    If cboProveedores.ListIndex > -1 Then
        condition = condition & " AND AdminComprasFacturasProveedores.id_proveedor = " & cboProveedores.ItemData(Me.cboProveedores.ListIndex)
    End If

    If Me.cboFantasia.ListIndex > -1 Then
        condition = condition & " AND AdminComprasFacturasProveedores.id_proveedor = " & cboFantasia.ItemData(Me.cboFantasia.ListIndex)
    End If

    If LenB(Me.txtComprobante) > 0 Then
        condition = condition & " AND AdminComprasFacturasProveedores.numero_factura like '%" & Trim(Me.txtComprobante.text) & "%'"
    End If

    If Me.cboEstado.ListIndex > -1 Then
        condition = condition & " AND AdminComprasFacturasProveedores.estado = " & Me.cboEstado.ItemData(Me.cboEstado.ListIndex)
    End If

    If Me.cboBoxFormaDePago.ListIndex > -1 Then
        condition = condition & " AND AdminComprasFacturasProveedores.forma_de_pago_cta_cte = " & Me.cboBoxFormaDePago.ItemData(Me.cboBoxFormaDePago.ListIndex)
    End If

    '#181
    If Me.cboCuentasContables.ListIndex > -1 Then
        condition = condition & " AND AdminComprasCuentasContables.id = '" & Me.cboCuentasContables.ItemData(Me.cboCuentasContables.ListIndex) & " '"
    End If

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    ' AGREGAR REGLA DE QUE SEA NUMERICO

    If Me.txtMontoDesde1 <> "" Then
        condition = condition & " AND (ROUND(AdminComprasFacturasProveedores.monto_neto + " _
                    & "AdminComprasFacturasProveedores.redondeo_iva +" _
                    & " AdminComprasFacturasProveedores.impuesto_interno +" _
                    & " (SELECT SUM(iva_calculado)         FROM sp.AdminComprasFacturasProveedoresIva acfpi " _
                    & " JOIN AdminComprasFacturasProveedores acfp ON acfpi.id_factura_proveedor=acfp.id " _
                    & " Where id_factura_proveedor = AdminComprasFacturasProveedores.Id " _
                    & " )                + " _
                    & " IF((SELECT SUM(valor)         FROM AdminComprasFacturasProveedoresPercepciones acfpp " _
                    & " JOIN AdminComprasFacturasProveedores acfp ON acfpp.id_factura_proveedor=acfp.id " _
                    & " Where id_factura_proveedor = AdminComprasFacturasProveedores.Id " _
                    & ") IS NULL,0," _
                    & " (SELECT SUM(valor)         FROM AdminComprasFacturasProveedoresPercepciones acfpp " _
                    & " JOIN AdminComprasFacturasProveedores acfp ON acfpp.id_factura_proveedor=acfp.id " _
                    & " Where id_factura_proveedor = AdminComprasFacturasProveedores.Id " _
                    & ")),2)                *             (IF (AdminComprasFacturasProveedores.tipo_doc_contable = 1,'-1','1'))) >= " & Me.txtMontoDesde1

    End If

    If Me.txtMontoHasta1 <> "" Then
        condition = condition & " AND (ROUND(AdminComprasFacturasProveedores.monto_neto + " _
                    & "AdminComprasFacturasProveedores.redondeo_iva +" _
                    & " AdminComprasFacturasProveedores.impuesto_interno +" _
                    & " (SELECT SUM(iva_calculado)         FROM sp.AdminComprasFacturasProveedoresIva acfpi " _
                    & " JOIN AdminComprasFacturasProveedores acfp ON acfpi.id_factura_proveedor=acfp.id " _
                    & " Where id_factura_proveedor = AdminComprasFacturasProveedores.Id " _
                    & " )                + " _
                    & " IF((SELECT SUM(valor)         FROM AdminComprasFacturasProveedoresPercepciones acfpp " _
                    & " JOIN AdminComprasFacturasProveedores acfp ON acfpp.id_factura_proveedor=acfp.id " _
                    & " Where id_factura_proveedor = AdminComprasFacturasProveedores.Id " _
                    & ") IS NULL,0," _
                    & " (SELECT SUM(valor)         FROM AdminComprasFacturasProveedoresPercepciones acfpp " _
                    & " JOIN AdminComprasFacturasProveedores acfp ON acfpp.id_factura_proveedor=acfp.id " _
                    & " Where id_factura_proveedor = AdminComprasFacturasProveedores.Id " _
                    & ")),2)                *             (IF (AdminComprasFacturasProveedores.tipo_doc_contable = 1,'-1','1'))) <= " & Me.txtMontoHasta1

    End If

    Dim ordenImporte As String

    If Me.cboOrdenImporte.ListIndex = 0 Then
        ordenImporte = "(ROUND(AdminComprasFacturasProveedores.monto_neto + " _
                       & "AdminComprasFacturasProveedores.redondeo_iva +" _
                       & " AdminComprasFacturasProveedores.impuesto_interno +" _
                       & " (SELECT SUM(iva_calculado)         FROM sp.AdminComprasFacturasProveedoresIva acfpi " _
                       & " JOIN AdminComprasFacturasProveedores acfp ON acfpi.id_factura_proveedor=acfp.id " _
                       & " Where id_factura_proveedor = AdminComprasFacturasProveedores.Id " _
                       & " )                + " _
                       & " IF((SELECT SUM(valor)         FROM AdminComprasFacturasProveedoresPercepciones acfpp " _
                       & " JOIN AdminComprasFacturasProveedores acfp ON acfpp.id_factura_proveedor=acfp.id " _
                       & " Where id_factura_proveedor = AdminComprasFacturasProveedores.Id " _
                       & ") IS NULL,0," _
                       & " (SELECT SUM(valor)         FROM AdminComprasFacturasProveedoresPercepciones acfpp " _
                       & " JOIN AdminComprasFacturasProveedores acfp ON acfpp.id_factura_proveedor=acfp.id " _
                       & " Where id_factura_proveedor = AdminComprasFacturasProveedores.Id " _
                       & ")),2)                *             (IF (AdminComprasFacturasProveedores.tipo_doc_contable = 1,'-1','1'))) ASC"
    ElseIf Me.cboOrdenImporte.ListIndex = 1 Then
        ordenImporte = "(ROUND(AdminComprasFacturasProveedores.monto_neto + " _
                       & "AdminComprasFacturasProveedores.redondeo_iva +" _
                       & " AdminComprasFacturasProveedores.impuesto_interno +" _
                       & " (SELECT SUM(iva_calculado)         FROM sp.AdminComprasFacturasProveedoresIva acfpi " _
                       & " JOIN AdminComprasFacturasProveedores acfp ON acfpi.id_factura_proveedor=acfp.id " _
                       & " Where id_factura_proveedor = AdminComprasFacturasProveedores.Id " _
                       & " )                + " _
                       & " IF((SELECT SUM(valor)         FROM AdminComprasFacturasProveedoresPercepciones acfpp " _
                       & " JOIN AdminComprasFacturasProveedores acfp ON acfpp.id_factura_proveedor=acfp.id " _
                       & " Where id_factura_proveedor = AdminComprasFacturasProveedores.Id " _
                       & ") IS NULL,0," _
                       & " (SELECT SUM(valor)         FROM AdminComprasFacturasProveedoresPercepciones acfpp " _
                       & " JOIN AdminComprasFacturasProveedores acfp ON acfpp.id_factura_proveedor=acfp.id " _
                       & " Where id_factura_proveedor = AdminComprasFacturasProveedores.Id " _
                       & ")),2)                *             (IF (AdminComprasFacturasProveedores.tipo_doc_contable = 1,'-1','1'))) DESC"
    End If


    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Set facturas = DAOFacturaProveedor.FindAll(condition, , ordenImporte, Permisos.AdminFaPVerSoloPropias)

    Dim F As clsFacturaProveedor
    Dim total As Double
    Dim totalneto As Double
    Dim totIva As Double
    Dim totalno As Double
    Dim totalpercep As Double
    Dim totalsaldo As Double

    Dim c As Integer

    total = 0

    For Each F In facturas

        If F.tipoDocumentoContable = tipoDocumentoContable.notaCredito Then c = -1 Else c = 1
        total = total + MonedaConverter.Convertir(F.total * c, F.moneda.Id, MonedaConverter.Patron.Id)
        totalneto = totalneto + MonedaConverter.Convertir(F.Monto * c - F.TotalNetoGravadoDiscriminado(0) * c, F.moneda.Id, MonedaConverter.Patron.Id)
        totalno = totalno + MonedaConverter.Convertir(F.TotalNetoGravadoDiscriminado(0) * c, F.moneda.Id, MonedaConverter.Patron.Id)
        totIva = totIva + MonedaConverter.Convertir(F.TotalIVA * c, F.moneda.Id, MonedaConverter.Patron.Id)

        'Agrega DNEMER 03/02/2021
        totalpercep = totalpercep + F.totalPercepciones * c

        '(Factura.Total - (Factura.NetoGravadoAbonadoGlobal + Factura.OtrosAbonadoGlobal)) * i)
        totalsaldo = totalsaldo + ((F.total - (F.NetoGravadoAbonadoGlobal + F.OtrosAbonadoGlobal)) * c)

    Next

    Me.lblTotal = "Total Filtrado: " & FormatCurrency(funciones.FormatearDecimales(total))
    Me.lblTotalNoGravadoFiltrado = "Total No Gravado: " & FormatCurrency(funciones.FormatearDecimales(totalno))
    Me.lblNetoGravadoFiltrado = "Total Neto Gravado: " & FormatCurrency(funciones.FormatearDecimales(totalneto))
    Me.lblTotalIVA = "Total IVA: " & FormatCurrency(funciones.FormatearDecimales(totIva))
    Me.lblTotalNeto = "Total Neto: " & FormatCurrency(funciones.FormatearDecimales(funciones.RedondearDecimales(totalneto) + funciones.RedondearDecimales(totalno)))

    'Agregar totalizador de Pendientes
    Me.lblTotalPendiente = "Total Saldo: " & FormatCurrency(funciones.FormatearDecimales(funciones.RedondearDecimales(totalsaldo))) & ""

    'Agregar totalizador de Percepciones
    Me.lblTotalPercepciones = "Total Percepciones: " & FormatCurrency(funciones.FormatearDecimales(totalpercep))

    grilla.ItemCount = facturas.count

    GridEXHelper.AutoSizeColumns Me.grilla, True

    Me.caption = "Cbtes. filtrados [Cantidad: " & facturas.count & "]"

End Sub



Private Sub Form_Resize()
    On Error Resume Next
    Me.grilla.Width = Me.ScaleWidth - 50
    Me.grilla.Height = Me.ScaleHeight - 1800

End Sub

Private Sub Form_Terminate()
    Set vFactura = Nothing
    Channel.RemoverSuscripcionTotal Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Channel.RemoverSuscripcionTotal Me
End Sub

Private Sub grilla_ColumnHeaderClick(ByVal Column As GridEX20.JSColumn)
    GridEXHelper.ColumnHeaderClick Me.grilla, Column
End Sub

Private Sub grilla_DblClick()
    verDetalle_Click
End Sub

Private Sub grilla_FetchIcon(ByVal rowIndex As Long, ByVal ColIndex As Integer, ByVal RowBookmark As Variant, ByVal IconIndex As GridEX20.JSRetInteger)
    If ColIndex = 15 And m_Archivos.item(Factura.Id) > 0 Then IconIndex = 1

End Sub

Private Sub grilla_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Me.grilla.ItemCount > 0 Then
        If Button = 2 Then
            SeleccionarFactura
            Me.finalizar.Enabled = (Factura.estado = EstadoFacturaProveedor.EnProceso)
            Me.editar.Enabled = (Factura.estado = EstadoFacturaProveedor.EnProceso)
            Me.mnuPagarEnEfectivo.Enabled = (Factura.estado = EstadoFacturaProveedor.Aprobada)
            Me.mnuEliminar.Enabled = (funciones.GetUserObj.usuario = "karinrz" Or funciones.GetUserObj.usuario = "nicolasba" Or funciones.GetUserObj.usuario = "diegonr" Or funciones.GetUserObj.usuario = "natalilo")

            If (Factura.estado = Saldada Or Factura.estado = pagoParcial) Then
                 
                If Factura.LiquidacionesCajaId = "-" Then
                    
                    Me.MnuVerOP.Enabled = True
                    Me.MnuVerOP.Visible = True
                    
                    Me.MnuVerLIQ.Enabled = False
                    Me.MnuVerLIQ.Visible = False
                    
                Else
                
                    Me.MnuVerLIQ.Enabled = True
                    Me.MnuVerLIQ.Visible = True
                    
                    Me.MnuVerOP.Enabled = False
                    Me.MnuVerOP.Visible = False
                    
                End If
            Else
                    Me.MnuVerOP.Enabled = False
                    Me.MnuVerOP.Visible = False
                    Me.MnuVerLIQ.Enabled = False
                    Me.MnuVerLIQ.Visible = False
            End If

            Me.PopupMenu menu

        End If
    End If
End Sub

Private Sub grilla_RowFormat(RowBuffer As GridEX20.JSRowData)
    On Error GoTo err1
    Set Factura = facturas(RowBuffer.rowIndex)

    If Factura.estado = EstadoFacturaProveedor.Aprobada Then
        RowBuffer.CellStyle(15) = "EstadoAprobado"
    ElseIf Factura.estado = EstadoFacturaProveedor.EnProceso Then
        RowBuffer.CellStyle(15) = " EstadoEnProceso"
    ElseIf Factura.estado = EstadoFacturaProveedor.Saldada Then
        RowBuffer.CellStyle(15) = "EstadoSaldado"
    End If
    Exit Sub
err1:
End Sub

Private Sub grilla_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)

    Set Factura = facturas.item(rowIndex)

    Dim i As Integer

    If Factura.tipoDocumentoContable = tipoDocumentoContable.notaCredito Then i = -1 Else i = 1

    With Factura

        If IsSomething(Factura.Proveedor) Then
            Values(1) = funciones.RazonSocialFormateada(Factura.Proveedor.RazonSocial)
        End If

        Values(2) = enums.EnumTipoDocumentoContableShort(Factura.tipoDocumentoContable)
        Values(3) = Factura.configFactura.TipoFactura
        Values(4) = Factura.numero
        Values(5) = Factura.FEcha
        Values(6) = Factura.moneda.NombreCorto
        Values(7) = Replace(FormatCurrency(funciones.FormatearDecimales(Factura.Monto - Factura.TotalNetoGravadoDiscriminado(0)) * i), "$", "")
        Values(8) = Replace(FormatCurrency(funciones.FormatearDecimales(Factura.TotalIVA) * i), "$", "")
        Values(9) = Replace(FormatCurrency(funciones.FormatearDecimales(Factura.TotalNetoGravadoDiscriminado(0)) * i), "$", "")
        Values(10) = Replace(FormatCurrency(funciones.FormatearDecimales(Factura.totalPercepciones) * i), "$", "")
        Values(11) = Replace(FormatCurrency(funciones.FormatearDecimales(Factura.ImpuestoInterno) * i), "$", "")
        Values(12) = Replace(FormatCurrency(funciones.FormatearDecimales(Factura.total) * i), "$", "")

        'ESTO MUESTRA TRUE O FALSE
        'Values(12) = (funciones.FormatearDecimales(Factura.Total) * i) > 2000000

        'ESTO MUESTRA SOLO LOS VALORES MAYORES A DOS MILLONES, LOS DEMAS LOS DEJA VACIOS

        If (funciones.FormatearDecimales(Factura.total) * i) > 2000000 Then
            Values(12) = Replace(FormatCurrency(funciones.FormatearDecimales(Factura.total) * i), "$", "")
        End If

        'Values(14) = "DESARROLLANDOSE..."
        'Values(14) = Replace(FormatCurrency(funciones.FormatearDecimales(Factura.Total - (Factura.NetoGravadoAbonadoGlobal + Factura.OtrosAbonadoGlobal)) * i), "$", "")
        
        Values(14) = Replace(FormatCurrency(funciones.FormatearDecimales(Factura.TotalPendiente - Factura.TotalAbonadoGlobal) * i), "$", "")
        If Factura.cuentasContables.count > 0 Then
            Values(13) = Factura.cuentasContables.item(1).cuentas.codigo
        End If

        Values(15) = enums.enumEstadoFacturaProveedor(Factura.estado)

        If Factura.FormaPagoCuentaCorriente Then
            Values(16) = "Cta. Cte."
        Else
            Values(16) = "Contado"
        End If

        If Factura.estado = EstadoFacturaProveedor.Saldada Or Factura.estado = EstadoFacturaProveedor.pagoParcial Then
            Values(17) = Factura.OrdenesPagoId
            Values(18) = Factura.LiquidacionesCajaId
        End If

        Values(19) = Factura.UsuarioCarga.usuario
        Values(20) = Factura.TipoCambio
        Values(21) = "(" & Val(m_Archivos.item(Factura.Id)) & ")"
        Values(22) = Factura.Id

    End With

End Sub


Private Property Get ISuscriber_id() As String
    ISuscriber_id = vId
End Property


Private Function ISuscriber_Notificarse(EVENTO As clsEventoObserver) As Variant
    If EVENTO.EVENTO = agregar_ Then
        facturas.Add EVENTO.Elemento
        Me.grilla.ItemCount = facturas.count
    ElseIf EVENTO.EVENTO = modificar_ Then
        Dim rectmp As clsFacturaProveedor
        Dim tmp As clsFacturaProveedor
        Set tmp = EVENTO.Elemento

        For i = facturas.count To 1 Step -1
            If facturas(i).Id = tmp.Id Then
                Set rectmp = facturas(i)
                rectmp.Id = tmp.Id
                rectmp.estado = tmp.estado
                rectmp.Proveedor = tmp.Proveedor
                rectmp.FEcha = tmp.FEcha
                rectmp.ImpuestoInterno = tmp.ImpuestoInterno
                rectmp.cuentasContables = tmp.cuentasContables
                rectmp.IvaAplicado = tmp.IvaAplicado
                rectmp.percepciones = tmp.percepciones
                rectmp.Redondeo = tmp.Redondeo
                rectmp.Monto = tmp.Monto
                rectmp.numero = tmp.numero
                rectmp.FormaPagoCuentaCorriente = tmp.FormaPagoCuentaCorriente
                Set rectmp.moneda = tmp.moneda
                Me.grilla.RefreshRowIndex i
                Exit For
            End If
        Next
    End If
End Function


Private Sub mnuArchivos_Click()
    Dim archi As New frmArchivos2
    archi.Origen = OrigenArchivos.OA_FacturaProveedor
    archi.ObjetoId = Factura.Id
    archi.caption = Factura.NumeroFormateado
    archi.Show

End Sub


Private Sub mnuEliminar_Click()
    If MsgBox("¿Está seguro de eliminar la " & Factura.NumeroFormateado & " de " & Factura.Proveedor.RazonSocial & "?", vbInformation + vbYesNo) = vbYes Then
        If DAOFacturaProveedor.Delete(Factura.Id) Then
            MsgBox "Factura eliminada.", vbInformation
            llenarGrilla
        Else
            MsgBox "No se pudo eliminar la factura.", vbCritical
        End If
    End If
End Sub


Private Sub mnuPagarEnEfectivo_Click()
    If MsgBox("¿Está seguro de abonar en efectivo el comprobante " & Factura.NumeroFormateado & " de " & Factura.moneda.NombreCorto & " " & Factura.total & "?", vbInformation + vbYesNo) = vbYes Then

        MsgBox "Se creará una OP con fecha " + CStr(Factura.FEcha)
        If IsDate(Factura.FEcha) Then
            If DAOFacturaProveedor.PagarEnEfectivo(Factura, Factura.FEcha, True) Then
                MsgBox "El pago de la factura ha sido registrado con la orden de pago Nº " & DAOOrdenPago.FindLast().Id & ".", vbInformation
                llenarGrilla
                Me.txtComprobante.SetFocus
            Else
                MsgBox "No se pudo registrar el pago de la factura.", vbCritical
            End If
        Else
            MsgBox "Debe ingresar una fecha de pago válida.", vbExclamation
        End If
    End If
End Sub


Private Sub mnuScan_Click()
    On Error Resume Next
    Dim archivos As New classArchivos
    If archivos.escanearDocumento(OrigenArchivos.OA_FacturaProveedor, Factura.Id) Then
        Set m_Archivos = DAOArchivo.GetCantidadArchivosPorReferencia(OA_FacturaProveedor)
        Me.grilla.RefreshRowIndex (Factura.Id)

    End If

End Sub

Private Sub MnuVerLIQ_Click()
    Dim f123 As New frmAdminComprasListaLIQSegunCbte
    f123.Factura = Factura
    f123.Show

End Sub

Private Sub MnuVerOP_Click()
    Dim f123 As New frmAdminComprasListaOPSegunCbte
    f123.Factura = Factura
    f123.Show

End Sub


Private Sub txtComprobante_GotFocus()
    foco Me.txtComprobante
End Sub


Private Sub verDetalle_Click()
    SeleccionarFactura
    Dim frm As frmAdminComprasNuevaFCProveedor
    Set frm = New frmAdminComprasNuevaFCProveedor

    frm.ver = True
    frm.Factura = Factura
    frm.Show
End Sub


Private Sub verHistorial_Click()
    If grilla.ItemCount > 0 Then
        SeleccionarFactura
        Factura.Historial = DaoFacturaProveedorHistorial.getAllByIdFactura(Factura.Id)
        frmHistoriales.lista = Factura.Historial
        frmHistoriales.Show
    End If
End Sub

Private Sub SeleccionarFactura()
    On Error Resume Next
    Set Factura = facturas.item(grilla.rowIndex(grilla.row))
End Sub
