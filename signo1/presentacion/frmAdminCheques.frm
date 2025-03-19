VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmAdminCheques 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administración de cheques"
   ClientHeight    =   9435
   ClientLeft      =   8250
   ClientTop       =   2265
   ClientWidth     =   15360
   Icon            =   "frmAdminCheques.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9435
   ScaleWidth      =   15360
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   9555
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15375
      _Version        =   786432
      _ExtentX        =   27120
      _ExtentY        =   16854
      _StockProps     =   68
      Appearance      =   10
      Color           =   128
      PaintManager.BoldSelected=   -1  'True
      PaintManager.ShowIcons=   -1  'True
      ItemCount       =   4
      Item(0).Caption =   "Cartera"
      Item(0).ControlCount=   2
      Item(0).Control(0)=   "Frame3"
      Item(0).Control(1)=   "grpResultados(0)"
      Item(1).Caption =   "Administrar Chequeras"
      Item(1).ControlCount=   7
      Item(1).Control(0)=   "grid_chequeras"
      Item(1).Control(1)=   "grid_cheques"
      Item(1).Control(2)=   "GroupBox1"
      Item(1).Control(3)=   "cboProveedores"
      Item(1).Control(4)=   "btnFiltrar"
      Item(1).Control(5)=   "Label6"
      Item(1).Control(6)=   "PushButton1"
      Item(2).Caption =   "Cheques Propios Utilizados"
      Item(2).ControlCount=   2
      Item(2).Control(0)=   "GroupBox2"
      Item(2).Control(1)=   "grpResultadosPropios(1)"
      Item(3).Caption =   "Cheques 3eros Utilizados"
      Item(3).ControlCount=   2
      Item(3).Control(0)=   "GroupBox3"
      Item(3).Control(1)=   "grpResultados3ros(2)"
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   375
         Left            =   -57040
         TabIndex        =   119
         Top             =   600
         Visible         =   0   'False
         Width           =   375
         _Version        =   786432
         _ExtentX        =   661
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "X"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnFiltrar 
         Height          =   375
         Left            =   -56320
         TabIndex        =   117
         Top             =   600
         Visible         =   0   'False
         Width           =   1575
         _Version        =   786432
         _ExtentX        =   2778
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Filtrar"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cboProveedores 
         Height          =   315
         Left            =   -60880
         TabIndex        =   116
         Top             =   630
         Visible         =   0   'False
         Width           =   3735
         _Version        =   786432
         _ExtentX        =   6588
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Text            =   "ComboBox1"
      End
      Begin VB.Frame grpResultados3ros 
         Caption         =   "Resultados"
         Height          =   5175
         Index           =   2
         Left            =   -69880
         TabIndex        =   103
         Top             =   3480
         Visible         =   0   'False
         Width           =   15135
         Begin GridEX20.GridEX grdCheques3eros 
            Height          =   4665
            Left            =   240
            TabIndex        =   104
            Top             =   360
            Width           =   14655
            _ExtentX        =   25850
            _ExtentY        =   8229
            Version         =   "2.0"
            BoundColumnIndex=   ""
            ReplaceColumnIndex=   ""
            ColumnAutoResize=   -1  'True
            MethodHoldFields=   -1  'True
            AllowColumnDrag =   0   'False
            AllowEdit       =   0   'False
            GroupByBoxVisible=   0   'False
            DataMode        =   99
            ColumnHeaderHeight=   285
            IntProp1        =   0
            IntProp2        =   0
            IntProp7        =   0
            ColumnsCount    =   11
            Column(1)       =   "frmAdminCheques.frx":000C
            Column(2)       =   "frmAdminCheques.frx":0150
            Column(3)       =   "frmAdminCheques.frx":0288
            Column(4)       =   "frmAdminCheques.frx":03A0
            Column(5)       =   "frmAdminCheques.frx":0508
            Column(6)       =   "frmAdminCheques.frx":0668
            Column(7)       =   "frmAdminCheques.frx":07B0
            Column(8)       =   "frmAdminCheques.frx":08F8
            Column(9)       =   "frmAdminCheques.frx":0A68
            Column(10)      =   "frmAdminCheques.frx":0BBC
            Column(11)      =   "frmAdminCheques.frx":0D14
            FormatStylesCount=   6
            FormatStyle(1)  =   "frmAdminCheques.frx":0E3C
            FormatStyle(2)  =   "frmAdminCheques.frx":0F74
            FormatStyle(3)  =   "frmAdminCheques.frx":1024
            FormatStyle(4)  =   "frmAdminCheques.frx":10D8
            FormatStyle(5)  =   "frmAdminCheques.frx":11B0
            FormatStyle(6)  =   "frmAdminCheques.frx":1268
            ImageCount      =   0
            PrinterProperties=   "frmAdminCheques.frx":1348
         End
      End
      Begin VB.Frame grpResultadosPropios 
         Caption         =   "Resultados"
         Height          =   5175
         Index           =   1
         Left            =   -69880
         TabIndex        =   102
         Top             =   3480
         Visible         =   0   'False
         Width           =   15135
         Begin GridEX20.GridEX gridChequesEmitidos 
            Height          =   4665
            Left            =   240
            TabIndex        =   105
            Top             =   360
            Width           =   14655
            _ExtentX        =   25850
            _ExtentY        =   8229
            Version         =   "2.0"
            BoundColumnIndex=   ""
            ReplaceColumnIndex=   ""
            ColumnAutoResize=   -1  'True
            MethodHoldFields=   -1  'True
            AllowColumnDrag =   0   'False
            GroupByBoxVisible=   0   'False
            DataMode        =   99
            ColumnHeaderHeight=   285
            IntProp1        =   0
            IntProp2        =   0
            IntProp7        =   0
            ColumnsCount    =   11
            Column(1)       =   "frmAdminCheques.frx":1520
            Column(2)       =   "frmAdminCheques.frx":16D0
            Column(3)       =   "frmAdminCheques.frx":17E8
            Column(4)       =   "frmAdminCheques.frx":1920
            Column(5)       =   "frmAdminCheques.frx":1A80
            Column(6)       =   "frmAdminCheques.frx":1BD8
            Column(7)       =   "frmAdminCheques.frx":1D40
            Column(8)       =   "frmAdminCheques.frx":1EA8
            Column(9)       =   "frmAdminCheques.frx":1FC8
            Column(10)      =   "frmAdminCheques.frx":20F8
            Column(11)      =   "frmAdminCheques.frx":2230
            FormatStylesCount=   6
            FormatStyle(1)  =   "frmAdminCheques.frx":2388
            FormatStyle(2)  =   "frmAdminCheques.frx":24C0
            FormatStyle(3)  =   "frmAdminCheques.frx":2570
            FormatStyle(4)  =   "frmAdminCheques.frx":2624
            FormatStyle(5)  =   "frmAdminCheques.frx":26FC
            FormatStyle(6)  =   "frmAdminCheques.frx":27B4
            ImageCount      =   0
            PrinterProperties=   "frmAdminCheques.frx":2894
         End
      End
      Begin VB.Frame grpResultados 
         Caption         =   "Resultados"
         Height          =   5175
         Index           =   0
         Left            =   120
         TabIndex        =   63
         Top             =   3480
         Width           =   15135
         Begin GridEX20.GridEX grid_cartera_cheques 
            Height          =   4665
            Left            =   240
            TabIndex        =   64
            Top             =   360
            Width           =   14655
            _ExtentX        =   25850
            _ExtentY        =   8229
            Version         =   "2.0"
            DefaultGroupMode=   1
            BoundColumnIndex=   ""
            ReplaceColumnIndex=   ""
            PreviewColumn   =   "observaciones"
            PreviewRowLines =   1
            ColumnAutoResize=   -1  'True
            ReadOnly        =   -1  'True
            MethodHoldFields=   -1  'True
            ContScroll      =   -1  'True
            AllowCardSizing =   0   'False
            AllowEdit       =   0   'False
            GroupByBoxVisible=   0   'False
            DataMode        =   99
            ColumnHeaderHeight=   285
            IntProp1        =   0
            IntProp2        =   0
            IntProp7        =   0
            ColumnsCount    =   9
            Column(1)       =   "frmAdminCheques.frx":2A6C
            Column(2)       =   "frmAdminCheques.frx":2C0C
            Column(3)       =   "frmAdminCheques.frx":2D78
            Column(4)       =   "frmAdminCheques.frx":2F00
            Column(5)       =   "frmAdminCheques.frx":30D4
            Column(6)       =   "frmAdminCheques.frx":3234
            Column(7)       =   "frmAdminCheques.frx":3390
            Column(8)       =   "frmAdminCheques.frx":3500
            Column(9)       =   "frmAdminCheques.frx":36CC
            FormatStylesCount=   6
            FormatStyle(1)  =   "frmAdminCheques.frx":386C
            FormatStyle(2)  =   "frmAdminCheques.frx":39A4
            FormatStyle(3)  =   "frmAdminCheques.frx":3A54
            FormatStyle(4)  =   "frmAdminCheques.frx":3B08
            FormatStyle(5)  =   "frmAdminCheques.frx":3BE0
            FormatStyle(6)  =   "frmAdminCheques.frx":3C98
            ImageCount      =   0
            PrinterProperties=   "frmAdminCheques.frx":3D78
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Búsqueda"
         Height          =   3015
         Left            =   120
         TabIndex        =   33
         Top             =   360
         Width           =   15135
         Begin XtremeSuiteControls.PushButton btnBorrarNumeroCartera 
            Height          =   315
            Left            =   2880
            TabIndex        =   115
            Top             =   480
            Width           =   375
            _Version        =   786432
            _ExtentX        =   661
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "X"
            UseVisualStyle  =   -1  'True
         End
         Begin VB.TextBox txtOrigen 
            Height          =   315
            Left            =   240
            TabIndex        =   65
            Top             =   2280
            Width           =   2535
         End
         Begin XtremeSuiteControls.PushButton btnBorrarOrigen 
            Height          =   315
            Left            =   2880
            TabIndex        =   61
            Top             =   2280
            Width           =   375
            _Version        =   786432
            _ExtentX        =   661
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "X"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton btnBorrarClasificacion 
            Height          =   315
            Left            =   2880
            TabIndex        =   57
            Top             =   1680
            Width           =   375
            _Version        =   786432
            _ExtentX        =   661
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "X"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton btnBorrarBanco 
            Height          =   315
            Left            =   2880
            TabIndex        =   56
            Top             =   1080
            Width           =   375
            _Version        =   786432
            _ExtentX        =   661
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "X"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.ComboBox cboClasificacion 
            Height          =   315
            Left            =   240
            TabIndex        =   55
            Top             =   1680
            Width           =   2535
            _Version        =   786432
            _ExtentX        =   4471
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Text            =   "cboClasificacion"
         End
         Begin XtremeSuiteControls.ComboBox cboBancoCartera 
            Height          =   315
            Left            =   240
            TabIndex        =   54
            Top             =   1080
            Width           =   2535
            _Version        =   786432
            _ExtentX        =   4471
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Text            =   "cboBanco"
         End
         Begin VB.Frame Frame 
            Height          =   2555
            Index           =   0
            Left            =   11160
            TabIndex        =   50
            Top             =   120
            Width           =   3855
            Begin XtremeSuiteControls.PushButton btnBuscarEnCartera 
               Default         =   -1  'True
               Height          =   495
               Index           =   0
               Left            =   120
               TabIndex        =   51
               Top             =   1920
               Width           =   1575
               _Version        =   786432
               _ExtentX        =   2778
               _ExtentY        =   873
               _StockProps     =   79
               Caption         =   "Buscar"
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
            Begin XtremeSuiteControls.PushButton btnExportarCartera 
               Height          =   495
               Index           =   1
               Left            =   2040
               TabIndex        =   52
               Top             =   1920
               Width           =   1575
               _Version        =   786432
               _ExtentX        =   2778
               _ExtentY        =   873
               _StockProps     =   79
               Caption         =   "Exportar"
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.ProgressBar ProgressBar 
               Height          =   375
               Index           =   0
               Left            =   120
               TabIndex        =   62
               Top             =   1440
               Width           =   3495
               _Version        =   786432
               _ExtentX        =   6165
               _ExtentY        =   661
               _StockProps     =   93
               Appearance      =   6
            End
         End
         Begin VB.TextBox txtNumeroChequeCartera 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   240
            TabIndex        =   34
            Top             =   480
            Width           =   2535
         End
         Begin XtremeSuiteControls.GroupBox GroFechaComprobante 
            Height          =   1215
            Index           =   1
            Left            =   5400
            TabIndex        =   36
            Top             =   120
            Width           =   4695
            _Version        =   786432
            _ExtentX        =   8281
            _ExtentY        =   2143
            _StockProps     =   79
            Caption         =   "Fecha Vencimiento"
            BackColor       =   16744576
            Appearance      =   4
            Begin XtremeSuiteControls.ComboBox cboRangosVtoCartera 
               Height          =   315
               Index           =   0
               Left            =   720
               TabIndex        =   39
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
            Begin XtremeSuiteControls.DateTimePicker dtpDesdeVtoCartera 
               Height          =   315
               Index           =   1
               Left            =   720
               TabIndex        =   37
               Top             =   720
               Width           =   1470
               _Version        =   786432
               _ExtentX        =   2593
               _ExtentY        =   556
               _StockProps     =   68
               CheckBox        =   -1  'True
               Format          =   1
               CurrentDate     =   45190.4376157407
            End
            Begin XtremeSuiteControls.DateTimePicker dtpHastaVtoCartera 
               Height          =   315
               Index           =   1
               Left            =   2925
               TabIndex        =   38
               Top             =   720
               Width           =   1470
               _Version        =   786432
               _ExtentX        =   2593
               _ExtentY        =   556
               _StockProps     =   68
               CheckBox        =   -1  'True
               Format          =   1
               CurrentDate     =   45190.4375810185
            End
            Begin XtremeSuiteControls.Label lblHasta 
               Height          =   195
               Index           =   1
               Left            =   2400
               TabIndex        =   42
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
            Begin XtremeSuiteControls.Label lblDesde 
               Height          =   195
               Index           =   1
               Left            =   165
               TabIndex        =   41
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
            Begin XtremeSuiteControls.Label lblRango 
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   40
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
         Begin XtremeSuiteControls.GroupBox GroFechaComprobante 
            Height          =   1215
            Index           =   0
            Left            =   5400
            TabIndex        =   43
            Top             =   1440
            Width           =   4695
            _Version        =   786432
            _ExtentX        =   8281
            _ExtentY        =   2143
            _StockProps     =   79
            Caption         =   "Fecha Recibido"
            BackColor       =   16744576
            Appearance      =   4
            Begin XtremeSuiteControls.ComboBox cboRangosRboCartera 
               Height          =   315
               Index           =   1
               Left            =   720
               TabIndex        =   46
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
            Begin XtremeSuiteControls.DateTimePicker dtpHastaRboCartera 
               Height          =   315
               Index           =   2
               Left            =   2925
               TabIndex        =   45
               Top             =   720
               Width           =   1470
               _Version        =   786432
               _ExtentX        =   2593
               _ExtentY        =   556
               _StockProps     =   68
               CheckBox        =   -1  'True
               Format          =   1
               CurrentDate     =   45190.4375578704
            End
            Begin XtremeSuiteControls.DateTimePicker dtpDesdeRboCartera 
               Height          =   315
               Index           =   2
               Left            =   720
               TabIndex        =   44
               Top             =   720
               Width           =   1470
               _Version        =   786432
               _ExtentX        =   2593
               _ExtentY        =   556
               _StockProps     =   68
               CheckBox        =   -1  'True
               Format          =   1
               CurrentDate     =   45190.4375231481
            End
            Begin XtremeSuiteControls.Label lblHasta 
               Height          =   195
               Index           =   0
               Left            =   2400
               TabIndex        =   49
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
            Begin XtremeSuiteControls.Label lblDesde 
               Height          =   195
               Index           =   0
               Left            =   165
               TabIndex        =   48
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
            Begin XtremeSuiteControls.Label lblRango 
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   47
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
         Begin VB.Label Label1 
            Caption         =   "Origen:"
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   60
            Top             =   2040
            Width           =   2535
         End
         Begin VB.Label Label1 
            Caption         =   "Clasificación:"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   59
            Top             =   1440
            Width           =   2535
         End
         Begin VB.Label Label1 
            Caption         =   "Banco:"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   58
            Top             =   870
            Width           =   2535
         End
         Begin VB.Label Label1 
            Caption         =   "Número:"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   53
            Top             =   240
            Width           =   2535
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox3 
         Height          =   3015
         Left            =   -69880
         TabIndex        =   28
         Top             =   360
         Visible         =   0   'False
         Width           =   15135
         _Version        =   786432
         _ExtentX        =   26696
         _ExtentY        =   5318
         _StockProps     =   79
         Caption         =   "Parámetros de búsqueda"
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.ComboBox cboClientes3erosUti 
            Height          =   315
            Left            =   240
            TabIndex        =   126
            Top             =   1680
            Width           =   3735
            _Version        =   786432
            _ExtentX        =   6588
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.PushButton PushButton4 
            Height          =   315
            Left            =   5040
            TabIndex        =   124
            Top             =   2280
            Width           =   375
            _Version        =   786432
            _ExtentX        =   661
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "X"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.ComboBox cboProveedores3eros 
            Height          =   315
            Left            =   2040
            TabIndex        =   123
            Top             =   2280
            Width           =   3015
            _Version        =   786432
            _ExtentX        =   5318
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.PushButton btnBorrarOPTerceros 
            Height          =   315
            Left            =   1440
            TabIndex        =   112
            Top             =   2280
            Width           =   375
            _Version        =   786432
            _ExtentX        =   661
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "X"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton btnBorrarOrigenTerceros 
            Height          =   315
            Left            =   3960
            TabIndex        =   111
            Top             =   1680
            Width           =   375
            _Version        =   786432
            _ExtentX        =   661
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "X"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton btnBorrarBancosTerceros 
            Height          =   315
            Left            =   2760
            TabIndex        =   110
            Top             =   1080
            Width           =   375
            _Version        =   786432
            _ExtentX        =   661
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "X"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton btnBorrarNumeroTerceros 
            Height          =   315
            Left            =   2760
            TabIndex        =   109
            Top             =   480
            Width           =   375
            _Version        =   786432
            _ExtentX        =   661
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "X"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.ComboBox cboBancos3ero 
            Height          =   315
            Left            =   240
            TabIndex        =   108
            Top             =   1080
            Width           =   2535
            _Version        =   786432
            _ExtentX        =   4471
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Text            =   "cboBancos3eros"
         End
         Begin VB.Frame Frame 
            Height          =   2775
            Index           =   2
            Left            =   11160
            TabIndex        =   98
            Top             =   120
            Width           =   3855
            Begin XtremeSuiteControls.PushButton btnBuscar 
               Height          =   495
               Index           =   0
               Left            =   120
               TabIndex        =   99
               Top             =   2160
               Width           =   1575
               _Version        =   786432
               _ExtentX        =   2778
               _ExtentY        =   873
               _StockProps     =   79
               Caption         =   "Buscar"
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
            Begin XtremeSuiteControls.PushButton btnExportar 
               Height          =   495
               Index           =   1
               Left            =   2040
               TabIndex        =   100
               Top             =   2160
               Width           =   1575
               _Version        =   786432
               _ExtentX        =   2778
               _ExtentY        =   873
               _StockProps     =   79
               Caption         =   "Exportar"
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.ProgressBar ProgressBar 
               Height          =   375
               Index           =   2
               Left            =   120
               TabIndex        =   101
               Top             =   1680
               Width           =   3495
               _Version        =   786432
               _ExtentX        =   6165
               _ExtentY        =   661
               _StockProps     =   93
               Appearance      =   6
            End
         End
         Begin VB.TextBox txtNumeroOP 
            Height          =   315
            Left            =   240
            TabIndex        =   30
            Top             =   2280
            Width           =   1185
         End
         Begin VB.TextBox txtNumeroCheque3ero 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   240
            TabIndex        =   29
            Top             =   480
            Width           =   2535
         End
         Begin XtremeSuiteControls.GroupBox GroFechaComprobante 
            Height          =   1215
            Index           =   4
            Left            =   6360
            TabIndex        =   80
            Top             =   240
            Width           =   4695
            _Version        =   786432
            _ExtentX        =   8281
            _ExtentY        =   2143
            _StockProps     =   79
            Caption         =   "Fecha Vencimiento"
            BackColor       =   16744576
            Appearance      =   4
            Begin XtremeSuiteControls.ComboBox cboRangosVtoTerceros 
               Height          =   315
               Index           =   2
               Left            =   720
               TabIndex        =   81
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
            Begin XtremeSuiteControls.DateTimePicker dtpDesdeVtoTerceros 
               Height          =   315
               Index           =   5
               Left            =   720
               TabIndex        =   82
               Top             =   720
               Width           =   1470
               _Version        =   786432
               _ExtentX        =   2593
               _ExtentY        =   556
               _StockProps     =   68
               CheckBox        =   -1  'True
               Format          =   1
               CurrentDate     =   45190.4376157407
            End
            Begin XtremeSuiteControls.DateTimePicker dtpHastaVtoTerceros 
               Height          =   315
               Index           =   5
               Left            =   2925
               TabIndex        =   83
               Top             =   720
               Width           =   1470
               _Version        =   786432
               _ExtentX        =   2593
               _ExtentY        =   556
               _StockProps     =   68
               CheckBox        =   -1  'True
               Format          =   1
               CurrentDate     =   45190.4375810185
            End
            Begin XtremeSuiteControls.Label lblRango 
               Height          =   195
               Index           =   4
               Left            =   120
               TabIndex        =   86
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
            Begin XtremeSuiteControls.Label lblDesde 
               Height          =   195
               Index           =   4
               Left            =   165
               TabIndex        =   85
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
            Begin XtremeSuiteControls.Label lblHasta 
               Height          =   195
               Index           =   4
               Left            =   2400
               TabIndex        =   84
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
         Begin XtremeSuiteControls.GroupBox GroFechaComprobante 
            Height          =   1215
            Index           =   5
            Left            =   6360
            TabIndex        =   87
            Top             =   1560
            Width           =   4695
            _Version        =   786432
            _ExtentX        =   8281
            _ExtentY        =   2143
            _StockProps     =   79
            Caption         =   "Fecha Emitido"
            BackColor       =   16744576
            Appearance      =   4
            Begin XtremeSuiteControls.ComboBox cboRangosRboEmitido 
               Height          =   315
               Index           =   2
               Left            =   720
               TabIndex        =   88
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
            Begin XtremeSuiteControls.DateTimePicker dtpHastaRboEmitido 
               Height          =   315
               Index           =   6
               Left            =   2925
               TabIndex        =   89
               Top             =   720
               Width           =   1470
               _Version        =   786432
               _ExtentX        =   2593
               _ExtentY        =   556
               _StockProps     =   68
               CheckBox        =   -1  'True
               Format          =   1
               CurrentDate     =   45190.4375578704
            End
            Begin XtremeSuiteControls.DateTimePicker dtpDesdeRboEmitido 
               Height          =   315
               Index           =   6
               Left            =   720
               TabIndex        =   90
               Top             =   720
               Width           =   1470
               _Version        =   786432
               _ExtentX        =   2593
               _ExtentY        =   556
               _StockProps     =   68
               CheckBox        =   -1  'True
               Format          =   1
               CurrentDate     =   45190.4375231481
            End
            Begin XtremeSuiteControls.Label lblRango 
               Height          =   195
               Index           =   5
               Left            =   120
               TabIndex        =   93
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
            Begin XtremeSuiteControls.Label lblDesde 
               Height          =   195
               Index           =   5
               Left            =   165
               TabIndex        =   92
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
            Begin XtremeSuiteControls.Label lblHasta 
               Height          =   195
               Index           =   5
               Left            =   2400
               TabIndex        =   91
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
         Begin XtremeSuiteControls.Label Label7 
            Height          =   180
            Left            =   2040
            TabIndex        =   125
            Top             =   2040
            Width           =   2175
            _Version        =   786432
            _ExtentX        =   3836
            _ExtentY        =   317
            _StockProps     =   79
            Caption         =   "Destino:"
         End
         Begin VB.Label Label 
            Caption         =   "Origen/Cliente:"
            Height          =   255
            Left            =   240
            TabIndex        =   107
            Top             =   1440
            Width           =   2535
         End
         Begin XtremeSuiteControls.Label lblOP 
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   106
            Top             =   2010
            Width           =   855
            _Version        =   786432
            _ExtentX        =   1508
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "O.P:"
         End
         Begin XtremeSuiteControls.Label lblOP 
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   32
            Top             =   840
            Width           =   2415
            _Version        =   786432
            _ExtentX        =   4260
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Banco:"
         End
         Begin XtremeSuiteControls.Label lblNumero 
            Height          =   135
            Left            =   240
            TabIndex        =   31
            Top             =   270
            Width           =   615
            _Version        =   786432
            _ExtentX        =   1085
            _ExtentY        =   238
            _StockProps     =   79
            Caption         =   "Número:"
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox2 
         Height          =   3015
         Left            =   -69880
         TabIndex        =   16
         Top             =   360
         Visible         =   0   'False
         Width           =   15135
         _Version        =   786432
         _ExtentX        =   26696
         _ExtentY        =   5318
         _StockProps     =   79
         Caption         =   "Parámetros de búsqueda"
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.PushButton PushButton2 
            Height          =   315
            Left            =   5060
            TabIndex        =   122
            Top             =   2280
            Width           =   375
            _Version        =   786432
            _ExtentX        =   661
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "X"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.ComboBox cboProveedoresPropios 
            Height          =   315
            Left            =   2040
            TabIndex        =   120
            Top             =   2280
            Width           =   3015
            _Version        =   786432
            _ExtentX        =   5318
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.PushButton btnBorrarNumeroPropios 
            Height          =   315
            Left            =   2760
            TabIndex        =   114
            Top             =   480
            Width           =   375
            _Version        =   786432
            _ExtentX        =   661
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "X"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton btnBorrarOPPropios 
            Height          =   315
            Left            =   1440
            TabIndex        =   113
            Top             =   2280
            Width           =   375
            _Version        =   786432
            _ExtentX        =   661
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "X"
            UseVisualStyle  =   -1  'True
         End
         Begin VB.Frame Frame 
            Height          =   2775
            Index           =   1
            Left            =   11160
            TabIndex        =   94
            Top             =   120
            Width           =   3855
            Begin XtremeSuiteControls.PushButton btnBuscarChePropios 
               Height          =   495
               Index           =   1
               Left            =   120
               TabIndex        =   95
               Top             =   2160
               Width           =   1575
               _Version        =   786432
               _ExtentX        =   2778
               _ExtentY        =   873
               _StockProps     =   79
               Caption         =   "Buscar"
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
            Begin XtremeSuiteControls.PushButton btnExportarChePropios 
               Height          =   495
               Index           =   0
               Left            =   2040
               TabIndex        =   96
               Top             =   2160
               Width           =   1575
               _Version        =   786432
               _ExtentX        =   2778
               _ExtentY        =   873
               _StockProps     =   79
               Caption         =   "Exportar"
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.ProgressBar ProgressBar 
               Height          =   375
               Index           =   1
               Left            =   120
               TabIndex        =   97
               Top             =   1680
               Width           =   3495
               _Version        =   786432
               _ExtentX        =   6165
               _ExtentY        =   661
               _StockProps     =   93
               Appearance      =   6
            End
         End
         Begin VB.TextBox txtIdOP 
            Height          =   285
            Left            =   240
            TabIndex        =   27
            Top             =   2280
            Width           =   1185
         End
         Begin VB.TextBox txtNroChequePropio 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   240
            TabIndex        =   25
            Top             =   480
            Width           =   2535
         End
         Begin XtremeSuiteControls.CheckBox chkIngresados 
            Height          =   315
            Left            =   240
            TabIndex        =   17
            Top             =   2640
            Width           =   1395
            _Version        =   786432
            _ExtentX        =   2461
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "Ingresados"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.ComboBox cboBancos1 
            Height          =   315
            Left            =   240
            TabIndex        =   18
            Top             =   1080
            Width           =   2535
            _Version        =   786432
            _ExtentX        =   4471
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.PushButton CMDsINCliente 
            Height          =   315
            Left            =   2760
            TabIndex        =   19
            Top             =   1080
            Width           =   375
            _Version        =   786432
            _ExtentX        =   661
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "X"
            BackColor       =   12632256
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.ComboBox cboChequera2 
            Height          =   315
            Left            =   240
            TabIndex        =   22
            Top             =   1680
            Width           =   4485
            _Version        =   786432
            _ExtentX        =   7911
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.PushButton PushButton3 
            Height          =   315
            Left            =   4730
            TabIndex        =   23
            Top             =   1680
            Width           =   375
            _Version        =   786432
            _ExtentX        =   661
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "X"
            BackColor       =   12632256
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.GroupBox GroFechaComprobante 
            Height          =   1215
            Index           =   2
            Left            =   6360
            TabIndex        =   66
            Top             =   240
            Width           =   4695
            _Version        =   786432
            _ExtentX        =   8281
            _ExtentY        =   2143
            _StockProps     =   79
            Caption         =   "Fecha Vencimiento"
            BackColor       =   16744576
            Appearance      =   4
            Begin XtremeSuiteControls.ComboBox cboRangosVtoPropios 
               Height          =   315
               Index           =   1
               Left            =   720
               TabIndex        =   67
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
            Begin XtremeSuiteControls.DateTimePicker dtpDesdeVtoPropios 
               Height          =   315
               Index           =   3
               Left            =   720
               TabIndex        =   68
               Top             =   720
               Width           =   1470
               _Version        =   786432
               _ExtentX        =   2593
               _ExtentY        =   556
               _StockProps     =   68
               CheckBox        =   -1  'True
               Format          =   1
               CurrentDate     =   45190.4376157407
            End
            Begin XtremeSuiteControls.DateTimePicker dtpHastaVtoPropios 
               Height          =   315
               Index           =   3
               Left            =   2925
               TabIndex        =   69
               Top             =   720
               Width           =   1470
               _Version        =   786432
               _ExtentX        =   2593
               _ExtentY        =   556
               _StockProps     =   68
               CheckBox        =   -1  'True
               Format          =   1
               CurrentDate     =   45190.4375810185
            End
            Begin XtremeSuiteControls.Label lblRango 
               Height          =   195
               Index           =   2
               Left            =   120
               TabIndex        =   72
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
            Begin XtremeSuiteControls.Label lblDesde 
               Height          =   195
               Index           =   2
               Left            =   165
               TabIndex        =   71
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
            Begin XtremeSuiteControls.Label lblHasta 
               Height          =   195
               Index           =   2
               Left            =   2400
               TabIndex        =   70
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
         Begin XtremeSuiteControls.GroupBox GroFechaComprobante 
            Height          =   1215
            Index           =   3
            Left            =   6360
            TabIndex        =   73
            Top             =   1560
            Width           =   4695
            _Version        =   786432
            _ExtentX        =   8281
            _ExtentY        =   2143
            _StockProps     =   79
            Caption         =   "Fecha Emitido"
            BackColor       =   16744576
            Appearance      =   4
            Begin XtremeSuiteControls.ComboBox cboRangosRboPropios 
               Height          =   315
               Index           =   0
               Left            =   720
               TabIndex        =   74
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
            Begin XtremeSuiteControls.DateTimePicker dtpHastaRboPropios 
               Height          =   315
               Index           =   4
               Left            =   2925
               TabIndex        =   75
               Top             =   720
               Width           =   1470
               _Version        =   786432
               _ExtentX        =   2593
               _ExtentY        =   556
               _StockProps     =   68
               CheckBox        =   -1  'True
               Format          =   1
               CurrentDate     =   45190.4375578704
            End
            Begin XtremeSuiteControls.DateTimePicker dtpDesdeRboPropios 
               Height          =   315
               Index           =   4
               Left            =   720
               TabIndex        =   76
               Top             =   720
               Width           =   1470
               _Version        =   786432
               _ExtentX        =   2593
               _ExtentY        =   556
               _StockProps     =   68
               CheckBox        =   -1  'True
               Format          =   1
               CurrentDate     =   45190.4375231481
            End
            Begin XtremeSuiteControls.Label lblRango 
               Height          =   195
               Index           =   3
               Left            =   120
               TabIndex        =   79
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
            Begin XtremeSuiteControls.Label lblDesde 
               Height          =   195
               Index           =   3
               Left            =   165
               TabIndex        =   78
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
            Begin XtremeSuiteControls.Label lblHasta 
               Height          =   195
               Index           =   3
               Left            =   2400
               TabIndex        =   77
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
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Proveedor:"
            Height          =   180
            Index           =   1
            Left            =   2040
            TabIndex        =   121
            Top             =   2040
            Width           =   1905
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "O.P:"
            Height          =   180
            Index           =   0
            Left            =   240
            TabIndex        =   26
            Top             =   2040
            Width           =   465
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Número:"
            Height          =   255
            Left            =   240
            TabIndex        =   24
            Top             =   240
            Width           =   705
         End
         Begin VB.Label Label8 
            Caption         =   "Chequera:"
            Height          =   240
            Left            =   240
            TabIndex        =   21
            Top             =   1440
            Width           =   4410
         End
         Begin VB.Label lblBanco 
            AutoSize        =   -1  'True
            Caption         =   "Banco:"
            Height          =   195
            Left            =   240
            TabIndex        =   20
            Top             =   870
            Width           =   510
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   2220
         Left            =   -69760
         TabIndex        =   1
         Top             =   6240
         Visible         =   0   'False
         Width           =   7365
         _Version        =   786432
         _ExtentX        =   12991
         _ExtentY        =   3916
         _StockProps     =   79
         Caption         =   "Crear Chequera"
         UseVisualStyle  =   -1  'True
         Begin VB.TextBox txtDesde 
            Height          =   285
            Left            =   990
            TabIndex        =   7
            Text            =   "0"
            Top             =   705
            Width           =   1035
         End
         Begin VB.TextBox txtHasta 
            Height          =   285
            Left            =   2910
            TabIndex        =   6
            Text            =   "0"
            Top             =   705
            Width           =   1020
         End
         Begin VB.TextBox txtNumero 
            Height          =   285
            Left            =   1005
            TabIndex        =   5
            Text            =   "0"
            Top             =   300
            Width           =   2955
         End
         Begin VB.TextBox txtObservaciones 
            Height          =   1080
            Left            =   4065
            MultiLine       =   -1  'True
            TabIndex        =   2
            Top             =   240
            Width           =   3120
         End
         Begin XtremeSuiteControls.ComboBox cboMonedas 
            Height          =   315
            Left            =   975
            TabIndex        =   3
            Top             =   1680
            Width           =   1515
            _Version        =   786432
            _ExtentX        =   2672
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Appearance      =   6
            Text            =   "ComboBox1"
            AutoComplete    =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton cmdCrear 
            Height          =   390
            Left            =   5640
            TabIndex        =   4
            Top             =   1560
            Width           =   1470
            _Version        =   786432
            _ExtentX        =   2593
            _ExtentY        =   688
            _StockProps     =   79
            Caption         =   "Crear"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.ComboBox cboBancos 
            Height          =   315
            Left            =   975
            TabIndex        =   8
            Top             =   1200
            Width           =   2970
            _Version        =   786432
            _ExtentX        =   5239
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Appearance      =   6
            Text            =   "ComboBox1"
            AutoComplete    =   -1  'True
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Desde"
            Height          =   270
            Left            =   315
            TabIndex        =   13
            Top             =   720
            Width           =   570
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Numero"
            Height          =   270
            Index           =   0
            Left            =   -45
            TabIndex        =   12
            Top             =   330
            Width           =   945
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Hasta"
            Height          =   240
            Left            =   2115
            TabIndex        =   11
            Top             =   720
            Width           =   675
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Bancos"
            Height          =   180
            Left            =   -30
            TabIndex        =   10
            Top             =   1267
            Width           =   945
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Moneda"
            Height          =   165
            Left            =   150
            TabIndex        =   9
            Top             =   1755
            Width           =   750
         End
      End
      Begin GridEX20.GridEX grid_cheques 
         Height          =   7350
         Left            =   -62200
         TabIndex        =   14
         Top             =   1095
         Visible         =   0   'False
         Width           =   7485
         _ExtentX        =   13203
         _ExtentY        =   12965
         Version         =   "2.0"
         PreviewRowIndent=   200
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         EmptyRows       =   -1  'True
         PreviewColumn   =   5
         PreviewRowLines =   1
         ColumnAutoResize=   -1  'True
         MethodHoldFields=   -1  'True
         RowHeaders      =   -1  'True
         DataMode        =   99
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   5
         Column(1)       =   "frmAdminCheques.frx":3F50
         Column(2)       =   "frmAdminCheques.frx":4068
         Column(3)       =   "frmAdminCheques.frx":417C
         Column(4)       =   "frmAdminCheques.frx":42B4
         Column(5)       =   "frmAdminCheques.frx":43C4
         FormatStylesCount=   7
         FormatStyle(1)  =   "frmAdminCheques.frx":4484
         FormatStyle(2)  =   "frmAdminCheques.frx":45BC
         FormatStyle(3)  =   "frmAdminCheques.frx":466C
         FormatStyle(4)  =   "frmAdminCheques.frx":4720
         FormatStyle(5)  =   "frmAdminCheques.frx":47F8
         FormatStyle(6)  =   "frmAdminCheques.frx":48B0
         FormatStyle(7)  =   "frmAdminCheques.frx":4990
         ImageCount      =   0
         PrinterProperties=   "frmAdminCheques.frx":4A4C
      End
      Begin GridEX20.GridEX grid_chequeras 
         Height          =   5490
         Left            =   -69745
         TabIndex        =   15
         Top             =   630
         Visible         =   0   'False
         Width           =   7340
         _ExtentX        =   12938
         _ExtentY        =   9684
         Version         =   "2.0"
         HoldSortSettings=   -1  'True
         DefaultGroupMode=   1
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         PreviewColumn   =   "observaciones"
         PreviewRowLines =   1
         ColumnAutoResize=   -1  'True
         MethodHoldFields=   -1  'True
         RowHeaders      =   -1  'True
         DataMode        =   99
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   6
         Column(1)       =   "frmAdminCheques.frx":4C24
         Column(2)       =   "frmAdminCheques.frx":4D3C
         Column(3)       =   "frmAdminCheques.frx":4E38
         Column(4)       =   "frmAdminCheques.frx":4F24
         Column(5)       =   "frmAdminCheques.frx":5020
         Column(6)       =   "frmAdminCheques.frx":511C
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmAdminCheques.frx":5244
         FormatStyle(2)  =   "frmAdminCheques.frx":537C
         FormatStyle(3)  =   "frmAdminCheques.frx":542C
         FormatStyle(4)  =   "frmAdminCheques.frx":54E0
         FormatStyle(5)  =   "frmAdminCheques.frx":55B8
         FormatStyle(6)  =   "frmAdminCheques.frx":5670
         ImageCount      =   0
         PrinterProperties=   "frmAdminCheques.frx":5750
      End
      Begin XtremeSuiteControls.Label Label6 
         Height          =   255
         Left            =   -62200
         TabIndex        =   118
         Top             =   660
         Visible         =   0   'False
         Width           =   1215
         _Version        =   786432
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Proveedor:"
         Alignment       =   1
      End
   End
   Begin GridEX20.GridEX gridBancos 
      Height          =   1845
      Left            =   480
      TabIndex        =   35
      Top             =   9000
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
      Column(1)       =   "frmAdminCheques.frx":5928
      Column(2)       =   "frmAdminCheques.frx":5A28
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminCheques.frx":5B18
      FormatStyle(2)  =   "frmAdminCheques.frx":5C50
      FormatStyle(3)  =   "frmAdminCheques.frx":5D00
      FormatStyle(4)  =   "frmAdminCheques.frx":5DB4
      FormatStyle(5)  =   "frmAdminCheques.frx":5E8C
      FormatStyle(6)  =   "frmAdminCheques.frx":5F44
      ImageCount      =   0
      PrinterProperties=   "frmAdminCheques.frx":6024
   End
   Begin VB.Menu veOP 
      Caption         =   "Ver OP"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuOpcionesChequeChequera 
      Caption         =   "mnuOpcionesChequeChequera"
      Visible         =   0   'False
      Begin VB.Menu mnuPasarCartera 
         Caption         =   "Pasar a cartera..."
      End
      Begin VB.Menu mnuAnularCheque 
         Caption         =   "Anular..."
      End
   End
End
Attribute VB_Name = "frmAdminCheques"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As Recordset
Dim cartera As Collection
Dim tmpChequera As chequera
Dim cheques1 As New Collection
Dim chequeras As Collection
Dim cheques2 As New Collection
Dim cheques3 As New Collection
Dim tmpCheque As cheque
Dim tmpCheque3eros As cheque
Dim bancos As Collection
Dim Banco As Banco
Private desde


Private Sub btnBorrarBanco_Click()
    Me.cboBancoCartera.ListIndex = -1
End Sub

Private Sub btnBorrarBancosTerceros_Click()
cboBancos3ero.ListIndex = -1
End Sub

Private Sub btnBorrarClasificacion_Click()
    Me.cboClasificacion.ListIndex = -1
End Sub

Private Sub btnBuscar_Click_1()

    Dim q As String
    Set cheques3 = New Collection

    q = "propio=0 and en_cartera=0 and orden_pago_origen!=0"

    
    If LenB(Me.txtNumeroCheque3ero) > 0 Then
        q = q & " and cheq.numero=" & Val(Me.txtNumeroCheque3ero)
    End If

    If LenB(Me.txtNumeroOP) > 0 Then
        q = q & " and cheq.orden_pago_origen=" & Val(Me.txtNumeroOP)
    End If


    Me.grdCheques3eros.ItemCount = 0

    Set cheques3 = New Collection

    Set cheques3 = DAOCheques.FindAll(q)

    Me.grdCheques3eros.ItemCount = cheques3.count

    Me.grpResultadosPropios(1).caption = "Resultados: [ " & cheques3.count & " ] cheques"

    GridEXHelper.AutoSizeColumns Me.grdCheques3eros

End Sub


Private Sub btnBuscarEnCartera_Click_1()
    MostrarCartera
End Sub


Private Sub btnBorrarNumeroCartera_Click()
txtNumeroChequeCartera = ""
End Sub

Private Sub btnBorrarNumeroPropios_Click()
txtNroChequePropio = ""
End Sub

Private Sub btnBorrarOPPropios_Click()
txtIdOP = ""
End Sub

Private Sub btnBorrarOPTerceros_Click()
txtNumeroOP = ""
End Sub

Private Sub btnBorrarOrigen_Click()
    Me.txtOrigen = ""
End Sub


Private Sub btnBorrarOrigenTerceros_Click()
    Me.cboClientes3erosUti.ListIndex = -1
End Sub


Private Sub btnBuscar_Click(Index As Integer)

    Dim q As String
    Set cheques3 = New Collection

    q = "propio=0 and en_cartera=0 and orden_pago_origen!=0"
    
    
    If Not IsNull(Me.dtpDesdeVtoTerceros(5)) Then
        q = q & " and fecha_vencimiento>=" & conectar.Escape(Format(Me.dtpDesdeVtoTerceros(5).value, "yyyy-mm-dd"))
    End If

    If Not IsNull(Me.dtpHastaVtoTerceros(5)) Then
        q = q & " and fecha_vencimiento<=" & conectar.Escape(Format(Me.dtpHastaVtoTerceros(5).value, "yyyy-mm-dd"))
    End If


    If Not IsNull(Me.dtpDesdeRboEmitido(6)) Then
        q = q & " and fecha_emision>=" & conectar.Escape(Format(Me.dtpDesdeRboEmitido(6).value, "yyyy-mm-dd"))
    End If

    If Not IsNull(Me.dtpHastaRboEmitido(6)) Then
        q = q & " and fecha_emision<=" & conectar.Escape(Format(Me.dtpHastaRboEmitido(6).value, "yyyy-mm-dd"))
    End If

    If Me.cboBancos3ero.ListIndex > -1 Then
        q = q & " and cheq.id_banco=" & Me.cboBancos3ero.ItemData(Me.cboBancos3ero.ListIndex)
    End If

    If LenB(Me.txtNumeroCheque3ero) > 0 Then
        q = q & " and cheq.numero like '%" & Trim(Me.txtNumeroCheque3ero.Text) & "%'"
    End If

    If LenB(Me.txtNumeroOP) > 0 Then
        q = q & " and cheq.orden_pago_origen=" & Val(Me.txtNumeroOP)
    End If
    
        If Me.cboClientes3erosUti.ListIndex <> -1 Then
        q = q & " and cheq.origen = '" & Me.cboClientes3erosUti.Text & "'"
    End If

    If Me.cboProveedores3eros.ListIndex <> -1 Then
        q = q & " AND prov.razon = '" & Me.cboProveedores3eros.Text & "'"
    End If
    

    Me.grdCheques3eros.ItemCount = 0

    Set cheques3 = New Collection

    Set cheques3 = DAOCheques.FindAllTercerosUti(q)

    Me.grdCheques3eros.ItemCount = cheques3.count

    Me.grpResultados3ros(2).caption = "Resultados: [ " & cheques3.count & " ] "

    GridEXHelper.AutoSizeColumns Me.grdCheques3eros
    
End Sub


Private Sub btnBuscarChePropios_Click(Index As Integer)
    
    Dim q As String
    Set cheques1 = New Collection

    q = "ingresado=" & Abs(Me.chkIngresados.value) & " and propio=1"
    
    
    If Not IsNull(Me.dtpDesdeVtoPropios(3)) Then
        q = q & " and fecha_vencimiento>=" & conectar.Escape(Format(Me.dtpDesdeVtoPropios(3).value, "yyyy-mm-dd"))
    End If

    If Not IsNull(Me.dtpHastaVtoPropios(3)) Then
        q = q & " and fecha_vencimiento<=" & conectar.Escape(Format(Me.dtpHastaVtoPropios(3).value, "yyyy-mm-dd"))
    End If


    If Not IsNull(Me.dtpDesdeRboPropios(4)) Then
        q = q & " and fecha_emision>=" & conectar.Escape(Format(Me.dtpDesdeRboPropios(4).value, "yyyy-mm-dd"))
    End If

    If Not IsNull(Me.dtpHastaRboPropios(4)) Then
        q = q & " and fecha_emision<=" & conectar.Escape(Format(Me.dtpHastaRboPropios(4).value, "yyyy-mm-dd"))
    End If


    If Me.cboBancos1.ListIndex > -1 Then
        q = q & " and cheqs.id_banco=" & Me.cboBancos1.ItemData(Me.cboBancos1.ListIndex)
    End If


    If Me.cboChequera2.ListIndex > -1 Then
        q = q & " and cheq.id_chequera=" & Me.cboChequera2.ItemData(Me.cboChequera2.ListIndex)
    End If


    If LenB(Me.txtNroChequePropio) > 0 Then
        q = q & " and cheq.numero like '%" & Trim(Me.txtNroChequePropio) & "%'"
    
    End If

    If LenB(Me.txtIdOP) > 0 Then
        q = q & " and cheq.orden_pago_origen=" & Val(Me.txtIdOP)
    End If
    
    
    If Me.cboProveedoresPropios.ListIndex <> -1 Then
        q = q & " AND cheq.origen = '" & Me.cboProveedoresPropios.Text & "'"
    End If
        
    Me.gridChequesEmitidos.ItemCount = 0
    q = q & "  order by fecha_vencimiento desc"
    
    Set cheques2 = New Collection
    
    Set cheques2 = DAOCheques.FindAll(q)

    For Each tmpCheque In cheques2
        If tmpCheque.Monto > 0 Then cheques1.Add tmpCheque


    Next tmpCheque

    Me.grpResultadosPropios(1).caption = "Resultados: [ " & cheques1.count & " ]"

    Me.gridChequesEmitidos.ItemCount = cheques1.count
    GridEXHelper.AutoSizeColumns Me.gridChequesEmitidos
    
End Sub

Private Sub btnBuscarEnCartera_Click(Index As Integer)
    MostrarCartera
End Sub

'EXPORTACION DE TERCEROS UTILIZADOS
Private Sub btnExportar_Click(Index As Integer)

'FUNCIÓN PARA EXPORTAR A EXCEL

    If (cheques3.count > 0) Then
        'INICIA EL PROGRESSBAR Y LO MUESTRA
        '    Me.ProgressBar(2).Visible = True
        '    Me.lblExportando.Visible = True

        'DEFINE EL VALOR MINIMO Y EL MAXIMO DEL PROGRESSBAR (CANTIDAD DE DATOS EN LA COLECCIÓN COL)
        Me.ProgressBar(2).min = 0
        Me.ProgressBar(2).max = cheques3.count


        'Dim xlApplication As New Excel.Application
        Dim xlApplication As Object
        Set xlApplication = CreateObject("Excel.Application")

        'Dim xlWorkbook As New Excel.Workbook
        Dim xlWorkbook As Object
        Set xlWorkbook = CreateObject("Excel.Application")

        'Dim xlWorksheet As New Excel.Worksheet
        Dim xlWorksheet As Object
        Set xlWorksheet = CreateObject("Excel.Application")


        Set xlWorkbook = xlApplication.Workbooks.Add

        Set xlWorksheet = xlWorkbook.Worksheets.item(1)

        xlWorksheet.Activate

        xlWorksheet.Cells(1, 1).value = "Reporte de Cheques de 3ros Utilizados"

        xlWorksheet.Columns(4).HorizontalAlignment = xlLeft
        xlWorksheet.Columns(7).HorizontalAlignment = xlLeft

        xlWorksheet.Cells(2, 1).value = "ID"
        xlWorksheet.Cells(2, 2).value = "Número"
        xlWorksheet.Cells(2, 3).value = "Importe"
        xlWorksheet.Cells(2, 4).value = "Fecha Emisión"
        xlWorksheet.Cells(2, 5).value = "Fecha Vencimiento"
        xlWorksheet.Cells(2, 6).value = "Banco"
        xlWorksheet.Cells(2, 7).value = "Origen"
        xlWorksheet.Cells(2, 8).value = "Recibo Origen"
        xlWorksheet.Cells(2, 9).value = "OP / LIQ / PCTA"
        xlWorksheet.Cells(2, 10).value = "Destino"
   
        Dim idx As Integer
        idx = 3

        Dim che As cheque

        'DEFINE EL CONTADOR DEL PROGRESSBAR Y LO INICIA EN 0
        Dim d As Long
        d = 0


        For Each che In cheques3

            Debug.Print

            xlWorksheet.Cells(idx, 1).value = che.Id
            xlWorksheet.Cells(idx, 2).value = che.numero
            xlWorksheet.Cells(idx, 3).value = funciones.FormatearDecimales(che.Monto)
            xlWorksheet.Cells(idx, 4).value = che.FechaEmision
            xlWorksheet.Cells(idx, 5).value = che.FechaVencimiento
            xlWorksheet.Cells(idx, 6).value = che.Banco.nombre
            xlWorksheet.Cells(idx, 7).value = che.OrigenDestino
            xlWorksheet.Cells(idx, 8).value = che.Recibo
            xlWorksheet.Cells(idx, 9).value = che.IdOrdenPagoOrigen
            xlWorksheet.Cells(idx, 10).value = che.destino

            idx = idx + 1

            'POR CADA ITERACION SUMA UN VALOR A LA VARIABLE D DEL PROGRESSBAR
            d = d + 1
            Me.ProgressBar(2).value = d

        Next

        xlWorksheet.Cells(idx, 3).Formula = "=SUM(c3:c" & idx - 1 & ")"
    
        'AUTOSIZE
        xlApplication.ScreenUpdating = False

        Dim wkSt As String

        wkSt = xlWorksheet.Name

        xlWorksheet.Cells.EntireColumn.AutoFit

        xlWorkbook.Sheets(wkSt).Select

        xlApplication.ScreenUpdating = True

        xlWorksheet.PageSetup.Orientation = xlLandscape
        xlWorksheet.PageSetup.BottomMargin = xlApplication.CentimetersToPoints(1)
        xlWorksheet.PageSetup.TopMargin = xlApplication.CentimetersToPoints(1)
        xlWorksheet.PageSetup.LeftMargin = xlApplication.CentimetersToPoints(1)
        xlWorksheet.PageSetup.RightMargin = xlApplication.CentimetersToPoints(1)

        Dim filename As String
        filename = funciones.GetTmpPath() & "tmp_info " & Hour(Now) & Minute(Now) & Second(Now) & " .xlsx"

        If Dir(filename) <> vbNullString Then Kill filename

        xlWorkbook.SaveAs filename

        xlWorkbook.Saved = True
        xlWorkbook.Close
        xlApplication.Quit

        funciones.ShellExecute 0, "open", filename, "", "", 0

        Set xlWorksheet = Nothing
        Set xlWorkbook = Nothing
        Set xlApplication = Nothing

        'REINICIA EL PROGRESSBAR Y LO OCULTA
        Me.ProgressBar(2).value = 0
        ' Me.ProgressBar(2).Visible = False
        '    Me.lblExportando.Visible = False
    Else
        MsgBox ("No hay resultados para exportar!")
    End If
    
End Sub

Private Sub btnExportarCartera_Click(Index As Integer)

    If (cartera.count > 0) Then
        Me.ProgressBar(0).min = 0
        Me.ProgressBar(0).max = cartera.count

        Dim xlApplication As Object
        Set xlApplication = CreateObject("Excel.Application")

        'Dim xlWorkbook As New Excel.Workbook
        Dim xlWorkbook As Object
        Set xlWorkbook = CreateObject("Excel.Application")

        'Dim xlWorksheet As New Excel.Worksheet
        Dim xlWorksheet As Object
        Set xlWorksheet = CreateObject("Excel.Application")


        Set xlWorkbook = xlApplication.Workbooks.Add

        Set xlWorksheet = xlWorkbook.Worksheets.item(1)

        xlWorksheet.Activate

        xlWorksheet.Cells(1, 1).value = "Reporte de Cartera"

        xlWorksheet.Columns(4).HorizontalAlignment = xlLeft
        xlWorksheet.Columns(7).HorizontalAlignment = xlLeft

        xlWorksheet.Cells(2, 1).value = "ID"
        xlWorksheet.Cells(2, 2).value = "NÚMERO"
        xlWorksheet.Cells(2, 3).value = "MONTO"
        xlWorksheet.Cells(2, 4).value = "VENCIMIENTO"
        xlWorksheet.Cells(2, 5).value = "ORIGEN"
        xlWorksheet.Cells(2, 6).value = "BANCO NOMBRE"
        xlWorksheet.Cells(2, 7).value = "CLASIFICACION"
        xlWorksheet.Cells(2, 8).value = "RECIBIDO"


        Dim idx As Integer
        idx = 3

        Dim che As cheque

        'DEFINE EL CONTADOR DEL PROGRESSBAR Y LO INICIA EN 0
        Dim d As Long
        d = 0


        For Each che In cartera

            Debug.Print

            xlWorksheet.Cells(idx, 1).value = che.Id
            xlWorksheet.Cells(idx, 2).value = che.numero
            xlWorksheet.Cells(idx, 3).value = che.Monto
            xlWorksheet.Cells(idx, 4).value = che.FechaVencimiento
            xlWorksheet.Cells(idx, 5).value = che.OrigenDestino
            xlWorksheet.Cells(idx, 6).value = che.Banco.nombre
            xlWorksheet.Cells(idx, 7).value = che.OrigenCheque
            xlWorksheet.Cells(idx, 8).value = che.FechaRecibido

            idx = idx + 1

            'POR CADA ITERACION SUMA UN VALOR A LA VARIABLE D DEL PROGRESSBAR
            d = d + 1
            '        progreso.value = d
            Me.ProgressBar(0).value = d

        Next

        '    xlWorksheet.Cells(idx, 5).Formula = "=SUM(E3:E" & idx - 1 & ")"

        'AUTOSIZE
        xlApplication.ScreenUpdating = False

        Dim wkSt As String

        wkSt = xlWorksheet.Name

        xlWorksheet.Cells.EntireColumn.AutoFit

        xlWorkbook.Sheets(wkSt).Select

        xlApplication.ScreenUpdating = True

        xlWorksheet.PageSetup.Orientation = xlLandscape
        xlWorksheet.PageSetup.BottomMargin = xlApplication.CentimetersToPoints(1)
        xlWorksheet.PageSetup.TopMargin = xlApplication.CentimetersToPoints(1)
        xlWorksheet.PageSetup.LeftMargin = xlApplication.CentimetersToPoints(1)
        xlWorksheet.PageSetup.RightMargin = xlApplication.CentimetersToPoints(1)

        Dim filename As String
        filename = funciones.GetTmpPath() & "tmp_info " & Hour(Now) & Minute(Now) & Second(Now) & " .xlsx"

        If Dir(filename) <> vbNullString Then Kill filename

        xlWorkbook.SaveAs filename

        xlWorkbook.Saved = True
        xlWorkbook.Close
        xlApplication.Quit

        funciones.ShellExecute 0, "open", filename, "", "", 0

        Set xlWorksheet = Nothing
        Set xlWorkbook = Nothing
        Set xlApplication = Nothing

        'REINICIA EL PROGRESSBAR Y LO OCULTA
        '    progreso.value = 0
        '    Me.progreso.Visible = False
        '    Me.lblExportando.Visible = False

        Me.ProgressBar(0).value = 0

    Else
        MsgBox ("No hay resultados para exportar")
    End If

End Sub



'EXPORTACION DE CHEQUES PROPIOS
Private Sub btnExportarChePropios_Click(Index As Integer)
'FUNCIÓN PARA EXPORTAR A EXCEL


If (cheques1.count > 0) Then


'INICIA EL PROGRESSBAR Y LO MUESTRA
Me.ProgressBar(1).Visible = True
'    Me.lblExportando.Visible = True


'DEFINE EL VALOR MINIMO Y EL MAXIMO DEL PROGRESSBAR (CANTIDAD DE DATOS EN LA COLECCIÓN COL)
Me.ProgressBar(1).min = 0
Me.ProgressBar(1).max = cheques1.count


'Dim xlApplication As New Excel.Application
    Dim xlApplication As Object
    Set xlApplication = CreateObject("Excel.Application")

    'Dim xlWorkbook As New Excel.Workbook
    Dim xlWorkbook As Object
    Set xlWorkbook = CreateObject("Excel.Application")

    'Dim xlWorksheet As New Excel.Worksheet
    Dim xlWorksheet As Object
    Set xlWorksheet = CreateObject("Excel.Application")


    Set xlWorkbook = xlApplication.Workbooks.Add

    Set xlWorksheet = xlWorkbook.Worksheets.item(1)

    xlWorksheet.Activate

    xlWorksheet.Cells(1, 1).value = "Reporte de Cheques Propios Utilizados"

    xlWorksheet.Columns(4).HorizontalAlignment = xlLeft
    xlWorksheet.Columns(7).HorizontalAlignment = xlLeft

    xlWorksheet.Cells(2, 1).value = "ID"
    xlWorksheet.Cells(2, 2).value = "Número"
    xlWorksheet.Cells(2, 3).value = "Importe"
    xlWorksheet.Cells(2, 4).value = "Fecha Emisión"
    xlWorksheet.Cells(2, 5).value = "Fecha Vencimiento"
    xlWorksheet.Cells(2, 6).value = "Banco"
    xlWorksheet.Cells(2, 7).value = "Destino"
    xlWorksheet.Cells(2, 8).value = "N OP"
    
    Dim idx As Integer
    idx = 3

    Dim che As cheque

    'DEFINE EL CONTADOR DEL PROGRESSBAR Y LO INICIA EN 0
    Dim d As Long
    d = 0


    For Each che In cheques1

        Debug.Print

        xlWorksheet.Cells(idx, 1).value = che.Id
        xlWorksheet.Cells(idx, 2).value = che.numero
        xlWorksheet.Cells(idx, 3).value = che.Monto
        xlWorksheet.Cells(idx, 4).value = che.FechaEmision
        xlWorksheet.Cells(idx, 5).value = che.FechaVencimiento
        xlWorksheet.Cells(idx, 6).value = che.Banco.nombre
        xlWorksheet.Cells(idx, 7).value = che.OrigenDestino
        xlWorksheet.Cells(idx, 8).value = che.IdOrdenPagoOrigen
        
        idx = idx + 1

        'POR CADA ITERACION SUMA UN VALOR A LA VARIABLE D DEL PROGRESSBAR
        d = d + 1
        Me.ProgressBar(1).value = d


    Next

    xlWorksheet.Cells(idx, 3).Formula = "=SUM(c3:c" & idx - 1 & ")"

    'AUTOSIZE
    xlApplication.ScreenUpdating = False

    Dim wkSt As String

    wkSt = xlWorksheet.Name

    xlWorksheet.Cells.EntireColumn.AutoFit

    xlWorkbook.Sheets(wkSt).Select

    xlApplication.ScreenUpdating = True

    xlWorksheet.PageSetup.Orientation = xlLandscape
    xlWorksheet.PageSetup.BottomMargin = xlApplication.CentimetersToPoints(1)
    xlWorksheet.PageSetup.TopMargin = xlApplication.CentimetersToPoints(1)
    xlWorksheet.PageSetup.LeftMargin = xlApplication.CentimetersToPoints(1)
    xlWorksheet.PageSetup.RightMargin = xlApplication.CentimetersToPoints(1)

    Dim filename As String
    filename = funciones.GetTmpPath() & "tmp_info " & Hour(Now) & Minute(Now) & Second(Now) & " .xlsx"

    If Dir(filename) <> vbNullString Then Kill filename

    xlWorkbook.SaveAs filename

    xlWorkbook.Saved = True
    xlWorkbook.Close
    xlApplication.Quit

    funciones.ShellExecute 0, "open", filename, "", "", 0

    Set xlWorksheet = Nothing
    Set xlWorkbook = Nothing
    Set xlApplication = Nothing

    'REINICIA EL PROGRESSBAR Y LO OCULTA
Me.ProgressBar(1).value = 0
Me.ProgressBar(1).Visible = False
    '    Me.lblExportando.Visible = False
    
    Else
    MsgBox ("No hay resultados para exportar!")
    
    End If
    
End Sub

Private Sub btnFiltrar_Click()
    On Error Resume Next
    Dim filter2 As String
    filter2 = "cheq.origen = """ & Me.cboProveedores.Text & """"
    Set tmpChequera.cheques = DAOCheques.FindAllByChequeraId(tmpChequera.Id, filter2)
    mostrarCheques
End Sub

Private Sub cboRangosVtoCartera_Click(Index As Integer)
    'funciones.CalculateDateRange Me.cboRangosVto(0), Me.dtpDesde(1), Me.dtpHasta(1)
    funciones.CalculateDateRange Me.cboRangosVtoCartera(0), Me.dtpDesdeVtoCartera(1), Me.dtpHastaVtoCartera(1)
End Sub

Private Sub cboRangosVtoPropios_Click(Index As Integer)
    funciones.CalculateDateRange Me.cboRangosVtoPropios(1), Me.dtpDesdeVtoPropios(3), Me.dtpHastaVtoPropios(3)
End Sub

Private Sub cboRangosVtoTerceros_Click(Index As Integer)
    funciones.CalculateDateRange Me.cboRangosVtoTerceros(2), Me.dtpDesdeVtoTerceros(5), Me.dtpHastaVtoTerceros(5)
End Sub
    


Private Sub cboRangosRboCartera_Click(Index As Integer)
    'funciones.CalculateDateRange Me.cboRangosRbo(1), Me.dtpDesde(2), Me.dtpHasta(2)
    funciones.CalculateDateRange Me.cboRangosRboCartera(1), Me.dtpDesdeRboCartera(2), Me.dtpHastaRboCartera(2)

End Sub

Private Sub cboRangosRboPropios_Click(Index As Integer)
    funciones.CalculateDateRange Me.cboRangosRboPropios(0), Me.dtpDesdeRboPropios(4), Me.dtpHastaRboPropios(4)
    
End Sub
    
Private Sub cboRangosRboEmitido_Click(Index As Integer)
    funciones.CalculateDateRange Me.cboRangosRboEmitido(2), Me.dtpDesdeRboEmitido(6), Me.dtpHastaRboEmitido(6)
    
End Sub




Private Sub cmdCrear_Click()
    Dim x As Long
    Dim col As Collection
    Dim id_banco As Long


    If MsgBox("zEstá seguro de crear la chequera?", vbQuestion + vbYesNo) = vbYes Then
        Dim chequera As New chequera
        If Me.cboBancos.ListIndex = -1 Then
            MsgBox "Seleccione un banco Correcto!", vbCritical, "Error"
            Exit Sub
        End If
        If Not IsNumeric(Me.txtNumero) Or Not IsNumeric(Me.txtDesde) Or Not IsNumeric(Me.txtHasta) Then
            MsgBox "Ingrese números válidos!", vbCritical, "Error"
            Exit Sub
        End If
        id_banco = Me.cboBancos.ItemData(Me.cboBancos.ListIndex)
        Set col = DAOChequeras.GetAll(DAOChequeras.CAMPO_NUMERO & "=" & Me.txtNumero & " AND id_banco=" & id_banco)
        If col.count > 0 Then
            MsgBox "El número de chequera de ese banco ya existe!", vbCritical, "Error"
            Exit Sub
        End If

        Set chequera.Banco = DAOBancos.GetById(id_banco)
        chequera.FechaCreacion = Now
        Set chequera.moneda = DAOMoneda.GetById(Me.cboMonedas.ItemData(Me.cboMonedas.ListIndex))
        chequera.numero = CLng(Me.txtNumero)
        chequera.NumeroDesde = CLng(Me.txtDesde)
        chequera.NumeroHasta = CLng(Me.txtHasta)
        chequera.observaciones = UCase(Me.txtObservaciones)
        Dim cheque As cheque
        For x = chequera.NumeroDesde To chequera.NumeroHasta
            Set cheque = New cheque
            cheque.numero = x
            cheque.EnCartera = False
            cheque.Propio = True
            cheque.Id = 0
            Set cheque.Banco = chequera.Banco
            Set cheque.moneda = chequera.moneda
            chequera.cheques.Add cheque

        Next
        If DAOChequeras.Guardar(chequera) Then
            MsgBox "Guardado Correctamente!", vbInformation, "Información"
            MostrarChequeras
        End If
    End If



End Sub

Private Sub CMDsINCliente_Click()
    Me.cboBancos1.ListIndex = -1
End Sub

Private Sub Form_Load()
    FormHelper.Customize Me
    
    GridEXHelper.CustomizeGrid Me.grid_chequeras, True, False
    GridEXHelper.CustomizeGrid Me.grid_cartera_cheques, True, True
    GridEXHelper.CustomizeGrid Me.grid_cheques, True, False
    GridEXHelper.CustomizeGrid Me.gridBancos, False, True
    GridEXHelper.CustomizeGrid Me.gridChequesEmitidos, False, False
    GridEXHelper.CustomizeGrid Me.grdCheques3eros, False, False
    Dim i As Integer
    
    i = 1
    
    'SOLAPA CARTERA
    DAOBancos.llenarComboXtremeSuite Me.cboBancoCartera
    Me.cboBancoCartera.ListIndex = -1
    
    DAOBancos.llenarComboXtremeSuite Me.cboBancos
    Me.cboBancos.ListIndex = -1
      
    Me.cboClasificacion.Clear
    Me.cboClasificacion.AddItem "Propios"
    Me.cboClasificacion.ItemData(cboClasificacion.NewIndex) = 1
    Me.cboClasificacion.AddItem "Terceros"
    Me.cboClasificacion.ItemData(cboClasificacion.NewIndex) = 2
    Me.cboClasificacion.AddItem "Terceros propio"
    Me.cboClasificacion.ItemData(cboClasificacion.NewIndex) = 3
    
    Me.cboClasificacion.ListIndex = -1
    
    DAOBancos.llenarComboXtremeSuite Me.cboBancos1
    cboBancos1.ListIndex = -1
    
    
    'SOLAPA ADMINISTRAR CHEQUERAS
    DAOMoneda.llenarComboXtremeSuite Me.cboMonedas
    
    Set bancos = DAOBancos.GetAll("id in (select idBanco from AdminConfigCuentas group by idBanco) ")

    cboBancos1.Clear
    For Each Banco In bancos
        cboBancos1.AddItem Banco.nombre
        cboBancos1.ItemData(cboBancos1.NewIndex) = Banco.Id
    Next

    Set bancos = DAOBancos.GetAll()
    
'    DAOProveedor.llenarComboXtremeSuite Me.cboProveedores, True, True, True
'    Me.cboProveedores.ListIndex = -1
    
    Me.grid_cheques.ItemCount = 0
    
    
    'SOLAPA CHEQUES PROPIOS UTILIZADOS
    DAOChequeras.llenarComboXtremeSuite Me.cboChequera2
    Me.cboChequera2.ListIndex = -1
    
    funciones.FillComboBoxDateRanges Me.cboRangosVtoPropios(1)
    Me.cboRangosVtoPropios(1) = i
    For i = 0 To Me.cboRangosVtoPropios(1).ListCount - 1
        If Me.cboRangosVtoPropios(1).ItemData(i) = DateRangeValue.DRV_YearCurrent Then Exit For
    Next i
    Me.cboRangosVtoPropios(1).ListIndex = -1


    funciones.FillComboBoxDateRanges Me.cboRangosRboPropios(0)
    Me.cboRangosRboPropios(0) = i
    For i = 0 To Me.cboRangosRboPropios(0).ListCount - 1
        If Me.cboRangosRboPropios(0).ItemData(i) = DateRangeValue.DRV_YearCurrent Then Exit For
    Next i
    Me.cboRangosRboPropios(0).ListIndex = -1
    
    DAOProveedor.llenarComboXtremeSuite Me.cboProveedoresPropios, True, True, True
    Me.cboProveedoresPropios.ListIndex = -1
    
        
    'SOLAPA CHEQUES 3EROS UTILIZADOS
    DAOBancos.llenarComboXtremeSuite Me.cboBancos3ero
    Me.cboBancos3ero.ListIndex = -1
    
    funciones.FillComboBoxDateRanges Me.cboRangosVtoTerceros(2)
    Me.cboRangosVtoTerceros(2) = i
    For i = 0 To Me.cboRangosVtoTerceros(2).ListCount - 1
        If Me.cboRangosVtoTerceros(2).ItemData(i) = DateRangeValue.DRV_YearCurrent Then Exit For
    Next i
    Me.cboRangosVtoTerceros(2).ListIndex = -1


    funciones.FillComboBoxDateRanges Me.cboRangosRboEmitido(2)
    Me.cboRangosRboEmitido(2) = i
    For i = 0 To Me.cboRangosRboEmitido(2).ListCount - 1
        If Me.cboRangosRboEmitido(2).ItemData(i) = DateRangeValue.DRV_YearCurrent Then Exit For
    Next i
    Me.cboRangosRboEmitido(2).ListIndex = -1
    
    
    Me.gridBancos.ItemCount = 0
    Me.gridChequesEmitidos.ItemCount = 0
    Me.grdCheques3eros.ItemCount = 0
    
    DAOProveedor.llenarComboXtremeSuite Me.cboProveedores3eros, True, True, True
    Me.cboProveedores3eros.ListIndex = -1
    
    DAOCliente.llenarComboXtremeSuite Me.cboClientes3erosUti, True, True, True
    Me.cboClientes3erosUti.ListIndex = -1
    
    '''''''''''''''''''''''''''''''''''''''''''''''
    'region FECHAS EN CARTERA
    
    
'    Me.dtpDesdeVtoCartera(1).value = Year(Now) & "-01-01"
'    desde = DateSerial(Year(Date), Month(Date), 1)   ' CDate(1 & "-" & Month(Now) & "-" & Year(Now))
'
'    Me.dtpDesdeVtoPropios(3).value = Year(Now) & "-01-01"
'    desde = DateSerial(Year(Date), Month(Date), 1)
'
'    Me.dtpDesdeVtoTerceros(5).value = Year(Now) & "-01-01"
'    desde = DateSerial(Year(Date), Month(Date), 1)
    
    funciones.FillComboBoxDateRanges Me.cboRangosVtoCartera(0)
    Me.cboRangosVtoCartera(0) = i
    For i = 0 To Me.cboRangosVtoCartera(0).ListCount - 1
        If Me.cboRangosVtoCartera(0).ItemData(i) = DateRangeValue.DRV_YearCurrent Then Exit For
    Next i
    Me.cboRangosVtoCartera(0).ListIndex = -1


    funciones.FillComboBoxDateRanges Me.cboRangosRboCartera(1)
    Me.cboRangosRboCartera(1) = i
    For i = 0 To Me.cboRangosRboCartera(1).ListCount - 1
        If Me.cboRangosRboCartera(1).ItemData(i) = DateRangeValue.DRV_YearCurrent Then Exit For
    Next i
    Me.cboRangosRboCartera(1).ListIndex = -1
    
'
'    Me.dtpDesde(1).value = Nothing
'    Me.dtpHasta(1).value = Nothing
'    Me.dtpDesde(2).value = Nothing
'    Me.dtpHasta(2).value = Nothing
    
    
    '''''''''''''''''''''''''''''''''''''''''''''''
    'endregion FECHAS EN CARTERA
    
    MostrarChequeras


    Set Me.grid_cartera_cheques.Columns("banco").DropDownControl = Me.gridBancos
    Me.gridBancos.ItemCount = bancos.count
    
    
    Me.grid_cartera_cheques.ItemCount = 0

    Dim idc As Long
    idc = chequeras.item(Me.grid_chequeras.RowIndex(Me.grid_chequeras.row)).Id

    Set tmpChequera = DAOChequeras.GetById(idc)
    Set tmpChequera.cheques = DAOCheques.FindAllByChequeraId(idc)

End Sub


Private Sub Form_Resize()
    Me.TabControl1.Width = Me.ScaleWidth
    Me.TabControl1.Height = Me.ScaleHeight
End Sub


Private Sub MostrarCartera()

    Dim filter2 As String
    Dim Orden As String


    filter2 = "1 = 1"
    
    If LenB(Me.txtOrigen.Text) > 0 Then
        filter2 = filter2 & " AND cheq.origen like '%" & Trim(Me.txtOrigen.Text) & "%'"
    End If

    If LenB(Me.txtNumeroChequeCartera.Text) > 0 Then
        filter2 = filter2 & " AND cheq.numero like '%" & Trim(Me.txtNumeroChequeCartera.Text) & "%'"
    End If

    If Not IsNull(Me.dtpDesdeVtoCartera(1).value) Then
        filter2 = filter2 & " AND cheq.fecha_vencimiento >= " & conectar.Escape(Me.dtpDesdeVtoCartera(1).value)
    End If

    If Not IsNull(Me.dtpHastaVtoCartera(1).value) Then
        filter2 = filter2 & " AND cheq.fecha_vencimiento <= " & conectar.Escape(dtpHastaVtoCartera(1).value)
    End If

    If Not IsNull(Me.dtpDesdeRboCartera(2).value) Then
        filter2 = filter2 & " AND cheq.fecha_recibido >= " & conectar.Escape(Me.dtpDesdeRboCartera(2).value)
    End If

    If Not IsNull(Me.dtpHastaRboCartera(2).value) Then
        filter2 = filter2 & " AND cheq.fecha_recibido <= " & conectar.Escape(Me.dtpHastaRboCartera(2).value)
    End If

    If Me.cboBancoCartera.ListIndex > -1 Then
        filter2 = filter2 & " and cheq.id_banco=" & Me.cboBancoCartera.ItemData(Me.cboBancoCartera.ListIndex)
    End If

    If Me.cboClasificacion.ListIndex > -1 Then
        If Me.cboClasificacion.ListIndex = 0 Then    'propio
            filter2 = filter2 & " AND cheq.propio = 1 AND cheq.teceros_propio = 0 "
        ElseIf Me.cboClasificacion.ListIndex = 1 Then    'terceros
            filter2 = filter2 & " AND cheq.propio = 0 AND cheq.teceros_propio = 0 "
        ElseIf Me.cboClasificacion.ListIndex = 2 Then    'terceros propio
            filter2 = filter2 & " AND cheq.propio = 0 AND cheq.teceros_propio = 1 "
        End If
    End If

    Orden = "cheq.id DESC"

    Set cartera = DAOCheques.FindAllEnCartera(filter2, Orden)

    Me.grid_cartera_cheques.ItemCount = 0
    Me.grid_cartera_cheques.ItemCount = cartera.count
    
    Me.grpResultados(0).caption = "Resultados: [ " & cartera.count & " ]"

End Sub


Private Sub MostrarChequeras()
    Set chequeras = DAOChequeras.GetAll
    Me.grid_chequeras.ItemCount = 0
    Me.grid_chequeras.ItemCount = chequeras.count
End Sub

Private Sub grdCheques3eros_ColumnHeaderClick(ByVal Column As GridEX20.JSColumn)
    GridEXHelper.ColumnHeaderClick Me.grdCheques3eros, Column
End Sub

Private Sub grdCheques3eros_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)

    Set tmpCheque3eros = cheques3.item(RowIndex)

    Values(1) = tmpCheque3eros.OrigenDestino
    Values(2) = tmpCheque3eros.Id
    Values(3) = tmpCheque3eros.Banco.nombre
    Values(4) = ""
    Values(5) = tmpCheque3eros.FechaEmision
    Values(6) = tmpCheque3eros.numero
    Values(7) = funciones.FormatearDecimales(tmpCheque3eros.Monto)
    Values(8) = tmpCheque3eros.FechaVencimiento
    Values(9) = tmpCheque3eros.Recibo
    Values(10) = tmpCheque3eros.IdOrdenPagoOrigen
    Values(11) = tmpCheque3eros.destino
    
End Sub

Private Sub grid_cartera_cheques_BeforeUpdate(ByVal Cancel As GridEX20.JSRetBoolean)
'validar


    Dim cond1 As Boolean, cond2 As Boolean
    Dim cond3 As Boolean
    cond1 = Not (IsDate(Me.grid_cartera_cheques.value(7)) And IsDate(Me.grid_cartera_cheques.value(3)))
    cond2 = Not (IsNumeric(Me.grid_cartera_cheques.value(2)) And IsNumeric(Me.grid_cartera_cheques.value(1)))
    cond3 = False    ' Not (IsNumeric(Me.grid_cartera_cheques.value(5)) And Val(Me.grid_cartera_cheques.value(5)) > 0)
    Cancel = cond1 Or cond2 Or cond3



End Sub

Private Sub grid_cartera_cheques_ColumnHeaderClick(ByVal Column As GridEX20.JSColumn)
    GridEXHelper.ColumnHeaderClick Me.grid_cartera_cheques, Column
End Sub


Private Sub grid_cartera_cheques_UnboundAddNew(ByVal NewRowBookmark As GridEX20.JSRetVariant, ByVal Values As GridEX20.JSRowData)
    On Error GoTo err1
    Set tmpCheque = New cheque
    Set tmpCheque.Banco = DAOBancos.GetById(Values(5))
    tmpCheque.EnCartera = True
    tmpCheque.FechaRecibido = Values(7)
    tmpCheque.FechaVencimiento = Values(3)
    Set tmpCheque.moneda = DAOMoneda.GetById(0)       ' reemplazar x un combo
    tmpCheque.Monto = Values(2)
    tmpCheque.numero = Values(1)
    tmpCheque.OrigenDestino = Values(4)
    tmpCheque.Propio = False

    If Not DAOCheques.Guardar(tmpCheque) Then GoTo err1
    cartera.Add tmpCheque, CStr(tmpCheque.Id)

    Exit Sub
err1:

End Sub

Private Sub grid_cartera_cheques_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    On Error GoTo err1
    Set tmpCheque = cartera.item(RowIndex)
    With Values
        .value(1) = tmpCheque.Id
        .value(2) = tmpCheque.numero
        .value(3) = funciones.FormatearDecimales(tmpCheque.Monto)
        .value(4) = tmpCheque.FechaVencimiento
        .value(5) = tmpCheque.OrigenDestino
        .value(6) = tmpCheque.Banco.nombre
        .value(7) = tmpCheque.OrigenCheque
        .value(8) = tmpCheque.FechaRecibido
        .value(9) = tmpCheque.observaciones

    End With




    Exit Sub
err1:


End Sub

Private Sub grid_cartera_cheques_UnboundUpdate(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    On Error GoTo err1
    Dim ant As String

    ant = tmpCheque.OrigenDestino

    Set tmpCheque = cartera.item(RowIndex)
    tmpCheque.OrigenDestino = Values(4)
    If Not DAOCheques.Guardar(tmpCheque) Then GoTo err1
    Exit Sub
err1:
    tmpCheque.OrigenDestino = ant
End Sub

Private Sub grid_chequeras_SelectionChange()
    On Error Resume Next
    Set tmpChequera.cheques = DAOCheques.FindAllByChequeraId(tmpChequera.Id)
    mostrarCheques
End Sub
Private Sub mostrarCheques()
    Me.grid_cheques.ItemCount = 0
    Me.grid_cheques.ItemCount = tmpChequera.cheques.count
End Sub
Private Sub grid_chequeras_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Set tmpChequera = chequeras.item(RowIndex)
    With Values
        .value(1) = tmpChequera.numero
        .value(2) = tmpChequera.FechaCreacion
        .value(3) = tmpChequera.Banco.nombre
        .value(4) = tmpChequera.NumeroDesde
        .value(5) = tmpChequera.NumeroHasta

    End With

End Sub


Private Sub grid_cheques_DblClick()
    If Me.grid_cheques.RowIndex(Me.grid_cheques.row) > 0 Then
        Set tmpCheque = tmpChequera.cheques(Me.grid_cheques.RowIndex(Me.grid_cheques.row))
        PasarACartera tmpCheque
    End If
End Sub

Private Sub PasarACartera(ch As cheque)
    If ch.EnCartera Then
        MsgBox "El cheque ya se encuentra en cartera.", vbInformation
    Else
        If ch.Utilizado Then
            MsgBox "El cheque ya fue utilizado.", vbInformation
        Else
            Dim f000 As New frmChequePropioACartera
            Set f000.cheque = ch
            Load f000
            f000.Show 1
            mostrarCheques
        End If
    End If
End Sub

Private Sub grid_cheques_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Set tmpCheque = tmpChequera.cheques(Me.grid_cheques.RowIndex(Me.grid_cheques.row))
        Me.mnuAnularCheque.Enabled = (tmpCheque.IdOrdenPagoOrigen <= 0) Or tmpCheque.estado = ChequeAnulado
        Me.PopupMenu Me.mnuOpcionesChequeChequera
    End If
End Sub


Private Sub grid_cheques_RowFormat(RowBuffer As GridEX20.JSRowData)
    On Error GoTo err1

    If tmpCheque.estado = ChequeAnulado Then RowBuffer.RowStyle = "anulado"
    Exit Sub
err1:

End Sub

Private Sub grid_cheques_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    On Error GoTo err1
    If Not IsSomething(tmpChequera.cheques) Or tmpChequera.cheques.count = 0 Then Set tmpChequera.cheques = DAOCheques.FindAllByChequeraId(tmpChequera.Id)
    Set tmpCheque = tmpChequera.cheques(RowIndex)
    With Values
        .value(1) = tmpCheque.numero
        .value(2) = IIf(tmpCheque.Utilizado, funciones.FormatearDecimales(tmpCheque.Monto), Empty)
        .value(3) = IIf(tmpCheque.Utilizado, tmpCheque.FechaVencimiento, Empty)
        .value(4) = IIf(tmpCheque.Utilizado, tmpCheque.OrigenDestino, Empty)
        '.value(5) = IIf(tmpCheque.Utilizado, tmpCheque.Observaciones, "DISPONIBLE")
        .value(5) = IIf(tmpCheque.estado = ChequeAnulado, "ANULADO", IIf(tmpCheque.Utilizado, "Utilizado en Orden de Pago Ns " & tmpCheque.IdOrdenPagoOrigen, "DISPONIBLE"))
    End With
    Exit Sub
err1:
End Sub

Private Sub gridBancos_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex <= bancos.count Then
        Set Banco = bancos.item(RowIndex)
        Values(1) = Banco.Id
        Values(2) = Banco.nombre
    End If
End Sub

Private Sub gridChequesEmitidos_ColumnHeaderClick(ByVal Column As GridEX20.JSColumn)
    GridEXHelper.ColumnHeaderClick Me.gridChequesEmitidos, Column

End Sub

Private Function buscarOP(chequeid As Long) As String
    Set rs = conectar.RSFactory("SELECT op.FECHA,opc.id_cheque FROM ordenes_pago_cheques opc INNER JOIN ordenes_pago op ON opc.id_orden_pago=op.id WHERE opc.id_cheque=" & chequeid)
    If Not rs.EOF And Not rs.BOF Then
        buscarOP = rs!FEcha
    End If
End Function

Private Sub gridChequesEmitidos_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Set tmpCheque = cheques1.item(RowIndex)
    
    Values(1) = ""
    Values(2) = tmpCheque.Banco.nombre
    Values(3) = tmpCheque.Id
    Values(4) = tmpCheque.FechaEmision
    Values(5) = tmpCheque.FechaVencimiento
    Values(6) = tmpCheque.numero
    Values(7) = funciones.FormatearDecimales(tmpCheque.Monto)
    Values(8) = tmpCheque.OrigenDestino
    Values(9) = ""
    Values(10) = tmpCheque.IdOrdenPagoOrigen
    
    If tmpCheque.NumeroLiquidacionCaja = 0 Then
        Values(11) = ""
    Else
        Values(10) = ""
        Values(8) = "VARIOS PROVEEDORES"
        Values(11) = tmpCheque.NumeroLiquidacionCaja
    End If


End Sub


Private Sub mnuPasarCartera_Click()
    grid_cheques_DblClick
End Sub

Private Sub btnBorrarNumeroTerceros_Click()
    txtNumeroCheque3ero = ""
End Sub

Private Sub PushButton1_Click()
    Me.cboProveedores.ListIndex = -1
End Sub

Private Sub PushButton2_Click()
    Me.cboProveedoresPropios.ListIndex = -1
End Sub

'
'Private Sub PushButton2_Click()
'    Dim elegidos As Boolean
'    Dim q As String
'
'    If Not IsNull(Me.dtpDesde) Then
'        q = "Desde " & Format(Me.dtpDesde, "dd-mm-yyyy") & Chr(10)
'    End If
'    If Not IsNull(Me.dtpHasta) Then
'        q = q & "Hasta " & Format(Me.dtpHasta, "dd-mm-yyyy") & Chr(10)
'
'    End If
'
'
'    If IsNull(Me.dtpHasta) And IsNull(Me.dtpDesde) Then
'        q = "PERIODO SIN ESPECIFICAR"
'    End If
'
'
'    With Me.gridChequesEmitidos.PrinterProperties
'        .FitColumns = True
'        .RepeatHeaders = True
'        .Orientation = jgexPPLandscape
'        .HeaderString(jgexHFCenter) = "Listado de cheques"
'        .FooterString(jgexHFCenter) = Now
'        .HeaderString(jgexHFLeft) = q
'
'    End With
'    Load frmPrintPreview
'    frmPrintPreview.Move Me.Left, Me.Top, Me.Width, Me.Height
'    gridChequesEmitidos.PrintPreview frmPrintPreview.GEXPreview1, elegidos
'    frmPrintPreview.Show 1
'
'End Sub

Private Sub PushButton3_Click()
    Me.cboChequera2.ListIndex = -1
End Sub

Private Sub PushButton4_Click()
    Me.cboProveedores3eros.ListIndex = -1
End Sub

Private Sub txtDesde_Validate(Cancel As Boolean)
    ValidarTextBox Me.txtDesde, Cancel
End Sub

Private Sub txtHasta_Validate(Cancel As Boolean)
    ValidarTextBox Me.txtHasta, Cancel
End Sub
Private Sub txtIdOP_GotFocus()
    foco Me.txtIdOP
End Sub
Private Sub txtNroChequePropio_GotFocus()
    foco Me.txtNroChequePropio
End Sub
Private Sub txtNumero_Validate(Cancel As Boolean)
    funciones.ValidarTextBox Me.txtNumero, Cancel
End Sub

