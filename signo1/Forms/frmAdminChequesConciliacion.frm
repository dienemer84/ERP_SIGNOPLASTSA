VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmAdminChequesConciliacion 
   Caption         =   "Conciliación de Cheques"
   ClientHeight    =   10425
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14550
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10425
   ScaleWidth      =   14550
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.PushButton cmdExportar 
      Height          =   255
      Left            =   10560
      TabIndex        =   0
      Top             =   2835
      Width           =   1935
      _Version        =   786432
      _ExtentX        =   3413
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Exportar Chequeras"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.GroupBox GroupBox4 
      Height          =   2700
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   12375
      _Version        =   786432
      _ExtentX        =   21828
      _ExtentY        =   4762
      _StockProps     =   79
      Caption         =   "Parámetros de búsqueda"
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
      Begin VB.TextBox TxtNumeroChequeEnChequera 
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   2535
      End
      Begin XtremeSuiteControls.GroupBox GroupBox5 
         Height          =   2415
         Left            =   8160
         TabIndex        =   2
         Top             =   120
         Width           =   4095
         _Version        =   786432
         _ExtentX        =   7223
         _ExtentY        =   4260
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.PushButton btnExportarEnChequera 
            Height          =   495
            Index           =   0
            Left            =   2160
            TabIndex        =   3
            Top             =   1800
            Width           =   1815
            _Version        =   786432
            _ExtentX        =   3201
            _ExtentY        =   873
            _StockProps     =   79
            Caption         =   "Exportar"
            Enabled         =   0   'False
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton btnBuscarEnChequera 
            Height          =   495
            Index           =   1
            Left            =   120
            TabIndex        =   4
            Top             =   1800
            Width           =   1815
            _Version        =   786432
            _ExtentX        =   3201
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
      End
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   315
         Left            =   2760
         TabIndex        =   5
         Top             =   480
         Width           =   375
         _Version        =   786432
         _ExtentX        =   661
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "X"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.GroupBox GroFechaComprobante 
         Height          =   1215
         Index           =   7
         Left            =   3360
         TabIndex        =   7
         Top             =   120
         Width           =   4695
         _Version        =   786432
         _ExtentX        =   8281
         _ExtentY        =   2143
         _StockProps     =   79
         Caption         =   "Fecha Vencimiento"
         BackColor       =   16744576
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   4
         Begin XtremeSuiteControls.ComboBox cboRangosVtoTerceros 
            Height          =   315
            Index           =   0
            Left            =   720
            TabIndex        =   8
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
            Index           =   0
            Left            =   720
            TabIndex        =   9
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
            Index           =   0
            Left            =   2925
            TabIndex        =   10
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
            Index           =   7
            Left            =   120
            TabIndex        =   13
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
            Index           =   7
            Left            =   165
            TabIndex        =   12
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
            Index           =   7
            Left            =   2400
            TabIndex        =   11
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
         Index           =   8
         Left            =   3360
         TabIndex        =   14
         Top             =   1320
         Width           =   4695
         _Version        =   786432
         _ExtentX        =   8281
         _ExtentY        =   2143
         _StockProps     =   79
         Caption         =   "Fecha Emitido"
         BackColor       =   16744576
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   4
         Begin XtremeSuiteControls.ComboBox cboRangosRboEmitido 
            Height          =   315
            Index           =   0
            Left            =   720
            TabIndex        =   15
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
            Index           =   0
            Left            =   2925
            TabIndex        =   16
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
            Index           =   0
            Left            =   720
            TabIndex        =   17
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
            Index           =   8
            Left            =   120
            TabIndex        =   20
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
            Index           =   8
            Left            =   165
            TabIndex        =   19
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
            Index           =   8
            Left            =   2400
            TabIndex        =   18
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
      Begin VB.Label Label6 
         Caption         =   "Número:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   2535
      End
   End
   Begin GridEX20.GridEX grid_cheques 
      Height          =   8535
      Left            =   120
      TabIndex        =   22
      Top             =   3120
      Width           =   12405
      _ExtentX        =   21881
      _ExtentY        =   15055
      Version         =   "2.0"
      PreviewRowIndent=   200
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      EmptyRows       =   -1  'True
      PreviewColumn   =   "monto"
      PreviewRowLines =   1
      ColumnAutoResize=   -1  'True
      MethodHoldFields=   -1  'True
      RowHeaders      =   -1  'True
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   11
      Column(1)       =   "frmAdminChequesConciliacion.frx":0000
      Column(2)       =   "frmAdminChequesConciliacion.frx":0164
      Column(3)       =   "frmAdminChequesConciliacion.frx":02BC
      Column(4)       =   "frmAdminChequesConciliacion.frx":040C
      Column(5)       =   "frmAdminChequesConciliacion.frx":0554
      Column(6)       =   "frmAdminChequesConciliacion.frx":0694
      Column(7)       =   "frmAdminChequesConciliacion.frx":07F8
      Column(8)       =   "frmAdminChequesConciliacion.frx":094C
      Column(9)       =   "frmAdminChequesConciliacion.frx":0A88
      Column(10)      =   "frmAdminChequesConciliacion.frx":0B48
      Column(11)      =   "frmAdminChequesConciliacion.frx":0CDC
      FormatStylesCount=   7
      FormatStyle(1)  =   "frmAdminChequesConciliacion.frx":0E5C
      FormatStyle(2)  =   "frmAdminChequesConciliacion.frx":0F94
      FormatStyle(3)  =   "frmAdminChequesConciliacion.frx":1044
      FormatStyle(4)  =   "frmAdminChequesConciliacion.frx":10F8
      FormatStyle(5)  =   "frmAdminChequesConciliacion.frx":11D0
      FormatStyle(6)  =   "frmAdminChequesConciliacion.frx":1288
      FormatStyle(7)  =   "frmAdminChequesConciliacion.frx":1368
      ImageCount      =   0
      PrinterProperties=   "frmAdminChequesConciliacion.frx":1424
   End
   Begin VB.Label Label13 
      Caption         =   "Label13"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   2835
      Width           =   8055
   End
End
Attribute VB_Name = "frmAdminChequesConciliacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

