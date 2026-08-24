VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmAdminCheques 
   Caption         =   "Administración de cheques"
   ClientHeight    =   10680
   ClientLeft      =   165
   ClientTop       =   2280
   ClientWidth     =   15360
   Icon            =   "frmAdminCheques.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   13980
   ScaleMode       =   0  'User
   ScaleWidth      =   26915.51
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   12315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   25935
      _Version        =   786432
      _ExtentX        =   45746
      _ExtentY        =   21722
      _StockProps     =   68
      Appearance      =   10
      Color           =   128
      PaintManager.BoldSelected=   -1  'True
      PaintManager.ShowIcons=   -1  'True
      ItemCount       =   4
      SelectedItem    =   1
      Item(0).Caption =   "Cartera"
      Item(0).ControlCount=   3
      Item(0).Control(0)=   "Frame3"
      Item(0).Control(1)=   "grid_cartera_cheques"
      Item(0).Control(2)=   "lbContadorChequesEnCartera"
      Item(1).Caption =   "Administrar Chequeras"
      Item(1).ControlCount=   9
      Item(1).Control(0)=   "grid_chequeras"
      Item(1).Control(1)=   "grid_cheques"
      Item(1).Control(2)=   "GroupBox1"
      Item(1).Control(3)=   "GroupBox4"
      Item(1).Control(4)=   "Label12"
      Item(1).Control(5)=   "Label13"
      Item(1).Control(6)=   "Label14"
      Item(1).Control(7)=   "cmdExportar"
      Item(1).Control(8)=   "cmdExportarChequeras"
      Item(2).Caption =   "Cheques Propios Utilizados"
      Item(2).ControlCount=   3
      Item(2).Control(0)=   "GroupBox2"
      Item(2).Control(1)=   "gridChequesEmitidos"
      Item(2).Control(2)=   "lbContadorChequesPropiosUtilizados"
      Item(3).Caption =   "Cheques 3eros Utilizados"
      Item(3).ControlCount=   3
      Item(3).Control(0)=   "GroupBox3"
      Item(3).Control(1)=   "grdCheques3eros"
      Item(3).Control(2)=   "lbContador3erosUtilizados"
      Begin XtremeSuiteControls.PushButton cmdExportarChequeras 
         Height          =   255
         Left            =   7920
         TabIndex        =   158
         Top             =   3120
         Width           =   1695
         _Version        =   786432
         _ExtentX        =   2990
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Exportar"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton cmdExportar 
         Height          =   255
         Left            =   20160
         TabIndex        =   157
         Top             =   3075
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
         Left            =   9720
         TabIndex        =   126
         Top             =   360
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
         Begin XtremeSuiteControls.GroupBox GroupBox6 
            Height          =   1575
            Left            =   120
            TabIndex        =   159
            Top             =   960
            Width           =   3135
            _Version        =   786432
            _ExtentX        =   5530
            _ExtentY        =   2778
            _StockProps     =   79
            Caption         =   "Conciliación de cheques"
            UseVisualStyle  =   -1  'True
            Begin VB.CheckBox chkOcultarIngresados 
               Caption         =   "Ocultar ingresados"
               Height          =   255
               Left            =   120
               TabIndex        =   162
               Top             =   360
               Value           =   1  'Checked
               Width           =   2895
            End
            Begin XtremeSuiteControls.PushButton PushButton5 
               Height          =   495
               Left            =   1800
               TabIndex        =   160
               Top             =   960
               Width           =   1215
               _Version        =   786432
               _ExtentX        =   2143
               _ExtentY        =   873
               _StockProps     =   79
               Caption         =   "Conciliar seleccionados"
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.DateTimePicker dtFechaConciliar 
               Height          =   315
               Left            =   120
               TabIndex        =   161
               Top             =   1080
               Width           =   1470
               _Version        =   786432
               _ExtentX        =   2593
               _ExtentY        =   556
               _StockProps     =   68
               Format          =   1
               CurrentDate     =   46258.5631828704
            End
         End
         Begin XtremeSuiteControls.GroupBox GroupBox5 
            Height          =   2415
            Left            =   8160
            TabIndex        =   150
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
               TabIndex        =   151
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
               TabIndex        =   152
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
            TabIndex        =   142
            Top             =   480
            Width           =   375
            _Version        =   786432
            _ExtentX        =   661
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "X"
            UseVisualStyle  =   -1  'True
         End
         Begin VB.TextBox TxtNumeroChequeEnChequera 
            Height          =   285
            Left            =   120
            TabIndex        =   141
            Top             =   480
            Width           =   2535
         End
         Begin XtremeSuiteControls.GroupBox GroFechaComprobante 
            Height          =   1215
            Index           =   7
            Left            =   3360
            TabIndex        =   127
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
               TabIndex        =   128
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
               TabIndex        =   129
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
               TabIndex        =   130
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
               Index           =   7
               Left            =   2400
               TabIndex        =   133
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
               Index           =   7
               Left            =   165
               TabIndex        =   132
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
               Index           =   7
               Left            =   120
               TabIndex        =   131
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
            Index           =   8
            Left            =   3360
            TabIndex        =   134
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
               TabIndex        =   135
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
               TabIndex        =   136
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
               TabIndex        =   137
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
               Index           =   8
               Left            =   2400
               TabIndex        =   140
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
               Index           =   8
               Left            =   165
               TabIndex        =   139
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
               Index           =   8
               Left            =   120
               TabIndex        =   138
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
         Begin VB.Label Label6 
            Caption         =   "Número:"
            Height          =   255
            Left            =   120
            TabIndex        =   144
            Top             =   240
            Width           =   2535
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Parámetros de búsqueda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Left            =   -69880
         TabIndex        =   86
         Top             =   360
         Visible         =   0   'False
         Width           =   15135
         Begin VB.TextBox txtNumeroChequeCartera 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   240
            TabIndex        =   98
            Top             =   480
            Width           =   2535
         End
         Begin VB.Frame Frame 
            Height          =   2775
            Index           =   0
            Left            =   11160
            TabIndex        =   94
            Top             =   120
            Width           =   3855
            Begin XtremeSuiteControls.PushButton btnBuscarEnCartera 
               Default         =   -1  'True
               Height          =   495
               Index           =   0
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
            Begin XtremeSuiteControls.PushButton btnExportarCartera 
               Height          =   495
               Index           =   1
               Left            =   2160
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
               Index           =   0
               Left            =   120
               TabIndex        =   97
               Top             =   1680
               Width           =   3615
               _Version        =   786432
               _ExtentX        =   6376
               _ExtentY        =   661
               _StockProps     =   93
               Appearance      =   6
            End
         End
         Begin VB.TextBox txtOrigen 
            Height          =   315
            Left            =   240
            TabIndex        =   88
            Top             =   2280
            Width           =   2535
         End
         Begin XtremeSuiteControls.PushButton btnBorrarNumeroCartera 
            Height          =   315
            Left            =   2880
            TabIndex        =   87
            Top             =   480
            Width           =   375
            _Version        =   786432
            _ExtentX        =   661
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "X"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton btnBorrarOrigen 
            Height          =   315
            Left            =   2880
            TabIndex        =   89
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
            TabIndex        =   90
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
            TabIndex        =   91
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
            TabIndex        =   92
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
            TabIndex        =   93
            Top             =   1080
            Width           =   2535
            _Version        =   786432
            _ExtentX        =   4471
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Text            =   "cboBanco"
         End
         Begin XtremeSuiteControls.GroupBox GroFechaComprobante 
            Height          =   1215
            Index           =   1
            Left            =   5400
            TabIndex        =   99
            Top             =   240
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
            Begin XtremeSuiteControls.ComboBox cboRangosVtoCartera 
               Height          =   315
               Index           =   0
               Left            =   720
               TabIndex        =   100
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
               TabIndex        =   101
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
               TabIndex        =   102
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
               Index           =   1
               Left            =   120
               TabIndex        =   105
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
               Index           =   1
               Left            =   165
               TabIndex        =   104
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
               Index           =   1
               Left            =   2400
               TabIndex        =   103
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
            Height          =   1335
            Index           =   0
            Left            =   5400
            TabIndex        =   106
            Top             =   1560
            Width           =   4695
            _Version        =   786432
            _ExtentX        =   8281
            _ExtentY        =   2355
            _StockProps     =   79
            Caption         =   "Fecha Recibido"
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
            Begin XtremeSuiteControls.ComboBox cboRangosRboCartera 
               Height          =   315
               Index           =   1
               Left            =   720
               TabIndex        =   107
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
               TabIndex        =   108
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
               TabIndex        =   109
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
               Index           =   0
               Left            =   120
               TabIndex        =   112
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
               Index           =   0
               Left            =   165
               TabIndex        =   111
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
               Index           =   0
               Left            =   2400
               TabIndex        =   110
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
         Begin VB.Label Label1 
            Caption         =   "Número:"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   116
            Top             =   240
            Width           =   2535
         End
         Begin VB.Label Label1 
            Caption         =   "Banco:"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   115
            Top             =   870
            Width           =   2535
         End
         Begin VB.Label Label1 
            Caption         =   "Clasificación:"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   114
            Top             =   1440
            Width           =   2535
         End
         Begin VB.Label Label1 
            Caption         =   "Origen:"
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   113
            Top             =   2040
            Width           =   2535
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox3 
         Height          =   3855
         Left            =   -69880
         TabIndex        =   28
         Top             =   360
         Visible         =   0   'False
         Width           =   15135
         _Version        =   786432
         _ExtentX        =   26696
         _ExtentY        =   6800
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
         Begin XtremeSuiteControls.ComboBox cboClientes3erosUti 
            Height          =   315
            Left            =   240
            TabIndex        =   85
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
            Left            =   5160
            TabIndex        =   83
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
            TabIndex        =   82
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
            TabIndex        =   76
            Top             =   2280
            Width           =   375
            _Version        =   786432
            _ExtentX        =   661
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "X"
            Enabled         =   0   'False
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton btnBorrarOrigenTerceros 
            Height          =   315
            Left            =   4080
            TabIndex        =   75
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
            Left            =   2880
            TabIndex        =   74
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
            Left            =   2880
            TabIndex        =   73
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
            TabIndex        =   72
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
            Height          =   3615
            Index           =   2
            Left            =   11160
            TabIndex        =   66
            Top             =   120
            Width           =   3855
            Begin XtremeSuiteControls.PushButton btnBuscar 
               Height          =   495
               Index           =   0
               Left            =   120
               TabIndex        =   67
               Top             =   3000
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
               Left            =   2160
               TabIndex        =   68
               Top             =   3000
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
               TabIndex        =   69
               Top             =   2400
               Width           =   3615
               _Version        =   786432
               _ExtentX        =   6376
               _ExtentY        =   661
               _StockProps     =   93
               Appearance      =   6
            End
         End
         Begin VB.TextBox txtNumeroOP 
            Enabled         =   0   'False
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
            TabIndex        =   48
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
               Index           =   2
               Left            =   720
               TabIndex        =   49
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
               TabIndex        =   50
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
               TabIndex        =   51
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
               TabIndex        =   54
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
               TabIndex        =   53
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
               TabIndex        =   52
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
            TabIndex        =   55
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
               Index           =   2
               Left            =   720
               TabIndex        =   56
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
               TabIndex        =   57
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
               TabIndex        =   58
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
               TabIndex        =   61
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
               TabIndex        =   60
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
               TabIndex        =   59
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
            Index           =   6
            Left            =   6360
            TabIndex        =   117
            Top             =   2520
            Width           =   4695
            _Version        =   786432
            _ExtentX        =   8281
            _ExtentY        =   2143
            _StockProps     =   79
            Caption         =   "Fecha Recepción"
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
            Begin XtremeSuiteControls.ComboBox cboRangosRboRecibido 
               Height          =   315
               Index           =   0
               Left            =   720
               TabIndex        =   118
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
            Begin XtremeSuiteControls.DateTimePicker dtpHastaRboRecibido 
               Height          =   315
               Index           =   0
               Left            =   2925
               TabIndex        =   119
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
            Begin XtremeSuiteControls.DateTimePicker dtpDesdeRboRecibido 
               Height          =   315
               Index           =   0
               Left            =   720
               TabIndex        =   120
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
               Index           =   6
               Left            =   120
               TabIndex        =   123
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
               Index           =   6
               Left            =   165
               TabIndex        =   122
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
               Index           =   6
               Left            =   2400
               TabIndex        =   121
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
            TabIndex        =   84
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
            TabIndex        =   71
            Top             =   1440
            Width           =   2535
         End
         Begin XtremeSuiteControls.Label lblOP 
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   70
            Top             =   2010
            Width           =   855
            _Version        =   786432
            _ExtentX        =   1508
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "O.P:"
            Enabled         =   0   'False
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
         Begin XtremeSuiteControls.PushButton PushButton2 
            Height          =   315
            Left            =   5160
            TabIndex        =   81
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
            TabIndex        =   79
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
            Left            =   2880
            TabIndex        =   78
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
            TabIndex        =   77
            Top             =   2280
            Width           =   375
            _Version        =   786432
            _ExtentX        =   661
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "X"
            Enabled         =   0   'False
            UseVisualStyle  =   -1  'True
         End
         Begin VB.Frame Frame 
            Height          =   2775
            Index           =   1
            Left            =   11160
            TabIndex        =   62
            Top             =   120
            Width           =   3855
            Begin XtremeSuiteControls.PushButton btnBuscarChePropios 
               Height          =   495
               Index           =   1
               Left            =   120
               TabIndex        =   63
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
               Left            =   2160
               TabIndex        =   64
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
               TabIndex        =   65
               Top             =   1680
               Width           =   3615
               _Version        =   786432
               _ExtentX        =   6376
               _ExtentY        =   661
               _StockProps     =   93
               Appearance      =   6
            End
         End
         Begin VB.TextBox txtIdOP 
            Enabled         =   0   'False
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
            Left            =   2880
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
            Left            =   4800
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
            TabIndex        =   34
            Top             =   240
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
            Begin XtremeSuiteControls.ComboBox cboRangosVtoPropios 
               Height          =   315
               Index           =   1
               Left            =   720
               TabIndex        =   35
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
               TabIndex        =   36
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
               TabIndex        =   37
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
            Begin XtremeSuiteControls.Label lblDesde 
               Height          =   195
               Index           =   2
               Left            =   165
               TabIndex        =   39
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
               TabIndex        =   38
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
            TabIndex        =   41
            Top             =   1560
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
            Begin XtremeSuiteControls.ComboBox cboRangosRboPropios 
               Height          =   315
               Index           =   0
               Left            =   720
               TabIndex        =   42
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
               TabIndex        =   43
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
            Begin XtremeSuiteControls.Label lblRango 
               Height          =   195
               Index           =   3
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
            Begin XtremeSuiteControls.Label lblDesde 
               Height          =   195
               Index           =   3
               Left            =   165
               TabIndex        =   46
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
               TabIndex        =   45
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
            Caption         =   "Destino:"
            Height          =   180
            Index           =   1
            Left            =   2040
            TabIndex        =   80
            Top             =   2040
            Width           =   1905
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "O.P:"
            Enabled         =   0   'False
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
         Height          =   2700
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   9525
         _Version        =   786432
         _ExtentX        =   16801
         _ExtentY        =   4762
         _StockProps     =   79
         Caption         =   "Crear Chequera"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
         Begin XtremeSuiteControls.ComboBox cboCuentaBancariaChequera 
            Height          =   315
            Left            =   960
            TabIndex        =   155
            Top             =   1635
            Width           =   2970
            _Version        =   786432
            _ExtentX        =   5239
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Appearance      =   6
            Text            =   "ComboBox1"
         End
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
            Width           =   5280
         End
         Begin XtremeSuiteControls.ComboBox cboMonedas 
            Height          =   315
            Left            =   975
            TabIndex        =   3
            Top             =   2040
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
            Height          =   510
            Left            =   7440
            TabIndex        =   4
            Top             =   2040
            Width           =   1935
            _Version        =   786432
            _ExtentX        =   3413
            _ExtentY        =   900
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
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            Caption         =   "Cta. Bcia."
            Height          =   255
            Left            =   165
            TabIndex        =   156
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Desde"
            Height          =   270
            Left            =   330
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
            Left            =   -45
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
            Top             =   2115
            Width           =   750
         End
      End
      Begin GridEX20.GridEX grid_cheques 
         Height          =   8535
         Left            =   9720
         TabIndex        =   14
         Top             =   3360
         Width           =   12405
         _ExtentX        =   21881
         _ExtentY        =   15055
         Version         =   "2.0"
         PreviewRowIndent=   200
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         EmptyRows       =   -1  'True
         PreviewColumn   =   6
         PreviewRowLines =   1
         ColumnAutoResize=   -1  'True
         MethodHoldFields=   -1  'True
         RowHeaders      =   -1  'True
         DataMode        =   99
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   8
         Column(1)       =   "frmAdminCheques.frx":000C
         Column(2)       =   "frmAdminCheques.frx":0178
         Column(3)       =   "frmAdminCheques.frx":02B8
         Column(4)       =   "frmAdminCheques.frx":041C
         Column(5)       =   "frmAdminCheques.frx":0570
         Column(6)       =   "frmAdminCheques.frx":06AC
         Column(7)       =   "frmAdminCheques.frx":076C
         Column(8)       =   "frmAdminCheques.frx":0900
         FormatStylesCount=   7
         FormatStyle(1)  =   "frmAdminCheques.frx":0A80
         FormatStyle(2)  =   "frmAdminCheques.frx":0BB8
         FormatStyle(3)  =   "frmAdminCheques.frx":0C68
         FormatStyle(4)  =   "frmAdminCheques.frx":0D1C
         FormatStyle(5)  =   "frmAdminCheques.frx":0DF4
         FormatStyle(6)  =   "frmAdminCheques.frx":0EAC
         FormatStyle(7)  =   "frmAdminCheques.frx":0F8C
         ImageCount      =   0
         PrinterProperties=   "frmAdminCheques.frx":1048
      End
      Begin GridEX20.GridEX grid_chequeras 
         Height          =   8490
         Left            =   120
         TabIndex        =   15
         Top             =   3360
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   14975
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
         ColumnsCount    =   8
         Column(1)       =   "frmAdminCheques.frx":1220
         Column(2)       =   "frmAdminCheques.frx":1384
         Column(3)       =   "frmAdminCheques.frx":14CC
         Column(4)       =   "frmAdminCheques.frx":1604
         Column(5)       =   "frmAdminCheques.frx":1760
         Column(6)       =   "frmAdminCheques.frx":18D0
         Column(7)       =   "frmAdminCheques.frx":1A40
         Column(8)       =   "frmAdminCheques.frx":1B88
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmAdminCheques.frx":1D1C
         FormatStyle(2)  =   "frmAdminCheques.frx":1E54
         FormatStyle(3)  =   "frmAdminCheques.frx":1F04
         FormatStyle(4)  =   "frmAdminCheques.frx":1FB8
         FormatStyle(5)  =   "frmAdminCheques.frx":2090
         FormatStyle(6)  =   "frmAdminCheques.frx":2148
         ImageCount      =   0
         PrinterProperties=   "frmAdminCheques.frx":2228
      End
      Begin GridEX20.GridEX grdCheques3eros 
         Height          =   4665
         Left            =   -69880
         TabIndex        =   124
         Top             =   4560
         Visible         =   0   'False
         Width           =   15135
         _ExtentX        =   26696
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
         ColumnsCount    =   15
         Column(1)       =   "frmAdminCheques.frx":2400
         Column(2)       =   "frmAdminCheques.frx":2544
         Column(3)       =   "frmAdminCheques.frx":267C
         Column(4)       =   "frmAdminCheques.frx":2794
         Column(5)       =   "frmAdminCheques.frx":28FC
         Column(6)       =   "frmAdminCheques.frx":2A5C
         Column(7)       =   "frmAdminCheques.frx":2BC4
         Column(8)       =   "frmAdminCheques.frx":2D0C
         Column(9)       =   "frmAdminCheques.frx":2E54
         Column(10)      =   "frmAdminCheques.frx":2FC4
         Column(11)      =   "frmAdminCheques.frx":3118
         Column(12)      =   "frmAdminCheques.frx":3270
         Column(13)      =   "frmAdminCheques.frx":33CC
         Column(14)      =   "frmAdminCheques.frx":3528
         Column(15)      =   "frmAdminCheques.frx":3670
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmAdminCheques.frx":3798
         FormatStyle(2)  =   "frmAdminCheques.frx":38D0
         FormatStyle(3)  =   "frmAdminCheques.frx":3980
         FormatStyle(4)  =   "frmAdminCheques.frx":3A34
         FormatStyle(5)  =   "frmAdminCheques.frx":3B0C
         FormatStyle(6)  =   "frmAdminCheques.frx":3BC4
         ImageCount      =   0
         PrinterProperties=   "frmAdminCheques.frx":3CA4
      End
      Begin GridEX20.GridEX gridChequesEmitidos 
         Height          =   5385
         Left            =   -69880
         TabIndex        =   125
         Top             =   3720
         Visible         =   0   'False
         Width           =   15135
         _ExtentX        =   26696
         _ExtentY        =   9499
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
         ColumnsCount    =   13
         Column(1)       =   "frmAdminCheques.frx":3E7C
         Column(2)       =   "frmAdminCheques.frx":402C
         Column(3)       =   "frmAdminCheques.frx":4144
         Column(4)       =   "frmAdminCheques.frx":427C
         Column(5)       =   "frmAdminCheques.frx":43DC
         Column(6)       =   "frmAdminCheques.frx":4534
         Column(7)       =   "frmAdminCheques.frx":469C
         Column(8)       =   "frmAdminCheques.frx":4804
         Column(9)       =   "frmAdminCheques.frx":4924
         Column(10)      =   "frmAdminCheques.frx":4A54
         Column(11)      =   "frmAdminCheques.frx":4B8C
         Column(12)      =   "frmAdminCheques.frx":4CE8
         Column(13)      =   "frmAdminCheques.frx":4E40
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmAdminCheques.frx":4F78
         FormatStyle(2)  =   "frmAdminCheques.frx":50B0
         FormatStyle(3)  =   "frmAdminCheques.frx":5160
         FormatStyle(4)  =   "frmAdminCheques.frx":5214
         FormatStyle(5)  =   "frmAdminCheques.frx":52EC
         FormatStyle(6)  =   "frmAdminCheques.frx":53A4
         ImageCount      =   0
         PrinterProperties=   "frmAdminCheques.frx":5484
      End
      Begin GridEX20.GridEX grid_cartera_cheques 
         Height          =   4665
         Left            =   -69880
         TabIndex        =   143
         Top             =   3720
         Visible         =   0   'False
         Width           =   15135
         _ExtentX        =   26696
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
         Column(1)       =   "frmAdminCheques.frx":565C
         Column(2)       =   "frmAdminCheques.frx":57FC
         Column(3)       =   "frmAdminCheques.frx":5968
         Column(4)       =   "frmAdminCheques.frx":5AF0
         Column(5)       =   "frmAdminCheques.frx":5CF4
         Column(6)       =   "frmAdminCheques.frx":5E54
         Column(7)       =   "frmAdminCheques.frx":5FB0
         Column(8)       =   "frmAdminCheques.frx":6120
         Column(9)       =   "frmAdminCheques.frx":631C
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmAdminCheques.frx":64BC
         FormatStyle(2)  =   "frmAdminCheques.frx":65F4
         FormatStyle(3)  =   "frmAdminCheques.frx":66A4
         FormatStyle(4)  =   "frmAdminCheques.frx":6758
         FormatStyle(5)  =   "frmAdminCheques.frx":6830
         FormatStyle(6)  =   "frmAdminCheques.frx":68E8
         ImageCount      =   0
         PrinterProperties=   "frmAdminCheques.frx":69C8
      End
      Begin VB.Label Label14 
         Caption         =   "* Las chequeras que tienen el tilde marcado no apareceran en el listado de chequeras."
         Height          =   255
         Left            =   120
         TabIndex        =   154
         Top             =   11880
         Width           =   7335
      End
      Begin VB.Label Label13 
         Caption         =   "Label13"
         Height          =   255
         Left            =   9720
         TabIndex        =   149
         Top             =   3075
         Width           =   8055
      End
      Begin VB.Label Label12 
         Caption         =   "Label12"
         Height          =   255
         Left            =   120
         TabIndex        =   148
         Top             =   3080
         Width           =   7335
      End
      Begin XtremeSuiteControls.Label lbContadorChequesEnCartera 
         Height          =   375
         Left            =   -69880
         TabIndex        =   147
         Top             =   3360
         Visible         =   0   'False
         Width           =   5415
         _Version        =   786432
         _ExtentX        =   9551
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "lbContadorChequesEnCartera"
      End
      Begin XtremeSuiteControls.Label lbContador3erosUtilizados 
         Height          =   375
         Left            =   -69880
         TabIndex        =   146
         Top             =   4200
         Visible         =   0   'False
         Width           =   6375
         _Version        =   786432
         _ExtentX        =   11245
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "lbContador3erosUtilizados"
      End
      Begin XtremeSuiteControls.Label lbContadorChequesPropiosUtilizados 
         Height          =   375
         Left            =   -69880
         TabIndex        =   145
         Top             =   3360
         Visible         =   0   'False
         Width           =   6375
         _Version        =   786432
         _ExtentX        =   11245
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "lbContadorChequesPropiosUtilizados"
      End
   End
   Begin GridEX20.GridEX gridBancos 
      Height          =   1845
      Left            =   480
      TabIndex        =   33
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
      Column(1)       =   "frmAdminCheques.frx":6BA0
      Column(2)       =   "frmAdminCheques.frx":6CA0
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminCheques.frx":6D90
      FormatStyle(2)  =   "frmAdminCheques.frx":6EC8
      FormatStyle(3)  =   "frmAdminCheques.frx":6F78
      FormatStyle(4)  =   "frmAdminCheques.frx":702C
      FormatStyle(5)  =   "frmAdminCheques.frx":7104
      FormatStyle(6)  =   "frmAdminCheques.frx":71BC
      ImageCount      =   0
      PrinterProperties=   "frmAdminCheques.frx":729C
   End
   Begin XtremeSuiteControls.Label Label11 
      Height          =   13575
      Left            =   7560
      TabIndex        =   153
      Top             =   -9720
      Width           =   8535
      _Version        =   786432
      _ExtentX        =   15055
      _ExtentY        =   23945
      _StockProps     =   79
      Caption         =   "* Las chequeras que tengan el tilde en usadas no se mostraran en el listado"
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
Dim chequera As Collection
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
Private cargandoChequera As Boolean
Private idChequeraMostrada As Long

Private Sub btnBorrarBanco_Click()
    Me.cboBancoCartera.ListIndex = -1
End Sub

Private Sub btnBorrarBancosTerceros_Click()
    Me.cboBancos3ero.ListIndex = -1
End Sub

Private Sub btnBorrarClasificacion_Click()
    Me.cboClasificacion.ListIndex = -1
End Sub

Private Sub btnBuscar_Click_1()

    Dim q As String

    Set cheques3 = New Collection

    q = "cheq.propio = 0 " _
      & "AND cheq.en_cartera = 0 " _
      & "AND (" _
      & "IFNULL(cheq.orden_pago_origen, 0) > 0 OR " _
      & "IFNULL(cheq.liquidacion_caja_origen, 0) > 0 OR " _
      & "IFNULL(cheq.pago_a_cuenta_origen, 0) > 0 OR " _
      & "IFNULL(cheq.movimiento_origen, 0) > 0" _
      & ")"

    If LenB(Me.txtNumeroCheque3ero.Text) > 0 Then
        q = q & " AND cheq.numero = " & _
                val(Me.txtNumeroCheque3ero.Text)
    End If

    If LenB(Me.txtNumeroOP.Text) > 0 Then
        q = q & " AND cheq.orden_pago_origen = " & _
                val(Me.txtNumeroOP.Text)
    End If

    Me.grdCheques3eros.ItemCount = 0

    Set cheques3 = DAOCheques.FindAllTercerosUti(q)

    If cheques3 Is Nothing Then
        Set cheques3 = New Collection
    End If

    Me.grdCheques3eros.ItemCount = cheques3.count

    Me.lbContador3erosUtilizados.caption = _
        "Cheques encontrados: [ " & cheques3.count & " ]"

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
    
    q = "propio=0 and en_cartera=0"
    
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
    
    
    If Not IsNull(Me.dtpDesdeRboRecibido(0)) Then
        q = q & " and fecha_recibido>=" & conectar.Escape(Format(Me.dtpDesdeRboRecibido(0).value, "yyyy-mm-dd"))
    End If

    If Not IsNull(Me.dtpHastaRboRecibido(0)) Then
        q = q & " and fecha_recibido<=" & conectar.Escape(Format(Me.dtpHastaRboRecibido(0).value, "yyyy-mm-dd"))
    End If

    If Me.cboBancos3ero.ListIndex > -1 Then
        q = q & " and cheq.id_banco=" & Me.cboBancos3ero.ItemData(Me.cboBancos3ero.ListIndex)
    End If

    If LenB(Me.txtNumeroCheque3ero) > 0 Then
        q = q & " and cheq.numero like '%" & Trim(Me.txtNumeroCheque3ero.Text) & "%'"
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

    Me.lbContador3erosUtilizados.caption = "Cheques encontrados: [ " & cheques3.count & " ]"

    GridEXHelper.AutoSizeColumns Me.grdCheques3eros
    
End Sub


Private Sub btnBuscarCartera_Click(Index As Integer)

End Sub

Private Sub btnBuscarChePropios_Click(Index As Integer)
    
    Dim q As String
    Set cheques1 = New Collection

    q = "ingresado=" & Abs(Me.chkIngresados.value) & " and propio=1 AND en_cartera= 0"
    
    
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

    Me.lbContadorChequesPropiosUtilizados.caption = "Cheques encontrados: [ " & cheques1.count & " ]"

    Me.gridChequesEmitidos.ItemCount = cheques1.count
    GridEXHelper.AutoSizeColumns Me.gridChequesEmitidos
    
End Sub


Private Sub btnBuscarEnChequera_Click(Index As Integer)
    BuscarChequeEnChequera
End Sub


Private Sub cmdExportar_Click()

    On Error GoTo err1

    Const xlCenter As Long = -4108
    Const xlLandscape As Long = 2
    Const xlOpenXMLWorkbook As Long = 51

    Dim xlApplication As Object
    Dim xlWorkbook As Object
    Dim xlWorksheet As Object

    Dim ch As cheque
    Dim fila As Long
    Dim filaEncabezado As Long
    Dim ultimaFila As Long
    Dim archivo As String

    Dim NombreBanco As String
    Dim nombreMoneda As String

    If tmpChequera Is Nothing Then
        MsgBox "Seleccione una chequera para exportar.", _
               vbExclamation, _
               "Exportar chequera"
        Exit Sub
    End If

    If tmpChequera.Cheques Is Nothing Then
        MsgBox "No hay cheques para exportar.", _
               vbInformation, _
               "Exportar chequera"
        Exit Sub
    End If

    If tmpChequera.Cheques.count = 0 Then
        MsgBox "No hay cheques para exportar.", _
               vbInformation, _
               "Exportar chequera"
        Exit Sub
    End If

    If Not tmpChequera.Banco Is Nothing Then
        NombreBanco = tmpChequera.Banco.nombre
    End If

    If Not tmpChequera.moneda Is Nothing Then
        nombreMoneda = tmpChequera.moneda.NombreCorto
    End If

    Me.MousePointer = vbHourglass

    Set xlApplication = CreateObject("Excel.Application")
    Set xlWorkbook = xlApplication.Workbooks.Add
    Set xlWorksheet = xlWorkbook.Worksheets.item(1)

    xlApplication.ScreenUpdating = False
    xlApplication.DisplayAlerts = False

    xlWorksheet.Name = "Chequera " & tmpChequera.numero

    'Título
    With xlWorksheet.Range("A1:H1")
        .Merge
        .value = "REPORTE DE CHEQUERA"
        .Font.Bold = True
        .Font.Size = 14
        .HorizontalAlignment = xlCenter
        .Interior.Color = &HD9EAD3
    End With

    'Datos generales
    xlWorksheet.Cells(2, 1).value = "Banco:"
    xlWorksheet.Cells(2, 2).value = NombreBanco

    xlWorksheet.Cells(2, 4).value = "Chequera N°:"
    xlWorksheet.Cells(2, 5).value = tmpChequera.numero

    xlWorksheet.Cells(2, 7).value = "Moneda:"
    xlWorksheet.Cells(2, 8).value = nombreMoneda

    xlWorksheet.Cells(3, 1).value = "Rango:"
    xlWorksheet.Cells(3, 2).value = _
        tmpChequera.NumeroDesde & " al " & tmpChequera.NumeroHasta

    xlWorksheet.Cells(3, 4).value = "Creación:"
    xlWorksheet.Cells(3, 5).value = tmpChequera.fechaCreacion
    xlWorksheet.Cells(3, 5).NumberFormat = "dd/mm/yyyy"

    xlWorksheet.Cells(3, 7).value = "Registros:"
    xlWorksheet.Cells(3, 8).value = tmpChequera.Cheques.count

    xlWorksheet.Range("A2:A3").Font.Bold = True
    xlWorksheet.Range("D2:D3").Font.Bold = True
    xlWorksheet.Range("G2:G3").Font.Bold = True

    'Encabezados
    filaEncabezado = 5

    xlWorksheet.Cells(filaEncabezado, 1).value = "Número"
    xlWorksheet.Cells(filaEncabezado, 2).value = "Monto"
    xlWorksheet.Cells(filaEncabezado, 3).value = "Vencimiento"
    xlWorksheet.Cells(filaEncabezado, 4).value = "Emisión"
    xlWorksheet.Cells(filaEncabezado, 5).value = "Destino"
    xlWorksheet.Cells(filaEncabezado, 6).value = "Uso"
    xlWorksheet.Cells(filaEncabezado, 7).value = "Ingresado"
    xlWorksheet.Cells(filaEncabezado, 8).value = "Fecha ingreso"

    With xlWorksheet.Range( _
            xlWorksheet.Cells(filaEncabezado, 1), _
            xlWorksheet.Cells(filaEncabezado, 8))

        .Font.Bold = True
        .Interior.Color = &HC0C0C0
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = 1
    End With

    'Datos
    fila = filaEncabezado + 1

    For Each ch In tmpChequera.Cheques

        xlWorksheet.Cells(fila, 1).value = ch.numero

        If ch.Utilizado Then

            xlWorksheet.Cells(fila, 2).value = ch.Monto

            If CDbl(ch.FechaVencimiento) > 0 Then
                xlWorksheet.Cells(fila, 3).value = _
                    ch.FechaVencimiento
            End If

            If CDbl(ch.FechaEmision) > 0 Then
                xlWorksheet.Cells(fila, 4).value = _
                    ch.FechaEmision
            End If

            xlWorksheet.Cells(fila, 5).value = _
                ch.OrigenDestino

        End If

        xlWorksheet.Cells(fila, 6).value = _
            DescripcionUsoCheque(ch)

        If ch.entro Then
            xlWorksheet.Cells(fila, 7).value = "SÍ"
        Else
            xlWorksheet.Cells(fila, 7).value = "NO"
        End If

        If CDbl(ch.FechaIngresoBanco) > 0 Then
            xlWorksheet.Cells(fila, 8).value = _
                ch.FechaIngresoBanco
        End If

        fila = fila + 1

    Next ch

    ultimaFila = fila - 1

    'Formato
    xlWorksheet.Range("B6:B" & ultimaFila).NumberFormat = _
        "#,##0.00"

    xlWorksheet.Range("C6:D" & ultimaFila).NumberFormat = _
        "dd/mm/yyyy"

    xlWorksheet.Range("H6:H" & ultimaFila).NumberFormat = _
        "dd/mm/yyyy"

    xlWorksheet.Range( _
        "A" & filaEncabezado & ":H" & ultimaFila).Borders.LineStyle = 1

    xlWorksheet.Range( _
        "A" & filaEncabezado & ":H" & ultimaFila).AutoFilter

    'Total
    xlWorksheet.Cells(fila, 1).value = "TOTAL"
    xlWorksheet.Cells(fila, 1).Font.Bold = True

    xlWorksheet.Cells(fila, 2).Formula = _
        "=SUM(B6:B" & ultimaFila & ")"

    xlWorksheet.Cells(fila, 2).Font.Bold = True
    xlWorksheet.Cells(fila, 2).NumberFormat = "#,##0.00"

    xlWorksheet.Columns("A:H").AutoFit

    xlWorksheet.PageSetup.Orientation = xlLandscape
    xlWorksheet.PageSetup.BottomMargin = _
        xlApplication.CentimetersToPoints(1)

    xlWorksheet.PageSetup.TopMargin = _
        xlApplication.CentimetersToPoints(1)

    xlWorksheet.PageSetup.LeftMargin = _
        xlApplication.CentimetersToPoints(1)

    xlWorksheet.PageSetup.RightMargin = _
        xlApplication.CentimetersToPoints(1)

    archivo = funciones.GetTmpPath() & _
              "Chequera_" & tmpChequera.numero & "_" & _
              Format$(Now, "yyyymmdd_hhnnss") & ".xlsx"

    If Dir$(archivo) <> vbNullString Then
        Kill archivo
    End If

    xlWorkbook.SaveAs archivo, xlOpenXMLWorkbook

    xlApplication.ScreenUpdating = True

    xlWorkbook.Close False
    xlApplication.Quit

    Set xlWorksheet = Nothing
    Set xlWorkbook = Nothing
    Set xlApplication = Nothing

    Me.MousePointer = vbDefault

    funciones.ShellExecute _
        0, "open", archivo, vbNullString, vbNullString, 1

    Exit Sub

err1:
    Me.MousePointer = vbDefault

    On Error Resume Next

    If Not xlWorkbook Is Nothing Then
        xlWorkbook.Close False
    End If

    If Not xlApplication Is Nothing Then
        xlApplication.Quit
    End If

    Set xlWorksheet = Nothing
    Set xlWorkbook = Nothing
    Set xlApplication = Nothing

    On Error GoTo 0

    MsgBox "No se pudo exportar la chequera." & _
           vbCrLf & Err.Description, _
           vbExclamation, _
           "Exportar a Excel"

End Sub

Private Sub cmdExportarChequeras_Click()

    On Error GoTo err1

    Const xlCenter As Long = -4108
    Const xlLandscape As Long = 2
    Const xlOpenXMLWorkbook As Long = 51

    Dim listaChequeras As Collection
    Dim ch As chequera
    Dim cuenta As CuentaBancaria

    Dim xlApplication As Object
    Dim xlWorkbook As Object
    Dim xlWorksheet As Object

    Dim fila As Long
    Dim filaEncabezado As Long
    Dim ultimaFila As Long
    Dim archivo As String

    Dim tipoCuentaTexto As String
    Dim estadoBanco As String
    Dim estadoMoneda As String
    Dim mensajeError As String

    'Traer todas las chequeras ordenadas por banco
    Set listaChequeras = DAOChequeras.GetAll( _
        "1 = 1 ORDER BY banco.nombre, chs.numero")

    If listaChequeras Is Nothing Then
        MsgBox "No se pudieron obtener las chequeras.", _
               vbExclamation, _
               "Exportar chequeras"
        Exit Sub
    End If

    If listaChequeras.count = 0 Then
        MsgBox "No hay chequeras para exportar.", _
               vbInformation, _
               "Exportar chequeras"
        Exit Sub
    End If

    Me.MousePointer = vbHourglass

    Set xlApplication = CreateObject("Excel.Application")
    Set xlWorkbook = xlApplication.Workbooks.Add
    Set xlWorksheet = xlWorkbook.Worksheets.item(1)

    xlApplication.ScreenUpdating = False
    xlApplication.DisplayAlerts = False

    xlWorksheet.Name = "Chequeras y cuentas"

    'Título
    With xlWorksheet.Range("A1:U1")
        .Merge
        .value = "CONTROL DE CHEQUERAS Y CUENTAS BANCARIAS"
        .Font.Bold = True
        .Font.Size = 14
        .HorizontalAlignment = xlCenter

    End With

    xlWorksheet.Cells(2, 1).value = "Fecha de exportación:"
    xlWorksheet.Cells(2, 2).value = Now
    xlWorksheet.Cells(2, 2).NumberFormat = "dd/mm/yyyy hh:mm"

    xlWorksheet.Cells(2, 4).value = "Cantidad de chequeras:"
    xlWorksheet.Cells(2, 5).value = listaChequeras.count

    xlWorksheet.Cells(2, 1).Font.Bold = True
    xlWorksheet.Cells(2, 4).Font.Bold = True

    With xlWorksheet.Range("A3:U3")
        .Merge
        .value = _
            "La validación compara el banco asignado a la " & _
            "chequera con el banco de la cuenta bancaria."
        .Font.Italic = True
    End With

    'Encabezados
    filaEncabezado = 5

    xlWorksheet.Cells(filaEncabezado, 1).value = "ID Chequera"
    xlWorksheet.Cells(filaEncabezado, 2).value = "N° Chequera"
    xlWorksheet.Cells(filaEncabezado, 3).value = "Fecha creación"
    xlWorksheet.Cells(filaEncabezado, 4).value = "Desde"
    xlWorksheet.Cells(filaEncabezado, 5).value = "Hasta"
    xlWorksheet.Cells(filaEncabezado, 6).value = "Usada/Antigua"

    xlWorksheet.Cells(filaEncabezado, 7).value = _
        "ID Banco Chequera"

    xlWorksheet.Cells(filaEncabezado, 8).value = _
        "Banco Chequera"

    xlWorksheet.Cells(filaEncabezado, 9).value = _
        "ID Cuenta Bancaria"

    xlWorksheet.Cells(filaEncabezado, 10).value = _
        "Número Cuenta"

    xlWorksheet.Cells(filaEncabezado, 11).value = _
        "Tipo Cuenta"

    xlWorksheet.Cells(filaEncabezado, 12).value = "CBU"

    xlWorksheet.Cells(filaEncabezado, 13).value = _
        "ID Banco Cuenta"

    xlWorksheet.Cells(filaEncabezado, 14).value = _
        "Banco de la Cuenta"

    xlWorksheet.Cells(filaEncabezado, 15).value = _
        "ID Moneda Chequera"

    xlWorksheet.Cells(filaEncabezado, 16).value = _
        "Moneda Chequera"

    xlWorksheet.Cells(filaEncabezado, 17).value = _
        "ID Moneda Cuenta"

    xlWorksheet.Cells(filaEncabezado, 18).value = _
        "Moneda Cuenta"

    xlWorksheet.Cells(filaEncabezado, 19).value = _
        "Validación Banco"

    xlWorksheet.Cells(filaEncabezado, 20).value = _
        "Validación Moneda"

    xlWorksheet.Cells(filaEncabezado, 21).value = _
        "Observaciones"

    With xlWorksheet.Range("A5:U5")
        .Font.Bold = True

        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = 1
        .WrapText = True
    End With

    'El número de cuenta y el CBU deben tratarse como texto
    xlWorksheet.Columns("J:J").NumberFormat = "@"
    xlWorksheet.Columns("L:L").NumberFormat = "@"

    fila = filaEncabezado + 1

    For Each ch In listaChequeras

        Set cuenta = Nothing
        tipoCuentaTexto = vbNullString
        estadoBanco = vbNullString
        estadoMoneda = vbNullString

        If Not ch.CuentaBancaria Is Nothing Then
            Set cuenta = ch.CuentaBancaria
        End If

        'Datos generales de la chequera
        xlWorksheet.Cells(fila, 1).value = ch.Id
        xlWorksheet.Cells(fila, 2).value = ch.numero
        xlWorksheet.Cells(fila, 3).value = ch.fechaCreacion
        xlWorksheet.Cells(fila, 4).value = ch.NumeroDesde
        xlWorksheet.Cells(fila, 5).value = ch.NumeroHasta

        If ch.usada Then
            xlWorksheet.Cells(fila, 6).value = "SÍ"
        Else
            xlWorksheet.Cells(fila, 6).value = "NO"
        End If

        'Banco asignado directamente a la chequera
        If Not ch.Banco Is Nothing Then
            xlWorksheet.Cells(fila, 7).value = ch.Banco.Id
            xlWorksheet.Cells(fila, 8).value = ch.Banco.nombre
        End If

        'Moneda de la chequera
        If Not ch.moneda Is Nothing Then
            xlWorksheet.Cells(fila, 15).value = ch.moneda.Id
            xlWorksheet.Cells(fila, 16).value = _
                ch.moneda.NombreCorto
        End If

        'Cuenta bancaria asociada
        If Not cuenta Is Nothing Then

            xlWorksheet.Cells(fila, 9).value = cuenta.Id

            xlWorksheet.Cells(fila, 10).value = _
                CStr(cuenta.numero)

            Select Case cuenta.TipoCuenta

                Case TipoCuentaBancaria.CuentaCorriente
                    tipoCuentaTexto = "Cuenta corriente"

                Case TipoCuentaBancaria.CajaAhorro
                    tipoCuentaTexto = "Caja de ahorro"

                Case Else
                    tipoCuentaTexto = "Sin definir"

            End Select

            xlWorksheet.Cells(fila, 11).value = _
                tipoCuentaTexto

            xlWorksheet.Cells(fila, 12).value = _
                CStr(cuenta.CBU)

            If Not cuenta.Banco Is Nothing Then
                xlWorksheet.Cells(fila, 13).value = _
                    cuenta.Banco.Id

                xlWorksheet.Cells(fila, 14).value = _
                    cuenta.Banco.nombre
            End If

            If Not cuenta.moneda Is Nothing Then
                xlWorksheet.Cells(fila, 17).value = _
                    cuenta.moneda.Id

                xlWorksheet.Cells(fila, 18).value = _
                    cuenta.moneda.NombreCorto
            End If

        End If

        '---------------------------------------------
        ' VALIDACIÓN DEL BANCO
        '---------------------------------------------
        If cuenta Is Nothing Then

            estadoBanco = "SIN CUENTA ASOCIADA"

        ElseIf ch.Banco Is Nothing Then

            estadoBanco = "CHEQUERA SIN BANCO"

        ElseIf cuenta.Banco Is Nothing Then

            estadoBanco = "CUENTA SIN BANCO"

        ElseIf ch.Banco.Id = cuenta.Banco.Id Then

            estadoBanco = "CORRECTO"

        Else

            estadoBanco = "REVISAR: BANCO DISTINTO"

        End If

        xlWorksheet.Cells(fila, 19).value = estadoBanco

        If estadoBanco = "CORRECTO" Then
            xlWorksheet.Cells(fila, 19).Interior.Color = Gray

        Else
            xlWorksheet.Cells(fila, 19).Interior.Color = Gray
        End If

        '---------------------------------------------
        ' VALIDACIÓN DE LA MONEDA
        '---------------------------------------------
        If cuenta Is Nothing Then

            estadoMoneda = "SIN CUENTA ASOCIADA"

        ElseIf ch.moneda Is Nothing Then

            estadoMoneda = "CHEQUERA SIN MONEDA"

        ElseIf cuenta.moneda Is Nothing Then

            estadoMoneda = "CUENTA SIN MONEDA"

        ElseIf ch.moneda.Id = cuenta.moneda.Id Then

            estadoMoneda = "CORRECTO"

        Else

            estadoMoneda = "REVISAR: MONEDA DISTINTA"

        End If

        xlWorksheet.Cells(fila, 20).value = estadoMoneda



        xlWorksheet.Cells(fila, 21).value = _
            ch.Observaciones

        fila = fila + 1

    Next ch

    ultimaFila = fila - 1

    'Formatos
    xlWorksheet.Range( _
        "C6:C" & ultimaFila).NumberFormat = "dd/mm/yyyy"

    xlWorksheet.Range( _
        "A5:U" & ultimaFila).Borders.LineStyle = 1

    xlWorksheet.Range( _
        "A5:U" & ultimaFila).AutoFilter

    xlWorksheet.Columns("A:U").AutoFit

    'Evitar columnas excesivamente anchas
    xlWorksheet.Columns("H:H").ColumnWidth = 25
    xlWorksheet.Columns("N:N").ColumnWidth = 25
    xlWorksheet.Columns("S:T").ColumnWidth = 26
    xlWorksheet.Columns("U:U").ColumnWidth = 35
    xlWorksheet.Columns("U:U").WrapText = True

    xlWorksheet.PageSetup.Orientation = xlLandscape
    xlWorksheet.PageSetup.Zoom = False
    xlWorksheet.PageSetup.FitToPagesWide = 1
    xlWorksheet.PageSetup.FitToPagesTall = False

    archivo = funciones.GetTmpPath() & _
              "Chequeras_Cuentas_" & _
              Format$(Now, "yyyymmdd_hhnnss") & ".xlsx"

    If Dir$(archivo) <> vbNullString Then
        Kill archivo
    End If

    xlWorkbook.SaveAs archivo, xlOpenXMLWorkbook

    xlApplication.ScreenUpdating = True

    xlWorkbook.Close False
    xlApplication.Quit

    Set xlWorksheet = Nothing
    Set xlWorkbook = Nothing
    Set xlApplication = Nothing

    Me.MousePointer = vbDefault

    funciones.ShellExecute _
        0, "open", archivo, vbNullString, vbNullString, 1

    Exit Sub

err1:
    mensajeError = Err.Description
    Me.MousePointer = vbDefault

    On Error Resume Next

    If Not xlWorkbook Is Nothing Then
        xlWorkbook.Close False
    End If

    If Not xlApplication Is Nothing Then
        xlApplication.Quit
    End If

    Set xlWorksheet = Nothing
    Set xlWorkbook = Nothing
    Set xlApplication = Nothing

    On Error GoTo 0

    MsgBox "No se pudieron exportar las chequeras." & _
           vbCrLf & mensajeError, _
           vbExclamation, _
           "Exportar a Excel"

End Sub

Private Sub TxtNumeroChequeEnChequera_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        BuscarChequeEnChequera
    End If

End Sub


Private Sub BuscarChequeEnChequera()

    If tmpChequera Is Nothing Then
        MsgBox "Seleccione una chequera.", _
               vbExclamation, _
               "Administración de cheques"
        Exit Sub
    End If

    MostrarChequera True

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
        
        Dim titulo As String
            titulo = "Reporte de Cheques de 3ros Utilizados"
        
        With xlWorksheet.Range("A1:K1")
            .Merge
            .value = titulo
            .Font.Bold = True
            .HorizontalAlignment = -4108 ' xlCenter
        End With

        xlWorksheet.Columns(4).HorizontalAlignment = xlLeft
        xlWorksheet.Columns(7).HorizontalAlignment = xlLeft

        Dim offset As Long
        offset = 3
        
        xlWorksheet.Cells(offset, 1).value = "ID"
        xlWorksheet.Cells(offset, 2).value = "Número"
        xlWorksheet.Cells(offset, 3).value = "Importe"
        xlWorksheet.Cells(offset, 4).value = "Fecha Emisión"
        xlWorksheet.Cells(offset, 5).value = "Fecha Vencimiento"
        xlWorksheet.Cells(offset, 6).value = "Fecha Recepción"
        xlWorksheet.Cells(offset, 7).value = "Banco"
        xlWorksheet.Cells(offset, 8).value = "Origen"
        xlWorksheet.Cells(offset, 9).value = "Recibo Origen"
        xlWorksheet.Cells(offset, 10).value = "Destino"
        xlWorksheet.Cells(offset, 11).value = "OP"
        xlWorksheet.Cells(offset, 12).value = "LIQUID CAJA"
        xlWorksheet.Cells(offset, 13).value = "PAGO A CTA"
        xlWorksheet.Cells(offset, 14).value = "MOV"
        
        xlWorksheet.Range(xlWorksheet.Cells(offset, 1), xlWorksheet.Cells(offset, 14)).Font.Bold = True
        xlWorksheet.Range(xlWorksheet.Cells(offset, 1), xlWorksheet.Cells(offset, 14)).Interior.Color = &HC0C0C0
        
        Dim idx As Integer
        idx = 4

        Dim che As cheque

        'DEFINE EL CONTADOR DEL PROGRESSBAR Y LO INICIA EN 0
        Dim d As Long
        d = 0


        For Each che In cheques3

            xlWorksheet.Cells(idx, 1).value = che.Id
            xlWorksheet.Cells(idx, 2).value = che.numero
            xlWorksheet.Cells(idx, 3).value = funciones.FormatearDecimales(che.Monto)
            xlWorksheet.Cells(idx, 4).value = che.FechaEmision
            xlWorksheet.Cells(idx, 5).value = che.FechaVencimiento
            xlWorksheet.Cells(idx, 6).value = che.FechaRecibido
            xlWorksheet.Cells(idx, 7).value = che.Banco.nombre
            xlWorksheet.Cells(idx, 8).value = che.OrigenDestino
            xlWorksheet.Cells(idx, 9).value = che.Recibo
            xlWorksheet.Cells(idx, 10).value = che.destino
            xlWorksheet.Cells(idx, 11).value = NullOrZeroToEmpty(che.IdOrdenPagoOrigen)
            xlWorksheet.Cells(idx, 12).value = NullOrZeroToEmpty(che.NumeroLiquidacionCaja)
            xlWorksheet.Cells(idx, 13).value = NullOrZeroToEmpty(che.NumeroPagoACuenta)
            xlWorksheet.Cells(idx, 14).value = NullOrZeroToEmpty(che.NumeroMovimiento)
            
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
        
        Dim titulo As String
            titulo = "Reporte de Cheques en Cartera"
    
        With xlWorksheet.Range("A1:H1")
            .Merge
            .value = titulo
            .Font.Bold = True
            .HorizontalAlignment = -4108 ' xlCenter
        End With

        xlWorksheet.Columns(4).HorizontalAlignment = xlLeft
        xlWorksheet.Columns(7).HorizontalAlignment = xlLeft
        
        Dim offset As Long
        offset = 3

        xlWorksheet.Cells(offset, 1).value = "Id"
        xlWorksheet.Cells(offset, 2).value = "Número"
        xlWorksheet.Cells(offset, 3).value = "Monto"
        xlWorksheet.Cells(offset, 4).value = "Vencimiento"
        xlWorksheet.Cells(offset, 5).value = "Origen"
        xlWorksheet.Cells(offset, 6).value = "Banco Nombre"
        xlWorksheet.Cells(offset, 7).value = "Clasificación"
        xlWorksheet.Cells(offset, 8).value = "Recibido"

        xlWorksheet.Range(xlWorksheet.Cells(offset, 1), xlWorksheet.Cells(offset, 8)).Font.Bold = True
        xlWorksheet.Range(xlWorksheet.Cells(offset, 1), xlWorksheet.Cells(offset, 8)).Interior.Color = &HC0C0C0

        Dim idx As Integer
        idx = 4

        Dim che As cheque

        'DEFINE EL CONTADOR DEL PROGRESSBAR Y LO INICIA EN 0
        Dim d As Long
        d = 0


        For Each che In cartera

            xlWorksheet.Cells(idx, 1).value = che.Id
            xlWorksheet.Cells(idx, 2).value = che.numero
            xlWorksheet.Cells(idx, 3).value = che.Monto
            xlWorksheet.Cells(idx, 4).value = che.FechaVencimiento
            xlWorksheet.Cells(idx, 5).value = che.OrigenDestino
            xlWorksheet.Cells(idx, 6).value = che.Banco.nombre
            xlWorksheet.Cells(idx, 7).value = che.OrigenCheque
            xlWorksheet.Cells(idx, 8).value = che.FechaRecibo

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
    
    Dim titulo As String
        titulo = "Reporte de Cheques Propios Utilizados"
    
    With xlWorksheet.Range("A1:K1")
        .Merge
        .value = titulo
        .Font.Bold = True
        .HorizontalAlignment = -4108 ' xlCenter
    End With

    xlWorksheet.Columns(4).HorizontalAlignment = xlLeft
    xlWorksheet.Columns(7).HorizontalAlignment = xlLeft
        
    Dim offset As Long
    offset = 3
        
    xlWorksheet.Cells(offset, 1).value = "ID"
    xlWorksheet.Cells(offset, 2).value = "Número"
    xlWorksheet.Cells(offset, 3).value = "Importe"
    xlWorksheet.Cells(offset, 4).value = "Fecha Emisión"
    xlWorksheet.Cells(offset, 5).value = "Fecha Vencimiento"
    xlWorksheet.Cells(offset, 6).value = "Banco"
    xlWorksheet.Cells(offset, 7).value = "Destino"
    xlWorksheet.Cells(offset, 8).value = "N OP"
    xlWorksheet.Cells(offset, 9).value = "N LIQ"
    xlWorksheet.Cells(offset, 10).value = "N PCTA"
    xlWorksheet.Cells(offset, 11).value = "N MOV"
   
    xlWorksheet.Range(xlWorksheet.Cells(offset, 1), xlWorksheet.Cells(offset, 11)).Font.Bold = True
    xlWorksheet.Range(xlWorksheet.Cells(offset, 1), xlWorksheet.Cells(offset, 11)).Interior.Color = &HC0C0C0
    
    Dim idx As Integer
    idx = 4

    Dim che As cheque

    'DEFINE EL CONTADOR DEL PROGRESSBAR Y LO INICIA EN 0
    Dim d As Long
    d = 0


    For Each che In cheques1

        xlWorksheet.Cells(idx, 1).value = che.Id
        xlWorksheet.Cells(idx, 2).value = che.numero
        xlWorksheet.Cells(idx, 3).value = che.Monto
        xlWorksheet.Cells(idx, 4).value = che.FechaEmision
        xlWorksheet.Cells(idx, 5).value = che.FechaVencimiento
        xlWorksheet.Cells(idx, 6).value = che.Banco.nombre
        xlWorksheet.Cells(idx, 7).value = che.OrigenDestino
        xlWorksheet.Cells(idx, 8).value = NullOrZeroToEmpty(che.IdOrdenPagoOrigen)
        xlWorksheet.Cells(idx, 9).value = NullOrZeroToEmpty(che.NumeroLiquidacionCaja)
        xlWorksheet.Cells(idx, 10).value = NullOrZeroToEmpty(che.NumeroPagoACuenta)
        xlWorksheet.Cells(idx, 11).value = NullOrZeroToEmpty(che.NumeroMovimiento)
        
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


Private Sub cboRangosVtoCartera_Click(Index As Integer)
    funciones.CalculateDateRange Me.cboRangosVtoCartera(0), Me.dtpDesdeVtoCartera(1), Me.dtpHastaVtoCartera(1)
End Sub

Private Sub cboRangosVtoPropios_Click(Index As Integer)
    funciones.CalculateDateRange Me.cboRangosVtoPropios(1), Me.dtpDesdeVtoPropios(3), Me.dtpHastaVtoPropios(3)
End Sub

Private Sub cboRangosVtoTerceros_Click(Index As Integer)

    Select Case Index
        Case 0
            funciones.CalculateDateRange _
                Me.cboRangosVtoTerceros(0), _
                Me.dtpDesdeVtoTerceros(0), _
                Me.dtpHastaVtoTerceros(0)

        Case 2
            funciones.CalculateDateRange _
                Me.cboRangosVtoTerceros(2), _
                Me.dtpDesdeVtoTerceros(5), _
                Me.dtpHastaVtoTerceros(5)
    End Select

End Sub
    
Private Sub cboRangosRboCartera_Click(Index As Integer)
    funciones.CalculateDateRange Me.cboRangosRboCartera(1), Me.dtpDesdeRboCartera(2), Me.dtpHastaRboCartera(2)
End Sub

Private Sub cboRangosRboPropios_Click(Index As Integer)
    funciones.CalculateDateRange Me.cboRangosRboPropios(0), Me.dtpDesdeRboPropios(4), Me.dtpHastaRboPropios(4)
End Sub
    
Private Sub cboRangosRboEmitido_Click(Index As Integer)

    Select Case Index
        Case 0
            funciones.CalculateDateRange _
                Me.cboRangosRboEmitido(0), _
                Me.dtpDesdeRboEmitido(0), _
                Me.dtpHastaRboEmitido(0)

        Case 2
            funciones.CalculateDateRange _
                Me.cboRangosRboEmitido(2), _
                Me.dtpDesdeRboEmitido(6), _
                Me.dtpHastaRboEmitido(6)
    End Select

End Sub

Private Sub cboRangosRboRecibido_Click(Index As Integer)
    funciones.CalculateDateRange Me.cboRangosRboRecibido(0), Me.dtpDesdeRboRecibido(0), Me.dtpHastaRboRecibido(0)
End Sub

Private Sub cmdCrear_Click()
    Dim x As Long
    Dim col As Collection
    Dim id_banco As Long


    If MsgBox("Está segur@ de crear la chequera?", vbQuestion + vbYesNo) = vbYes Then
        
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
        
        If Me.cboCuentaBancariaChequera.ListIndex = -1 Then
        
            MsgBox "Debe seleccionar la cuenta bancaria " & _
                   "a la que pertenece la chequera.", _
                   vbExclamation, _
                   "Chequeras"
        
            Exit Sub
        
        End If


        Set chequera.Banco = DAOBancos.GetById(id_banco)
        
        Set chequera.CuentaBancaria = _
        DAOCuentaBancaria.FindById( _
            Me.cboCuentaBancariaChequera.ItemData( _
                Me.cboCuentaBancariaChequera.ListIndex))
            
        chequera.fechaCreacion = Now
        Set chequera.moneda = DAOMoneda.GetById(Me.cboMonedas.ItemData(Me.cboMonedas.ListIndex))
        chequera.numero = CLng(Me.txtNumero)
        chequera.NumeroDesde = CLng(Me.txtDesde)
        chequera.NumeroHasta = CLng(Me.txtHasta)
        chequera.Observaciones = UCase(Me.txtObservaciones)
        Dim cheque As cheque
        For x = chequera.NumeroDesde To chequera.NumeroHasta
            Set cheque = New cheque
            cheque.numero = x
            cheque.EnCartera = False
            cheque.Propio = True
            cheque.Id = 0
            Set cheque.Banco = chequera.Banco
            Set cheque.moneda = chequera.moneda
            chequera.Cheques.Add cheque

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
    
    GridEXHelper.CustomizeGrid Me.grid_chequeras, True, True
    GridEXHelper.CustomizeGrid Me.grid_cartera_cheques, True, True
    GridEXHelper.CustomizeGrid Me.grid_cheques, True, True
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
    
    
    
    '''''''''''''''''''''''''''
    
    dtpDesdeVtoCartera(1).value = Year(Now) & "-01-01"
    desde = DateSerial(Year(Date), Month(Date), 1)   ' CDate(1 & "-" & Month(Now) & "-" & Year(Now))
    funciones.FillComboBoxDateRanges Me.cboRangosVtoCartera(0)
    Me.cboRangosVtoCartera(0).ListIndex = i
    For i = 0 To Me.cboRangosVtoCartera(0).ListCount - 1
        If Me.cboRangosVtoCartera(0).ItemData(i) = DateRangeValue.DRV_YearCurrent Then Exit For
    Next i
    Me.cboRangosVtoCartera(0).ListIndex = i


    dtpDesdeRboCartera(2).value = Year(Now) & "-01-01"
    desde = DateSerial(Year(Date), Month(Date), 1)   ' CDate(1 & "-" & Month(Now) & "-" & Year(Now))
    funciones.FillComboBoxDateRanges Me.cboRangosRboCartera(1)
    Me.cboRangosRboCartera(1).ListIndex = i
    For i = 0 To Me.cboRangosRboCartera(1).ListCount - 1
        If Me.cboRangosRboCartera(1).ItemData(i) = DateRangeValue.DRV_YearCurrent Then Exit For
    Next i
    Me.cboRangosRboCartera(1).ListIndex = i
   
    Me.lbContadorChequesEnCartera.caption = "Resultados: [ " & 0 & " ]"

    ''''''''''''''''''''''''''''
    
    'SOLAPA ADMINISTRAR CHEQUERAS
    DAOMoneda.llenarComboXtremeSuite Me.cboMonedas
    
    Set bancos = DAOBancos.GetAll("id in (select idBanco from AdminConfigCuentas group by idBanco) ")

    cboBancos1.Clear
    For Each Banco In bancos
        cboBancos1.AddItem Banco.nombre
        cboBancos1.ItemData(cboBancos1.NewIndex) = Banco.Id
    Next

    Set bancos = DAOBancos.GetAll()
    
    
    Me.grid_cheques.ItemCount = 0
    
        
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'SOLAPA CHEQUES PROPIOS UTILIZADOS
    DAOChequeras.llenarComboXtremeSuite Me.cboChequera2
    Me.cboChequera2.ListIndex = -1
    
    
    dtpDesdeVtoPropios(3).value = Year(Now) & "-01-01"
    desde = DateSerial(Year(Date), Month(Date), 1)   ' CDate(1 & "-" & Month(Now) & "-" & Year(Now))
    
    funciones.FillComboBoxDateRanges Me.cboRangosVtoPropios(1)
    Me.cboRangosVtoPropios(1) = i
    For i = 0 To Me.cboRangosVtoPropios(1).ListCount - 1
        If Me.cboRangosVtoPropios(1).ItemData(i) = DateRangeValue.DRV_YearCurrent Then Exit For
    Next i
    Me.cboRangosVtoPropios(1).ListIndex = i



    dtpDesdeRboPropios(4).value = Year(Now) & "-01-01"
    desde = DateSerial(Year(Date), Month(Date), 1)   ' CDate(1 & "-" & Month(Now) & "-" & Year(Now))
    
    funciones.FillComboBoxDateRanges Me.cboRangosRboPropios(0)
    Me.cboRangosRboPropios(0) = i
    For i = 0 To Me.cboRangosRboPropios(0).ListCount - 1
        If Me.cboRangosRboPropios(0).ItemData(i) = DateRangeValue.DRV_YearCurrent Then Exit For
    Next i
    Me.cboRangosRboPropios(0).ListIndex = i

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    funciones.FillComboBoxDateRanges Me.cboRangosRboPropios(0)
    Me.cboRangosRboPropios(0) = i
    For i = 0 To Me.cboRangosRboPropios(0).ListCount - 1
        If Me.cboRangosRboPropios(0).ItemData(i) = DateRangeValue.DRV_YearCurrent Then Exit For
    Next i
    Me.cboRangosRboPropios(0).ListIndex = -1
    
'''    DAOProveedor.llenarComboXtremeSuite Me.cboProveedoresPropios, True, True, True

    Call DAOProveedor.llenarComboProveedores(cboProveedoresPropios)
    
    Me.cboProveedoresPropios.ListIndex = -1
    
    Me.lbContadorChequesPropiosUtilizados.caption = "Resultados: [ " & cheques1.count & " ]"
    
    'SOLAPA CHEQUES 3EROS UTILIZADOS
    DAOBancos.llenarComboXtremeSuite Me.cboBancos3ero
    Me.cboBancos3ero.ListIndex = -1
    
    dtpDesdeVtoTerceros(5).value = Year(Now) & "-01-01"
    desde = DateSerial(Year(Date), Month(Date), 1)   ' CDate(1 & "-" & Month(Now) & "-" & Year(Now))
    
    
    funciones.FillComboBoxDateRanges Me.cboRangosVtoTerceros(2)
    Me.cboRangosVtoTerceros(2) = i
    For i = 0 To Me.cboRangosVtoTerceros(2).ListCount - 1
        If Me.cboRangosVtoTerceros(2).ItemData(i) = DateRangeValue.DRV_YearCurrent Then Exit For
    Next i
    Me.cboRangosVtoTerceros(2).ListIndex = -1



    dtpDesdeRboEmitido(6).value = Year(Now) & "-01-01"
    desde = DateSerial(Year(Date), Month(Date), 1)   ' CDate(1 & "-" & Month(Now) & "-" & Year(Now))
    

    funciones.FillComboBoxDateRanges Me.cboRangosRboEmitido(2)
    Me.cboRangosRboEmitido(2) = i
    For i = 0 To Me.cboRangosRboEmitido(2).ListCount - 1
        If Me.cboRangosRboEmitido(2).ItemData(i) = DateRangeValue.DRV_YearCurrent Then Exit For
    Next i
    Me.cboRangosRboEmitido(2).ListIndex = -1
    
    
    dtpDesdeRboRecibido(0).value = Year(Now) & "-01-01"
    desde = DateSerial(Year(Date), Month(Date), 1)   ' CDate(1 & "-" & Month(Now) & "-" & Year(Now))
    

    funciones.FillComboBoxDateRanges Me.cboRangosRboRecibido(0)
    Me.cboRangosRboRecibido(0) = i
    For i = 0 To Me.cboRangosRboRecibido(0).ListCount - 1
        If Me.cboRangosRboRecibido(0).ItemData(i) = DateRangeValue.DRV_YearCurrent Then Exit For
    Next i
    Me.cboRangosRboRecibido(0).ListIndex = -1
    
    
    Me.gridBancos.ItemCount = 0
    Me.gridChequesEmitidos.ItemCount = 0
    Me.grdCheques3eros.ItemCount = 0
    
'''    DAOProveedor.llenarComboXtremeSuite Me.cboProveedores3eros, True, True, True
    
    Call DAOProveedor.llenarComboProveedores(cboProveedores3eros)
    
    Me.cboProveedores3eros.ListIndex = -1
    
    DAOCliente.llenarComboXtremeSuite Me.cboClientes3erosUti, True, True, True
    Me.cboClientes3erosUti.ListIndex = -1
    
    Me.lbContador3erosUtilizados.caption = "Resultados: [ " & cheques1.count & " ]"
    
    '''''''''''''''''''''''''''''''''''''''''''''''
    'region FECHAS EN CARTERA
    
   
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
    
'''''''''''''''''''''''''''''''''''
   
    funciones.FillComboBoxDateRanges Me.cboRangosVtoTerceros(0)
    Me.cboRangosVtoTerceros(0) = i
    For i = 0 To Me.cboRangosVtoTerceros(0).ListCount - 1
        If Me.cboRangosVtoTerceros(0).ItemData(i) = DateRangeValue.DRV_YearCurrent Then Exit For
    Next i
    Me.cboRangosVtoTerceros(0).ListIndex = -1


    funciones.FillComboBoxDateRanges Me.cboRangosRboEmitido(0)
    Me.cboRangosRboEmitido(0) = i
    For i = 0 To Me.cboRangosRboEmitido(0).ListCount - 1
        If Me.cboRangosRboEmitido(0).ItemData(i) = DateRangeValue.DRV_YearCurrent Then Exit For
    Next i
    Me.cboRangosRboEmitido(0).ListIndex = -1
    
'''''''''''''''''''''''''''''''''''

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

    If Not chequeras Is Nothing Then
        If chequeras.count > 0 Then
            Set tmpChequera = chequeras.item(1)
            idChequeraMostrada = tmpChequera.Id
    
            MostrarChequera
        End If
    End If

End Sub


Private Sub Form_Resize()
    On Error Resume Next
    
    Dim margen As Long
    Dim espacioInferior As Long
    Dim masmargen As Long
    
    margen = 120
    espacioInferior = 120
    masmargen = 2000
    
    With Me.TabControl1
        .Left = 0
        .Top = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight
    End With
    
    With Me.gridChequesEmitidos
'''       .Left = Me.TabControl1.Left + margen
       .Top = Me.TabControl1.ScaleHeight + Me.TabControl1.Top + margen
       .Width = Me.TabControl1.Width - margen
       .Height = Me.TabControl1.Height / 1.75
        
    End With
    
    With Me.grdCheques3eros
'''       .Left = Me.TabControl1.Left + margen
       .Top = Me.TabControl1.ScaleHeight + Me.TabControl1.Top + margen
       .Width = Me.TabControl1.Width - margen
       .Height = Me.TabControl1.Height / 1.75
        
    End With
    
    With Me.grid_cartera_cheques
       .Top = Me.TabControl1.ScaleHeight + Me.TabControl1.Top + margen
       .Width = Me.TabControl1.Width - margen
       .Height = Me.TabControl1.Height / 1.75
    End With
    
    With Me.grid_chequeras
       .Top = Me.TabControl1.ScaleHeight + Me.TabControl1.Top + margen
'''       .Width = Me.TabControl1.Width - margen
       .Height = Me.TabControl1.Height / 1.75
    End With
    
    With Me.grid_cheques
       .Top = Me.TabControl1.ScaleHeight + Me.TabControl1.Top + margen
'''       .Width = Me.TabControl1.Width - margen
       .Height = Me.TabControl1.Height / 1.75
    End With
    
End Sub


Private Sub MostrarChequera( _
    Optional ByVal AvisarChequeNoEncontrado As Boolean = False)

    Dim filter2 As String
    Dim numeroBuscado As String
    Dim resultado As Collection

    If tmpChequera Is Nothing Then Exit Sub

    filter2 = "1 = 1"
    numeroBuscado = Trim$(Me.TxtNumeroChequeEnChequera.Text)

    If LenB(numeroBuscado) > 0 Then
        filter2 = filter2 & _
                  " AND cheq.numero LIKE " & _
                  conectar.Escape("%" & numeroBuscado & "%")
    End If

    'Filtros propios de la solapa Administrar Chequeras
    If Not IsNull(Me.dtpDesdeVtoTerceros(0).value) Then
        filter2 = filter2 & _
                  " AND cheq.fecha_vencimiento >= " & _
                  conectar.Escape(Format( _
                      Me.dtpDesdeVtoTerceros(0).value, _
                      "yyyy-mm-dd"))
    End If

    If Not IsNull(Me.dtpHastaVtoTerceros(0).value) Then
        filter2 = filter2 & _
                  " AND cheq.fecha_vencimiento <= " & _
                  conectar.Escape(Format( _
                      Me.dtpHastaVtoTerceros(0).value, _
                      "yyyy-mm-dd"))
    End If

    If Not IsNull(Me.dtpDesdeRboEmitido(0).value) Then
        filter2 = filter2 & _
                  " AND cheq.fecha_emision >= " & _
                  conectar.Escape(Format( _
                      Me.dtpDesdeRboEmitido(0).value, _
                      "yyyy-mm-dd"))
    End If

    If Not IsNull(Me.dtpHastaRboEmitido(0).value) Then
        filter2 = filter2 & _
                  " AND cheq.fecha_emision <= " & _
                  conectar.Escape(Format( _
                      Me.dtpHastaRboEmitido(0).value, _
                      "yyyy-mm-dd"))
    End If

    Set resultado = DAOCheques.FindAllByChequeraId( _
                        tmpChequera.Id, filter2)

    If resultado Is Nothing Then
        Set resultado = New Collection
    End If

    Set tmpChequera.Cheques = resultado

    Me.grid_cheques.ItemCount = 0
    Me.grid_cheques.ItemCount = tmpChequera.Cheques.count
    Me.grid_cheques.Refresh

    Me.Label13.caption = _
        "Cheques mostrados: [ " & _
        tmpChequera.Cheques.count & " ]"

    If AvisarChequeNoEncontrado Then
        If LenB(numeroBuscado) > 0 And _
           tmpChequera.Cheques.count = 0 Then

            MsgBox "Cheque no encontrado.", _
                   vbInformation, _
                   "Administración de cheques"

            Me.TxtNumeroChequeEnChequera.SetFocus
            Me.TxtNumeroChequeEnChequera.SelStart = 0
            Me.TxtNumeroChequeEnChequera.SelLength = _
                Len(Me.TxtNumeroChequeEnChequera.Text)
        End If
    End If

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
        filter2 = filter2 & " AND rec.fecha >= " & conectar.Escape(Me.dtpDesdeRboCartera(2).value)
    End If

    If Not IsNull(Me.dtpHastaRboCartera(2).value) Then
        filter2 = filter2 & " AND rec.fecha <= " & conectar.Escape(Me.dtpHastaRboCartera(2).value)
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
    
    Me.lbContadorChequesEnCartera.caption = "Cheques encontrados: [ " & cartera.count & " ]"

End Sub


Private Sub MostrarChequeras()
    Set chequeras = DAOChequeras.GetAll
    Me.grid_chequeras.ItemCount = 0
    Me.grid_chequeras.ItemCount = chequeras.count
    
   Me.Label12.caption = "Chequeras mostradas: [ " & chequeras.count & " ]"
    
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
    Values(6) = tmpCheque3eros.FechaRecibido
    Values(7) = tmpCheque3eros.numero
    Values(8) = funciones.FormatearDecimales(tmpCheque3eros.Monto)
    Values(9) = tmpCheque3eros.FechaVencimiento
    Values(10) = tmpCheque3eros.Recibo
    Values(15) = tmpCheque3eros.destino
    Values(11) = tmpCheque3eros.IdOrdenPagoOrigen
    Values(12) = tmpCheque3eros.NumeroLiquidacionCaja
    Values(13) = tmpCheque3eros.NumeroPagoACuenta
    Values(14) = tmpCheque3eros.NumeroMovimiento
    
    
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
        .value(8) = tmpCheque.FechaRecibo
        .value(9) = tmpCheque.Observaciones

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
    CargarChequeraSeleccionada
End Sub


Private Sub grid_chequeras_Click()
    CargarChequeraSeleccionada
End Sub


Private Sub mostrarCheques()
    Me.grid_cheques.ItemCount = 0

    If tmpChequera Is Nothing Then Exit Sub
    If tmpChequera.Cheques Is Nothing Then Exit Sub

    Me.grid_cheques.ItemCount = tmpChequera.Cheques.count
End Sub


Private Sub grid_chequeras_UnboundReadData( _
    ByVal RowIndex As Long, _
    ByVal Bookmark As Variant, _
    ByVal Values As GridEX20.JSRowData)

    On Error GoTo err1

    If chequeras Is Nothing Then Exit Sub
    If RowIndex < 1 Or RowIndex > chequeras.count Then Exit Sub

    Dim chFila As chequera
    Set chFila = chequeras.item(RowIndex)

    With Values

        .value(1) = chFila.numero
        .value(2) = chFila.fechaCreacion

        'Banco
        If Not chFila.Banco Is Nothing Then
            .value(3) = chFila.Banco.nombre
        Else
            .value(3) = vbNullString
        End If

        'Cuenta bancaria
        If Not chFila.CuentaBancaria Is Nothing Then
            .value(4) = chFila.CuentaBancaria.numero
        Else
            .value(4) = vbNullString
        End If

        .value(5) = chFila.NumeroDesde
        .value(6) = chFila.NumeroHasta

        .value(8) = Abs(CInt(chFila.usada))

    End With

    Exit Sub

err1:
    Debug.Print "grid_chequeras_UnboundReadData: " & _
                Err.Number & " - " & Err.Description

End Sub


Private Sub grid_chequeras_UnboundUpdate(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    
    On Error GoTo err1

    If RowIndex <= 0 Then Exit Sub
    If RowIndex > chequeras.count Then Exit Sub

    Dim ch As chequera
    Dim usadaAnterior As Boolean
    Dim nuevaUsada As Boolean

    Set ch = chequeras.item(RowIndex)

    usadaAnterior = ch.usada

    If IsNull(Values(8)) Or IsEmpty(Values(8)) Then
        nuevaUsada = False
    Else
        nuevaUsada = CBool(Values(8))
    End If

    If Not DAOChequeras.ActualizarUsada(ch.Id, nuevaUsada) Then
        GoTo err1
    End If

    'Actualizar también el objeto de la colección
    ch.usada = nuevaUsada

    Exit Sub

err1:
    ch.usada = usadaAnterior

    MsgBox "No se pudo actualizar el estado de la chequera." & vbCrLf & _
           Err.Description, _
           vbCritical, _
           "Error"

    MostrarChequeras


End Sub

Private Sub grid_cheques_DblClick()
    If Me.grid_cheques.RowIndex(Me.grid_cheques.row) > 0 Then
        Set tmpCheque = tmpChequera.Cheques(Me.grid_cheques.RowIndex(Me.grid_cheques.row))
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
        Set tmpCheque = tmpChequera.Cheques(Me.grid_cheques.RowIndex(Me.grid_cheques.row))
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

Private Sub grid_cheques_UnboundReadData( _
    ByVal RowIndex As Long, _
    ByVal Bookmark As Variant, _
    ByVal Values As GridEX20.JSRowData)

    On Error GoTo err1

    If tmpChequera Is Nothing Then Exit Sub
    If tmpChequera.Cheques Is Nothing Then Exit Sub

    If RowIndex < 1 Then Exit Sub
    If RowIndex > tmpChequera.Cheques.count Then Exit Sub

    Dim chFila As cheque
    Set chFila = tmpChequera.Cheques.item(RowIndex)

    With Values
        .value(1) = chFila.numero

        If chFila.Utilizado Then
            .value(2) = funciones.FormatearDecimales(chFila.Monto)
            .value(3) = chFila.FechaVencimiento
            .value(4) = chFila.FechaEmision
            .value(5) = chFila.OrigenDestino
        Else
            .value(2) = Empty
            .value(3) = Empty
            .value(4) = Empty
            .value(5) = Empty
        End If

            .value(6) = DescripcionUsoCheque(chFila)
            
            .value(7) = Abs(CInt(chFila.entro))

            If CDbl(chFila.FechaIngresoBanco) > 0 Then
                .value(8) = chFila.FechaIngresoBanco
            Else
                .value(8) = Empty
            End If
            
    End With

    Exit Sub

err1:
    Debug.Print "grid_cheques_UnboundReadData: " & _
                Err.Number & " - " & Err.Description
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
    
    If tmpCheque.NumeroLiquidacionCaja <> 0 Then
        Values(8) = "VARIOS PROVEEDORES"
        Values(10) = ""
        Values(11) = tmpCheque.NumeroLiquidacionCaja
        Values(12) = ""
        Values(13) = ""
    ElseIf tmpCheque.NumeroPagoACuenta <> 0 Then
        Values(8) = "PAGO A CUENTA"
        Values(10) = ""
        Values(11) = ""
        Values(12) = tmpCheque.NumeroPagoACuenta
        Values(13) = ""
    ElseIf tmpCheque.IdOrdenPagoOrigen <> 0 Then
        Values(8) = tmpCheque.OrigenDestino
        Values(10) = tmpCheque.IdOrdenPagoOrigen
        Values(11) = ""
        Values(12) = ""
        Values(13) = ""
    ElseIf tmpCheque.NumeroMovimiento <> 0 Then
        Values(8) = "MOVIMIENTO"
        Values(10) = ""
        Values(11) = ""
        Values(12) = ""
        Values(13) = tmpCheque.NumeroMovimiento
    Else
        Values(8) = ""
        Values(10) = ""
        Values(11) = ""
        Values(12) = ""
        Values(13) = ""
    End If

End Sub


Private Sub mnuPasarCartera_Click()
    grid_cheques_DblClick
End Sub


Private Sub btnBorrarNumeroTerceros_Click()
    txtNumeroCheque3ero = ""
End Sub


Private Sub PushButton1_Click()
    Me.TxtNumeroChequeEnChequera.Text = vbNullString

    MostrarChequera

    Me.TxtNumeroChequeEnChequera.SetFocus

End Sub


'''Private Sub PushButton1_Click()
'''    Me.cboProveedores.ListIndex = -1
'''End Sub

Private Sub PushButton2_Click()
    Me.cboProveedoresPropios.ListIndex = -1
End Sub

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


Private Sub AjustarGrid(ByVal grd As Object, ByVal ctrlSuperior As Object, _
                        ByVal margen As Long, ByVal espacioInferior As Long)
    Dim topGrid As Long
    Dim altoGrid As Long
    
    topGrid = ctrlSuperior.Top + ctrlSuperior.Height + margen
    
    grd.Left = margen
    grd.Top = topGrid
    grd.Width = Me.ScaleWidth - (margen * 2)
    
    altoGrid = Me.ScaleHeight - topGrid - espacioInferior
    If altoGrid > 300 Then
        grd.Height = altoGrid
    End If
End Sub


Private Sub CargarChequeraSeleccionada( _
    Optional ByVal Forzar As Boolean = False)

    On Error GoTo err1

    If cargandoChequera Then Exit Sub
    If chequeras Is Nothing Then Exit Sub
    If chequeras.count = 0 Then Exit Sub

    Dim fila As Long

    fila = Me.grid_chequeras.RowIndex( _
                Me.grid_chequeras.row)

    If fila < 1 Or fila > chequeras.count Then Exit Sub

    Dim chSeleccionada As chequera
    Set chSeleccionada = chequeras.item(fila)

    'Evita ejecutar dos consultas por el mismo clic
    If Not Forzar Then
        If idChequeraMostrada = chSeleccionada.Id Then
            Exit Sub
        End If
    End If

    cargandoChequera = True
    
    'La búsqueda anterior no debe aplicarse a la nueva chequera
    Me.TxtNumeroChequeEnChequera.Text = vbNullString
    
    Set tmpChequera = chSeleccionada

    MostrarChequera

    idChequeraMostrada = tmpChequera.Id

salir:
    cargandoChequera = False
    Exit Sub

err1:
    Me.grid_cheques.ItemCount = 0
    Me.grid_cheques.Refresh

    MsgBox "No se pudieron cargar los cheques de la chequera." & _
           vbCrLf & Err.Description, _
           vbExclamation, _
           "Administración de cheques"

    Resume salir
End Sub


Private Function DescripcionUsoCheque(ByVal ch As cheque) As String

    Dim numeroLiquidacion As Long

    If ch Is Nothing Then
        DescripcionUsoCheque = vbNullString
        Exit Function
    End If

    If ch.estado = ChequeAnulado Then
        DescripcionUsoCheque = "ANULADO"
        Exit Function
    End If

    'Primero se verifica Movimiento porque actualmente
    'puede haber registros históricos con ambos campos cargados.
    If ch.NumeroMovimiento > 0 Then

        DescripcionUsoCheque = _
            "Utilizado en Movimiento de Caja y Bancos N° " & _
            ch.NumeroMovimiento

    ElseIf ch.NumeroPagoACuenta > 0 Then

        DescripcionUsoCheque = _
            "Utilizado en Pago a Cuenta N° " & _
            ch.NumeroPagoACuenta

    ElseIf ch.NumeroLiquidacionCaja > 0 Or _
           ch.IdLiquidacionCajaOrigen > 0 Then

        numeroLiquidacion = ch.NumeroLiquidacionCaja

        If numeroLiquidacion = 0 Then
            numeroLiquidacion = ch.IdLiquidacionCajaOrigen
        End If

        DescripcionUsoCheque = _
            "Utilizado en Liquidación de Caja N° " & _
            numeroLiquidacion

    ElseIf ch.IdOrdenPagoOrigen > 0 Then

        DescripcionUsoCheque = _
            "Utilizado en Orden de Pago N° " & _
            ch.IdOrdenPagoOrigen

    ElseIf ch.Utilizado Then

        DescripcionUsoCheque = _
            "UTILIZADO - ORIGEN NO IDENTIFICADO"

    Else

        DescripcionUsoCheque = "DISPONIBLE"

    End If

End Function


Private Sub grid_cheques_UnboundUpdate( _
    ByVal RowIndex As Long, _
    ByVal Bookmark As Variant, _
    ByVal Values As GridEX20.JSRowData)

    On Error GoTo err1

    If tmpChequera Is Nothing Then Exit Sub
    If tmpChequera.Cheques Is Nothing Then Exit Sub

    If RowIndex < 1 Or _
       RowIndex > tmpChequera.Cheques.count Then
        Exit Sub
    End If

    Dim ch As cheque
    Dim nuevoIngresado As Boolean
    Dim nuevaFecha As Date
    Dim respuesta As VbMsgBoxResult
    Dim mensaje As String
    Dim estadoAnterior As Boolean

    Set ch = tmpChequera.Cheques.item(RowIndex)

    estadoAnterior = ch.entro

    If IsNull(Values(7)) Or IsEmpty(Values(7)) Then
        nuevoIngresado = False
    Else
        nuevoIngresado = CBool(Values(7))
    End If

    'Definir fecha
    If nuevoIngresado Then

        If IsDate(Values(8)) Then
            nuevaFecha = CDate(Values(8))
        Else
            nuevaFecha = Date
        End If

    Else
        nuevaFecha = 0
    End If

    'Validar solamente cuando cambia el tilde
    If nuevoIngresado <> estadoAnterior Then

        If nuevoIngresado Then

            'Opcional: impedir marcar cheques disponibles
            If Not ch.Utilizado Then
                MsgBox "El cheque N° " & ch.numero & _
                       " todavía no fue utilizado." & vbCrLf & _
                       "No se puede registrar su ingreso al banco.", _
                       vbExclamation, _
                       "Administración de cheques"

                MostrarChequera
                Exit Sub
            End If

            mensaje = _
                "¿Confirma que el cheque N° " & ch.numero & _
                " ingresó al banco?" & vbCrLf & vbCrLf & _
                "Fecha de ingreso: " & _
                Format$(nuevaFecha, "dd/mm/yyyy")

        Else

            mensaje = _
                "¿Confirma que desea quitar la marca de ingresado " & _
                "del cheque N° " & ch.numero & "?" & _
                vbCrLf & vbCrLf & _
                "También se eliminará la fecha de ingreso."

        End If

        respuesta = MsgBox( _
                        mensaje, _
                        vbQuestion + vbYesNo + vbDefaultButton1, _
                        "Confirmar modificación")

        If respuesta <> vbYes Then
            MostrarChequera
            Exit Sub
        End If

    End If

    'Validar fecha
    If nuevoIngresado Then
        If nuevaFecha > Date Then
            MsgBox "La fecha de ingreso no puede ser posterior a hoy.", _
                   vbExclamation, _
                   "Fecha incorrecta"

            MostrarChequera
            Exit Sub
        End If
    End If

    'Guardar
    If Not DAOCheques.ActualizarIngresoBanco( _
                ch.Id, _
                nuevoIngresado, _
                nuevaFecha) Then

        Err.Raise vbObjectError + 1001, _
                  "grid_cheques_UnboundUpdate", _
                  "No se pudo actualizar el cheque."
    End If

    'Actualizar el objeto
    ch.entro = nuevoIngresado
    ch.FechaIngresoBanco = nuevaFecha

    Exit Sub

err1:
    MsgBox "No se pudo guardar el ingreso del cheque." & _
           vbCrLf & Err.Description, _
           vbExclamation, _
           "Administración de cheques"

    MostrarChequera

End Sub

Private Sub cboBancos_Click()

    Dim cuentas As Collection
    Dim cuenta As CuentaBancaria
    Dim IdBanco As Long

    Me.cboCuentaBancariaChequera.Clear

    If Me.cboBancos.ListIndex = -1 Then Exit Sub

    IdBanco = Me.cboBancos.ItemData( _
                    Me.cboBancos.ListIndex)

    Set cuentas = DAOCuentaBancaria.FindAll( _
                    "c.idBanco = " & IdBanco)

    For Each cuenta In cuentas

        Me.cboCuentaBancariaChequera.AddItem _
            cuenta.DescripcionFormateada

        Me.cboCuentaBancariaChequera.ItemData( _
            Me.cboCuentaBancariaChequera.NewIndex) = cuenta.Id

    Next cuenta

    If Me.cboCuentaBancariaChequera.ListCount = 1 Then
        Me.cboCuentaBancariaChequera.ListIndex = 0
    End If

End Sub

