VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmAdminFacturasEmitidas 
   BackColor       =   &H00FF8080&
   Caption         =   "Comprobantes Emitidos"
   ClientHeight    =   7470
   ClientLeft      =   60
   ClientTop       =   4725
   ClientWidth     =   9420
   Icon            =   "frmFacturasEmitidas.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7470
   ScaleWidth      =   9420
   WindowState     =   2  'Maximized
   Begin VB.Frame grpTotalizadores 
      Caption         =   "Totales"
      Height          =   1335
      Index           =   0
      Left            =   14760
      TabIndex        =   47
      Top             =   120
      Width           =   7095
      Begin XtremeSuiteControls.Label lblPercepciones_Dolar 
         Height          =   195
         Left            =   3360
         TabIndex        =   55
         Top             =   720
         Width           =   2235
         _Version        =   786432
         _ExtentX        =   3942
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Total Percepciones: U$S 00,00"
         BackColor       =   -2147483633
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblIVA_Dolar 
         Height          =   195
         Left            =   3360
         TabIndex        =   54
         Top             =   480
         Width           =   1515
         _Version        =   786432
         _ExtentX        =   2672
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Total IVA: U$S 00,00"
         BackColor       =   -2147483633
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblNG_Dolar 
         Height          =   195
         Left            =   3360
         TabIndex        =   53
         Top             =   240
         Width           =   1500
         _Version        =   786432
         _ExtentX        =   2646
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Total NG: U$S 00,00"
         BackColor       =   -2147483633
         AutoSize        =   -1  'True
      End
      Begin VB.Label lblTotalDolares 
         AutoSize        =   -1  'True
         Caption         =   "Total: U$S 00,00 "
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
         Left            =   3360
         TabIndex        =   52
         Top             =   960
         Width           =   1530
      End
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         Caption         =   "Total: $ 00,00"
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
         TabIndex        =   51
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lblTotalPercepciones 
         AutoSize        =   -1  'True
         Caption         =   "Total Percepciones: $ 00,00"
         Height          =   195
         Left            =   240
         TabIndex        =   50
         Top             =   720
         Width           =   2010
      End
      Begin VB.Label lblTotalIVA 
         AutoSize        =   -1  'True
         Caption         =   "Total IVA: $ 00,00"
         Height          =   195
         Left            =   240
         TabIndex        =   49
         Top             =   480
         Width           =   1290
      End
      Begin VB.Label lblTotalNeto 
         AutoSize        =   -1  'True
         Caption         =   "Total NG: $ 00,00"
         Height          =   195
         Left            =   240
         TabIndex        =   48
         Top             =   240
         Width           =   1275
      End
   End
   Begin VB.Frame grpBotones 
      Height          =   855
      Index           =   1
      Left            =   14760
      TabIndex        =   41
      Top             =   1680
      Width           =   7095
      Begin XtremeSuiteControls.PushButton btnBuscar 
         Default         =   -1  'True
         Height          =   420
         Left            =   120
         TabIndex        =   42
         Top             =   240
         Width           =   1245
         _Version        =   786432
         _ExtentX        =   2196
         _ExtentY        =   741
         _StockProps     =   79
         Caption         =   "Buscar"
         BackColor       =   12632256
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
      Begin XtremeSuiteControls.PushButton btnImprimir 
         Height          =   420
         Left            =   2880
         TabIndex        =   43
         Top             =   240
         Width           =   1245
         _Version        =   786432
         _ExtentX        =   2196
         _ExtentY        =   741
         _StockProps     =   79
         Caption         =   "Imprimir"
         BackColor       =   12632256
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnExportar 
         Height          =   420
         Left            =   1440
         TabIndex        =   44
         ToolTipText     =   "Exporta s?lo pendientes"
         Top             =   240
         Width           =   1245
         _Version        =   786432
         _ExtentX        =   2196
         _ExtentY        =   741
         _StockProps     =   79
         Caption         =   "Exportar"
         BackColor       =   12632256
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ProgressBar barProgreso 
         Height          =   420
         Left            =   5040
         TabIndex        =   45
         Top             =   240
         Visible         =   0   'False
         Width           =   1935
         _Version        =   786432
         _ExtentX        =   3413
         _ExtentY        =   741
         _StockProps     =   93
         BackColor       =   12632256
         Appearance      =   6
         BarColor        =   65280
      End
      Begin XtremeSuiteControls.PushButton btnReducirVentana 
         Height          =   420
         HelpContextID   =   1
         Left            =   4320
         TabIndex        =   46
         ToolTipText     =   "Amplia la grilla de comprobantes o Reestablece su tamaño."
         Top             =   240
         Width           =   495
         _Version        =   786432
         _ExtentX        =   873
         _ExtentY        =   741
         _StockProps     =   79
         Caption         =   "R"
         UseVisualStyle  =   -1  'True
         BorderGap       =   10
      End
   End
   Begin XtremeSuiteControls.GroupBox grpFiltrosPrincipal 
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   14535
      _Version        =   786432
      _ExtentX        =   25638
      _ExtentY        =   4471
      _StockProps     =   79
      Caption         =   "Filtros"
      BackColor       =   12632256
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.PushButton btnClearMoneda 
         Height          =   285
         Left            =   8090
         TabIndex        =   59
         Top             =   360
         Width           =   375
         _Version        =   786432
         _ExtentX        =   661
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "X"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cboMoneda 
         Height          =   360
         Left            =   6720
         TabIndex        =   57
         Top             =   300
         Width           =   1275
         _Version        =   786432
         _ExtentX        =   2249
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
         Appearance      =   6
      End
      Begin XtremeSuiteControls.CheckBox chkAgruparCbtes 
         Height          =   255
         Left            =   12360
         TabIndex        =   40
         Top             =   2040
         Width           =   2055
         _Version        =   786432
         _ExtentX        =   3625
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Agrupar cbtes asociados"
         BackColor       =   12632256
         Appearance      =   6
      End
      Begin XtremeSuiteControls.PushButton btnClearTipo 
         Height          =   285
         Left            =   3000
         TabIndex        =   39
         Top             =   255
         Width           =   375
         _Version        =   786432
         _ExtentX        =   661
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "X"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cboTipo 
         Height          =   315
         Left            =   1650
         TabIndex        =   37
         Top             =   240
         Width           =   1290
         _Version        =   786432
         _ExtentX        =   2275
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Appearance      =   6
         Text            =   "cboTipo"
         DropDownItemCount=   4
      End
      Begin VB.TextBox txtID 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1620
         TabIndex        =   35
         Top             =   1920
         Width           =   1335
      End
      Begin XtremeSuiteControls.PushButton PushButton2 
         Height          =   285
         Left            =   9165
         TabIndex        =   33
         Top             =   1965
         Width           =   375
         _Version        =   786432
         _ExtentX        =   661
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "X"
         BackColor       =   12632256
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cboOrdenImporte 
         Height          =   315
         Left            =   6720
         TabIndex        =   31
         Top             =   1920
         Width           =   2355
         _Version        =   786432
         _ExtentX        =   4154
         _ExtentY        =   556
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
         Appearance      =   6
      End
      Begin XtremeSuiteControls.CheckBox chkboxVerIds 
         Height          =   255
         Left            =   11160
         TabIndex        =   30
         Top             =   2040
         Width           =   1215
         _Version        =   786432
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Mostrar Id's"
         BackColor       =   12632256
         Appearance      =   6
      End
      Begin XtremeSuiteControls.CheckBox chkCredito 
         Height          =   255
         Left            =   3600
         TabIndex        =   23
         Top             =   270
         Width           =   1695
         _Version        =   786432
         _ExtentX        =   2990
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "de Crédito (MiPyme)"
         BackColor       =   12632256
         Appearance      =   6
      End
      Begin VB.TextBox txtReferencia 
         Height          =   315
         Left            =   1620
         TabIndex        =   15
         Top             =   1530
         Width           =   3490
      End
      Begin XtremeSuiteControls.ComboBox cboClientes 
         Height          =   315
         Left            =   1620
         TabIndex        =   14
         Top             =   1125
         Width           =   3510
         _Version        =   786432
         _ExtentX        =   6191
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Appearance      =   6
      End
      Begin VB.TextBox txtNroFactura 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1620
         TabIndex        =   1
         Top             =   680
         Width           =   1290
      End
      Begin VB.TextBox txtRemitoAplicado 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3960
         TabIndex        =   2
         Top             =   1920
         Width           =   1170
      End
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   285
         Left            =   5190
         TabIndex        =   3
         Top             =   1140
         Width           =   375
         _Version        =   786432
         _ExtentX        =   661
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "X"
         BackColor       =   12632256
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.GroupBox grpFechaEmision 
         Height          =   1650
         Left            =   9720
         TabIndex        =   4
         Top             =   240
         Width           =   4695
         _Version        =   786432
         _ExtentX        =   8281
         _ExtentY        =   2910
         _StockProps     =   79
         Caption         =   "Fecha Emision"
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.DateTimePicker dtpDesde 
            Height          =   315
            Left            =   840
            TabIndex        =   5
            Top             =   735
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
            Left            =   3015
            TabIndex        =   6
            Top             =   750
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
            Left            =   840
            TabIndex        =   7
            Top             =   360
            Width           =   3645
            _Version        =   786432
            _ExtentX        =   6429
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Style           =   2
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.Label Label6 
            Height          =   195
            Left            =   2520
            TabIndex        =   10
            Top             =   795
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
            Left            =   240
            TabIndex        =   9
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
         Begin XtremeSuiteControls.Label Label7 
            Height          =   195
            Left            =   240
            TabIndex        =   8
            Top             =   405
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
      Begin XtremeSuiteControls.ComboBox cboPuntosVenta 
         Height          =   360
         Left            =   3585
         TabIndex        =   17
         Top             =   660
         Width           =   1530
         _Version        =   786432
         _ExtentX        =   2699
         _ExtentY        =   635
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   2
         Appearance      =   6
         Text            =   "cboMoneda"
         DropDownItemCount=   3
      End
      Begin XtremeSuiteControls.PushButton PushButton3 
         Height          =   285
         Left            =   5190
         TabIndex        =   18
         Top             =   705
         Width           =   375
         _Version        =   786432
         _ExtentX        =   661
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "X"
         BackColor       =   12632256
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cboEstados 
         Height          =   360
         Left            =   6720
         TabIndex        =   20
         Top             =   705
         Width           =   2355
         _Version        =   786432
         _ExtentX        =   4154
         _ExtentY        =   635
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   2
         Appearance      =   6
         Text            =   "cboMoneda"
         DropDownItemCount=   3
      End
      Begin XtremeSuiteControls.PushButton PushButton4 
         Height          =   285
         Left            =   9165
         TabIndex        =   22
         Top             =   740
         Width           =   375
         _Version        =   786432
         _ExtentX        =   661
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "X"
         BackColor       =   12632256
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cboEstadosSaldada 
         Height          =   360
         Left            =   6720
         TabIndex        =   24
         Top             =   1110
         Width           =   2355
         _Version        =   786432
         _ExtentX        =   4154
         _ExtentY        =   635
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   2
         Appearance      =   6
         Text            =   "cboMoneda"
         DropDownItemCount=   3
      End
      Begin XtremeSuiteControls.PushButton PushButton5 
         Height          =   285
         Left            =   9165
         TabIndex        =   25
         Top             =   1140
         Width           =   375
         _Version        =   786432
         _ExtentX        =   661
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "X"
         BackColor       =   12632256
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cboEstadoAfip 
         Height          =   360
         Left            =   6720
         TabIndex        =   27
         Top             =   1515
         Width           =   2355
         _Version        =   786432
         _ExtentX        =   4154
         _ExtentY        =   635
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   2
         Appearance      =   6
         Text            =   "cboMoneda"
         DropDownItemCount=   3
      End
      Begin XtremeSuiteControls.PushButton cmdLimpiarCboEstadoAfip 
         Height          =   285
         Left            =   9165
         TabIndex        =   28
         Top             =   1560
         Width           =   375
         _Version        =   786432
         _ExtentX        =   661
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "X"
         BackColor       =   12632256
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkObservaciones 
         Height          =   225
         Index           =   1
         Left            =   9720
         TabIndex        =   56
         Top             =   2040
         Width           =   1695
         _Version        =   786432
         _ExtentX        =   2990
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "Observaciones"
         BackColor       =   12632256
         Appearance      =   6
         Value           =   1
      End
      Begin XtremeSuiteControls.Label Label 
         Height          =   495
         Left            =   5800
         TabIndex        =   58
         Top             =   240
         Width           =   855
         _Version        =   786432
         _ExtentX        =   1508
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Moneda"
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label Label11 
         Height          =   255
         Index           =   2
         Left            =   570
         TabIndex        =   38
         Top             =   270
         Width           =   960
         _Version        =   786432
         _ExtentX        =   1693
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Tipo Cbte"
         BackColor       =   12632256
         Alignment       =   1
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label11 
         Height          =   375
         Index           =   1
         Left            =   915
         TabIndex        =   36
         Top             =   1890
         Width           =   615
         _Version        =   786432
         _ExtentX        =   1085
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "ID CBte"
         BackColor       =   12632256
         Alignment       =   1
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label11 
         Height          =   375
         Index           =   0
         Left            =   5415
         TabIndex        =   32
         Top             =   1920
         Width           =   1215
         _Version        =   786432
         _ExtentX        =   2143
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Orden Importe"
         BackColor       =   12632256
         Alignment       =   1
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label14 
         Height          =   285
         Left            =   5475
         TabIndex        =   29
         Top             =   1560
         Width           =   1155
         _Version        =   786432
         _ExtentX        =   2037
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "Estado AFIP"
         BackColor       =   12632256
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label Label12 
         Height          =   285
         Left            =   5955
         TabIndex        =   26
         Top             =   1140
         Width           =   675
         _Version        =   786432
         _ExtentX        =   1191
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "Saldada"
         BackColor       =   12632256
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label Label10 
         Height          =   285
         Left            =   6075
         TabIndex        =   21
         Top             =   720
         Width           =   555
         _Version        =   786432
         _ExtentX        =   979
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "Estado"
         BackColor       =   12632256
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label Label9 
         Height          =   285
         Left            =   3300
         TabIndex        =   19
         Top             =   720
         Width           =   375
         _Version        =   786432
         _ExtentX        =   661
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "PV"
         BackColor       =   12632256
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OC / Referencia"
         Height          =   195
         Left            =   360
         TabIndex        =   16
         Top             =   1560
         Width           =   1170
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nro Cbte"
         Height          =   270
         Left            =   30
         TabIndex        =   13
         Top             =   735
         Width           =   1500
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente"
         Height          =   225
         Left            =   270
         TabIndex        =   12
         Top             =   1170
         Width           =   1260
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Rto"
         Height          =   255
         Left            =   3240
         TabIndex        =   11
         Top             =   1980
         Width           =   660
      End
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   240
      Top             =   8400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin GridEX20.GridEX gridComprobantesEmitidos 
      Height          =   6495
      Left            =   120
      TabIndex        =   34
      Top             =   2640
      Width           =   23805
      _ExtentX        =   41989
      _ExtentY        =   11456
      Version         =   "2.0"
      PreviewRowIndent=   100
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      PreviewColumn   =   "preview"
      PreviewRowLines =   1
      RowHeight       =   26
      ColumnAutoResize=   -1  'True
      MethodHoldFields=   -1  'True
      ContScroll      =   -1  'True
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      ImageCount      =   1
      ImagePicture1   =   "frmFacturasEmitidas.frx":000C
      RowHeaders      =   -1  'True
      DataMode        =   99
      CardSpacing     =   16
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   26
      Column(1)       =   "frmFacturasEmitidas.frx":0326
      Column(2)       =   "frmFacturasEmitidas.frx":04C6
      Column(3)       =   "frmFacturasEmitidas.frx":05DA
      Column(4)       =   "frmFacturasEmitidas.frx":0716
      Column(5)       =   "frmFacturasEmitidas.frx":0856
      Column(6)       =   "frmFacturasEmitidas.frx":09AA
      Column(7)       =   "frmFacturasEmitidas.frx":0B0E
      Column(8)       =   "frmFacturasEmitidas.frx":0D4E
      Column(9)       =   "frmFacturasEmitidas.frx":0E96
      Column(10)      =   "frmFacturasEmitidas.frx":1016
      Column(11)      =   "frmFacturasEmitidas.frx":1112
      Column(12)      =   "frmFacturasEmitidas.frx":1212
      Column(13)      =   "frmFacturasEmitidas.frx":1372
      Column(14)      =   "frmFacturasEmitidas.frx":14C6
      Column(15)      =   "frmFacturasEmitidas.frx":160E
      Column(16)      =   "frmFacturasEmitidas.frx":1766
      Column(17)      =   "frmFacturasEmitidas.frx":18AE
      Column(18)      =   "frmFacturasEmitidas.frx":19F6
      Column(19)      =   "frmFacturasEmitidas.frx":1ADA
      Column(20)      =   "frmFacturasEmitidas.frx":1C2A
      Column(21)      =   "frmFacturasEmitidas.frx":1D6A
      Column(22)      =   "frmFacturasEmitidas.frx":1EE2
      Column(23)      =   "frmFacturasEmitidas.frx":2052
      Column(24)      =   "frmFacturasEmitidas.frx":21B2
      Column(25)      =   "frmFacturasEmitidas.frx":22EA
      Column(26)      =   "frmFacturasEmitidas.frx":243E
      FormatStylesCount=   16
      FormatStyle(1)  =   "frmFacturasEmitidas.frx":2562
      FormatStyle(2)  =   "frmFacturasEmitidas.frx":269A
      FormatStyle(3)  =   "frmFacturasEmitidas.frx":274A
      FormatStyle(4)  =   "frmFacturasEmitidas.frx":27FE
      FormatStyle(5)  =   "frmFacturasEmitidas.frx":28D6
      FormatStyle(6)  =   "frmFacturasEmitidas.frx":298E
      FormatStyle(7)  =   "frmFacturasEmitidas.frx":2A6E
      FormatStyle(8)  =   "frmFacturasEmitidas.frx":2AFA
      FormatStyle(9)  =   "frmFacturasEmitidas.frx":2BDA
      FormatStyle(10) =   "frmFacturasEmitidas.frx":2C8A
      FormatStyle(11) =   "frmFacturasEmitidas.frx":2D3E
      FormatStyle(12) =   "frmFacturasEmitidas.frx":2DEE
      FormatStyle(13) =   "frmFacturasEmitidas.frx":2E9E
      FormatStyle(14) =   "frmFacturasEmitidas.frx":2F52
      FormatStyle(15) =   "frmFacturasEmitidas.frx":302A
      FormatStyle(16) =   "frmFacturasEmitidas.frx":310E
      ImageCount      =   1
      ImagePicture(1) =   "frmFacturasEmitidas.frx":31EE
      PrinterProperties=   "frmFacturasEmitidas.frx":3508
   End
   Begin XtremeSuiteControls.TaskDialog taskDialog 
      Left            =   14955
      Top             =   750
      _Version        =   786432
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   0
      WindowTitle     =   "TaskDialog1"
   End
   Begin VB.Menu mnuFacturas 
      Caption         =   "armnuFacturas"
      Visible         =   0   'False
      Begin VB.Menu NRO 
         Caption         =   "nro"
         Enabled         =   0   'False
      End
      Begin VB.Menu editar 
         Caption         =   "Editar..."
      End
      Begin VB.Menu separador2 
         Caption         =   "-"
      End
      Begin VB.Menu aprobarFactura 
         Caption         =   "Aprobar localmente..."
      End
      Begin VB.Menu mnuEnviarAfip 
         Caption         =   "Enviar a AFIP..."
      End
      Begin VB.Menu separador 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAprobarEnviar 
         Caption         =   "Aprobar localmente y Enviar a AFIP..."
      End
      Begin VB.Menu mnuDesaprobarFactura 
         Caption         =   "Desaprobar..."
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu sepa3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRechazo 
         Caption         =   "Rechazo de comprobantes (FCE)"
      End
      Begin VB.Menu ImprimirFactura 
         Caption         =   "Imprimir..."
      End
      Begin VB.Menu AnularFactura 
         Caption         =   "Anular"
      End
      Begin VB.Menu desAnular 
         Caption         =   "Quitar Anulaci?n"
         Visible         =   0   'False
      End
      Begin VB.Menu aplicar 
         Caption         =   "Aplicar Recibo..."
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu aplicarNCaFC 
         Caption         =   "Aplicar NC a Factura o ND..."
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu aplicarNDaFC 
         Caption         =   "Aplicar ND a Factura o NC..."
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu o 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCrearCopiaFactura 
         Caption         =   "Crear copia a partir de comprobante"
      End
      Begin VB.Menu mnuFechaPagoPropuesta 
         Caption         =   "Establecer Fecha Pago Propuesta"
      End
      Begin VB.Menu mnuFechaEntrega 
         Caption         =   "Establecer Fecha Entrega..."
      End
      Begin VB.Menu sdf 
         Caption         =   "-"
      End
      Begin VB.Menu verHistorialFactura 
         Caption         =   "Ver Historial..."
      End
      Begin VB.Menu mnuArchivos 
         Caption         =   "Archivos Asociados..."
      End
      Begin VB.Menu LineaUlt 
         Caption         =   "-"
      End
      Begin VB.Menu MnuVerRecibo 
         Caption         =   "Ver Recibos de Cobro"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditarCampos 
         Caption         =   "Editar Datos..."
      End
      Begin VB.Menu verFactura 
         Caption         =   "Ver Detalle..."
      End
      Begin VB.Menu archivos 
         Caption         =   "Archivos Asociados..."
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu scanear 
         Caption         =   "Adquirir..."
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmAdminFacturasEmitidas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements ISuscriber

Dim vId As String
Dim facturas As Collection
Dim Factura As Factura
Dim m_Archivos As Dictionary


Private Sub AnularFactura_Click()
    Dim r As Long
    r = Me.gridComprobantesEmitidos.rowIndex(Me.gridComprobantesEmitidos.row)
    If MsgBox("¿Desea anular el comprobante seleccionado?", vbYesNo, "Confirmacion") = vbYes Then

        If DAOFactura.Anular(Factura) Then
            MsgBox "Comprobante anulado con éxito!", vbInformation, "Información"
            Me.gridComprobantesEmitidos.RefreshRowIndex r
        Else
            MsgBox "Hubo un error. No se anulo el comprobante!", vbCritical, "Error"
        End If

    End If
End Sub

Private Sub aplicarNCaFC_Click()
    On Error GoTo err1
    If MsgBox("¿Seguro de aplicar comprobante?", vbYesNo, "Confirmación") = vbYes Then
        'seleccionar factura para aplicar
        Set Selecciones.Factura = Nothing
        Dim F As New frmAdminFacturasNCElegirFC

        F.idCliente = Factura.cliente.Id
        F.TiposDocs.Add tipoDocumentoContable.Factura

        If Factura.TipoDocumento = tipoDocumentoContable.NotaCredito Then
            F.TiposDocs.Add tipoDocumentoContable.notaDebito
        End If
        If Factura.TipoDocumento = tipoDocumentoContable.notaDebito Then
            F.TiposDocs.Add tipoDocumentoContable.NotaCredito
        End If

        F.EstadosDocs.Add EstadoFacturaCliente.Aprobada
        F.Show 1

        If IsSomething(Selecciones.Factura) Then
            If DAOFactura.aplicarNCaFC(Selecciones.Factura.Id, Factura.Id) Then
                llenarGrilla

                MsgBox "Comprobantes Vinculados: " & Factura.NumeroFormateado & " | " & Selecciones.Factura.NumeroFormateado & vbNewLine & "" _
                     & "Aplicación exitosa!", vbInformation, "Información"

            End If
        End If
    End If
    Exit Sub
err1:
    MsgBox Err.Description, vbCritical, "Error"
End Sub


Private Sub aplicarNDaFC_Click()
    On Error GoTo err1
    If MsgBox("¿Seguro de aplicar la ND a un comprobante?", vbYesNo, "Confirmación") = vbYes Then
        'seleccionar factura para aplicar
        Set Selecciones.Factura = Nothing
        Dim F As New frmAdminFacturasNCElegirFC

        F.idCliente = Factura.cliente.Id
        F.TiposDocs.Add tipoDocumentoContable.Factura
        
        If Factura.TipoDocumento = tipoDocumentoContable.notaDebito Then
            F.TiposDocs.Add tipoDocumentoContable.NotaCredito
        End If

        F.EstadosDocs.Add EstadoFacturaCliente.Aprobada
        F.Show 1

        If IsSomething(Selecciones.Factura) Then
            If DAOFactura.aplicarNotaDebitoaFC(Selecciones.Factura.Id, Factura.Id) Then
                llenarGrilla

                MsgBox "Comprobantes Vinculados: " & Factura.NumeroFormateado & " | " & Selecciones.Factura.NumeroFormateado & vbNewLine & "" _
                     & "Aplicación exitosa!", vbInformation, "Información"

            End If
        End If
    End If
    Exit Sub
err1:
    MsgBox Err.Description, vbCritical, "Error"
End Sub

Private Sub aprobarFactura_Click()
    On Error GoTo err1
    Dim g As Long
    Dim msgadicional As String
    msgadicional = ""
    If MsgBox("¿Desea aprobar localmente el comprobante?", vbYesNo + vbQuestion, "Confirmacion") = vbYes Then
        g = Me.gridComprobantesEmitidos.rowIndex(Me.gridComprobantesEmitidos.row)
        If DAOFactura.aprobarV2(Factura, True, False) Then

            If Factura.Tipo.PuntoVenta.EsElectronico And Not Factura.Tipo.PuntoVenta.CaeManual And Not Factura.AprobadaAFIP Then
                msgadicional = "Esta factura deberá enviarse a la afip"
            End If
            If Factura.Tipo.PuntoVenta.EsElectronico And Factura.Tipo.PuntoVenta.CaeManual And Not Factura.AprobadaAFIP Then
                msgadicional = "Recuerde agregar al comprobante: CAE y fecha de vencimiento del CAE "
            End If

            Dim msg As String
            msg = "Comprobante aprobado con éxito!"
            If IsSomething(Factura.CaeSolicitarResponse) Then
                If LenB(Factura.CaeSolicitarResponse.observaciones) > 5 Then

                    msg = msg & Chr(10) & Factura.CaeSolicitarResponse.observaciones
                End If

                If LenB(msgadicional) > 0 Then
                    msg = msg & Chr(10) & msgadicional
                End If

            End If
            MsgBox msg, vbInformation, "Información"

            Me.gridComprobantesEmitidos.RefreshRowIndex g
            Me.txtNroFactura.SetFocus

        Else
            GoTo err1
        End If
    End If
    Exit Sub
err1:
    'MsgBox "Factura no aprobada, compruebe:" & vbNewLine & "Si la factura es de anticipo, compruebe que el valor de la misma sea el mismo que el anticipo de la OT." & vbNewLine & "Que el detalle del remito no este ya facturado." & vbNewLine & Err.Description, vbCritical

    MsgBox Err.Description, vbCritical, Err.Source
    Me.gridComprobantesEmitidos.RefreshRowIndex g
End Sub

Private Sub archivos_Click()
    Dim F As New frmArchivos2
    F.Origen = 101
    F.ObjetoId = Factura.Id
    F.caption = "Comprobante " & Factura.GetShortDescription(False, True)
    F.Show
End Sub


Private Sub btnClearMoneda_Click()
    Me.cboMoneda.ListIndex = -1
End Sub

Private Sub btnClearTipo_Click()
    Me.cboTipo.ListIndex = 3
End Sub

Private Sub btnExportar_Click()

'FUNCIÓN PARA EXPORTAR A EXCEL


'INICIA EL PROGRESSBAR Y LO MUESTRA
    Me.barProgreso.Visible = True

    'DEFINE EL VALOR MINIMO Y EL MAXIMO DEL PROGRESSBAR (CANTIDAD DE DATOS EN LA COLECCIÓN COL)
    barProgreso.min = 0
    barProgreso.max = facturas.count


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

    xlWorksheet.Cells(1, 1).value = "Reporte de comprobantes emitidos"

    '    If (id > 0) Then
    '        xlWorksheet.Cells(1, 2).value = DAOCliente.BuscarPorID(id).razon
    '    Else
    '        xlWorksheet.Cells(1, 2).value = "Todos"
    '    End If

    xlWorksheet.Columns(4).HorizontalAlignment = xlLeft
    xlWorksheet.Columns(12).HorizontalAlignment = xlLeft

    xlWorksheet.Cells(2, 1).value = "Comprobante"
    xlWorksheet.Cells(2, 2).value = "Emision"

    xlWorksheet.Cells(2, 3).value = "Moneda"
    xlWorksheet.Cells(2, 4).value = "Cotización"

    xlWorksheet.Cells(2, 5).value = "Detalle"

    xlWorksheet.Cells(2, 6).value = "Neto Gravado $ ARS"
    xlWorksheet.Cells(2, 7).value = "Neto Gravado U$S"

    xlWorksheet.Cells(2, 8).value = "Percepciones $ ARS"
    xlWorksheet.Cells(2, 9).value = "Percepciones U$S"

    xlWorksheet.Cells(2, 10).value = "IVA $ ARS"
    xlWorksheet.Cells(2, 11).value = "IVA U$S"

    xlWorksheet.Cells(2, 12).value = "Exento $ ARS"
    xlWorksheet.Cells(2, 13).value = "Exento U$S"

    xlWorksheet.Cells(2, 14).value = "Importe en $ ARS"
    xlWorksheet.Cells(2, 15).value = "Importe en U$S"

    xlWorksheet.Cells(2, 16).value = "Vencimiento"
    xlWorksheet.Cells(2, 17).value = "Atraso / Estado"
    xlWorksheet.Cells(2, 18).value = "Entrega"
    xlWorksheet.Cells(2, 19).value = "Atraso / Dias"
    xlWorksheet.Cells(2, 20).value = "Cliente"
    xlWorksheet.Cells(2, 21).value = "Cuit"
    xlWorksheet.Cells(2, 22).value = "Observaciones"
    xlWorksheet.Cells(2, 23).value = "Observaciones Cancela"
    xlWorksheet.Cells(2, 24).value = "Recibos Asociados"
    xlWorksheet.Cells(2, 25).value = "ID"
    xlWorksheet.Cells(2, 26).value = "ID_Asociacion"
    
    xlWorksheet.Range("A2:Y2").Font.Bold = True

    Dim idx As Integer
    idx = 3

    Dim fac As Factura

    'DEFINE EL CONTADOR DEL PROGRESSBAR Y LO INICIA EN 0
    Dim d As Long
    d = 0


    For Each fac In facturas

        xlWorksheet.Cells(idx, 1).value = fac.GetShortDescription(False, True)
        xlWorksheet.Cells(idx, 2).value = fac.FechaEmision
        xlWorksheet.Cells(idx, 3).value = fac.moneda.NombreCorto
        xlWorksheet.Cells(idx, 4).value = fac.CambioAPatron
        xlWorksheet.Cells(idx, 5).value = fac.OrdenCompra

        If fac.moneda.Cambio = 1 Then

            If fac.TipoDocumento = tipoDocumentoContable.NotaCredito Then
                ' VAN TODOS NEGATIVOS SI ES NOTA DE CREDITO
                xlWorksheet.Cells(idx, 6).value = (((fac.TotalEstatico.TotalNetoGravado * fac.CambioAPatron) + fac.TotalEstatico.TotalIVADiscrimandoONo) - fac.TotalEstatico.TotalIVA) * -1
                xlWorksheet.Cells(idx, 8).value = fac.TotalEstatico.TotalPercepcionesIB * fac.CambioAPatron * -1
                xlWorksheet.Cells(idx, 10).value = fac.TotalEstatico.TotalIVA * fac.CambioAPatron * -1
                xlWorksheet.Cells(idx, 12).value = fac.TotalEstatico.TotalExento * fac.CambioAPatron * -1
                xlWorksheet.Cells(idx, 14).value = fac.TotalEstatico.total * fac.CambioAPatron * -1
                'dolares
                xlWorksheet.Cells(idx, 7).value = "0"
                'dolares
                xlWorksheet.Cells(idx, 9).value = "0"
                'dolares
                xlWorksheet.Cells(idx, 11).value = ""
                'dolares
                xlWorksheet.Cells(idx, 13).value = "0"
                'dolares
                xlWorksheet.Cells(idx, 15).value = "0"

            Else
                ' VAN TODOS POSITIVOS AL NO SER NOTA DE CREDITO
                xlWorksheet.Cells(idx, 6).value = (((fac.TotalEstatico.TotalNetoGravado * fac.CambioAPatron) + fac.TotalEstatico.TotalIVADiscrimandoONo) - fac.TotalEstatico.TotalIVA)
                xlWorksheet.Cells(idx, 8).value = fac.TotalEstatico.TotalPercepcionesIB * fac.CambioAPatron
                xlWorksheet.Cells(idx, 10).value = fac.TotalEstatico.TotalIVA * fac.CambioAPatron
                xlWorksheet.Cells(idx, 12).value = fac.TotalEstatico.TotalExento * fac.CambioAPatron
                xlWorksheet.Cells(idx, 14).value = fac.TotalEstatico.total * fac.CambioAPatron
                'dolares
                xlWorksheet.Cells(idx, 7).value = "0"
                'dolares
                xlWorksheet.Cells(idx, 9).value = "0"
                'dolares
                xlWorksheet.Cells(idx, 11).value = "0"
                'dolares
                xlWorksheet.Cells(idx, 13).value = "0"
                'dolares
                xlWorksheet.Cells(idx, 15).value = "0"
            End If

        Else
            If fac.TipoDocumento = tipoDocumentoContable.NotaCredito Then
                ' VAN TODOS NEGATIVOS SI ES NOTA DE CREDITO
                xlWorksheet.Cells(idx, 6).value = fac.TotalEstatico.TotalNetoGravado * fac.CambioAPatron * -1
                xlWorksheet.Cells(idx, 8).value = fac.TotalEstatico.TotalPercepcionesIB * fac.CambioAPatron * -1
                xlWorksheet.Cells(idx, 10).value = fac.TotalEstatico.TotalIVA * fac.CambioAPatron * -1
                xlWorksheet.Cells(idx, 12).value = fac.TotalEstatico.TotalExento * fac.CambioAPatron * -1
                xlWorksheet.Cells(idx, 14).value = fac.TotalEstatico.total * fac.CambioAPatron * -1
                'dolares
                xlWorksheet.Cells(idx, 7).value = fac.TotalEstatico.TotalNetoGravado * -1
                'dolares
                xlWorksheet.Cells(idx, 9).value = fac.TotalEstatico.TotalPercepcionesIB * -1
                'dolares
                xlWorksheet.Cells(idx, 11).value = fac.TotalEstatico.TotalIVA * -1
                'dolares
                xlWorksheet.Cells(idx, 13).value = fac.TotalEstatico.TotalExento * -1
                'dolares
                xlWorksheet.Cells(idx, 15).value = fac.TotalEstatico.total * -1

            Else
                ' VAN TODOS POSITIVOS AL NO SER NOTA DE CREDITO
                xlWorksheet.Cells(idx, 6).value = fac.TotalEstatico.TotalNetoGravado * fac.CambioAPatron
                xlWorksheet.Cells(idx, 8).value = fac.TotalEstatico.TotalPercepcionesIB * fac.CambioAPatron
                xlWorksheet.Cells(idx, 10).value = fac.TotalEstatico.TotalIVA * fac.CambioAPatron
                xlWorksheet.Cells(idx, 12).value = fac.TotalEstatico.TotalExento * fac.CambioAPatron
                xlWorksheet.Cells(idx, 14).value = fac.TotalEstatico.total * fac.CambioAPatron
                'dolares
                xlWorksheet.Cells(idx, 7).value = fac.TotalEstatico.TotalNetoGravado
                'dolares
                xlWorksheet.Cells(idx, 9).value = fac.TotalEstatico.TotalPercepcionesIB
                'dolares
                xlWorksheet.Cells(idx, 11).value = fac.TotalEstatico.TotalIVA
                'dolares
                xlWorksheet.Cells(idx, 13).value = fac.TotalEstatico.TotalExento
                'dolares
                xlWorksheet.Cells(idx, 15).value = fac.TotalEstatico.total
            End If

        End If


        'xlWorksheet.Cells(idx, 16).value = fac.Vencimiento
        xlWorksheet.Cells(idx, 17).value = fac.StringDiasAtraso

        If xlWorksheet.Cells(idx, 17).value = "En Edición" Then
            xlWorksheet.Cells(idx, 17).Interior.ColorIndex = 46    ' naranja
        End If

        If (fac.DiferenciaDiasEntrega <> -1) Then
            xlWorksheet.Cells(idx, 18).value = Format(fac.FechaEntrega, "dd/mm/yyyy")
            xlWorksheet.Cells(idx, 19).value = fac.DiferenciaDiasEntrega & " dias"
        Else
            xlWorksheet.Cells(idx, 18).value = "no definida"
            xlWorksheet.Cells(idx, 19).value = 0
        End If

        xlWorksheet.Cells(idx, 20).value = fac.cliente.razon
        xlWorksheet.Cells(idx, 21).value = fac.cliente.Cuit
        xlWorksheet.Cells(idx, 22).value = fac.observaciones
        xlWorksheet.Cells(idx, 23).value = fac.observaciones_cancela
        xlWorksheet.Cells(idx, 24).value = fac.RecibosAplicadosId
        
        xlWorksheet.Cells(idx, 25).value = fac.Id
        xlWorksheet.Cells(idx, 26).value = fac.idAsociacion
        idx = idx + 1

        'POR CADA ITERACION SUMA UN VALOR A LA VARIABLE D DEL PROGRESSBAR
        d = d + 1
        barProgreso.value = d


    Next

    xlWorksheet.Cells(idx + 1, 5).value = "Totales: "
    xlWorksheet.Cells(idx + 1, 5).HorizontalAlignment = xlRight

    xlWorksheet.Cells(idx + 1, 6).Formula = "=SUM(F3:F" & idx - 1 & ")"
    xlWorksheet.Cells(idx + 1, 7).Formula = "=SUM(G3:G" & idx - 1 & ")"
    xlWorksheet.Cells(idx + 1, 8).Formula = "=SUM(H3:H" & idx - 1 & ")"
    xlWorksheet.Cells(idx + 1, 9).Formula = "=SUM(I3:I" & idx - 1 & ")"
    xlWorksheet.Cells(idx + 1, 10).Formula = "=SUM(J3:J" & idx - 1 & ")"
    xlWorksheet.Cells(idx + 1, 11).Formula = "=SUM(K3:K" & idx - 1 & ")"
    xlWorksheet.Cells(idx + 1, 12).Formula = "=SUM(L3:L" & idx - 1 & ")"
    xlWorksheet.Cells(idx + 1, 13).Formula = "=SUM(M3:M" & idx - 1 & ")"
    xlWorksheet.Cells(idx + 1, 14).Formula = "=SUM(N3:N" & idx - 1 & ")"
    xlWorksheet.Cells(idx + 1, 15).Formula = "=SUM(O3:O" & idx - 1 & ")"

    xlWorksheet.Cells(idx + 1, 6).Font.Bold = True
    xlWorksheet.Cells(idx + 1, 7).Font.Bold = True
    xlWorksheet.Cells(idx + 1, 8).Font.Bold = True
    xlWorksheet.Cells(idx + 1, 9).Font.Bold = True
    xlWorksheet.Cells(idx + 1, 10).Font.Bold = True
    xlWorksheet.Cells(idx + 1, 11).Font.Bold = True
    xlWorksheet.Cells(idx + 1, 12).Font.Bold = True
    xlWorksheet.Cells(idx + 1, 13).Font.Bold = True
    xlWorksheet.Cells(idx + 1, 14).Font.Bold = True
    xlWorksheet.Cells(idx + 1, 15).Font.Bold = True

    xlWorksheet.Range("F3:O15").NumberFormat = "#,##0.00"
    xlWorksheet.Range("F3:O" & idx + 1).HorizontalAlignment = xlRight
    xlWorksheet.Range("F3:O" & idx + 1).NumberFormat = "#,##0.00"

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

    xlWorksheet.Activate
    xlWorksheet.Range("A3").Select
    xlWorksheet.Application.ActiveWindow.FreezePanes = True

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
    barProgreso.value = 0
    Me.barProgreso.Visible = False

End Sub

Private Sub btnReducirVentana_Click()
'SE REDUCE
    If btnReducirVentana.caption = "R" Then
        Me.grpTotalizadores(0).Visible = False
        Me.grpFiltrosPrincipal.Visible = False
        
        Me.btnBuscar.Visible = False
        
        Me.grpBotones(1).Top = 120
        Me.gridComprobantesEmitidos.Top = Me.grpBotones(1).Top + 960
        
        Me.gridComprobantesEmitidos.Height = Me.ScaleHeight - 1300
        
        Me.gridComprobantesEmitidos.GroupByBoxVisible = False
        
        Me.gridComprobantesEmitidos.PreviewRowLines = 0
                
        Me.btnReducirVentana.caption = "A"

        
'SE REESTABLECE
    ElseIf btnReducirVentana.caption = "A" Then
        Me.grpTotalizadores(0).Visible = True
        Me.grpFiltrosPrincipal.Visible = True

        Me.btnBuscar.Visible = True

        Me.grpBotones(1).Top = 1680
        Me.gridComprobantesEmitidos.Top = Me.grpBotones(1).Top + 960

        Me.gridComprobantesEmitidos.Height = Me.ScaleHeight - 2900
        
        Me.gridComprobantesEmitidos.PreviewRowLines = 1
                
        Me.btnReducirVentana.caption = "R"
        
    End If
    
End Sub

Private Sub cboRangos_Click()
    funciones.CalculateDateRange Me.cboRangos, Me.dtpDesde, Me.dtpHasta
End Sub

Private Sub chkAgruparCbtes_Click()
    agruparAsociados
End Sub

Private Sub agruparAsociados()
    If Me.chkAgruparCbtes Then
        Me.gridComprobantesEmitidos.Groups.Add 26, jgexSortDescending
    Else
        Me.gridComprobantesEmitidos.Groups.Clear
    End If

End Sub


' 1451- AGREGO FUNCION DE MOSTRAR ID U OCULTAR
Private Sub chkboxVerIds_Click()
    verIds
    
End Sub

Private Sub verIds()
    If Me.chkboxVerIds Then
        Me.gridComprobantesEmitidos.Columns(24).Visible = True
        Me.gridComprobantesEmitidos.Columns(24).Width = 800
    Else
        Me.gridComprobantesEmitidos.Columns(24).Visible = False
    End If
    
End Sub


Private Sub chkVerObservaciones_Click()
    verObservaciones
End Sub


Private Sub verObservaciones()
    If Me.chkObservaciones(1) Then
        Me.gridComprobantesEmitidos.PreviewRowLines = 1
    Else
        Me.gridComprobantesEmitidos.PreviewRowLines = 0
    End If
End Sub


Private Sub btnBuscar_Click()
    llenarGrilla

End Sub

Private Sub btnImprimir_Click()
    With Me.gridComprobantesEmitidos.PrinterProperties
        .FitColumns = True
        .RepeatHeaders = True
        .Orientation = jgexPPLandscape
        .HeaderString(jgexHFCenter) = "Emitidos"
        .FooterString(jgexHFCenter) = Now
        '202
        .FooterDistance = 1500
        .FooterString(jgexHFLeft) = lblTotalNeto & Chr(10) & lblTotalIVA & Chr(10) & lblTotalPercepciones & Chr(10) & lblTotal
        '202

    End With
    Load frmPrintPreview
    frmPrintPreview.Move Me.Left, Me.Top, Me.Width, Me.Height
    gridComprobantesEmitidos.PrintPreview frmPrintPreview.GEXPreview1
    frmPrintPreview.Show 1
End Sub



Private Sub chkObservaciones_Click(Index As Integer)
verObservaciones
End Sub

Private Sub cmdLimpiarCboEstadoAfip_Click()
    Me.cboEstadoAfip.ListIndex = -1
End Sub


Private Sub editar_Click()

    Dim f_c3h3 As New frmAdminFacturasEdicion
    f_c3h3.idFactura = Factura.Id
    f_c3h3.Show

End Sub

Private Sub Form_Load()
    FormHelper.Customize Me

    Me.gridComprobantesEmitidos.ItemCount = 0
    GridEXHelper.CustomizeGrid Me.gridComprobantesEmitidos, True    ', True

    Me.gridComprobantesEmitidos.GroupByBoxVisible = False
        
    DAOCliente.llenarComboXtremeSuite Me.cboClientes, False, True, False
    Me.cboClientes.ListIndex = -1
    
    DAOMoneda.llenarComboXtremeSuite Me.cboMoneda
    Me.cboMoneda.ListIndex = -1
    
    vId = funciones.CreateGUID

    Channel.AgregarSuscriptor Me, FacturaCliente_

    'Modificaci?n 15/05/20 (Se muestran todos los comprobanes sin filtrar por punto de venta)
    DAOPuntoVenta.llenarComboXtremeSuite Me.cboPuntosVenta, False

    cboEstados.Clear
    cboEstados.AddItem "En Edición"
    cboEstados.ItemData(cboEstados.NewIndex) = 1
    cboEstados.AddItem "Aprobadas"
    cboEstados.ItemData(cboEstados.NewIndex) = 2
    cboEstados.AddItem "Anuladas"
    cboEstados.ItemData(cboEstados.NewIndex) = 3
    cboEstados.AddItem "Cancela NC"
    cboEstados.ItemData(cboEstados.NewIndex) = 4
    cboEstados.AddItem "Cancela NC Parcial"
    cboEstados.ItemData(cboEstados.NewIndex) = 5
    cboEstados.AddItem "Aplicada de ND"
    cboEstados.ItemData(cboEstados.NewIndex) = 6
    
    Me.cboEstadosSaldada.Clear
    cboEstadosSaldada.AddItem "No Saldada"
    cboEstadosSaldada.ItemData(cboEstadosSaldada.NewIndex) = 0
    cboEstadosSaldada.AddItem "Total"
    cboEstadosSaldada.ItemData(cboEstadosSaldada.NewIndex) = 1
    cboEstadosSaldada.AddItem "Parcial"
    cboEstadosSaldada.ItemData(cboEstadosSaldada.NewIndex) = 2
    cboEstadosSaldada.AddItem "Cancela NC"
    cboEstadosSaldada.ItemData(cboEstadosSaldada.NewIndex) = 3
    cboEstadosSaldada.AddItem "Cancela NC Parcial"
    cboEstadosSaldada.ItemData(cboEstadosSaldada.NewIndex) = 4

    Me.cboTipo.Clear
    cboTipo.AddItem "Todos"
'    cboTipo.ItemData(cboTipo.NewIndex) = 0
    cboTipo.AddItem "FC"
    cboTipo.ItemData(cboTipo.NewIndex) = 1
    cboTipo.AddItem "NC"
    cboTipo.ItemData(cboTipo.NewIndex) = 2
    cboTipo.AddItem "ND"
    cboTipo.ItemData(cboTipo.NewIndex) = 3
    
    cboTipo.ListIndex = 3
    
    Me.cboEstadoAfip.Clear
    cboEstadoAfip.AddItem "Sólo informadas"
    cboEstadoAfip.ItemData(cboEstadoAfip.NewIndex) = 0
    cboEstadoAfip.AddItem "Sólo no informadas"
    cboEstadoAfip.ItemData(cboEstadoAfip.NewIndex) = 1

    Me.cboOrdenImporte.Clear
    cboOrdenImporte.AddItem "Ascendente"
    cboOrdenImporte.ItemData(cboOrdenImporte.NewIndex) = 0
    cboOrdenImporte.AddItem "Descendente"
    cboOrdenImporte.ItemData(cboOrdenImporte.NewIndex) = 1

    Dim i As Integer
    funciones.FillComboBoxDateRanges Me.cboRangos
    For i = 0 To Me.cboRangos.ListCount - 1
        If Me.cboRangos.ItemData(i) = DateRangeValue.DRV_YearCurrent Then Exit For
    Next i

    Me.cboRangos.ListIndex = i
    llenarGrilla
    verObservaciones

    ''Me.caption = caption & "(" & Name & ")"

End Sub

Private Sub llenarGrilla()
    Dim filtro As String
    Set m_Archivos = DAOArchivo.GetCantidadArchivosPorReferencia(OA_factura)

    Me.gridComprobantesEmitidos.ItemCount = 0
    filtro = "1=1"

    If Me.cboClientes.ListIndex >= 0 Then
        filtro = filtro & " and idCliente=" & cboClientes.ItemData(Me.cboClientes.ListIndex)
    End If

    If Me.cboPuntosVenta.ListIndex >= 0 Then
        filtro = filtro & " and pv.id=" & cboPuntosVenta.ItemData(Me.cboPuntosVenta.ListIndex)
    End If

    If Me.cboEstados.ListIndex >= 0 Then
        filtro = filtro & " and AdminFacturas.estado=" & cboEstados.ItemData(Me.cboEstados.ListIndex)
    End If

    If Me.cboEstadosSaldada.ListIndex >= 0 Then
        filtro = filtro & " and AdminFacturas.saldada=" & cboEstadosSaldada.ItemData(Me.cboEstadosSaldada.ListIndex)
    End If

    If Me.chkCredito.value > 0 Then
        filtro = filtro & " and AdminFacturas.EsCredito=" & Me.chkCredito.value
    End If

    If LenB(Me.txtReferencia.text) > 0 Then
        filtro = filtro & " and AdminFacturas.OrdenCompra like '%" & Trim(Me.txtReferencia.text) & "%'"
    End If
    
    If LenB(Me.txtNroFactura) > 0 And IsNumeric(Me.txtNroFactura) Then
        filtro = filtro & " and nroFactura=" & Me.txtNroFactura
    End If
    
    If LenB(Me.txtID) > 0 And IsNumeric(Me.txtID) Then
        filtro = filtro & " and AdminFacturas.id=" & Me.txtID
    End If

    If Not IsNull(Me.dtpDesde.value) Then
        filtro = filtro & " AND AdminFacturas.FechaEmision >= " & conectar.Escape(Me.dtpDesde.value)
    End If

    If Not IsNull(Me.dtpHasta.value) Then
        filtro = filtro & " AND AdminFacturas.FechaEmision <= " & conectar.Escape(Me.dtpHasta.value)
    End If

    If LenB(Me.txtRemitoAplicado.text) > 0 Then
        filtro = filtro & " and AdminFacturas.id IN (SELECT fd.idFactura FROM AdminFacturasDetalleNueva fd INNER JOIN entregas e ON e.id = fd.idEntrega INNER JOIN remitos r ON r.id = e.Remito WHERE r.numero = " & Me.txtRemitoAplicado.text & ")"
    End If

    If Me.cboEstadoAfip.ListIndex = 0 Then
        filtro = filtro & " and AdminFacturas.aprobacion_afip=1"
    End If

    If Me.cboEstadoAfip.ListIndex = 1 Then
        filtro = filtro & " and AdminFacturas.aprobacion_afip=0"
    End If
    
    If Me.cboTipo.ListIndex = 4 Then ' Todos
        filtro = filtro & ""
    ElseIf Me.cboTipo.ListIndex = 0 Then ' FC
        filtro = filtro & " and AdminFacturas.id_tipo_discriminado IN (1, 4, 7, 10, 14, 15, 21)"
    ElseIf Me.cboTipo.ListIndex = 1 Then ' NC
        filtro = filtro & " and AdminFacturas.id_tipo_discriminado IN (2, 5, 8, 11, 13, 16, 22)"
    ElseIf Me.cboTipo.ListIndex = 2 Then ' ND
        filtro = filtro & " and AdminFacturas.id_tipo_discriminado IN (3, 6, 9, 12, 17)"
    End If

    If Me.cboMoneda.ListIndex <> -1 Then
        If Me.cboMoneda.ListIndex = 0 Then
            filtro = filtro & " and AdminFacturas.idMoneda = 00000000000"
        ElseIf Me.cboMoneda.ListIndex = 1 Then
            filtro = filtro & " and AdminFacturas.idMoneda = 00000000002"
        ElseIf Me.cboMoneda.ListIndex = 2 Then
            filtro = filtro & " and AdminFacturas.idMoneda = 00000000003"
        ElseIf Me.cboMoneda.ListIndex = 3 Then
            filtro = filtro & " and AdminFacturas.idMoneda = 00000000001"
        End If
    End If


    Dim ordenImporte As String

    If Me.cboOrdenImporte.ListIndex = 0 Then
        ordenImporte = "AdminFacturas.total_estatico * AdminFacturas.cambio_a_patron ASC"
    ElseIf Me.cboOrdenImporte.ListIndex = 1 Then
        ordenImporte = "AdminFacturas.total_estatico * AdminFacturas.cambio_a_patron DESC"
    End If


    Set facturas = DAOFactura.FindAll(filtro, , , ordenImporte)

    Dim F As Factura
    Dim c As Integer

    For Each F In facturas
        Dim total As Double
        Dim Percepcion As Double
        Dim totalPercepcionesIIBB As Double
        Dim TotalIVATodo As Double
        Dim totalNG As Double
        
        Dim totalDolares As Double
        Dim totalPercepcionesIIBBDolar As Double
        Dim totalIVADolar As Double
        Dim totalNGDolar As Double
        
        
        If F.TipoDocumento = tipoDocumentoContable.NotaCredito Then c = -1 Else c = 1

        total = total + MonedaConverter.ConvertirForzado2(F.TotalEstatico.total * c, MonedaConverter.Patron.Id, F.moneda.Id, F.CambioAPatron)
        TotalIVATodo = TotalIVATodo + MonedaConverter.ConvertirForzado2(F.TotalEstatico.TotalIVADiscrimandoONo * c, MonedaConverter.Patron.Id, F.moneda.Id, F.CambioAPatron)
        totalNG = totalNG + MonedaConverter.ConvertirForzado2(F.TotalEstatico.TotalNetoGravado * c, MonedaConverter.Patron.Id, F.moneda.Id, F.CambioAPatron)
        Percepcion = F.TotalEstatico.TotalPercepcionesIB * c
        totalPercepcionesIIBB = totalPercepcionesIIBB + MonedaConverter.ConvertirForzado2(F.TotalEstatico.TotalPercepcionesIB * c, MonedaConverter.Patron.Id, F.moneda.Id, F.CambioAPatron)
        
        ' Moneda.id = 1 >> DOLAR
        If F.moneda.Id = 1 Then
                totalDolares = totalDolares + F.TotalEstatico.total * c
                totalPercepcionesIIBBDolar = totalPercepcionesIIBBDolar + F.TotalEstatico.TotalPercepcionesIB * c
                totalIVADolar = totalIVADolar + F.TotalEstatico.TotalIVADiscrimandoONo * c
                totalNGDolar = totalNGDolar + F.TotalEstatico.TotalNetoGravado * c
        End If
        
    Next

    Me.lblTotal = "Total: " & FormatCurrency(funciones.FormatearDecimales(total))
    Me.lblTotalPercepciones = "Total Percepciones: " & FormatCurrency(funciones.FormatearDecimales(totalPercepcionesIIBB))
    Me.lblTotalIVA = "Total IVA: " & FormatCurrency(funciones.FormatearDecimales(TotalIVATodo))
    Me.lblTotalNeto = "Total NG: " & FormatCurrency(funciones.FormatearDecimales(totalNG))

    Me.lblTotalDolares = "Total: U$S " & Replace(FormatCurrency(totalDolares), "$", "")
    Me.lblPercepciones_Dolar.caption = "Total Percepciones: U$S " & Replace(FormatCurrency(funciones.FormatearDecimales(totalPercepcionesIIBBDolar)), "$", "")
    Me.lblIVA_Dolar.caption = "Total IVA: U$S " & Replace(FormatCurrency(funciones.FormatearDecimales(totalIVADolar)), "$", "")
    Me.lblNG_Dolar.caption = "Total NG: U$S " & Replace(FormatCurrency(funciones.FormatearDecimales(totalNGDolar)), "$", "")
    
    Me.gridComprobantesEmitidos.ItemCount = 0
    Me.gridComprobantesEmitidos.ItemCount = facturas.count

    ' 1451- AGREGO FUNCION DE MOSTRAR ID U OCULTAR
    Me.gridComprobantesEmitidos.Columns(24).Visible = False

    Me.caption = "Cbtes. Venta [Cant: " & facturas.count & "]"

End Sub


Private Sub Form_Resize()
    On Error Resume Next
    Me.gridComprobantesEmitidos.Width = Me.ScaleWidth - 400
    Me.gridComprobantesEmitidos.Height = Me.ScaleHeight - 3200
    Me.grpFiltrosPrincipal.Width = Me.gridComprobantesEmitidos.Width

End Sub


Private Sub Form_Terminate()
    Channel.RemoverSuscripcionTotal Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Channel.RemoverSuscripcionTotal Me
End Sub

Private Sub gridComprobantesEmitidos_BeforePrintPage(ByVal PageNumber As Long, ByVal nPages As Long)
    gridComprobantesEmitidos.PrinterProperties.FooterString(jgexHFRight) = "Página" & PageNumber & " de " & nPages
End Sub

Private Sub gridComprobantesEmitidos_ColumnHeaderClick(ByVal Column As GridEX20.JSColumn)
    GridEXHelper.ColumnHeaderClick Me.gridComprobantesEmitidos, Column
End Sub

Private Sub gridComprobantesEmitidos_DblClick()
    verFactura_Click
End Sub

Private Sub gridComprobantesEmitidos_FetchIcon(ByVal rowIndex As Long, ByVal ColIndex As Integer, ByVal RowBookmark As Variant, ByVal IconIndex As GridEX20.JSRetInteger)
    If ColIndex = 20 And m_Archivos.item(Factura.Id) > 0 Then IconIndex = 1
End Sub

Private Sub gridComprobantesEmitidos_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If facturas.count > 0 Then
        SeleccionarFactura
        If Button = 2 Then
            Me.NRO.caption = "[ Nro. " & Format(Factura.numero, "0000") & " ]"

            If Factura.Tipo.PuntoVenta.CaeManual Then
                Me.mnuEnviarAfip.caption = "Cargar CAE manualmente"
            Else
                Me.mnuEnviarAfip.caption = "Informar a AFIP"
            End If

'''            'Aplicar a Factura o ND...
'''            If Factura.TipoDocumento = tipoDocumentoContable.NotaCredito Then
'''                Me.aplicarNCaFC.caption = "Aplicar NC a Factura o ND..."
'''                Me.aplicarNCaFC.Enabled = True
'''                Me.aplicarNCaFC.Enabled = False
'''            End If
'''
'''            If Factura.TipoDocumento = tipoDocumentoContable.notaDebito Then
'''                Me.aplicarNDaFC.caption = "Aplicar ND a Factura o NC..."
'''                Me.aplicarNCaFC.Enabled = False
'''                Me.aplicarNCaFC.Enabled = True
'''            End If


            ' Si el estado del comprobante es EN PROCESO
            If Factura.estado = EstadoFacturaCliente.EnProceso Then   'no se aprob? localmente
                Me.aplicarNCaFC.Enabled = False
                Me.aplicarNCaFC.Visible = False
                Me.aplicarNDaFC.Visible = False
                Me.aplicarNDaFC.Enabled = False
                Me.editar.Enabled = True
                Me.editar.Visible = True
                Me.desAnular.Visible = False
                Me.AnularFactura.Visible = False
                Me.AnularFactura.Enabled = False
                Me.aprobarFactura.Enabled = Permisos.AdminFacturasAprobaciones
                Me.aprobarFactura.Visible = True
                Me.mnuEnviarAfip.Visible = False
                Me.ImprimirFactura.Enabled = False
                Me.ImprimirFactura.Visible = False
                Me.mnuDesaprobarFactura.Visible = False
                Me.aplicar.Enabled = False
                Me.aplicar.Visible = False
                Me.mnuFechaPagoPropuesta.Enabled = True
                Me.mnuFechaPagoPropuesta.Visible = True
                Me.mnuFechaEntrega.Enabled = True
                Me.mnuFechaEntrega.Visible = True

                Me.mnuEditarCampos.Visible = False
                Me.mnuEditarCampos.Enabled = False


                'opci?n combinada solo v?lida para comprobantes electr?nicos no aprobados localmente
                '23-08-2020
                If Factura.esCredito Then
                    If Factura.TipoDocumento = tipoDocumentoContable.Factura Then
                        Me.mnuAprobarEnviar.Visible = Factura.Tipo.PuntoVenta.EsElectronico And Permisos.AdminFacturasAprobaciones
                        Me.mnuAprobarEnviar.Enabled = Factura.Tipo.PuntoVenta.EsElectronico And Permisos.AdminFacturasAprobaciones
                    Else
                        Me.mnuAprobarEnviar.Visible = False
                        Me.mnuAprobarEnviar.Enabled = False
                    End If
                Else
                    Me.mnuAprobarEnviar.Visible = Factura.Tipo.PuntoVenta.EsElectronico And Permisos.AdminFacturasAprobaciones And Not Factura.Tipo.PuntoVenta.CaeManual
                    Me.mnuAprobarEnviar.Enabled = Factura.Tipo.PuntoVenta.EsElectronico And Permisos.AdminFacturasAprobaciones And Not Factura.Tipo.PuntoVenta.CaeManual
                End If

            End If


            ' Si el comprobante NO EST? EN PROCESO
            If Factura.estado <> EstadoFacturaCliente.EnProceso And Factura.estado <> EstadoFacturaCliente.Anulada Then     'se aprobo localmente y no est? anulada
                Me.editar.Enabled = False
                Me.editar.Visible = False
                Me.desAnular.Visible = False
                Me.aprobarFactura.Enabled = False
                Me.aprobarFactura.Visible = False

                Me.mnuFechaEntrega.Enabled = True
                Me.mnuFechaEntrega.Visible = True
                Me.mnuFechaPagoPropuesta.Enabled = True
                Me.mnuFechaPagoPropuesta.Visible = True

                'opci?n combinada solo v?lida para comprobantes electr?nicos no aprobados localmente
                '23-08-2020
                Me.mnuAprobarEnviar.Visible = False
                Me.mnuAprobarEnviar.Enabled = False

                Me.mnuEditarCampos.Visible = True
                Me.mnuEditarCampos.Enabled = True

                'Aplicar a Factura o ND...
                If Factura.TipoDocumento = tipoDocumentoContable.NotaCredito Then
                    Me.aplicarNCaFC.caption = "Aplicar NC a Factura o ND..."
                    Me.aplicarNCaFC.Enabled = True
                    Me.aplicarNCaFC.Visible = True
                    Me.aplicarNDaFC.Enabled = False
                End If

                If Factura.TipoDocumento = tipoDocumentoContable.notaDebito Then
                    Me.aplicarNDaFC.caption = "Aplicar ND a Factura o NC..."
                    Me.aplicarNCaFC.Enabled = False
                    Me.aplicarNDaFC.Enabled = True
                    Me.aplicarNDaFC.Visible = True
                End If
                
                If Factura.TipoDocumento = tipoDocumentoContable.Factura Then
                    Me.aplicarNCaFC.Visible = False
                    Me.aplicarNDaFC.Visible = False
                End If
                
'                Me.aplicarNCaFC.Visible = True
'                Me.aplicarNCaFC.Enabled = True
'
'                Me.aplicarNDaFC.Visible = True
'                Me.aplicarNDaFC.Enabled = True

                If Factura.Tipo.PuntoVenta.EsElectronico Then
                    Me.AnularFactura.Visible = False    'si es electronico no se puede anular comprobante
                    Me.AnularFactura.Enabled = False

                    If Factura.AprobadaAFIP Then
                        Me.mnuEnviarAfip.Enabled = False
                        Me.mnuEnviarAfip.Visible = False
                        Me.aplicar.Visible = False
                        Me.aplicar.Enabled = False

                    Else
                        Me.mnuEnviarAfip.Visible = True
                        Me.mnuEnviarAfip.Enabled = True
                        Me.aplicarNCaFC.Visible = (Factura.TipoDocumento = tipoDocumentoContable.NotaCredito Or Factura.TipoDocumento = tipoDocumentoContable.notaDebito) And (Factura.estado = EstadoFacturaCliente.Aprobada)
                        Me.aplicarNCaFC.Enabled = (Factura.TipoDocumento = tipoDocumentoContable.NotaCredito Or Factura.TipoDocumento = tipoDocumentoContable.notaDebito) And (Factura.estado = EstadoFacturaCliente.Aprobada)

                    End If

                Else
                    Me.mnuEnviarAfip.Enabled = False
                    Me.mnuEnviarAfip.Visible = False

                    Me.AnularFactura.Visible = True
                    Me.AnularFactura.Enabled = True
                    Me.mnuDesaprobarFactura.Visible = False

'                    Me.aplicarNCaFC.Enabled = (Factura.TipoDocumento = tipoDocumentoContable.NotaCredito Or Factura.TipoDocumento = tipoDocumentoContable.notaDebito) And (Factura.estado = EstadoFacturaCliente.Aprobada)

                End If

                Me.ImprimirFactura.Enabled = True
                Me.ImprimirFactura.Visible = True

                'si es FCE muestro el form para cambiar el estado de rechazo
                Me.mnuRechazo.Visible = Factura.esCredito

            End If

            If Factura.estado = EstadoFacturaCliente.Anulada Then
                Me.mnuFechaEntrega.Enabled = False
                Me.editar.Enabled = False
                Me.AnularFactura.Visible = False
                Me.aprobarFactura.Enabled = False
                Me.ImprimirFactura.Enabled = False
                Me.aplicar.Enabled = False
                Me.aplicarNCaFC.Enabled = False
                Me.aplicarNDaFC.Enabled = False

            End If


            If Factura.estado = EstadoFacturaCliente.CanceladaNC Then
                Me.editar.Enabled = False
                Me.AnularFactura.Enabled = False
                Me.AnularFactura.Visible = False
                Me.aprobarFactura.Enabled = False
                Me.ImprimirFactura.Enabled = True
                Me.aplicar.Enabled = False
                Me.aplicarNCaFC.Enabled = False
                Me.aplicarNDaFC.Enabled = False

            End If
            
            If Factura.estado = EstadoFacturaCliente.CanceladaNCParcial Then
                Me.editar.Enabled = False
                Me.AnularFactura.Enabled = False
                Me.AnularFactura.Visible = False
                Me.aprobarFactura.Enabled = False
                Me.ImprimirFactura.Enabled = True
                Me.aplicar.Enabled = False
                Me.aplicarNCaFC.Enabled = False
                Me.aplicarNDaFC.Enabled = False

            End If

            If Factura.estado = EstadoFacturaCliente.AplicadaND Then
                Me.editar.Enabled = False
                Me.AnularFactura.Enabled = False
                Me.AnularFactura.Visible = False
                Me.aprobarFactura.Enabled = False
                Me.ImprimirFactura.Enabled = True
                Me.aplicar.Enabled = False
                Me.aplicarNCaFC.Enabled = False
                Me.aplicarNDaFC.Enabled = False

            End If

            If Factura.estado = EstadoFacturaCliente.AplicadaACbte Then
                Me.editar.Enabled = False
                Me.AnularFactura.Enabled = False
                Me.AnularFactura.Visible = False
                Me.aprobarFactura.Enabled = False
                Me.ImprimirFactura.Enabled = True
                Me.aplicar.Enabled = False
                Me.aplicarNCaFC.Enabled = False
                Me.aplicarNDaFC.Enabled = False

            End If



            Me.archivos.Enabled = Permisos.SistemaArchivosVer
            Me.separador.Visible = Me.mnuEnviarAfip.Visible Or Me.aprobarFactura
            Me.sepa3.Visible = Me.mnuDesaprobarFactura.Visible Or Me.mnuAprobarEnviar.Visible

            If Factura.Saldado <> NoSaldada Then
                Me.mnuFechaEntrega.Enabled = False
                Me.mnuFechaEntrega.Visible = False
                Me.mnuFechaPagoPropuesta.Enabled = False
                Me.mnuFechaPagoPropuesta.Visible = False
                
                Me.MnuVerRecibo.Enabled = True
                Me.MnuVerRecibo.Visible = True
                
            
            End If

            Me.PopupMenu Me.mnuFacturas

        End If
    End If
End Sub

Private Sub gridComprobantesEmitidos_RowFormat(RowBuffer As GridEX20.JSRowData)
    On Error GoTo err1
    Set Factura = facturas.item(RowBuffer.rowIndex)

    If Factura.estado = EstadoFacturaCliente.Anulada Then
        RowBuffer.RowStyle = "anulada"
    Else
        If Factura.estado = EstadoFacturaCliente.EnProceso Then
            RowBuffer.CellStyle(12) = "pendiente"
        ElseIf Factura.estado = EstadoFacturaCliente.Aprobada Then
            RowBuffer.CellStyle(12) = "aprobada"
        End If

        '        If factura.Saldado = TipoSaldadoFactura.NoSaldada Or factura.Saldado = TipoSaldadoFactura.SaldadoParcial Or factura.Saldado = TipoSaldadoFactura.notaCredito Then
        '            If factura.EstaAtrasada Then
        '                RowBuffer.CellStyle(16) = "no_saldada"
        '            Else
        '                RowBuffer.CellStyle(16) = "no_vencida"
        '            End If
        '        ElseIf factura.Saldado = saldadoTotal Then
        '            RowBuffer.CellStyle(16) = "saldada"
        '        End If

        'Nemer agrega formato especial a la Celda
        If Factura.AprobadaAFIP = True Then
            RowBuffer.CellStyle(13) = "informadaAfip"
        Else
            RowBuffer.CellStyle(13) = "No_informadaAfip"
        End If

    End If
    Exit Sub
err1:

End Sub

Private Sub gridComprobantesEmitidos_SelectionChange()
    SeleccionarFactura
End Sub

Private Sub SeleccionarFactura()
    On Error Resume Next
    Set Factura = facturas.item(Me.gridComprobantesEmitidos.rowIndex(Me.gridComprobantesEmitidos.row))

End Sub

Private Sub gridComprobantesEmitidos_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    On Error GoTo err1
    Set Factura = facturas.item(rowIndex)


    Values(1) = Factura.GetShortDescription(True, False)    'enums.EnumTipoDocumentoContable(Factura.TipoDocumento)

    If IsSomething(Factura.Tipo) Then
        Values(2) = Factura.Tipo.TipoFactura.Tipo
    End If

    Values(3) = Factura.Tipo.PuntoVenta.PuntoVenta

    If Factura.esCredito Then
        Values(4) = "(FCE)"
    Else
        Values(4) = ""
    End If


    If Factura.Tipo.PuntoVenta.EsElectronico And Not Factura.AprobadaAFIP And Factura.estado <> EstadoFacturaCliente.EnProceso Then
        Values(5) = "Nro. Pendiente"
    Else
        Values(5) = Factura.NumeroFormateado
    End If

    If Factura.estado = EstadoFacturaCliente.EnProceso Then
        Values(5) = "Nro. Pendiente"
    End If

    Values(6) = Factura.FechaEmision

    'MONTO BASE
    Values(7) = Replace(FormatCurrency(funciones.FormatearDecimales(Factura.TotalEstatico.total)), "$", "")

    If Factura.moneda.Id = 0 Then
        Values(8) = Factura.moneda.NombreCorto
    Else
        Values(8) = Factura.moneda.NombreCorto & " " & Factura.CambioAPatron
    End If

    'MONTO TOTAL
    Values(9) = Replace(FormatCurrency(funciones.FormatearDecimales(Factura.TotalEstatico.total * Factura.CambioAPatron)), "$", "")

    Values(10) = Factura.OrdenCompra
    Values(11) = Factura.cliente.razon

    Values(12) = enums.EnumEstadoDocumentoContable(Factura.estado)

    If Factura.AprobadaAFIP = True Then
        Values(13) = "Informada"
    Else
        Values(13) = "No Informada"
    End If


    Values(14) = EnumTipoSaldadoFactura(Factura.Saldado)

    If Factura.Vencimiento < Factura.FechaEmision Then
        Values(15) = "No definido"
    Else
        Values(15) = Factura.Vencimiento
    End If

    Values(16) = Factura.StringDiasAtraso


    Values(17) = Factura.usuarioCreador.usuario


    'Values(18) = Factura.observaciones

    If Factura.Tipo.PuntoVenta.EsElectronico Or Factura.Tipo.PuntoVenta.CaeManual Then

        If Factura.estado = EstadoFacturaCliente.EnProceso Then
            Values(18) = "Comprobante en proceso"
        Else

            If LenB(Factura.CAE) <= 2 Then

                'Values(18) = "/ CAE no definido"

                Values(18) = "/ CAE no definido" & " / " & Factura.observaciones & " / " & Factura.observaciones_cancela

            Else

                Values(18) = "CAE: " & Factura.CAE & " / " & Factura.observaciones & " / " & Factura.observaciones_cancela
            End If
        End If
    End If

    If IsSomething(Factura.UsuarioAprobacion) Then
        Values(19) = Factura.UsuarioAprobacion.usuario
    Else
        Values(19) = vbNullString
    End If


    If CDbl(Factura.FechaPropuestaPago) > 0 Then Values(20) = Factura.FechaPropuestaPago

    If Factura.DiferenciaDiasEntrega = -1 Then
        Values(20) = "Defina fecha"
    Else

        If CDbl(Factura.FechaEntrega) > 0 And Factura.estado <> EstadoFacturaCliente.Anulada Then
            If Factura.Saldado = NoSaldada Then Values(21) = Format(Factura.FechaEntrega, "dd/mm/yyyy") & " (" & Factura.DiferenciaDiasEntrega & ")"

            '       If Factura.Saldado = SaldadoTotal Then
            '      Values(18) = "Saldada"
            '    Else
            '           Values(18) = "Anulada"
            '    End If
            '        Values(18) = Factura.FechaEntrega
            '    End If

        Else
            If Factura.estado = EstadoFacturaCliente.Anulada Then
                Values(21) = "Anulada"
            Else
                Values(21) = Factura.FechaEntrega
            End If
        End If

    End If


    Values(22) = Factura.TasaAjusteMensual

    Values(23) = "(" & Val(m_Archivos.item(Factura.Id)) & ")"

    Values(24) = Factura.Id

    Values(25) = Factura.RecibosAplicadosId
    
    Values(26) = Factura.idAsociacion
    
    verIds

    Exit Sub
err1:
End Sub

Private Sub ImprimirFactura_Click()

    On Error GoTo err451:
    Dim clasea As New classAdministracion
    Dim veces As Long


    If Factura.Tipo.PuntoVenta.EsElectronico Or Factura.Tipo.PuntoVenta.CaeManual Then
        veces = clasea.facturaImpresa(Factura.Id)
        If veces > 0 Then
            If MsgBox("Este comprobante ya fue generarlo" & Chr(10) & "¿Desea volver a generarlo?", vbYesNo, "Confirmación") = vbYes Then
                'DAOFactura.GenerarPdf (Factura.id)
                DAOFactura.VerFacturaElectronicaParaImpresion (Factura.Id)
            End If
        Else
            DAOFactura.VerFacturaElectronicaParaImpresion (Factura.Id)
        End If
    Else

        veces = clasea.facturaImpresa(Factura.Id)
        If veces = 0 Or veces = -1 Then
            If MsgBox("'¿Desea imprimir este comprobante?", vbYesNo, "Confirmación") = vbYes Then
                CD.Flags = cdlPDUseDevModeCopies
                CD.Copies = 3
                CD.ShowPrinter
                Dim i As Long
                For i = 1 To CD.Copies
                    DAOFactura.Imprimir Factura.Id
                Next
            End If

        ElseIf veces > 0 Then
            If MsgBox("Este comprobante ya fue impreso." & Chr(10) & "¿Desea volver a imprimirlo?", vbYesNo, "Confirmación") = vbYes Then
                CD.Flags = cdlPDUseDevModeCopies
                CD.Copies = 3
                CD.ShowPrinter

                For i = 1 To CD.Copies
                    DAOFactura.Imprimir Factura.Id
                Next i
            End If

        End If
    End If
    Exit Sub
err451:

End Sub

Private Property Get ISuscriber_id() As String
    ISuscriber_id = vId
End Property

Private Function ISuscriber_Notificarse(EVENTO As clsEventoObserver) As Variant
    Dim tmp As Factura
    If EVENTO.EVENTO = agregar_ Then
        llenarGrilla
        Me.gridComprobantesEmitidos.Refresh
    ElseIf EVENTO.EVENTO = modificar_ Then
        Set tmp = EVENTO.Elemento

        Dim i As Long
        For i = facturas.count To 1 Step -1

            If facturas(i).Id = tmp.Id Then

                '                Set Factura = facturas(i)
                '                Factura.Id = tmp.Id
                '                Factura.Detalles = tmp.Detalles
                '                Factura.estado = tmp.estado
                '                Factura.OrdenCompra = tmp.OrdenCompra
                '                Factura.estado = tmp.estado
                '                Factura.Observaciones = tmp.Observaciones
                '                Factura.TasaAjusteMensual = tmp.TasaAjusteMensual
                '                Set Factura.Cliente = tmp.Cliente

                facturas.remove i
                If facturas.count > 0 Then
                    If i = 1 Then    'ver esto cuand oes un solo item
                        facturas.Add tmp, CStr(tmp.Id), 1
                    ElseIf (i - 1) = facturas.count Then
                        facturas.Add tmp, CStr(tmp.Id), , i - 1
                    Else
                        facturas.Add tmp, CStr(tmp.Id), i
                    End If
                Else
                    facturas.Add tmp, CStr(tmp.Id)
                End If

                'DAOFactura.FindById(tmp.Id, True)

                Me.gridComprobantesEmitidos.RefreshRowIndex i
                Exit For

            End If

        Next

    End If


End Function

'Private Sub mnuAplicarANC_Click()
'  If MsgBox("?Seguro de aplicar a FC a NC?", vbYesNo, "Confirmaci?n") = vbYes Then
'        'seleccionar factura para aplicar
'        Set Selecciones.Factura = Nothing
'          Dim F As New frmAdminFacturasNCElegirFC
'
'        F.idCliente = Factura.cliente.id
'            F.TiposDocs.Add tipoDocumentoContable.notaCredito
'            F.EstadosDocs.Add EstadoFacturaCliente.Aprobada
'            F.Show 1
'
'        If IsSomething(Selecciones.Factura) Then
'            If DAOFactura.aplicarNCaFC(Factura.id, Selecciones.Factura.id) Then
'                MsgBox "Aplicaci?n existosa!", vbInformation, "Informaci?n"
'            Else
'                MsgBox "Se produjo un error, se abortan los cambios!", vbCritical, "Error"
'            End If
'        End If
'    End If
'End Sub



Private Sub mnuAprobarEnviar_Click()
    On Error GoTo err1
    Dim g As Long
    Dim msgadicional As String
    msgadicional = ""
    If MsgBox("¿Desea aprobar localmente el comprobante e informarlo a AFIP?", vbYesNo + vbQuestion, "Confirmacion") = vbYes Then
        g = Me.gridComprobantesEmitidos.rowIndex(Me.gridComprobantesEmitidos.row)
        If DAOFactura.aprobarV2(Factura, True, True) Then



            If Factura.Tipo.PuntoVenta.EsElectronico And Not Factura.Tipo.PuntoVenta.CaeManual And Not Factura.AprobadaAFIP Then
                msgadicional = "Esta factura deberá enviarse a la afip"
            End If
            If Factura.Tipo.PuntoVenta.EsElectronico And Factura.Tipo.PuntoVenta.CaeManual And Not Factura.AprobadaAFIP Then
                msgadicional = "Recuerde agregar al comprobante: CAE y fecha de vencimiento del CAE "
            End If

            Dim msg As String
            msg = "Comprobante aprobado con exito!"
            If IsSomething(Factura.CaeSolicitarResponse) Then
                If LenB(Factura.CaeSolicitarResponse.observaciones) > 5 Then

                    msg = msg & Chr(10) & Factura.CaeSolicitarResponse.observaciones
                End If

                If LenB(msgadicional) > 0 Then
                    msg = msg & Chr(10) & msgadicional
                End If

            End If
            MsgBox msg, vbInformation, "Información"

            Me.gridComprobantesEmitidos.RefreshRowIndex g
            Me.txtNroFactura.SetFocus
        Else
            GoTo err1
        End If
    End If
    Exit Sub
err1:
    'MsgBox "Factura no aprobada, compruebe:" & vbNewLine & "Si la factura es de anticipo, compruebe que el valor de la misma sea el mismo que el anticipo de la OT." & vbNewLine & "Que el detalle del remito no este ya facturado." & vbNewLine & Err.Description, vbCritical

    MsgBox Err.Description, vbCritical, Err.Source
    Me.gridComprobantesEmitidos.RefreshRowIndex g
End Sub

Private Sub mnuArchivos_Click()
    Dim archi As New frmArchivos2

    archi.Origen = OrigenArchivos.OA_factura
    archi.ObjetoId = Factura.Id
    archi.caption = Factura.GetShortDescription(False, True)
    archi.Show

End Sub

Private Sub mnuCrearCopiaFactura_Click()
    Me.taskDialog.Reset
    Me.taskDialog.MessageBoxStyle = True
    Me.taskDialog.WindowTitle = "Copia fiel de Comprobante"
    Me.taskDialog.MainInstructionText = "¿De que tipo es el nuevo comprobante?"
    Me.taskDialog.ContentText = "Elija el tipo de comprobante para el nuevo comprobante."
    taskDialog.RelativePosition = False

    Me.taskDialog.CommonButtons = 0
    taskDialog.CommonButtons = taskDialog.CommonButtons Or xtpTaskButtonOk
    taskDialog.CommonButtons = taskDialog.CommonButtons Or xtpTaskButtonCancel

    taskDialog.DefaultRadioButton = -1
    taskDialog.AddRadioButton "Factura", tipoDocumentoContable.Factura
    taskDialog.AddRadioButton "Nota de Débito", tipoDocumentoContable.notaDebito
    taskDialog.AddRadioButton "Nota de Crédito", tipoDocumentoContable.NotaCredito


    taskDialog.MainIcon = xtpTaskIconInformation

    If taskDialog.ShowDialog = xtpTaskButtonOk Then
        If Me.taskDialog.DefaultRadioButton = -1 Then
            MsgBox "Debe seleccionar un tipo para el nuevo comprobante.", vbExclamation + vbOKOnly
        Else
            Dim newFact As Factura
            Set newFact = DAOFactura.CrearCopiaFiel(Factura, Me.taskDialog.DefaultRadioButton)
            If IsSomething(newFact) Then
                MsgBox "Se creó un nuevo comprobante (" & newFact.GetShortDescription(False, True) & ")", vbInformation + vbOKOnly
            Else
                MsgBox "Hubo un error al copiar la factura.", vbCritical + vbOKOnly
            End If
        End If
    End If



End Sub

Private Sub mnuDesaprobarFactura_Click()

    On Error GoTo err1
    Dim g As Long

    If MsgBox("¿Desea desaprobar localmente el comprobante?", vbYesNo + vbQuestion, "Confirmacion") = vbYes Then
        g = Me.gridComprobantesEmitidos.rowIndex(Me.gridComprobantesEmitidos.row)
        If DAOFactura.desaprobar(Factura) Then
            MsgBox "Comprobante desaprobado con éxito!", vbInformation, "Información"
            Me.gridComprobantesEmitidos.RefreshRowIndex g
            Me.txtNroFactura.SetFocus
        Else
            GoTo err1
        End If
    End If
    Exit Sub
err1:
    MsgBox "Factura no aprobada, compruebe:" & vbNewLine & "Si la factura es de anticipo, compruebe que el valor de la misma sea el mismo que el anticipo de la OT." & vbNewLine & "Que el detalle del remito no este ya facturado." & vbNewLine & Err.Description, vbCritical
End Sub



Private Sub mnuEditarCampos_Click()
    Dim f_ADFE As New frmAdminFacturasEditarDatos
    f_ADFE.idFactura = Factura.Id
    f_ADFE.Show

End Sub

Private Sub mnuEnviarAfip_Click()
    On Error GoTo err1
    Dim g As Long

    If Not Factura.Tipo.PuntoVenta.EsElectronico Then
        Err.Raise 300, "Informar AFIP", "No puede informar un comprobante de un PV no catalogado como electrónico."
    End If

    If Factura.Tipo.PuntoVenta.EsElectronico And Factura.AprobadaAFIP Then
        Err.Raise 302, "Informar AFIP", "No puede informar un comprobante que ya fue informado."
    End If

    If Factura.Tipo.PuntoVenta.CaeManual Then

        Dim gg As Long
        gg = Me.gridComprobantesEmitidos.rowIndex(Me.gridComprobantesEmitidos.row)

        Dim F As New frmAdminFacturasAprobarSinAfip
        Set F.Factura = Factura
        F.Show 1

        Me.gridComprobantesEmitidos.RefreshRowIndex gg

    Else
        If MsgBox("¿Desea informar  el comprobante?", vbYesNo + vbQuestion, "Confirmacion") = vbYes Then
            g = Me.gridComprobantesEmitidos.rowIndex(Me.gridComprobantesEmitidos.row)
            If DAOFactura.aprobarV2(Factura, False, True) Then

                Dim msg As String
                msg = "Comprobante informado con éxito!"
                If IsSomething(Factura.CaeSolicitarResponse) Then
                    If LenB(Factura.CaeSolicitarResponse.observaciones) > 5 Then

                        msg = msg & Chr(10) & Factura.CaeSolicitarResponse.observaciones
                    End If
                End If
                MsgBox msg, vbInformation, "Información"

                Me.gridComprobantesEmitidos.RefreshRowIndex g
                Me.txtNroFactura.SetFocus
            Else
                GoTo err1
            End If
        End If
    End If


    Exit Sub
err1:

    MsgBox Err.Description, vbCritical, Err.Source
    Me.gridComprobantesEmitidos.RefreshRowIndex g

End Sub

Private Sub mnuFechaEntrega_Click()
    Dim fechaAnterior As String
    Dim fechaPosterior As String
    Dim nuevaFecha As Date
    Dim Update As Boolean

    On Error GoTo ErrorHandler ' Etiqueta de manejo de errores

    If CDbl(Factura.FechaEntrega) > 0 Then fechaAnterior = Factura.FechaEntrega

    fechaPosterior = InputBox("Establezca fecha de entrega", "Fecha de Entrega", fechaAnterior)

    If LenB(fechaPosterior) = 0 Then
        nuevaFecha = #1/1/2005#
        Update = True
    Else
        If IsDate(fechaPosterior) Then
            nuevaFecha = CDate(fechaPosterior)
            Update = True
        Else
            Err.Raise vbObjectError + 9999, "mnuFechaEntrega_Click", "La fecha no es válida." ' Lanza un error personalizado
        End If
    End If

    If Update Then
        Factura.FechaEntrega = nuevaFecha
        If DAOFactura.Guardar(Factura) Then
            Me.gridComprobantesEmitidos.RefreshRowIndex (Me.gridComprobantesEmitidos.row)
        Else
            Err.Raise vbObjectError + 9998, "mnuFechaEntrega_Click", "Error al guardar la factura." ' Lanza un error personalizado
        End If
    End If

    Exit Sub ' Salir del manejo de errores
ErrorHandler:
    MsgBox "Se ha producido un error: " & Err.Description, vbOKOnly + vbExclamation, "Error"
    Err.Clear ' Limpia el objeto de error
End Sub


Private Sub mnuFechaPagoPropuesta_Click()
    Dim fechaAnterior As String
    Dim fechaPosterior As String
    Dim nuevaFecha As Date
    Dim Update As Boolean

    If CDbl(Factura.FechaPropuestaPago) > 0 Then fechaAnterior = Factura.FechaPropuestaPago

    fechaPosterior = InputBox("Establezca fecha de pago propuesta", "Fecha de Pago", fechaAnterior)

    If LenB(fechaPosterior) = 0 Then
        Update = (MsgBox("¿Desea dejar en blanco la fecha de pago propuesta?", vbYesNo + vbQuestion) = vbYes)
    Else
        If IsDate(fechaPosterior) Then
            nuevaFecha = CDate(fechaPosterior)
            Update = True
        Else
            MsgBox "La fecha no es válida.", vbOKOnly + vbExclamation, "Fecha"
        End If
    End If

    If Update Then
        Factura.FechaPropuestaPago = nuevaFecha
        If DAOFactura.Guardar(Factura) Then
            Me.gridComprobantesEmitidos.ReBind
        Else
            MsgBox "Error al guardar la factura.", vbOKOnly + vbCritical, "Error"
        End If
    End If

End Sub

Private Sub mnuRechazo_Click()
    Dim F As New frmAdminFacturaRechazoAfip
    Set F.Factura = Factura
    F.Show
End Sub


Private Sub MnuVerRecibo_Click()
    Dim Fa As New frmAdminVentasListaRECSegunCbte
    Set Fa.vFactura = Factura
    Fa.Show
End Sub


Private Sub PushButton1_Click()
    Me.cboClientes.ListIndex = -1
End Sub



Private Sub PushButton2_Click()
    Me.cboOrdenImporte.ListIndex = -1
End Sub

Private Sub PushButton3_Click()
    Me.cboPuntosVenta.ListIndex = -1
End Sub

Private Sub PushButton4_Click()
    Me.cboEstados.ListIndex = -1
End Sub



Private Sub PushButton5_Click()
    Me.cboEstadosSaldada.ListIndex = -1
End Sub

Private Sub scanear_Click()
    On Error Resume Next
    Dim archivos As New classArchivos
    If archivos.escanearDocumento(OrigenArchivos.OA_factura, Factura.Id) Then
        Set m_Archivos = DAOArchivo.GetCantidadArchivosPorReferencia(OA_factura)
        Me.gridComprobantesEmitidos.RefreshRowIndex (Factura.Id)
    End If
End Sub

'Private Sub txtOrdenCompra_GotFocus()
'    foco Me.txtOrdenCompra
'End Sub


Private Sub verFactura_Click()
    Dim f_c3h3 As New frmAdminFacturasEdicion
    f_c3h3.ReadOnly = True
    f_c3h3.idFactura = Factura.Id
    f_c3h3.Show

End Sub

Private Sub verHistorialFactura_Click()
    Set Factura.Historial = DAOFacturaHistorial.getAllByIdFactura(Factura.Id)
    frmHistoriales.lista = Factura.Historial
    frmHistoriales.Show
End Sub
