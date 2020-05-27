VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~3.OCX"
Begin VB.Form frmFacturaEdicion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Factura"
   ClientHeight    =   10470
   ClientLeft      =   3945
   ClientTop       =   2385
   ClientWidth     =   11670
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFacturaEdicion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10470
   ScaleWidth      =   11670
   Begin XtremeSuiteControls.GroupBox grpTotales 
      Height          =   1575
      Left            =   9000
      TabIndex        =   42
      Top             =   8760
      Width           =   2580
      _Version        =   786432
      _ExtentX        =   4551
      _ExtentY        =   2778
      _StockProps     =   79
      Caption         =   "Totales"
      UseVisualStyle  =   -1  'True
      Begin VB.Label lblIVATot 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "999.999.999"
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
         Left            =   1380
         TabIndex        =   50
         Top             =   945
         Width           =   1080
      End
      Begin VB.Label lblPercepciones 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "999.999.999"
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
         Left            =   1380
         TabIndex        =   49
         Top             =   690
         Width           =   1080
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Percepciones "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   165
         TabIndex        =   48
         Top             =   690
         Width           =   1020
      End
      Begin VB.Label lblSubTotal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "999.999.999"
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
         Left            =   1380
         TabIndex        =   47
         Top             =   435
         Width           =   1080
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Subtotal "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   555
         TabIndex        =   46
         Top             =   435
         Width           =   630
      End
      Begin VB.Label lblIva2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "IVA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   900
         TabIndex        =   45
         Top             =   945
         Width           =   255
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Total "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   780
         TabIndex        =   44
         Top             =   1170
         Width           =   405
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "999.999.999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1380
         TabIndex        =   43
         Top             =   1170
         Width           =   1080
      End
   End
   Begin XtremeSuiteControls.PushButton btnItemRemito 
      Height          =   360
      Left            =   120
      TabIndex        =   10
      Top             =   9000
      Width           =   2055
      _Version        =   786432
      _ExtentX        =   3625
      _ExtentY        =   635
      _StockProps     =   79
      Caption         =   "Item de Remito..."
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton lblBuscandoPercepcion 
      Height          =   285
      Left            =   120
      TabIndex        =   41
      Top             =   0
      Width           =   11415
      _Version        =   786432
      _ExtentX        =   20135
      _ExtentY        =   503
      _StockProps     =   79
      Caption         =   "Buscando Percepcion..."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
   End
   Begin XtremeSuiteControls.GroupBox grpDatos 
      Height          =   5535
      Left            =   120
      TabIndex        =   16
      Top             =   360
      Width           =   11475
      _Version        =   786432
      _ExtentX        =   20241
      _ExtentY        =   9763
      _StockProps     =   79
      Caption         =   "Datos"
      UseVisualStyle  =   -1  'True
      Begin VB.ComboBox txtCondObs 
         Height          =   315
         ItemData        =   "frmFacturaEdicion.frx":000C
         Left            =   1040
         List            =   "frmFacturaEdicion.frx":001C
         TabIndex        =   68
         Top             =   3550
         Width           =   3255
      End
      Begin VB.Frame frmFC 
         Caption         =   "Factura de Crédito"
         Enabled         =   0   'False
         Height          =   1455
         Left            =   120
         TabIndex        =   61
         Top             =   3960
         Visible         =   0   'False
         Width           =   11260
         Begin VB.ComboBox cboConceptosAIncluir 
            Height          =   315
            ItemData        =   "frmFacturaEdicion.frx":007C
            Left            =   5450
            List            =   "frmFacturaEdicion.frx":0089
            TabIndex        =   73
            Top             =   360
            Width           =   3015
         End
         Begin MSComCtl2.DTPicker dtFechaPagoCreditoHasta 
            Height          =   315
            Left            =   2280
            TabIndex        =   70
            Top             =   1050
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   16449537
            CurrentDate     =   43967
         End
         Begin MSComCtl2.DTPicker dtFechaPagoCreditoDesde 
            Height          =   315
            Left            =   2280
            TabIndex        =   69
            Top             =   690
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   16449537
            CurrentDate     =   43967
         End
         Begin VB.ComboBox cboCuentasCBU 
            Height          =   315
            ItemData        =   "frmFacturaEdicion.frx":00BA
            Left            =   4680
            List            =   "frmFacturaEdicion.frx":00BC
            Style           =   2  'Dropdown List
            TabIndex        =   66
            Top             =   840
            Visible         =   0   'False
            Width           =   6210
         End
         Begin MSComCtl2.DTPicker dtFechaPagoCredito 
            Height          =   315
            Left            =   2280
            TabIndex        =   63
            Top             =   330
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   16449537
            CurrentDate     =   43967
         End
         Begin VB.Line Line5 
            X1              =   3720
            X2              =   3720
            Y1              =   1200
            Y2              =   840
         End
         Begin VB.Line Line4 
            X1              =   3600
            X2              =   3720
            Y1              =   840
            Y2              =   840
         End
         Begin VB.Line Line3 
            X1              =   3600
            X2              =   3720
            Y1              =   1200
            Y2              =   1200
         End
         Begin VB.Label lblConceptosAIncluir 
            Caption         =   "Conceptos a incluir:"
            Height          =   255
            Left            =   3960
            TabIndex        =   74
            Top             =   420
            Width           =   1455
         End
         Begin XtremeSuiteControls.Label lblFechaPagoCreditoHasta 
            Height          =   255
            Left            =   1740
            TabIndex        =   72
            Top             =   1080
            Width           =   615
            _Version        =   786432
            _ExtentX        =   1085
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Hasta:"
         End
         Begin XtremeSuiteControls.Label lblFechaPagoCreditoDesde 
            Height          =   255
            Left            =   360
            TabIndex        =   71
            Top             =   720
            Width           =   1935
            _Version        =   786432
            _ExtentX        =   3413
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Periodo Facturado desde:"
         End
         Begin VB.Label LblCBU 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "CBU:"
            Height          =   195
            Left            =   4245
            TabIndex        =   67
            Top             =   900
            Width           =   360
         End
         Begin VB.Label lblVerCbu 
            Caption         =   "NO INFORMADO"
            Height          =   195
            Left            =   4920
            TabIndex        =   65
            Top             =   900
            Width           =   2415
         End
         Begin VB.Label lblFechaPagoCredito 
            Caption         =   "Fecha de Vto. para el Pago:"
            Height          =   255
            Left            =   240
            TabIndex        =   64
            Top             =   360
            Width           =   2055
         End
         Begin VB.Label lblCbuCredito 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "CBU:"
            Height          =   195
            Left            =   240
            TabIndex        =   62
            Top             =   1845
            Visible         =   0   'False
            Width           =   360
         End
      End
      Begin VB.TextBox txtTasaAjuste 
         Height          =   300
         Left            =   5640
         TabIndex        =   6
         Top             =   3150
         Width           =   840
      End
      Begin VB.TextBox txtDiasVenc 
         Height          =   300
         Left            =   2205
         TabIndex        =   5
         Top             =   3150
         Width           =   840
      End
      Begin VB.TextBox txtReferencia 
         Height          =   300
         Left            =   1040
         TabIndex        =   4
         Top             =   2750
         Width           =   5415
      End
      Begin XtremeSuiteControls.GroupBox grpPercep 
         Height          =   1215
         Left            =   6840
         TabIndex        =   33
         Top             =   2760
         Width           =   4545
         _Version        =   786432
         _ExtentX        =   8017
         _ExtentY        =   2143
         _StockProps     =   79
         Caption         =   "Percepciones IIBB"
         UseVisualStyle  =   -1  'True
         Begin VB.TextBox txtPercepcion 
            Height          =   300
            Left            =   1530
            TabIndex        =   8
            Top             =   780
            Width           =   2715
         End
         Begin XtremeSuiteControls.ComboBox cboPadron 
            Height          =   315
            Left            =   1530
            TabIndex        =   7
            Top             =   360
            Width           =   2715
            _Version        =   786432
            _ExtentX        =   4789
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Style           =   2
            Appearance      =   6
            Text            =   "cboMoneda"
            DropDownItemCount=   3
         End
         Begin XtremeSuiteControls.Label Label17 
            Height          =   195
            Left            =   600
            TabIndex        =   36
            Top             =   840
            Width           =   840
            _Version        =   786432
            _ExtentX        =   1482
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Percepcion:"
            AutoSize        =   -1  'True
         End
         Begin XtremeSuiteControls.Label lblPadron 
            Height          =   195
            Left            =   240
            TabIndex        =   35
            Top             =   420
            Width           =   1215
            _Version        =   786432
            _ExtentX        =   2143
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Padron a utilizar:"
            AutoSize        =   -1  'True
         End
         Begin VB.Label lblVencido 
            Alignment       =   2  'Center
            BackColor       =   &H000000EE&
            Caption         =   "PADRON VENCIDO"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   1560
            TabIndex        =   34
            Top             =   0
            Visible         =   0   'False
            Width           =   2670
         End
      End
      Begin VB.TextBox txtNumero 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   7995
         TabIndex        =   1
         Text            =   "999999"
         Top             =   765
         Width           =   3090
      End
      Begin XtremeSuiteControls.ComboBox cboCliente 
         Height          =   315
         Left            =   1095
         TabIndex        =   0
         Top             =   255
         Width           =   4515
         _Version        =   786432
         _ExtentX        =   7964
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Text            =   "cboCliente"
      End
      Begin XtremeSuiteControls.DateTimePicker dtpFecha 
         Height          =   405
         Left            =   8040
         TabIndex        =   2
         Top             =   1320
         Width           =   3090
         _Version        =   786432
         _ExtentX        =   5450
         _ExtentY        =   714
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   1
         CurrentDate     =   40234.4177546296
      End
      Begin XtremeSuiteControls.ComboBox cboMoneda 
         Height          =   405
         Left            =   8010
         TabIndex        =   3
         Top             =   1875
         Width           =   3090
         _Version        =   786432
         _ExtentX        =   5450
         _ExtentY        =   714
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
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
      Begin XtremeSuiteControls.ComboBox cboTiposFactura 
         Height          =   360
         Left            =   8040
         TabIndex        =   56
         Top             =   240
         Width           =   3090
         _Version        =   786432
         _ExtentX        =   5450
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
      Begin VB.Label lblEsCredito 
         Caption         =   "FACTURA DE CRÉDITO ELECTRÓNICA MiPyMES (FCE)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7320
         TabIndex        =   60
         Top             =   2385
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Punto:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7260
         TabIndex        =   57
         Top             =   293
         Width           =   705
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Provincia:"
         Height          =   195
         Left            =   300
         TabIndex        =   53
         Top             =   1840
         Width           =   705
      End
      Begin VB.Label lblProvincia 
         AutoSize        =   -1  'True
         Caption         =   "2343242"
         Height          =   195
         Left            =   1080
         TabIndex        =   52
         Top             =   1840
         Width           =   630
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "% Tasa ajuste mensual:"
         Height          =   195
         Left            =   3840
         TabIndex        =   51
         Top             =   3195
         Width           =   1740
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Condicion:"
         Height          =   195
         Left            =   210
         TabIndex        =   39
         Top             =   3600
         Width           =   750
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cant Días Vencimiento FF:"
         Height          =   195
         Left            =   240
         TabIndex        =   38
         Top             =   3195
         Width           =   1875
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Referencia:"
         Height          =   195
         Left            =   120
         TabIndex        =   37
         Top             =   2805
         Width           =   840
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFDBBF&
         DrawMode        =   9  'Not Mask Pen
         X1              =   11280
         X2              =   0
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Moneda:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7050
         TabIndex        =   32
         Top             =   1935
         Width           =   915
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7275
         TabIndex        =   31
         Top             =   1380
         Width           =   690
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Numero:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7020
         TabIndex        =   30
         Top             =   825
         Width           =   945
      End
      Begin VB.Label lblNCND 
         AutoSize        =   -1  'True
         Caption         =   "N/D"
         Height          =   195
         Left            =   6105
         TabIndex        =   29
         Top             =   960
         Width           =   270
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFDBBF&
         DrawMode        =   9  'Not Mask Pen
         X1              =   6720
         X2              =   6720
         Y1              =   240
         Y2              =   3960
      End
      Begin VB.Label lblTipoFactura 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   24.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   5895
         TabIndex        =   28
         Top             =   270
         Width           =   645
      End
      Begin VB.Line Line 
         BorderColor     =   &H00FFDBBF&
         DrawMode        =   9  'Not Mask Pen
         X1              =   5715
         X2              =   5715
         Y1              =   240
         Y2              =   2640
      End
      Begin VB.Label lblCodPostal 
         AutoSize        =   -1  'True
         Caption         =   "2343242"
         Height          =   195
         Left            =   1080
         TabIndex        =   27
         Top             =   2145
         Width           =   630
      End
      Begin VB.Label lblLocalidad 
         AutoSize        =   -1  'True
         Caption         =   "HHHHHH"
         Height          =   195
         Left            =   1080
         TabIndex        =   26
         Top             =   1530
         Width           =   630
      End
      Begin VB.Label lblDireccion 
         AutoSize        =   -1  'True
         Caption         =   "RIVAD 3242"
         Height          =   195
         Left            =   1080
         TabIndex        =   25
         Top             =   1230
         Width           =   870
      End
      Begin VB.Label lblIVA 
         AutoSize        =   -1  'True
         Caption         =   "23"
         Height          =   195
         Left            =   1080
         TabIndex        =   24
         Top             =   930
         Width           =   180
      End
      Begin VB.Label lblCuit 
         AutoSize        =   -1  'True
         Caption         =   "23-30279550-9"
         Height          =   195
         Left            =   1080
         TabIndex        =   23
         Top             =   630
         Width           =   1110
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cod Postal:"
         Height          =   195
         Left            =   180
         TabIndex        =   22
         Top             =   2145
         Width           =   825
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Localidad:"
         Height          =   195
         Left            =   285
         TabIndex        =   21
         Top             =   1530
         Width           =   720
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "CUIT:"
         Height          =   195
         Left            =   585
         TabIndex        =   20
         Top             =   630
         Width           =   420
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Direccion:"
         Height          =   195
         Left            =   300
         TabIndex        =   19
         Top             =   1230
         Width           =   705
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "IVA:"
         Height          =   195
         Left            =   690
         TabIndex        =   18
         Top             =   930
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   450
         TabIndex        =   17
         Top             =   300
         Width           =   555
      End
   End
   Begin XtremeSuiteControls.GroupBox grpDetalles 
      Height          =   2820
      Left            =   120
      TabIndex        =   40
      Top             =   5880
      Width           =   11475
      _Version        =   786432
      _ExtentX        =   20241
      _ExtentY        =   4974
      _StockProps     =   79
      Caption         =   "Detalles (Cant: 0)"
      UseVisualStyle  =   -1  'True
      Begin GridEX20.GridEX gridDetalles 
         Height          =   2460
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   11250
         _ExtentX        =   19844
         _ExtentY        =   4339
         Version         =   "2.0"
         PreviewRowIndent=   300
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         PreviewColumn   =   "origen"
         PreviewRowLines =   1
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
         ColumnsCount    =   10
         Column(1)       =   "frmFacturaEdicion.frx":00BE
         Column(2)       =   "frmFacturaEdicion.frx":01F6
         Column(3)       =   "frmFacturaEdicion.frx":02EA
         Column(4)       =   "frmFacturaEdicion.frx":0402
         Column(5)       =   "frmFacturaEdicion.frx":0522
         Column(6)       =   "frmFacturaEdicion.frx":0662
         Column(7)       =   "frmFacturaEdicion.frx":0796
         Column(8)       =   "frmFacturaEdicion.frx":08BE
         Column(9)       =   "frmFacturaEdicion.frx":09EE
         Column(10)      =   "frmFacturaEdicion.frx":0AFE
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmFacturaEdicion.frx":0BF6
         FormatStyle(2)  =   "frmFacturaEdicion.frx":0D1E
         FormatStyle(3)  =   "frmFacturaEdicion.frx":0DCE
         FormatStyle(4)  =   "frmFacturaEdicion.frx":0E82
         FormatStyle(5)  =   "frmFacturaEdicion.frx":0F5A
         FormatStyle(6)  =   "frmFacturaEdicion.frx":1012
         ImageCount      =   0
         PrinterProperties=   "frmFacturaEdicion.frx":10F2
      End
   End
   Begin XtremeSuiteControls.PushButton btnGuardar 
      Height          =   360
      Left            =   6840
      TabIndex        =   15
      Top             =   9960
      Width           =   2055
      _Version        =   786432
      _ExtentX        =   3625
      _ExtentY        =   635
      _StockProps     =   79
      Caption         =   "Guardar"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton PushButton1 
      Height          =   360
      Left            =   120
      TabIndex        =   11
      Top             =   9480
      Width           =   2055
      _Version        =   786432
      _ExtentX        =   3625
      _ExtentY        =   635
      _StockProps     =   79
      Caption         =   "Seleccionar OT..."
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton btnItemsDescuentoAnticipo 
      Height          =   360
      Left            =   6840
      TabIndex        =   13
      Top             =   9000
      Width           =   2055
      _Version        =   786432
      _ExtentX        =   3625
      _ExtentY        =   635
      _StockProps     =   79
      Caption         =   "Generar Items Anticipo OT"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton PushButton2 
      Height          =   360
      Left            =   120
      TabIndex        =   12
      Top             =   9960
      Width           =   2055
      _Version        =   786432
      _ExtentX        =   3625
      _ExtentY        =   635
      _StockProps     =   79
      Caption         =   "Crear Item de concepto..."
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton cmdNueva 
      Height          =   360
      Left            =   6840
      TabIndex        =   14
      Top             =   9480
      Width           =   2055
      _Version        =   786432
      _ExtentX        =   3625
      _ExtentY        =   635
      _StockProps     =   79
      Caption         =   "Nueva"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.ComboBox cboMonedaAjuste 
      Height          =   405
      Left            =   2760
      TabIndex        =   54
      Top             =   9480
      Width           =   2550
      _Version        =   786432
      _ExtentX        =   4498
      _ExtentY        =   714
      _StockProps     =   77
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
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
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "23-30279550-9"
      Height          =   195
      Left            =   480
      TabIndex        =   59
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label txtDetallesCAE 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   135
      TabIndex        =   58
      Top             =   10440
      Width           =   5385
   End
   Begin VB.Label lblAjuste 
      AutoSize        =   -1  'True
      Caption         =   "Ajuste a"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2760
      TabIndex        =   55
      Top             =   9000
      Width           =   870
   End
   Begin VB.Menu mnuDetalles 
      Caption         =   "mnuDetalles"
      Visible         =   0   'False
      Begin VB.Menu mnuAplicarDetalleRemito 
         Caption         =   "Aplicar detalle de remito"
      End
   End
End
Attribute VB_Name = "frmFacturaEdicion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tipos As Collection
Dim Tipo As clsTipoFactura
Implements ISuscriber
Private Factura As Factura
Private dataLoading As Boolean
Private detalle As FacturaDetalle
Private suscId As String

Public NuevoTipoDocumento As tipoDocumentoContable
Public EsAnticipo As Boolean

Public ReadOnly As Boolean

Private detaFactRemito As FacturaDetalle

Public Property Let idFactura(value As Long)
    Set Factura = DAOFactura.FindById(value, True, True)
End Property

Private Sub btnGuardar_Click()
On Error GoTo err1
    If Me.gridDetalles.EditMode = jgexEditModeOn Then
        MsgBox "Todavia esta editando algun detalle de la factura.", vbExclamation
        Exit Sub
    End If

    If Not Factura.cliente.CUITValido Or Not Factura.cliente.ValidoRemitoFactura Then
        MsgBox "El cliente no es valido para poder facturar.", vbExclamation + vbOKOnly
        Exit Sub
    End If

    If LenB(Factura.numero) = 0 Or _
       LenB(Factura.OrdenCompra) = 0 Or _
       Factura.CantDiasPago = 0 Then

        MsgBox "La factura debe poseer Nº, referencia y dias de venc de forma de pago.", vbExclamation + vbOKOnly
    Else

        If EsAnticipo And Factura.OTsFacturadasAnticipo.count = 0 Then
            MsgBox "Se produjo un error en la asociación del anticipo.", vbExclamation + vbOKOnly
            Exit Sub
        End If



        If Me.cboMonedaAjuste.ListIndex = -1 Then
            Factura.IdMonedaAjuste = 0
            Factura.TipoCambioAjuste = 1

        Else
            Dim mon As clsMoneda
            Set mon = DAOMoneda.GetById(Me.cboMonedaAjuste.ItemData(Me.cboMonedaAjuste.ListIndex))

            Factura.IdMonedaAjuste = mon.id
            Factura.TipoCambioAjuste = mon.Cambio

        End If



        If Factura.EsAnticipo Or EsAnticipo Then
            'If Factura.DetallesMismaOT > 0 Then


            Dim Ot As OrdenTrabajo
            For Each Ot In Factura.OTsFacturadasAnticipo
                If Ot.Anticipo > 0 And Not Ot.AnticipoFacturado And Not Factura.AnticipoDescontado Then
                    MsgBox "Deberá tener un item que sea por descuento de anticipo. Por favor rehaga la factura!", vbCritical + vbOKOnly, "Error"
                    Exit Sub
                End If
            Next Ot

            'End If
        End If

        Dim deta As FacturaDetalle
        'Dim ot As OrdenTrabajo
        For Each deta In Factura.Detalles
            If IsSomething(deta.detalleRemito) Then
                If deta.detalleRemito.idpedido <> 0 Then
                    If deta.OtIdAnticipo = 0 Then
                        If IsSomething(deta.detalleRemito) Then



                            Set Ot = DAOOrdenTrabajo.FindById(deta.detalleRemito.idpedido)

                            If Ot.EsHija Then
                                Dim id_marco As Long
                                id_marco = Ot.OTMarcoIdPadre
                                Set Ot = DAOOrdenTrabajo.FindById(id_marco)
                            End If
                            If IsSomething(Ot) Then
                                If Ot.Anticipo > 0 Then
                                    If Not IsSomething(Factura.DetalleAnticipoOT(Ot.id)) Then
                                        MsgBox "No existe en la factura un detalle que certifique el descuento por anticipo de OT Nº " & Ot.IdFormateado & vbNewLine & "Haga click en el boton ""Generar Items Anticipo OT"" para generar el detalle de anticipo para la factura actual.", vbExclamation
                                        Exit Sub
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next deta
        
        Factura.CBU = Me.cboCuentasCBU.text
        Dim c As CuentaBancaria
        
        Set c = DAOCuentaBancaria.FindById(Me.cboCuentasCBU.ItemData(Me.cboCuentasCBU.ListIndex))
        
        If IsSomething(c) Then
          Factura.CBU = c.CBU
        End If
        
        If DAOFactura.Save(Factura, True) Then
            MsgBox "La " & StrConv(Factura.TipoDocumentoDescription, vbProperCase) & " ha sido guardada.", vbOKOnly + vbInformation
        Else
           Err.Raise "9999", "Guardando factura", Err.Description
        End If
    End If
Exit Sub
err1:
       MsgBox "Ocurrió un error al guardar. Controle datos y que el Nº no este ya asignado. o Bien" & Chr(10) & Err.Description, vbCritical + vbOKOnly
End Sub

Private Sub btnItemRemito_Click()
    If IsSomething(Factura.cliente) Then
        Dim idEntrega As Long
        Dim f11 As New frmPlaneamientoRemitosListaProceso
        f11.idCliMostrar = Factura.cliente.id
        f11.mostrar = 2
     
        Set Selecciones.RemitoElegido = Nothing
        f11.Show 1
        If IsSomething(Selecciones.RemitoElegido) Then
            Dim frm As New frmPlaneamientoRemitoVer
            Set frm.Remito = Selecciones.RemitoElegido    'DAORemitoS.FindById(Selecciones.RemitoElegido.Id)
               frm.IdFormSuscriber = suscId
            Me.txtReferencia = Selecciones.RemitoElegido.detalle
            frm.ParaFacturar = True
            frm.grilla.MultiSelect = True
            frm.MostrarInfoAdministracion = True
            frm.Show
            frm.grilla.Columns(6).Visible = True
        End If
    Else
        MsgBox "Debe seleccionar un cliente para poder agregar un item de remito.", vbInformation + vbOKOnly
    End If

End Sub

Private Sub AgregarEntregas(col As Collection)
    'Dim tmp As Variant
    Dim redeta As remitoDetalle
    Dim Ot As OrdenTrabajo
    For Each redeta In col
        'Set redeta = DAORemitoSDetalle.FindById(CLng(tmp))
        If IsSomething(redeta) Then
            For Each detalle In Factura.Detalles
                If detalle.DetalleRemitoId = redeta.id Then
                    GoTo prox    'ya existe en la factura esa entrega aplicada, pasamos a la proxima
                End If
            Next

            Set detalle = New FacturaDetalle
            Set detalle.Factura = Factura
            detalle.idFactura = Factura.id
            detalle.Cantidad = redeta.Cantidad
            detalle.detalle = redeta.VerElemento
            detalle.PorcentajeDescuento = 0
            detalle.Bruto = redeta.Valor
            Set Ot = DAOOrdenTrabajo.FindById(redeta.idpedido)
            If IsSomething(Ot) Then
                detalle.Bruto = MonedaConverter.Convertir(redeta.Valor, Ot.moneda.id, Factura.moneda.id)
            End If
            detalle.IvaAplicado = True
            detalle.IBAplicado = True

            detalle.AplicadoARemito = True
            Set detalle.detalleRemito = redeta
            detalle.DetalleRemitoId = redeta.id

            Factura.Detalles.Add detalle


        End If
prox:
    Next redeta

    CargarDetalles
    Totalizar
End Sub

Private Sub btnItemsDescuentoAnticipo_Click()
    Dim detalle As FacturaDetalle
    Dim detalleAnticipo As FacturaDetalle
    Dim Ot As OrdenTrabajo
    Dim facturaAnticipo As Factura

    Factura.RemoveDetallesAnticipoOT

    For Each detalle In Factura.Detalles
        If detalle.OtIdAnticipo = 0 Then    'no es de descuento de anticipo de ot
            If Not detalle.OrigenEsConcepto Then
                If IsSomething(detalle.detalleRemito) Then


                    Set Ot = DAOOrdenTrabajo.FindById(detalle.detalleRemito.idpedido)

                    If Ot.EsHija Then


                        Set Ot = DAOOrdenTrabajo.FindById(Ot.OTMarcoIdPadre)
                    End If


                    If IsSomething(Ot) Then
                        If Ot.Anticipo > 0 Then
                            Set detalleAnticipo = Nothing
                            Set detalleAnticipo = Factura.DetalleAnticipoOT(Ot.id)
                            If Not IsSomething(detalleAnticipo) Then
                                Set detalleAnticipo = New FacturaDetalle
                                detalleAnticipo.OtIdAnticipo = Ot.id
                                Factura.Detalles.Add detalleAnticipo
                                detalleAnticipo.PorcentajeDescuento = Ot.Anticipo



                                detalleAnticipo.IvaAplicado = Factura.EstaDiscriminada      'True


                                detalleAnticipo.IBAplicado = True
                                detalleAnticipo.Cantidad = -1



                                Set detalleAnticipo.Factura = Factura



                                Set facturaAnticipo = DAOFactura.FindById(Ot.AnticipoFacturadoIdFactura)




                                If IsSomething(facturaAnticipo) Then
                                    detalleAnticipo.detalle = "ANTICIPO SEGÚN " & facturaAnticipo.GetShortDescription(False, True) & " de OT Nº " & Ot.IdFormateado
                                Else
                                    'no hay factura asociada, habria que seleccionar una factura, asociarla y volver a realizar el proceso
                                    MsgBox "No hay factura de anticipo asociada a la OT Nº " & Ot.IdFormateado & "." & vbNewLine & "Realice la asociacion desde el listado de OT (click derecho).", vbExclamation
                                    Exit Sub
                                End If
                            End If
                            detalleAnticipo.Bruto = detalleAnticipo.Bruto + funciones.RedondearDecimales(detalle.Total * Factura.moneda.Cambio)
                        End If
                    End If
                End If
            End If
        End If
    Next

    CargarDetalles
    Totalizar

End Sub

Private Sub cboCliente_Click()
    If IsSomething(Factura) And Me.cboCliente.ListIndex <> -1 And Not dataLoading Then
        Set Factura.cliente = DAOCliente.BuscarPorID(Me.cboCliente.ItemData(Me.cboCliente.ListIndex))
        Factura.Detalles = New Collection

        Set Factura.TipoIVA = Factura.cliente.TipoIVA

        Dim tipos As New Collection

        Set tipos = DAOTipoFacturaDiscriminado.FindAllByFilter("id_iva= " & Factura.TipoIVA.idIVA & " and tipo_documento=" & Factura.TipoDocumento)
        Dim Tipo As clsTipoFacturaDiscriminado
        Me.cboTiposFactura.Clear


Dim id_Default As Long
id_Default = 0
Dim nidx As Long
        'lleno el combo de tipos de factura y dejo el default marcado
        For Each Tipo In tipos
            cboTiposFactura.AddItem Tipo.descripcion
            nidx = cboTiposFactura.NewIndex
            cboTiposFactura.ItemData(nidx) = Tipo.id
            If Tipo.PuntoVenta.default Then id_Default = nidx
            
        Next Tipo


    'pos on default pv
    If cboTiposFactura.ListCount > 0 Then
   'cboTiposFactura.ListIndex = id_Default
   cboTiposFactura.ListIndex = id_Default
            
    End If




        'esto hay que ponerlo en onclick del cbotipos 26-12-12
        '     Set Factura.Tipo = DAOTipoFactura.FindFirstByFilter("id IN (select TipoFactura FROM AdminConfigFacturas where idIVA = " & Factura.TipoIVA.idIVA & ")")

        Factura.AlicuotaAplicada = Factura.TipoIVA.alicuota
        Set Factura.cliente = DAOCliente.BuscarPorID(Factura.cliente.id)
        Dim classA As New classAdministracion
    'Set Factura.Tipo = DAOTipoFacturaDiscriminado.FindById(id_Default)
        If IsSomething(Factura.Tipo.TipoFactura) Then
            Factura.EstaDiscriminada = Factura.Tipo.TipoFactura.Discrimina
            Me.lblTipoFactura.caption = Factura.Tipo.TipoFactura.Tipo


            'paso esto al evento click del cboTipos 26-12-12
            '        If Factura.Id = 0 Then 'agregado para q no cambie el nro de factura cuando estoy en edicion yu eliko otro cliente
            '            Me.txtNumero.text = Format(DAOFactura.proximaFactura(Factura.tipo.Id), "0000")
            '        End If


        Else
            Me.lblTipoFactura.caption = vbNullString
            Me.txtNumero.text = 0
        End If


        Me.lblNCND.Visible = (Factura.TipoDocumento <> tipoDocumentoContable.Factura)
        Me.lblNCND.caption = Factura.GetShortDescription(True, True)


        CargarDetalles
        MostrarCliente
        MostrarPercepcionIIBB
        LimpiarTotales
    End If
End Sub

Private Sub MostrarPercepcionIIBB()
    Me.lblBuscandoPercepcion.Visible = False
    Dim tabla As String
    If Me.cboPadron.ListIndex = 0 Then
        tabla = "IIBB2_Percepcion"
    Else
        tabla = "IIBB2_PercepcionAnt"
    End If

    Me.txtPercepcion.text = 0
    Me.lblVencido.Visible = False

    If Factura.cliente.CUITValido Then
        Me.lblBuscandoPercepcion.Visible = True
        DoEvents
        Dim rs As Recordset
        Set rs = conectar.RSFactory("select * from sp_permisos." & tabla & " where cuit='" & Factura.cliente.Cuit & "'")
        Me.lblBuscandoPercepcion.Visible = False
        DoEvents
        If IsSomething(rs) Then
            If Not rs.EOF And Not rs.BOF Then
                Me.lblVencido.Visible = (Now() > CDate(ConvertirAFechaAfip(rs!FechaHasta)))
                Me.txtPercepcion.text = rs!alicuota
                Factura.AlicuotaPercepcionesIIBB = (rs!alicuota / 100) + 1
            End If
        End If
    End If
End Sub





Private Sub cboMoneda_Click()
    If IsSomething(Factura) And Me.cboMoneda.ListIndex <> -1 And Not dataLoading Then
        Set Factura.moneda = DAOMoneda.GetById(Me.cboMoneda.ItemData(Me.cboMoneda.ListIndex))
    End If
End Sub
Private Sub cboPadron_Click()
    
    If IsSomething(Factura.cliente) And Not dataLoading Then
         MostrarPercepcionIIBB
    End If
End Sub
Private Sub LimpiarTotales()
    Me.lblSubTotal.caption = funciones.FormatearDecimales(0)
    Me.lblPercepciones.caption = funciones.FormatearDecimales(0)
    Me.lblIVATot.caption = funciones.FormatearDecimales(0)
    Me.lblTotal.caption = funciones.FormatearDecimales(0)
End Sub

Private Sub cboTiposFactura_Click()
    Dim id As Long
    id = Me.cboTiposFactura.ItemData(Me.cboTiposFactura.ListIndex)
    Set Factura.Tipo = DAOTipoFacturaDiscriminado.FindById(id)


    '1 11 19
'    Me.lblCbuCredito.Visible = Factura.Tipo.PuntoVenta.EsCredito
    Me.frmFC.Enabled = Factura.Tipo.PuntoVenta.EsCredito
'    Me.lblFechaPagoCredito.Visible = Factura.Tipo.PuntoVenta.EsCredito
'    Me.dtFechaPagoCredito.Visible = Factura.Tipo.PuntoVenta.EsCredito
    
    Me.cboCuentasCBU.Visible = Factura.Tipo.PuntoVenta.EsCredito
    Me.lblEsCredito.Visible = Factura.Tipo.PuntoVenta.EsCredito
    
    Me.frmFC.Enabled = Factura.Tipo.PuntoVenta.EsCredito
    Me.frmFC.Visible = Factura.Tipo.PuntoVenta.EsCredito
    
    Me.lblEsCredito.caption = Factura.DescripcionCreditoAdicional
    
    Me.lblVerCbu.Visible = True
    If Not Factura.Tipo.PuntoVenta.EsCredito Then
        Me.lblVerCbu = "NO INFORMADO"
    End If


    If Factura.id = 0 Then    'agregado para q no cambie el nro de factura cuando estoy en edicion yu elijo otro cliente
 '       Me.txtNumero.Enabled = Not Factura.Tipo.PuntoVenta.EsElectronico
 '       If Factura.Tipo.PuntoVenta.EsElectronico Then


 '           Dim Ult As String
 '          Me.txtNumero.text = "0000"    'ERPHelper.GetUltimoAutorizado(Factura.Tipo.PuntoVenta.PuntoVenta, Factura.Tipo.id)
'Else
 Me.txtNumero.text = Format(DAOFactura.proximaFactura(Factura.Tipo.id), "00000000") 'NuevoTipoDocumento, Factura.Tipo.TipoFactura.id), "0000")
'        End If
        Else
'        If Factura.Tipo.PuntoVenta.EsElectronico Then
'           Me.txtNumero.text = "0000"
'        Else
            Me.txtNumero.text = Format(DAOFactura.proximaFactura(Factura.Tipo.id), "00000000") 'NuevoTipoDocumento, Factura.Tipo.TipoFactura.id), "0000")
'        End If
        End If
    

End Sub

Private Sub cmdNueva_Click()
    Dim frm2 As New frmFacturaEdicion
    frm2.Show
End Sub



Private Sub dtFechaPagoCredito_Change()
   If Not dataLoading Then
        Factura.FechaPago = Me.dtFechaPagoCredito.value
    End If
End Sub





'Aca da error 16/05/20

'Private Sub dtFechaPagoCredito_Change()
'   If Not dataLoading Then
'        Factura.FechaPago = Me.dtFechaPagoCredito.value
'    End If
'End Sub

'Private Sub dtFechaPagoCredito_Change()
'   If Not dataLoading Then
'        Factura.FechaPago = Me.dtFechaPagoCredito.value
'    End If
'End Sub


Private Sub dtpFecha_Change()
    If Not dataLoading Then
        Factura.FechaEmision = Me.dtpFecha.value
    End If
End Sub
Private Sub Form_Load()
    Customize Me
    dataLoading = True
    DAOCliente.llenarComboXtremeSuite Me.cboCliente
    cboCliente.ListIndex = -1
    DAOMoneda.llenarComboXtremeSuite Me.cboMoneda
    DAOMoneda.llenarComboXtremeSuite Me.cboMonedaAjuste, True
    DAOCuentaBancaria.llenarComboCBU Me.cboCuentasCBU
    'Me.cboCuentasCBU.Visible = False
    If Not IsSomething(Factura) Then
        Set Factura = New Factura
        Factura.Detalles = New Collection
        Set Factura.Tipo = New clsTipoFacturaDiscriminado
        Factura.Tipo.TipoDoc = NuevoTipoDocumento
        Me.caption = "Nueva " & StrConv(Factura.TipoDocumentoDescription, vbProperCase)
        Me.dtpFecha.value = Now

        If Me.cboMoneda.ListIndex <> -1 Then
            Set Factura.moneda = DAOMoneda.GetById(Me.cboMoneda.ItemData(Me.cboMoneda.ListIndex))
        End If
    Else
        Me.caption = Factura.GetShortDescription(False, True)
    End If

    suscId = funciones.CreateGUID
    Channel.AgregarSuscriptor Me, TipoSuscripcion.FacturarRemitosDetalle_, True
    Me.lblBuscandoPercepcion.Visible = False
    GridEXHelper.CustomizeGrid Me.gridDetalles, , True

    Me.cboPadron.Clear
    cboPadron.AddItem "Actual"
    Me.cboPadron.ItemData(Me.cboPadron.NewIndex) = 0
    cboPadron.AddItem "Anterior"
    Me.cboPadron.ItemData(Me.cboPadron.NewIndex) = 1
    Me.cboPadron.ListIndex = 0

    Me.gridDetalles.ItemCount = 0

    Me.lblNCND.Visible = (Factura.TipoDocumento <> tipoDocumentoContable.Factura)
    Me.lblNCND.caption = Factura.GetShortDescription(True, True)

    If Factura.id = 0 Then
        Factura.FechaEmision = Now
        Factura.estado = EstadoFacturaCliente.EnProceso
        LimpiarFactura
        LimpiarCliente
        LimpiarTotales
    Else
        CargarFactura
    End If



    dataLoading = False


    Me.grpDatos.Enabled = Not ReadOnly
    Me.cboMonedaAjuste.Enabled = Not ReadOnly
    'Me.grpDetalles.Enabled = Not ReadOnly
    If ReadOnly Then
        Me.gridDetalles.EditMode = jgexEditModeOff
        Me.gridDetalles.AllowAddNew = False
        Me.gridDetalles.ReadOnly = True

        Dim mon_ajuste As clsMoneda
        Set mon_ajuste = DAOMoneda.GetById(Factura.IdMonedaAjuste)
        If IsSomething(mon_ajuste) Then
            Me.lblAjuste.caption = "Ajuste a " & mon_ajuste.NombreCorto & " " & Factura.TipoCambioAjuste
            Me.cboMonedaAjuste.Visible = False
        End If
        Dim colu As JSColumn
        For Each colu In Me.gridDetalles.Columns
            colu.EditType = jgexEditNone
        Next colu
    End If

    If EsAnticipo Or Factura.EsAnticipo Then
        Me.gridDetalles.Columns(1).EditType = jgexEditNone
        Me.gridDetalles.AllowDelete = False
        Factura.origenFacturado = OrigenFacturadoAnticipoOT
    End If


    Me.PushButton1.Enabled = Factura.EsAnticipo Or EsAnticipo Or Factura.origenFacturado = OrigenFacturadoAnticipoOT
    Me.PushButton2.Enabled = Factura.EsAnticipo Or EsAnticipo Or Factura.origenFacturado = OrigenFacturadoAnticipoOT
    Me.btnGuardar.Enabled = Not ReadOnly Or EsAnticipo
    Me.btnItemRemito.Enabled = Not ReadOnly And Not EsAnticipo
End Sub

Private Sub LimpiarFactura()
    Me.txtNumero.text = vbNullString
    Me.lblTipoFactura.caption = vbNullString
    'Me.lblNCND.caption = vbNullString
    Me.txtReferencia.text = vbNullString
    Me.txtDiasVenc.text = 0
    
    Me.txtCondObs.text = vbNullString
    
End Sub

Private Sub LimpiarCliente()
    Me.lblCuit.caption = vbNullString
    Me.lblIVA.caption = vbNullString
    Me.lblDireccion.caption = vbNullString
    Me.lblLocalidad.caption = vbNullString
    Me.lblProvincia.caption = vbNullString
    Me.lblCodPostal.caption = vbNullString
End Sub

Private Sub CargarFactura()
    If Not IsSomething(Factura) Then Exit Sub
    Me.cboTiposFactura.Enabled = Not (Factura.estado = EstadoFacturaCliente.Aprobada)
    Me.txtNumero.Locked = (Factura.estado = EstadoFacturaCliente.Aprobada) 'Or Factura.Tipo.PuntoVenta.EsElectronico


    If Factura.estado = EstadoFacturaCliente.Aprobada And Factura.Tipo.PuntoVenta.EsElectronico Then

        If LenB(Factura.CAE) > 0 Then
            Me.txtDetallesCAE.caption = "CAE " & Factura.CAE & " | CAE VTO " & Factura.CAEVto
            Me.txtNumero.Locked = True
        Else
            Me.txtDetallesCAE.caption = ""
        End If
    Else

        Me.txtDetallesCAE.caption = ""
    End If


    If IsSomething(Factura.cliente) Then
        Me.cboCliente.ListIndex = funciones.PosIndexCbo(Factura.cliente.id, Me.cboCliente)
        MostrarCliente
    Else
        LimpiarCliente
    End If
    Me.cboMoneda.ListIndex = funciones.PosIndexCbo(Factura.moneda.id, Me.cboMoneda)

    Me.cboMonedaAjuste.ListIndex = funciones.PosIndexCbo(Factura.IdMonedaAjuste, Me.cboMonedaAjuste)

    If Factura.id = 0 Then
        'creo que aaca no entra nunca
        Dim classA As New classAdministracion
        Me.txtNumero.text = Format(DAOFactura.proximaFactura(Factura.Tipo.id)) 'NuevoTipoDocumento, Factura.Tipo.TipoFactura.id), "0000")
    Else
        Set tipos = DAOTipoFacturaDiscriminado.FindAllByFilter("id_iva=" & Factura.TipoIVA.idIVA & " and tipo_Documento=" & Factura.TipoDocumento)    'acft.id IN (select TipoFactura FROM AdminConfigFacturas where idIVA = " & Factura.TipoIVA.idIVA & ")")

        Me.cboTiposFactura.Clear
        Dim T

        'lleno el combo de tipos de factura y dejo el primero x default

        For Each T In tipos
            cboTiposFactura.AddItem T.PuntoVenta.descripcion
            cboTiposFactura.ItemData(cboTiposFactura.NewIndex) = T.id
        Next T

        Me.cboTiposFactura.ListIndex = funciones.PosIndexCbo(Factura.Tipo.id, Me.cboTiposFactura)

        Me.txtNumero.text = Factura.numero
    End If
    Me.dtpFecha.value = Factura.FechaEmision
    Me.txtPercepcion.text = Round((Factura.AlicuotaPercepcionesIIBB - 1) * 100, 2)
    Me.txtDiasVenc.text = Factura.CantDiasPago
    Me.txtReferencia.text = Factura.OrdenCompra
    Me.txtCondObs.text = Factura.observaciones
    Me.lblTipoFactura.caption = Factura.Tipo.TipoFactura.Tipo

    
    Me.txtTasaAjuste.text = Factura.TasaAjusteMensual
   ' Me.txtCbuCredito = Factura.CBU
    
   Dim c As CuentaBancaria
   If Factura.Tipo.PuntoVenta.EsCredito And LenB(Factura.CBU) > 0 Then
   
      
      
        Set c = DAOCuentaBancaria.FindByCBU(Factura.CBU)
      
      If ReadOnly Then
      
            Me.cboCuentasCBU.Visible = IsSomething(c)
            Me.lblVerCbu.Visible = Not IsSomething(c)
      
            If IsSomething(c) Then
                 Me.cboCuentasCBU.ListIndex = funciones.PosIndexCbo(c.id, Me.cboCuentasCBU)
            Else
                Me.lblVerCbu = Factura.CBU
            End If
      
      Else
        Me.lblVerCbu.Visible = False
             If IsSomething(c) Then
                 Me.cboCuentasCBU.ListIndex = funciones.PosIndexCbo(c.id, Me.cboCuentasCBU)
            End If
        
      
      End If
          Me.dtFechaPagoCredito = Factura.FechaPago
    Else
       Me.cboCuentasCBU.Visible = False
            Me.lblVerCbu.Visible = True
            Me.lblVerCbu = "NO INFORMADO"
      End If
          
      '''TODO: SEGUIR ACA 1-5-2020
'          If IsSomething(c) Then
'          Me.lblVerCbu.Visible = False
'          Me.cboCuentasCBU.Visible = True
'            Set c = DAOCuentaBancaria.FindByCBU(Factura.CBU)
'             Me.cboCuentasCBU.ListIndex = funciones.PosIndexCbo(c.id, Me.cboCuentasCBU)
'          Else
'            Me.lblVerCbu.Visible = True
'            Me.cboCuentasCBU.Visible = False
'             Me.lblVerCbu = Factura.CBU
'          End If
'        End If
    

    CargarDetalles
    Totalizar
End Sub

Private Sub Totalizar()
    Me.lblSubTotal.caption = funciones.FormatearDecimales(Factura.TotalSubTotal)
    Me.lblPercepciones.caption = funciones.FormatearDecimales(Factura.totalPercepciones)
    Me.lblIVATot.caption = funciones.FormatearDecimales(Factura.TotalIVA)
    Me.lblTotal.caption = funciones.FormatearDecimales(Factura.Total)

    GridEXHelper.AutoSizeColumns Me.gridDetalles
End Sub

Private Sub CargarDetalles()
    Me.gridDetalles.ItemCount = 0
    Me.gridDetalles.ItemCount = Factura.Detalles.count
    ActualizarCantDetalles
End Sub

Private Sub ActualizarCantDetalles()
    Me.grpDetalles.caption = "Detalles (Cant: " & Me.gridDetalles.ItemCount & ")"
End Sub


Private Sub MostrarCliente()
    On Error Resume Next
    If Factura Is Nothing Then Exit Sub
    If Factura.cliente Is Nothing Then Exit Sub
    Me.lblCuit.caption = Factura.cliente.Cuit
    Me.lblIVA.caption = Factura.cliente.TipoIVA.detalle
    Me.lblDireccion.caption = Factura.cliente.Domicilio
    Me.lblLocalidad.caption = Factura.cliente.localidad.nombre
    Me.lblCodPostal.caption = Factura.cliente.localidad.cp

    Me.lblProvincia = Factura.cliente.provincia.nombre

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Channel.RemoverSuscripcionTotal Me
End Sub

Private Sub gridDetalles_AfterDelete()
    ActualizarCantDetalles
End Sub

Private Sub gridDetalles_AfterUpdate()
    ActualizarCantDetalles
End Sub

Private Sub gridDetalles_BeforeUpdate(ByVal Cancel As GridEX20.JSRetBoolean)
    If Me.gridDetalles.row = -1 Then    'es nuevo
        Me.gridDetalles.value(7) = True
        Me.gridDetalles.value(8) = True
    End If

    Cancel = Not IsNumeric(Me.gridDetalles.value(1)) Or Not IsNumeric(Me.gridDetalles.value(3)) Or Not IsNumeric(Me.gridDetalles.value(4))
End Sub

Private Sub gridDetalles_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 And ReadOnly And Me.gridDetalles.HitTest(x, y) = jgexHitTestConstants.jgexHTCell Then
        Dim row As Long: row = Me.gridDetalles.RowFromPoint(x, y)
        If row > 0 Then
            Set detalle = Factura.Detalles.item(Me.gridDetalles.RowIndex(row))
            If IsSomething(detalle) Then
                If (Not detalle.OrigenEsConcepto And detalle.AplicadoARemito) Or detalle.OrigenEsConcepto Then
                    Me.PopupMenu Me.mnuDetalles
                    'Else
                    '    MsgBox "El detalle debe ser de concepto para poder aplicar un detalle de remito.", vbExclamation
                End If
            End If
        End If
    End If
End Sub

Private Sub gridDetalles_SelectionChange()
    If ReadOnly Then
        Me.gridDetalles.EditMode = jgexEditModeOff
        Me.gridDetalles.AllowAddNew = False
        Me.gridDetalles.ReadOnly = True
        Exit Sub
    End If


    Dim it As Long
    it = Me.gridDetalles.RowIndex(gridDetalles.row)
    If it > 0 Then
        Set detalle = Factura.Detalles.item(it)

        If detalle.OrigenEsConcepto Then
            gridDetalles.Columns(1).EditType = jgexEditTextBox
        Else
            gridDetalles.Columns(1).EditType = jgexEditNone
        End If
    Else
        gridDetalles.Columns(1).EditType = jgexEditTextBox
    End If

End Sub

Private Sub gridDetalles_UnboundAddNew(ByVal NewRowBookmark As GridEX20.JSRetVariant, ByVal Values As GridEX20.JSRowData)
    Set detalle = New FacturaDetalle
    Set detalle.Factura = Factura
    detalle.idFactura = Factura.id
    detalle.Cantidad = Values(1)
    detalle.detalle = Values(2)
    detalle.PorcentajeDescuento = Values(3)
    detalle.Bruto = Values(4)
    detalle.IvaAplicado = Values(7)
    detalle.IBAplicado = Values(8)

    Factura.Detalles.Add detalle

    Totalizar
End Sub

Private Sub gridDetalles_UnboundDelete(ByVal RowIndex As Long, ByVal Bookmark As Variant)
    If RowIndex > 0 And Factura.Detalles.count > 0 Then
        Factura.Detalles.remove RowIndex
        Totalizar
    End If
End Sub

Private Sub gridDetalles_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex <= Factura.Detalles.count Then
        Set detalle = Factura.Detalles.item(RowIndex)
        Values(1) = funciones.FormatearDecimales(detalle.Cantidad)
        Values(2) = detalle.detalle
        Values(3) = funciones.FormatearDecimales(detalle.PorcentajeDescuento)
        Values(4) = funciones.FormatearDecimales(detalle.Bruto)
        Values(5) = funciones.FormatearDecimales(detalle.SubTotal)
        Values(6) = funciones.FormatearDecimales(detalle.Total)
        Values(7) = detalle.IvaAplicado
        Values(8) = detalle.IBAplicado
        Values(9) = detalle.VerOrigen
        Values(10) = detalle.idprovincia
    End If
End Sub

Private Sub gridDetalles_UnboundUpdate(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex > 0 And Factura.Detalles.count > 0 Then
        Set detalle = Factura.Detalles.item(RowIndex)

        detalle.Cantidad = Values(1)
        detalle.detalle = Values(2)
        detalle.PorcentajeDescuento = Values(3)
        detalle.Bruto = Values(4)
        detalle.IvaAplicado = Values(7)
        detalle.IBAplicado = Values(8)
        Totalizar
    End If
End Sub


Private Property Get ISuscriber_id() As String
    ISuscriber_id = suscId
End Property

Private Function ISuscriber_Notificarse(EVENTO As clsEventoObserver) As Variant
    On Error GoTo err1



Dim A As frmPlaneamientoRemitoVer

Set A = EVENTO.Originador

If A.IdFormSuscriber <> ISuscriber_id Then Exit Function

    If EVENTO.Tipo = FacturarRemitosDetalle_ Then
        If ReadOnly Then

            'aplicacion de detalle remito post facturacion
            If Not IsSomething(EVENTO.Elemento) Then Exit Function
            If EVENTO.Elemento.count > 0 Then
                Dim redeta As remitoDetalle
                Set redeta = EVENTO.Elemento(1)
                If redeta.Facturado Then
                    MsgBox "El detalle del remito ya se encuentra facturado.", vbExclamation
                Else
                    If redeta.facturable Then

                        If IsSomething(detaFactRemito) Then
                            detaFactRemito.DetalleRemitoId = redeta.id
                            detaFactRemito.AplicadoARemito = True
                            redeta.Facturado = True
                            Dim transactionResult As Boolean: transactionResult = True
                            Dim q As String
                            conectar.BeginTransaction

                            q = "INSERT INTO AdminFacturasDetalleAplicacionRemitos (idFacturaDetalle, idRemitoDetalle, cantidadAplicada) VALUES (" & detaFactRemito.id & ", " & redeta.id & "  ,  " & redeta.Cantidad & ")"
                            transactionResult = transactionResult And conectar.execute(q)

                            Dim remi As Remito
                            Set remi = DAORemitoS.FindById(detaFactRemito.detalleRemito.Remito)
                            remi.EstadoFacturado = DAORemitoS.AnalizarEstadoFacturado(remi.id)



                            transactionResult = transactionResult And DAORemitoS.Guardar(remi, False, False)
                            transactionResult = transactionResult And DAOFacturaDetalles.Guardar(detaFactRemito)
                            transactionResult = transactionResult And DAORemitoSDetalle.Guardar(redeta)
                            transactionResult = transactionResult And DAODetalleOrdenTrabajo.SaveCantidad(redeta.idDetallePedido, redeta.Cantidad, CantidadFacturada_, redeta.Valor, Factura.id, Factura.moneda.id, Factura.CambioAPatron, Factura.TipoCambioAjuste)

                            If transactionResult Then

                                conectar.CommitTransaction
                                remi.EstadoFacturado = DAORemitoS.AnalizarEstadoFacturado(remi.id)
                                DAORemitoS.Guardar remi, False, False
                                CargarDetalles
                                Totalizar
                                MsgBox "El detalle de la factura ha sido actualizado como remitado.", vbInformation + vbOKOnly
                            Else
                                conectar.RollBackTransaction
                                detaFactRemito.AplicadoARemito = False
                                detaFactRemito.DetalleRemitoId = -1
                                redeta.Facturado = False
                                MsgBox "No se pudieron guardar los cambios.", vbCritical + vbOKOnly
                            End If
                        End If


                    Else
                        MsgBox "El detalle del remito no es facturable.", vbExclamation
                    End If
                End If

            End If

        Else
            AgregarEntregas EVENTO.Elemento
        End If

    End If
    Exit Function

err1:
End Function








Private Sub mnuAplicarDetalleRemito_Click()

    Set detaFactRemito = Nothing

    On Error Resume Next
    Dim f11 As New frmPlaneamientoRemitosListaProceso
    f11.idCliMostrar = Factura.cliente.id
    f11.mostrar = 2
    Set Selecciones.RemitoElegido = Nothing
    f11.Show 1

    If IsSomething(Selecciones.RemitoElegido) Then
        Dim frm As frmPlaneamientoRemitoVer

        Set frm = New frmPlaneamientoRemitoVer
        frm.IdFormSuscriber = ISuscriber_id
        frm.Usable = False
        frm.editar = False
        Set frm.Remito = Selecciones.RemitoElegido    'DAORemitoS.FindById(Selecciones.RemitoElegido.Id)
        frm.ParaFacturar = True
        frm.grilla.MultiSelect = False
        frm.MostrarInfoAdministracion = True

        Set detaFactRemito = Factura.Detalles(Me.gridDetalles.RowIndex(Me.gridDetalles.row))

        frm.Show
    End If


End Sub

Private Sub PushButton1_Click()

    If IsSomething(Factura.cliente) Then
        Set Selecciones.OrdenTrabajo = Nothing
        Set frmPlaneamientoPedidosSeleccion.cliente = Factura.cliente
        frmPlaneamientoPedidosSeleccion.MostrarAnticipo = True
        frmPlaneamientoPedidosSeleccion.Show 1

        Dim Ot As OrdenTrabajo
        If IsSomething(Selecciones.OrdenTrabajo) Then
            If Not funciones.BuscarEnColeccion(Factura.OTsFacturadasAnticipo, CStr(Selecciones.OrdenTrabajo.id)) Then
                Set Ot = DAOOrdenTrabajo.FindById(Selecciones.OrdenTrabajo.id)
                Set Ot.Detalles = DAODetalleOrdenTrabajo.FindAllByOrdenTrabajo(Ot.id, True, True, True)

                Factura.OTsFacturadasAnticipo.Add Ot, CStr(Ot.id)

                Factura.Detalles = New Collection
                Factura.OrdenCompra = vbNullString
                Me.txtReferencia = "FACTURA POR ANTICIPO OT"
                Me.txtCondObs = vbNullString
                Me.txtDiasVenc = 0

                Dim deta As FacturaDetalle

                For Each Ot In Factura.OTsFacturadasAnticipo
                    Me.txtReferencia.text = Me.txtReferencia.text & " " & Ot.IdFormateado
                    'Factura.OrdenCompra = Factura.OrdenCompra & " | " & ot.Descripcion


                    '                    If IsSomething(Ot.detalles) Then
                    '                       If Ot.detalles.count = 0 Then Set Ot.detalles = DAODetalleOrdenTrabajo.FindAllByOrdenTrabajo(Ot.Id, True, True, True)
                    '
                    '                    Else
                    '                        Set Ot.detalles = DAODetalleOrdenTrabajo.FindAllByOrdenTrabajo(Ot.Id, True, True, True)
                    '                    End If

                    Set deta = Factura.DetalleFacturaAnticipoOt(funciones.RedondearDecimales(Ot.Anticipo))
                    If Not IsSomething(deta) Then
                        Set deta = New FacturaDetalle
                        deta.Cantidad = 1
                        deta.DescuentoAnticipo = True
                        deta.IvaAplicado = True
                        deta.IBAplicado = True
                        deta.PorcentajeDescuento = 0
                        Set deta.Factura = Factura
                        deta.detalle = "ANTICIPO " & funciones.RedondearDecimales(Ot.Anticipo) & "% | OT"
                        Factura.Detalles.Add deta
                    End If
                    deta.detalle = deta.detalle & " " & Ot.IdFormateado

            'bug #2
                    deta.Bruto = deta.Bruto + funciones.RedondearDecimales((Ot.Total * Ot.moneda.Cambio * Ot.Anticipo) / 100)

                    '   deta.Bruto = MonedaConverter.Convertir(deta.Bruto, Ot.Moneda.Id, Factura.Moneda.Id)

                Next Ot


                CargarDetalles
                Totalizar
            End If
        End If
    Else
        MsgBox "Debe seleccionar un cliente para poder operar.", vbExclamation
    End If
End Sub



Private Sub txtCondObs_Change()
    If Not dataLoading Then
        Factura.observaciones = Me.txtCondObs.text
    End If

End Sub

Private Sub txtDiasVenc_Change()
    If Not dataLoading Then
        Factura.CantDiasPago = Val(Me.txtDiasVenc.text)
    End If
End Sub

Private Sub txtNumero_Change()
    If Not dataLoading Then
        Factura.numero = Me.txtNumero
    End If
End Sub

Private Sub txtPercepcion_Change()
    On Error GoTo E
    If Not dataLoading Then
        If LenB(Me.txtPercepcion.text) = 0 Then
            Factura.AlicuotaPercepcionesIIBB = 0
        Else
            Factura.AlicuotaPercepcionesIIBB = 1 + (CDbl(Me.txtPercepcion.text) / 100)
        End If
        Totalizar
    End If

    Exit Sub
E:
    Factura.AlicuotaPercepcionesIIBB = 0
    Me.txtPercepcion.text = 0
End Sub



Private Sub txtReferencia_Change()
    If Not dataLoading Then
        Factura.OrdenCompra = Me.txtReferencia.text
    End If
End Sub

Private Sub txtTasaAjuste_Change()
    If Not dataLoading Then
        Factura.TasaAjusteMensual = Val(Me.txtTasaAjuste.text)
    End If
End Sub
