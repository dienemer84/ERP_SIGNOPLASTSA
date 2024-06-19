VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmAdminFacturasEdicion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Factura"
   ClientHeight    =   11205
   ClientLeft      =   3945
   ClientTop       =   2385
   ClientWidth     =   17775
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
   ScaleHeight     =   11205
   ScaleWidth      =   17775
   Begin VB.Frame Frame 
      Caption         =   "Datos del cliente seleccionado"
      Height          =   2655
      Left            =   120
      TabIndex        =   81
      Top             =   1680
      Width           =   5535
      Begin VB.Label lblIdImpositivo 
         Caption         =   "22222222222"
         Height          =   255
         Left            =   1320
         TabIndex        =   98
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label lblCuitPais 
         Caption         =   "11111111"
         Height          =   255
         Left            =   1320
         TabIndex        =   97
         Top             =   1680
         Width           =   3255
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "ID Impositivo:"
         Enabled         =   0   'False
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   96
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Cuit Pais:"
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   95
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Provincia:"
         Height          =   195
         Left            =   480
         TabIndex        =   93
         Top             =   1200
         Width           =   705
      End
      Begin VB.Label lblProvincia 
         AutoSize        =   -1  'True
         Caption         =   "2343242"
         Height          =   195
         Left            =   1365
         TabIndex        =   92
         Top             =   1200
         Width           =   630
      End
      Begin VB.Label lblCodPostal 
         AutoSize        =   -1  'True
         Caption         =   "2343242"
         Height          =   195
         Left            =   1365
         TabIndex        =   91
         Top             =   1440
         Width           =   630
      End
      Begin VB.Label lblLocalidad 
         AutoSize        =   -1  'True
         Caption         =   "HHHHHH"
         Height          =   195
         Left            =   1365
         TabIndex        =   90
         Top             =   960
         Width           =   630
      End
      Begin VB.Label lblDireccion 
         Caption         =   "RIVAD 3242"
         Height          =   195
         Left            =   1365
         TabIndex        =   89
         Top             =   720
         Width           =   4095
      End
      Begin VB.Label lblIVA 
         AutoSize        =   -1  'True
         Caption         =   "23"
         Height          =   195
         Left            =   1365
         TabIndex        =   88
         Top             =   480
         Width           =   180
      End
      Begin VB.Label lblCuit 
         AutoSize        =   -1  'True
         Caption         =   "23-30279550-9"
         Height          =   195
         Left            =   1365
         TabIndex        =   87
         Top             =   240
         Width           =   1110
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cod Postal:"
         Height          =   195
         Left            =   360
         TabIndex        =   86
         Top             =   1440
         Width           =   825
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Localidad:"
         Height          =   195
         Left            =   480
         TabIndex        =   85
         Top             =   960
         Width           =   720
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "CUIT:"
         Height          =   195
         Left            =   765
         TabIndex        =   84
         Top             =   255
         Width           =   420
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Direccion:"
         Height          =   195
         Left            =   480
         TabIndex        =   83
         Top             =   720
         Width           =   705
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "IVA:"
         Height          =   195
         Index           =   0
         Left            =   870
         TabIndex        =   82
         Top             =   480
         Width           =   315
      End
   End
   Begin XtremeSuiteControls.PushButton btnExportarContenido 
      Height          =   615
      Left            =   15480
      TabIndex        =   80
      Top             =   8280
      Width           =   2055
      _Version        =   786432
      _ExtentX        =   3625
      _ExtentY        =   1085
      _StockProps     =   79
      Caption         =   "Exportar Detalle de Cbte a Excel"
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
   Begin VB.Frame grpDatosCliente 
      Caption         =   "Cliente"
      Height          =   1575
      Left            =   120
      TabIndex        =   75
      Top             =   120
      Width           =   5535
      Begin XtremeSuiteControls.PushButton btnCrearCliente 
         Height          =   375
         Left            =   240
         TabIndex        =   78
         Top             =   960
         Width           =   1935
         _Version        =   786432
         _ExtentX        =   3413
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Crear Cliente"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cboCliente 
         Height          =   315
         Left            =   240
         TabIndex        =   76
         Top             =   480
         Width           =   4995
         _Version        =   786432
         _ExtentX        =   8811
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Text            =   "cboCliente"
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Seleccionar Cliente:"
         Height          =   195
         Left            =   240
         TabIndex        =   77
         Top             =   240
         Width           =   1410
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Período de Servicio / Producto"
      Enabled         =   0   'False
      Height          =   1695
      Left            =   11640
      TabIndex        =   58
      Top             =   120
      Width           =   6015
      Begin MSComCtl2.DTPicker dtFechaPagoCreditoHasta 
         Height          =   405
         Left            =   2640
         TabIndex        =   59
         Top             =   1095
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   714
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   67764225
         CurrentDate     =   43967
      End
      Begin MSComCtl2.DTPicker dtFechaPagoCreditoDesde 
         Height          =   405
         Left            =   2640
         TabIndex        =   60
         Top             =   645
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   714
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   67764225
         CurrentDate     =   43967
      End
      Begin VB.Line Line8 
         BorderColor     =   &H80000010&
         X1              =   4080
         X2              =   4320
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line9 
         BorderColor     =   &H80000010&
         X1              =   4080
         X2              =   4320
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Line Line10 
         BorderColor     =   &H80000010&
         X1              =   4320
         X2              =   4320
         Y1              =   1320
         Y2              =   840
      End
      Begin VB.Label lblPeriodoFacturadoT 
         Caption         =   "Período Facturado"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   63
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label lblPeriodoFacturadoD 
         Caption         =   "Desde:"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   62
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblPeriodoFacturadoH 
         Caption         =   "Hasta:"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1980
         TabIndex        =   61
         Top             =   1170
         Width           =   615
      End
   End
   Begin VB.Frame frmFCE1 
      Caption         =   "Factura Eléctronica"
      Height          =   1095
      Left            =   11640
      TabIndex        =   53
      Top             =   9000
      Visible         =   0   'False
      Width           =   6015
      Begin MSComCtl2.DTPicker dtFechaServDesde1 
         Height          =   405
         Left            =   1680
         TabIndex        =   54
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   714
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   67764225
         CurrentDate     =   43983
      End
      Begin MSComCtl2.DTPicker dtFechaServHasta1 
         Height          =   405
         Left            =   4080
         TabIndex        =   55
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   714
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   67764225
         CurrentDate     =   43983
      End
      Begin VB.Label lblFechaServDesde1 
         Caption         =   "Servicio Desde:"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   57
         Top             =   435
         Width           =   1815
      End
      Begin VB.Label lblFechaServHasta1 
         Caption         =   "Hasta:"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3420
         TabIndex        =   56
         Top             =   435
         Width           =   975
      End
   End
   Begin VB.Frame frmFC 
      Caption         =   "Factura de Crédito Eléctronica"
      Enabled         =   0   'False
      Height          =   2655
      Left            =   11640
      TabIndex        =   35
      Top             =   1800
      Width           =   6015
      Begin VB.ComboBox cboOpcional27 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmFacturaEdicion.frx":000C
         Left            =   120
         List            =   "frmFacturaEdicion.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   65
         Top             =   1560
         Width           =   5850
      End
      Begin VB.ComboBox cboCuentasCBU 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmFacturaEdicion.frx":0010
         Left            =   120
         List            =   "frmFacturaEdicion.frx":0012
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   720
         Width           =   5850
      End
      Begin VB.Label Label22 
         Caption         =   "Opción de Transferencia"
         Height          =   195
         Left            =   120
         TabIndex        =   64
         Top             =   1320
         Width           =   5595
      End
      Begin VB.Label LblCBU 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "CBU:"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   39
         Top             =   360
         Width           =   420
      End
      Begin VB.Label lblVerCbu 
         Caption         =   "NO INFORMADO"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   38
         Top             =   780
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Label lblCbuCredito 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "CBU:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2160
         TabIndex        =   37
         Top             =   360
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.Label lblEsCredito 
         Caption         =   "FACTURA DE CRÉDITO ELECTRÓNICA MiPyMES (FCE)"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   2160
         Visible         =   0   'False
         Width           =   3975
      End
   End
   Begin VB.Frame frmTextoAdicional 
      Caption         =   "Texto Adicional (Limite de 300 caracteres)"
      Height          =   2895
      Left            =   11640
      TabIndex        =   33
      Top             =   4440
      Width           =   6015
      Begin VB.TextBox txtTextoAdicional 
         Height          =   2175
         Left            =   120
         MaxLength       =   300
         MultiLine       =   -1  'True
         TabIndex        =   34
         Top             =   240
         Width           =   5775
      End
      Begin VB.Label lblCaracteresRestantes 
         Caption         =   "Caracteres restantes: "
         Height          =   255
         Left            =   120
         TabIndex        =   79
         Top             =   2520
         Width           =   3375
      End
   End
   Begin XtremeSuiteControls.GroupBox grpTotales 
      Height          =   1575
      Left            =   11640
      TabIndex        =   19
      Top             =   7320
      Width           =   3780
      _Version        =   786432
      _ExtentX        =   6667
      _ExtentY        =   2778
      _StockProps     =   79
      Caption         =   "Totales"
      UseVisualStyle  =   -1  'True
      Begin VB.Label lblTC 
         Caption         =   "Tipo Cambio:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   94
         Top             =   1200
         Width           =   2535
      End
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
         Left            =   1620
         TabIndex        =   27
         Top             =   705
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
         Left            =   1620
         TabIndex        =   26
         Top             =   450
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
         Left            =   405
         TabIndex        =   25
         Top             =   450
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
         Left            =   1620
         TabIndex        =   24
         Top             =   195
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
         Left            =   795
         TabIndex        =   23
         Top             =   195
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
         Left            =   1140
         TabIndex        =   22
         Top             =   705
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
         Left            =   1020
         TabIndex        =   21
         Top             =   930
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
         Left            =   1620
         TabIndex        =   20
         Top             =   930
         Width           =   1080
      End
   End
   Begin XtremeSuiteControls.PushButton btnItemRemito 
      Height          =   360
      Left            =   120
      TabIndex        =   6
      Top             =   10200
      Width           =   2055
      _Version        =   786432
      _ExtentX        =   3625
      _ExtentY        =   635
      _StockProps     =   79
      Caption         =   "Item de Remito..."
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.GroupBox grpDatos 
      Height          =   4335
      Left            =   5760
      TabIndex        =   12
      Top             =   120
      Width           =   5835
      _Version        =   786432
      _ExtentX        =   10292
      _ExtentY        =   7646
      _StockProps     =   79
      Caption         =   "Datos del Comprobante"
      Appearance      =   4
      Begin XtremeSuiteControls.CheckBox chkEsCredito 
         Height          =   375
         Left            =   2640
         TabIndex        =   49
         Top             =   780
         Width           =   255
         _Version        =   786432
         _ExtentX        =   450
         _ExtentY        =   661
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
      End
      Begin VB.ComboBox cboConceptosAIncluir 
         Enabled         =   0   'False
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
         ItemData        =   "frmFacturaEdicion.frx":0014
         Left            =   2520
         List            =   "frmFacturaEdicion.frx":0021
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Top             =   1250
         Width           =   3090
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
         Left            =   2520
         TabIndex        =   1
         Text            =   "999999"
         Top             =   1800
         Width           =   3090
      End
      Begin XtremeSuiteControls.DateTimePicker dtpFecha 
         Height          =   405
         Left            =   2520
         TabIndex        =   2
         Top             =   2400
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
      Begin XtremeSuiteControls.ComboBox cboTiposFactura 
         Height          =   360
         Left            =   2520
         TabIndex        =   0
         Top             =   240
         Width           =   3090
         _Version        =   786432
         _ExtentX        =   5450
         _ExtentY        =   635
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
      Begin XtremeSuiteControls.ComboBox cboMoneda 
         Height          =   405
         Left            =   2520
         TabIndex        =   3
         Top             =   3000
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
         Text            =   "cboMoneda"
         DropDownItemCount=   3
         EnableMarkup    =   -1  'True
      End
      Begin MSComCtl2.DTPicker dtFechaPagoCredito 
         Height          =   405
         Left            =   3915
         TabIndex        =   52
         Top             =   3645
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   714
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   67764225
         CurrentDate     =   43967
      End
      Begin VB.Label lblFechaPagoCredito 
         Caption         =   "Fecha de Vto. para el Pago:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   51
         Top             =   3720
         Width           =   2535
      End
      Begin VB.Label Label23 
         Caption         =   "DE CRÉDITO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   50
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label lblConceptosAIncluir 
         Caption         =   "Concepto:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   48
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label13 
         Caption         =   "MiPyMES:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   46
         Top             =   825
         Width           =   1095
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFDBBF&
         DrawMode        =   9  'Not Mask Pen
         X1              =   14640
         X2              =   3120
         Y1              =   9000
         Y2              =   9000
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
         Left            =   1740
         TabIndex        =   30
         Top             =   285
         Width           =   705
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
         Left            =   1530
         TabIndex        =   17
         Top             =   3060
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
         Left            =   1755
         TabIndex        =   16
         Top             =   2460
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
         Left            =   1500
         TabIndex        =   15
         Top             =   1860
         Width           =   945
      End
      Begin VB.Label lblNCND 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "N/D"
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   960
         Width           =   615
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFDBBF&
         DrawMode        =   9  'Not Mask Pen
         X1              =   1080
         X2              =   1080
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
         Left            =   240
         TabIndex        =   13
         Top             =   270
         Width           =   645
      End
   End
   Begin XtremeSuiteControls.GroupBox grpDetalles 
      Height          =   3795
      Left            =   120
      TabIndex        =   18
      Top             =   6240
      Width           =   11475
      _Version        =   786432
      _ExtentX        =   20241
      _ExtentY        =   6694
      _StockProps     =   79
      Caption         =   "Detalles (Cant: 0)"
      Appearance      =   2
      Begin GridEX20.GridEX gridDetalles 
         Height          =   3315
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   11250
         _ExtentX        =   19844
         _ExtentY        =   5847
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
         Column(1)       =   "frmFacturaEdicion.frx":0052
         Column(2)       =   "frmFacturaEdicion.frx":018A
         Column(3)       =   "frmFacturaEdicion.frx":027E
         Column(4)       =   "frmFacturaEdicion.frx":0396
         Column(5)       =   "frmFacturaEdicion.frx":04B6
         Column(6)       =   "frmFacturaEdicion.frx":05F6
         Column(7)       =   "frmFacturaEdicion.frx":072A
         Column(8)       =   "frmFacturaEdicion.frx":0852
         Column(9)       =   "frmFacturaEdicion.frx":0982
         Column(10)      =   "frmFacturaEdicion.frx":0A92
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmFacturaEdicion.frx":0B8A
         FormatStyle(2)  =   "frmFacturaEdicion.frx":0CB2
         FormatStyle(3)  =   "frmFacturaEdicion.frx":0D62
         FormatStyle(4)  =   "frmFacturaEdicion.frx":0E16
         FormatStyle(5)  =   "frmFacturaEdicion.frx":0EEE
         FormatStyle(6)  =   "frmFacturaEdicion.frx":0FA6
         ImageCount      =   0
         PrinterProperties=   "frmFacturaEdicion.frx":1086
      End
   End
   Begin XtremeSuiteControls.PushButton btnGuardar 
      Height          =   600
      Left            =   15480
      TabIndex        =   11
      Top             =   7440
      Width           =   2055
      _Version        =   786432
      _ExtentX        =   3625
      _ExtentY        =   1058
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
   Begin XtremeSuiteControls.PushButton btnSeleccionarOT 
      Height          =   360
      Left            =   120
      TabIndex        =   7
      Top             =   10680
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
      Left            =   9120
      TabIndex        =   9
      Top             =   10680
      Width           =   2415
      _Version        =   786432
      _ExtentX        =   4260
      _ExtentY        =   635
      _StockProps     =   79
      Caption         =   "Generar Items Anticipo OT"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton btnCrearItemConcepto 
      Height          =   360
      Left            =   9120
      TabIndex        =   8
      Top             =   10200
      Width           =   2415
      _Version        =   786432
      _ExtentX        =   4260
      _ExtentY        =   635
      _StockProps     =   79
      Caption         =   "Crear Item de concepto..."
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton cmdNueva 
      Height          =   360
      Left            =   11640
      TabIndex        =   10
      Top             =   10200
      Width           =   2415
      _Version        =   786432
      _ExtentX        =   4260
      _ExtentY        =   635
      _StockProps     =   79
      Caption         =   "Nueva"
      Enabled         =   0   'False
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.ComboBox cboMonedaAjuste 
      Height          =   405
      Left            =   3840
      TabIndex        =   28
      Top             =   10200
      Visible         =   0   'False
      Width           =   2550
      _Version        =   786432
      _ExtentX        =   4498
      _ExtentY        =   714
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
   Begin XtremeSuiteControls.GroupBox grpPercep 
      Height          =   1695
      Left            =   6720
      TabIndex        =   40
      Top             =   4440
      Width           =   4875
      _Version        =   786432
      _ExtentX        =   8599
      _ExtentY        =   2990
      _StockProps     =   79
      Caption         =   "Percepciones IIBB"
      Appearance      =   4
      Begin VB.TextBox txtPercepcion 
         Height          =   300
         Left            =   1560
         TabIndex        =   41
         Top             =   840
         Width           =   2715
      End
      Begin XtremeSuiteControls.ComboBox cboPadron 
         Height          =   315
         Left            =   1560
         TabIndex        =   42
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
         Left            =   1440
         TabIndex        =   45
         Top             =   1320
         Visible         =   0   'False
         Width           =   2670
      End
      Begin XtremeSuiteControls.Label lblPadron 
         Height          =   195
         Left            =   240
         TabIndex        =   44
         Top             =   420
         Width           =   1215
         _Version        =   786432
         _ExtentX        =   2143
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Padron a utilizar:"
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label17 
         Height          =   195
         Left            =   480
         TabIndex        =   43
         Top             =   900
         Width           =   840
         _Version        =   786432
         _ExtentX        =   1482
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Percepcion:"
         AutoSize        =   -1  'True
      End
   End
   Begin XtremeSuiteControls.GroupBox grpInfo 
      Height          =   1695
      Left            =   120
      TabIndex        =   66
      Top             =   4440
      Width           =   6495
      _Version        =   786432
      _ExtentX        =   11456
      _ExtentY        =   2990
      _StockProps     =   79
      Caption         =   "Detalles"
      UseVisualStyle  =   -1  'True
      Begin VB.ComboBox txtCondObs 
         Height          =   315
         ItemData        =   "frmFacturaEdicion.frx":1256
         Left            =   1035
         List            =   "frmFacturaEdicion.frx":1266
         TabIndex        =   70
         Top             =   1200
         Width           =   5295
      End
      Begin VB.TextBox txtTasaAjuste 
         Height          =   300
         Left            =   5160
         TabIndex        =   69
         Top             =   720
         Width           =   1200
      End
      Begin VB.TextBox txtDiasVenc 
         Height          =   300
         Left            =   2160
         TabIndex        =   68
         Top             =   720
         Width           =   1080
      End
      Begin VB.TextBox txtReferencia 
         Height          =   300
         Left            =   1395
         TabIndex        =   67
         Top             =   240
         Width           =   4935
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "% Tasa ajuste mensual:"
         Height          =   195
         Left            =   3350
         TabIndex        =   74
         Top             =   780
         Width           =   1740
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Condicion:"
         Height          =   195
         Left            =   240
         TabIndex        =   73
         Top             =   1260
         Width           =   750
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cant Días Vencimiento FF:"
         Height          =   195
         Left            =   240
         TabIndex        =   72
         Top             =   780
         Width           =   1875
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "OC / Referencia:"
         Height          =   195
         Left            =   120
         TabIndex        =   71
         Top             =   300
         Width           =   1215
      End
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "23-30279550-9"
      Height          =   195
      Left            =   600
      TabIndex        =   32
      Top             =   240
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
      Left            =   11760
      TabIndex        =   31
      Top             =   10800
      Width           =   5385
   End
   Begin VB.Label lblAjuste 
      AutoSize        =   -1  'True
      Caption         =   "Ajuste a"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3000
      TabIndex        =   29
      Top             =   10260
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Menu mnuDetalles 
      Caption         =   "mnuDetalles"
      Visible         =   0   'False
      Begin VB.Menu mnuAplicarDetalleRemito 
         Caption         =   "Aplicar detalle de remito"
      End
   End
End
Attribute VB_Name = "frmAdminFacturasEdicion"
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
'Private ErrorAfip As Boolean
Public NuevoTipoDocumento As tipoDocumentoContable
Public EsAnticipo As Boolean

Public ReadOnly As Boolean

Private detaFactRemito As FacturaDetalle

Public Property Let idFactura(value As Long)
    Set Factura = DAOFactura.FindById(value, True, True)
End Property

Private Sub btnCrearCliente_Click()

        frmVentasClienteNuevo.Show 1
        
        CargarClientesEnCbo
        

End Sub

Private Sub btnExportarContenido_Click()

    Dim xlApplication As Object
    Set xlApplication = CreateObject("Excel.Application")

    Dim xlWorkbook As Object
    Set xlWorkbook = CreateObject("Excel.Application")

    Dim xlWorksheet As Object
    Set xlWorksheet = CreateObject("Excel.Application")

    Set xlWorkbook = xlApplication.Workbooks.Add

    Set xlWorksheet = xlWorkbook.Worksheets.item(1)

    xlWorksheet.Activate

    xlWorksheet.Cells(1, 1).value = "Detalle de Cbte " + Factura.GetShortDescription(False, False)

    xlWorksheet.Cells(2, 1).value = "Cantidad"
    xlWorksheet.Cells(2, 2).value = "Detalle"
    xlWorksheet.Cells(2, 3).value = "% Descuento"
    xlWorksheet.Cells(2, 4).value = "U Bruto"
    xlWorksheet.Cells(2, 5).value = "U Neto"
    xlWorksheet.Cells(2, 6).value = "Total"
    xlWorksheet.Cells(2, 7).value = "IVA"
    xlWorksheet.Cells(2, 8).value = "IIBB"

    xlWorksheet.Range("A2:I2").Font.Bold = True

    Dim idx As Integer
    idx = 3

    Dim deta As FacturaDetalle

    'DEFINE EL CONTADOR DEL PROGRESSBAR Y LO INICIA EN 0
    Dim d As Long
    d = 0

    For Each deta In Factura.Detalles
        xlWorksheet.Cells(idx, 1).value = deta.Cantidad
        xlWorksheet.Cells(idx, 2).value = deta.detalle
        xlWorksheet.Cells(idx, 3).value = deta.PorcentajeDescuento
        xlWorksheet.Cells(idx, 4).value = deta.Bruto

        xlWorksheet.Cells(idx, 5).value = deta.NetoGravado
        xlWorksheet.Cells(idx, 6).value = deta.total

        If deta.IvaAplicado Then
            xlWorksheet.Cells(idx, 7).value = "SI"
        Else
            xlWorksheet.Cells(idx, 7).value = "NO"
        End If

        If deta.IBAplicado Then
            xlWorksheet.Cells(idx, 8).value = "SI"
        Else
            xlWorksheet.Cells(idx, 8).value = "NO"
        End If

        idx = idx + 1

        'POR CADA ITERACION SUMA UN VALOR A LA VARIABLE D DEL PROGRESSBAR
        d = d + 1

    Next
    
    xlWorksheet.Columns(1).ColumnWidth = 8 ' Puedes ajustar el valor según tus necesidades
    xlWorksheet.Cells(idx + 1, 3).value = "Totales: "
    xlWorksheet.Cells(idx + 1, 3).HorizontalAlignment = xlRight

    xlWorksheet.Cells(idx + 1, 4).Formula = "=SUM(D3:D" & idx - 1 & ")"
    xlWorksheet.Cells(idx + 1, 5).Formula = "=SUM(E3:E" & idx - 1 & ")"
    xlWorksheet.Cells(idx + 1, 6).Formula = "=SUM(F3:F" & idx - 1 & ")"

    xlWorksheet.Cells(idx + 1, 4).Font.Bold = True
    xlWorksheet.Cells(idx + 1, 5).Font.Bold = True
    xlWorksheet.Cells(idx + 1, 6).Font.Bold = True

    xlWorksheet.Range("D3:E15").NumberFormat = "#,##0.00"
    xlWorksheet.Range("D3:F" & idx + 1).HorizontalAlignment = xlRight
    xlWorksheet.Range("D3:F" & idx + 1).NumberFormat = "#,##0.00"
    
    
    xlWorksheet.Range("A" & idx + 3 & ":A" & idx + 6).HorizontalAlignment = xlRight
    xlWorksheet.Range("A" & idx + 3 & ":A" & idx + 6).Font.Bold = True
    
    xlWorksheet.Cells(idx + 3, 1).value = "Subtotal"
    xlWorksheet.Cells(idx + 3, 2).value = Me.lblSubTotal.caption
    
    xlWorksheet.Cells(idx + 4, 1).value = "Percepciones"
    xlWorksheet.Cells(idx + 4, 2).value = Me.lblPercepciones.caption
    
    xlWorksheet.Cells(idx + 5, 1).value = "IVA"
    xlWorksheet.Cells(idx + 5, 2).value = Me.lblIVATot.caption
    
    xlWorksheet.Cells(idx + 6, 1).value = "Total"
    xlWorksheet.Cells(idx + 6, 2).value = Me.lblTotal.caption
    
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

End Sub

Private Sub btnGuardar_Click()

    On Error GoTo err1

    If Me.gridDetalles.EditMode = jgexEditModeOn Then
        MsgBox "Todavia esta editando algun detalle de la factura.", vbExclamation
        Exit Sub
    End If



    If Not Factura.Cliente.CUITValido Or Not Factura.Cliente.ValidoRemitoFactura Then
        MsgBox "El cliente no es valido para poder facturar.", vbExclamation + vbOKOnly
        Exit Sub
    End If

    Factura.observaciones = Me.txtCondObs.text

    If LenB(Factura.numero) = 0 Or _
       LenB(Factura.OrdenCompra) = 0 Or _
       Factura.observaciones = "" Or _
       Factura.CantDiasPago = 0 Then
        MsgBox "El Comprobante debe poseer Nº, referencia, cantidad dias de vto. de FF y condición cargada.", vbExclamation + vbOKOnly
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
            Factura.IdMonedaAjuste = mon.Id
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
                                    If Not IsSomething(Factura.DetalleAnticipoOT(Ot.Id)) Then
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

        Factura.observaciones = Me.txtCondObs.text
        Factura.TextoAdicional = Me.txtTextoAdicional
        Factura.FechaServDesde = Me.dtFechaServDesde1.value
        Factura.FechaServHasta = Me.dtFechaServHasta1.value
        Factura.fechaPago = Me.dtFechaPagoCredito.value
        Factura.esCredito = Me.chkEsCredito.value

        If Factura.esCredito Then

            If Me.cboOpcional27.ListIndex < 0 Then Err.Raise "Para FCE es obligatorio informar opcional 27 con valor SCA ó ADC"
            Factura.Opcional27 = Me.cboOpcional27.ItemData(Me.cboOpcional27.ListIndex)
        Else
            Factura.Opcional27 = 0

        End If

        ASociarConcepto

        If DAOFactura.Save(Factura, True) Then
            MsgBox "La " & StrConv(Factura.TipoDocumentoDescription, vbProperCase) & " ha sido guardada.", vbOKOnly + vbInformation
'            Unload Me
        Else
            Err.Raise "9999", "Guardando factura", Err.Description
        End If

    End If

    Exit Sub
err1:
    MsgBox "Ocurrió un error al guardar." & Chr(10) & "Controle: " & Chr(10) & "- Que todos los datos estén cargados." & Chr(10) & "- Que el Nº de cbte. no esté ya asignado." & Chr(10) & "- Que se haya seleccionado OPCIÓN DE TRANSFERENCIA." & Chr(10) & "ERROR: " & Err.Description, vbCritical + vbOKOnly
End Sub


Private Sub btnItemRemito_Click()
    If IsSomething(Factura.Cliente) Then
        '        Dim idEntrega As Long
        Dim f11 As New frmPlaneamientoRemitosListaProceso
        f11.idCliMostrar = Factura.Cliente.Id
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
                If detalle.DetalleRemitoId = redeta.Id Then
                    GoTo prox    'ya existe en la factura esa entrega aplicada, pasamos a la proxima
                End If
            Next

            Set detalle = New FacturaDetalle
            Set detalle.Factura = Factura
            detalle.idFactura = Factura.Id
            detalle.Cantidad = redeta.Cantidad
            detalle.detalle = redeta.VerElemento
            detalle.PorcentajeDescuento = 0
            detalle.Bruto = redeta.Valor
            Set Ot = DAOOrdenTrabajo.FindById(redeta.idpedido)
            If IsSomething(Ot) Then
                detalle.Bruto = MonedaConverter.Convertir(redeta.Valor, Ot.moneda.Id, Factura.moneda.Id)

                If Ot.moneda.Id <> Factura.moneda.Id Then
                    'MsgBox ("Tenemos monedas distintas")
                    MsgBox ("La Moneda de la OT incluída es: " & Ot.moneda.NombreCorto & vbCrLf & "" _
                          & "La Moneda del Comprobante que se está cargando es: " & Factura.moneda.NombreCorto & vbCrLf & "" _
                          & "Se procede a realizar la conversión correspondiente." & vbCrLf & "" _
                          & "El importe de la Moneda del comprobante es de: " & Factura.moneda.MonedaCambio.Cambio)

                Else
                    'MsgBox ("Tenemos las mismas monedas")

                End If

            End If
            detalle.IvaAplicado = True
            detalle.IBAplicado = True

            detalle.AplicadoARemito = True
            Set detalle.detalleRemito = redeta
            detalle.DetalleRemitoId = redeta.Id

            'detalle.detalleRemito = redeta.VerOrigen



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
                            Set detalleAnticipo = Factura.DetalleAnticipoOT(Ot.Id)
                            If Not IsSomething(detalleAnticipo) Then
                                Set detalleAnticipo = New FacturaDetalle
                                detalleAnticipo.OtIdAnticipo = Ot.Id
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
                            detalleAnticipo.Bruto = detalleAnticipo.Bruto + funciones.RedondearDecimales(detalle.total * Factura.moneda.Cambio)
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

        Set Factura.Cliente = DAOCliente.BuscarPorID(Me.cboCliente.ItemData(Me.cboCliente.ListIndex))
        Factura.Detalles = New Collection

        Set Factura.TipoIVA = Factura.Cliente.TipoIVA

        Dim tipos As New Collection

        Set tipos = DAOTipoFacturaDiscriminado.FindAllByFilter("id_iva= " & Factura.TipoIVA.idIVA & " and tipo_documento=" & Factura.TipoDocumento)

        Dim Tipo As clsTipoFacturaDiscriminado

        Me.cboTiposFactura.Enabled = True
        Me.Label6.Enabled = True

        Me.Label14.Enabled = True
        Me.Label15.Enabled = True
        Me.Label16.Enabled = True
        Me.txtNumero.Enabled = True    'Not factura.Tipo.PuntoVenta.EsElectronico
        Me.dtpFecha.Enabled = True
        Me.cboMoneda.Enabled = True

        Me.grpPercep.Enabled = True
        Me.grpInfo.Enabled = True

        Me.cboConceptosAIncluir.Enabled = True

        Me.cboTiposFactura.Clear


        Dim id_Default As Long
        id_Default = 0
        Dim nidx As Long
        'lleno el combo de tipos de factura y dejo el default marcado
        For Each Tipo In tipos
            cboTiposFactura.AddItem Tipo.descripcion
            nidx = cboTiposFactura.NewIndex
            cboTiposFactura.ItemData(nidx) = Tipo.Id
            If Tipo.PuntoVenta.default Then id_Default = nidx

        Next Tipo


        'pos on default pv
        If cboTiposFactura.ListCount > 0 Then
            'cboTiposFactura.ListIndex = id_Default
            cboTiposFactura.ListIndex = id_Default

        End If

        Factura.AlicuotaAplicada = Factura.TipoIVA.alicuota
        
        Set Factura.Cliente = DAOCliente.BuscarPorID(Factura.Cliente.Id)

        If IsSomething(Factura.Tipo.TipoFactura) Then
            Factura.EstaDiscriminada = Factura.Tipo.TipoFactura.Discrimina
            Me.lblTipoFactura.caption = Factura.Tipo.TipoFactura.Tipo
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

        Set Selecciones.OrdenTrabajo = Nothing

    End If
End Sub

Private Sub MostrarPercepcionIIBB()
'Me.lblBuscandoPercepcion.Visible = False
    Dim tabla As String
    If Me.cboPadron.ListIndex = 0 Then
        tabla = "IIBB2_Percepcion"
    Else
        tabla = "IIBB2_PercepcionAnt"
    End If

    Me.txtPercepcion.text = 0
    Me.lblVencido.Visible = False

    If Factura.Cliente.CUITValido Then
        'Me.lblBuscandoPercepcion.Visible = True
        DoEvents
        Dim rs As Recordset
        Set rs = conectar.RSFactory("SELECT * FROM sp_permisos." & tabla & " WHERE cuit='" & Factura.Cliente.Cuit & "'")
        'Me.lblBuscandoPercepcion.Visible = False
        DoEvents
        If IsSomething(rs) Then
            If Not rs.EOF And Not rs.BOF Then
                'Me.lblVencido.Visible = (Now() > CDate(ConvertirAFechaAfip(rs!FechaHasta)))
                Me.lblVencido.Visible = Format(Now, "dd/mm/yyyy") > CDate(ConvertirAFechaAfip(rs!FechaHasta))
                'Me.lblVencido.Visible = False
                Me.txtPercepcion.text = rs!alicuota
                Factura.AlicuotaPercepcionesIIBB = (rs!alicuota / 100) + 1
            End If
        End If
    End If
End Sub


Private Sub ASociarConcepto()
    If IsSomething(Factura) And Me.cboConceptosAIncluir.ListIndex <> -1 And Not dataLoading Then
        Factura.ConceptoIncluir = Me.cboConceptosAIncluir.ItemData(Me.cboConceptosAIncluir.ListIndex)
    End If
End Sub


Private Sub ConceptosIncuir()
    If IsSomething(Factura) And Me.cboConceptosAIncluir.ListIndex <> -1 And Not dataLoading Then

        ASociarConcepto

        'Me.lblFechaPagoCredito.Enabled = Factura.EsCredito Or (Factura.ConceptoIncluir = ConceptoProductoServicio Or Factura.ConceptoIncluir = ConceptoServicio)
        'Me.dtFechaPagoCredito.Enabled = Factura.EsCredito Or (Factura.ConceptoIncluir = ConceptoProductoServicio Or Factura.ConceptoIncluir = ConceptoServicio)

        Me.lblPeriodoFacturadoT.Enabled = Factura.ConceptoIncluir = ConceptoProductoServicio Or Factura.ConceptoIncluir = ConceptoServicio

        Me.lblPeriodoFacturadoD.Enabled = Factura.ConceptoIncluir = ConceptoProductoServicio Or Factura.ConceptoIncluir = ConceptoServicio
        Me.dtFechaPagoCreditoDesde.Enabled = Factura.ConceptoIncluir = ConceptoProductoServicio Or Factura.ConceptoIncluir = ConceptoServicio

        Me.lblPeriodoFacturadoH.Enabled = Factura.ConceptoIncluir = ConceptoProductoServicio Or Factura.ConceptoIncluir = ConceptoServicio
        Me.dtFechaPagoCreditoHasta.Enabled = Factura.ConceptoIncluir = ConceptoProductoServicio Or Factura.ConceptoIncluir = ConceptoServicio


    End If

    'fce_nemer_03062020_#133
    ' If Factura.ConceptoIncluir = ConceptoProducto Then
    'Me.lblFechaServDesde.Enabled = False
    'Me.dtFechaServDesde.Enabled = False
    'Me.lblFechaServHasta.Enabled = False
    ' Me.dtFechaServHasta.Enabled = False
    ' End If

End Sub


Private Sub cboConceptosAIncluir_Click()
    ConceptosIncuir

End Sub

Private Sub cboMoneda_Click()
    If IsSomething(Factura) And Me.cboMoneda.ListIndex <> -1 And Not dataLoading Then
        Set Factura.moneda = DAOMoneda.GetById(Me.cboMoneda.ItemData(Me.cboMoneda.ListIndex))
    End If
End Sub

Private Sub cboMonedaAjuste_Click()
    If IsSomething(Factura) And Me.cboMoneda.ListIndex <> -1 And Not dataLoading Then
        Factura.TipoCambioAjuste = DAOMoneda.GetById(Me.cboMonedaAjuste.ItemData(Me.cboMonedaAjuste.ListIndex)).Cambio
    End If
End Sub


Private Sub cboPadron_Click()

    If IsSomething(Factura.Cliente) And Not dataLoading Then
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

    If Me.cboTiposFactura.ListIndex = -1 Then Exit Sub
 
    Dim Id As Long

    Id = Me.cboTiposFactura.ItemData(Me.cboTiposFactura.ListIndex)

    Set Factura.Tipo = DAOTipoFacturaDiscriminado.FindById(Id)


    '1 11 19
    '    Me.lblCbuCredito.Visible = Factura.Tipo.PuntoVenta.EsCredito
    Me.Frame1.Enabled = Factura.esCredito
    Me.frmFC.Enabled = Factura.esCredito
    Me.Label22.Enabled = Factura.esCredito

    'Me.dtFechaPagoCredito.Enabled = Factura.EsCredito Or Factura.Tipo.PuntoVenta.CaeManual

    Me.dtFechaPagoCreditoDesde.Enabled = Factura.esCredito
    Me.dtFechaPagoCreditoHasta.Enabled = Factura.esCredito

    Me.cboCuentasCBU.Enabled = Factura.esCredito

    'Me.lblFechaPagoCredito.Enabled = Factura.EsCredito Or (Factura.ConceptoIncluir = ConceptoProductoServicio Or Factura.ConceptoIncluir = ConceptoServicio)
    'Me.dtFechaPagoCredito.Enabled = Factura.EsCredito Or (Factura.ConceptoIncluir = ConceptoProductoServicio Or Factura.ConceptoIncluir = ConceptoServicio)

    Me.LblCBU.Enabled = Factura.esCredito

    'fce_nemer_02062020_#113
    Me.lblPeriodoFacturadoT.Enabled = Factura.esCredito
    Me.lblPeriodoFacturadoD.Enabled = Factura.esCredito
    Me.lblPeriodoFacturadoH.Enabled = Factura.esCredito

    'fce_nemer_03062020_#133
    'Me.lblFechaPagoCredito.Enabled = Factura.Tipo.PuntoVenta.EsElectronico
    'Me.dtFechaPagoCredito.Enabled = Factura.Tipo.PuntoVenta.EsElectronico


    Me.lblEsCredito.caption = Factura.DescripcionCreditoAdicional

    Me.lblVerCbu.Visible = True
    If Not Factura.esCredito Then
        Me.lblVerCbu = "NO INFORMADO"
    End If


    If Factura.Id = 0 Then    'agregado para q no cambie el nro de factura cuando estoy en edicion yu elijo otro cliente
        '       Me.txtNumero.Enabled = Not Factura.Tipo.PuntoVenta.EsElectronico
        '       If Factura.Tipo.PuntoVenta.EsElectronico Then


        '           Dim Ult As String
        '          Me.txtNumero.text = "0000"    'ERPHelper.GetUltimoAutorizado(Factura.Tipo.PuntoVenta.PuntoVenta, Factura.Tipo.id)
        'Else


        Me.txtNumero.text = Format(DAOFactura.proximaFactura(Factura), "00000000")    'NuevoTipoDocumento, Factura.Tipo.TipoFactura.id), "0000")
        Me.txtNumero.Enabled = Not Factura.Tipo.PuntoVenta.EsElectronico Or Factura.Tipo.PuntoVenta.CaeManual


        '        End If
    Else
        If Factura.estado <> EstadoFacturaCliente.EnProceso Then
            Me.txtNumero.text = Format(Factura.numero, "00000000")   'Factura.NumeroFormateado

        Else
            If Factura.Tipo.PuntoVenta.CaeManual Then
                Me.txtNumero.text = Format(Factura.numero, "00000000")
            Else
                Me.txtNumero.text = Format(DAOFactura.proximaFactura(Factura), "00000000")
            End If
        End If
        '        If Factura.Tipo.PuntoVenta.EsElectronico Then
        '           Me.txtNumero.text = "0000"
        '        Else
        'Me.txtNumero.text = Format(DAOFactura.proximaFactura(factura.Tipo.id), "00000000") 'NuevoTipoDocumento, Factura.Tipo.TipoFactura.id), "0000")
        '        End If
    End If

    Me.txtNumero.Enabled = Not Factura.Tipo.PuntoVenta.EsElectronico Or Factura.Tipo.PuntoVenta.CaeManual

    ValidarEsCredito
End Sub

Private Sub chkEsCredito_Click()

    Factura.esCredito = Me.chkEsCredito.value

    ValidarEsCredito
    cboTiposFactura_Click
End Sub

Private Sub cmdNueva_Click()
    Dim frm2 As New frmAdminFacturasEdicion
    frm2.Show
End Sub


Private Sub dtFechaPagoCredito_Change()
    If Not dataLoading Then
        Factura.fechaPago = Me.dtFechaPagoCredito.value
    End If

    Me.txtDiasVenc = DateDiff("d", Me.dtpFecha, Me.dtFechaPagoCredito)

End Sub

'fce_nemer_28052020
Private Sub dtFechaPagoCreditoDesde_Change()
    If Not dataLoading Then
        Factura.FechaVtoDesde = Me.dtFechaPagoCreditoDesde.value
    End If
End Sub

'fce_nemer_28052020
Private Sub dtFechaPagoCreditoHasta_Change()
    If Not dataLoading Then
        Factura.FechaVtoHasta = Me.dtFechaPagoCreditoHasta.value
    End If
End Sub

'fce_nemer_02062020_#113
'Private Sub dtFechaServDesde_Change()
'   If Not dataLoading Then
'        Factura.FechaServDesde = Me.dtFechaServDesde.value
'    End If
'End Sub

'fce_nemer_02062020_#113
'Private Sub dtFechaServHasta_Change()
'   If Not dataLoading Then
'        Factura.FechaServHasta = Me.dtFechaServHasta.value
'    End If
'End Sub


Private Sub dtpFecha_Change()
    If Not dataLoading Then

        Factura.FechaEmision = Me.dtpFecha.value

        'fce_nemer_02062020_#113
        'Me.dtFechaServDesde.value = Me.dtpFecha.value
        'Me.dtFechaServHasta.value = Me.dtpFecha.value

        'fce_nemer_09062020
        txtDiasVenc_LostFocus

        Me.txtDiasVenc = DateDiff("d", Me.dtpFecha, Me.dtFechaPagoCredito)

    End If
End Sub


Private Sub Form_Load()
    Customize Me
    dataLoading = True
    
    CargarClientesEnCbo
    
    DAOMoneda.llenarComboXtremeSuite Me.cboMoneda
    DAOMoneda.llenarComboXtremeSuite Me.cboMonedaAjuste, True
    DAOCuentaBancaria.llenarComboCBU Me.cboCuentasCBU
    'Me.cboCuentasCBU.Visible = False

    'opcional 27

    Me.cboOpcional27.Clear
    Me.cboOpcional27.AddItem "TRANSFERENCIA AL SISTEMA DE CIRCULACION ABIERTA"
    Me.cboOpcional27.ItemData(Me.cboOpcional27.NewIndex) = 1

    Me.cboOpcional27.AddItem "AGENTE DE DEPOSITO COLECTIVO "
    Me.cboOpcional27.ItemData(Me.cboOpcional27.NewIndex) = 2

    If Not IsSomething(Factura) Then
        Set Factura = New Factura
        Factura.Detalles = New Collection
        Set Factura.Tipo = New clsTipoFacturaDiscriminado

        Me.cboConceptosAIncluir.ListIndex = funciones.PosIndexCbo(1, Me.cboConceptosAIncluir)

        Factura.Tipo.TipoDoc = NuevoTipoDocumento
        Me.caption = "Nueva " & StrConv(Factura.TipoDocumentoDescription, vbProperCase)
        Me.dtpFecha.value = Now

        Me.dtFechaPagoCredito.value = Now

        'fce_nemer_28052020
        Me.dtFechaPagoCreditoDesde.value = Now
        Me.dtFechaPagoCreditoHasta.value = Now

        'fce_nemer_02062020_#113
        'Me.dtFechaServDesde.value = Factura.FechaEmision
        'Me.dtFechaServHasta.value = Factura.FechaEmision


        '#218 consultar con karin cual quiere dejar por default
        Me.cboOpcional27.ListIndex = -1


        If Me.cboMoneda.ListIndex <> -1 Then
            Set Factura.moneda = DAOMoneda.GetById(Me.cboMoneda.ItemData(Me.cboMoneda.ListIndex))
        End If
    Else
        Me.caption = Factura.GetShortDescription(False, True)
    End If

    suscId = funciones.CreateGUID
    Channel.AgregarSuscriptor Me, TipoSuscripcion.FacturarRemitosDetalle_, True
    ' Me.lblBuscandoPercepcion.Visible = False
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

    If Factura.Id = 0 Then
        Factura.FechaEmision = Now

        Factura.fechaPago = Now

        'fce_nemer_28052020
        Factura.FechaVtoDesde = Now
        Factura.FechaVtoHasta = Now

        'fce_nemer_02062020_#113
        'Me.dtFechaServDesde.value = Factura.FechaEmision
        'Me.dtFechaServHasta.value = Factura.FechaEmision

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

        Me.caption = "Anticipo " & Me.caption
        Me.gridDetalles.Columns(1).EditType = jgexEditNone
        Me.gridDetalles.AllowDelete = False
        Factura.origenFacturado = OrigenFacturadoAnticipoOT
    End If


    Me.btnSeleccionarOT.Enabled = Factura.EsAnticipo Or EsAnticipo Or Factura.origenFacturado = OrigenFacturadoAnticipoOT
    Me.btnCrearItemConcepto.Enabled = Factura.EsAnticipo Or EsAnticipo Or Factura.origenFacturado = OrigenFacturadoAnticipoOT
    Me.btnGuardar.Enabled = Not ReadOnly Or EsAnticipo
    Me.btnItemRemito.Enabled = Not ReadOnly And Not EsAnticipo

    'fce_nemer_16062020
    Me.frmTextoAdicional.Enabled = Not ReadOnly
    Me.txtTextoAdicional.Enabled = Not ReadOnly
    Me.lblFechaPagoCredito.Enabled = Not ReadOnly
    Me.dtFechaPagoCredito.Enabled = Not ReadOnly
    Me.grpPercep.Enabled = Not ReadOnly
    Me.frmFC.Enabled = Not ReadOnly
    'Me.frmFCE.Enabled = Not ReadOnly
    Me.Label22.Enabled = Not ReadOnly
    Me.Frame1.Enabled = Not ReadOnly
    Me.lblConceptosAIncluir.Enabled = Not ReadOnly
    Me.cboConceptosAIncluir.Enabled = Not ReadOnly
    Me.Label6.Enabled = Not ReadOnly
    Me.cboTiposFactura.Enabled = Not ReadOnly
    Me.Label14.Enabled = Not ReadOnly

    If IsSomething(Factura) And IsSomething(Factura.Tipo) And IsSomething(Factura.Tipo.PuntoVenta) Then
        Me.txtNumero.Enabled = Not ReadOnly And (Not Factura.Tipo.PuntoVenta.EsElectronico Or Factura.Tipo.PuntoVenta.CaeManual)
    Else
        Me.txtNumero.Enabled = Not ReadOnly    'And Not factura.Tipo.PuntoVenta.EsElectronico
    End If

    '  Me.txtNumero.Enabled = Not ReadOnly 'And Not factura.Tipo.PuntoVenta.EsElectronico
    Me.Label15.Enabled = Not ReadOnly
    Me.dtpFecha.Enabled = Not ReadOnly
    Me.Label16.Enabled = Not ReadOnly
    Me.cboMoneda.Enabled = Not ReadOnly
    Me.grpInfo.Enabled = Not ReadOnly
    Me.txtReferencia.Enabled = Not ReadOnly
    Me.txtDiasVenc.Enabled = Not ReadOnly
    Me.txtTasaAjuste.Enabled = Not ReadOnly
    Me.txtCondObs.Enabled = Not ReadOnly
    Me.Label18.Enabled = Not ReadOnly
    Me.Label19.Enabled = Not ReadOnly
    Me.Label11.Enabled = Not ReadOnly
    Me.Label20.Enabled = Not ReadOnly
    Me.lblPadron.Enabled = Not ReadOnly
    Me.Label17.Enabled = Not ReadOnly
    Me.cboPadron.Enabled = Not ReadOnly
    Me.txtPercepcion.Enabled = Not ReadOnly
    Me.cboCliente.Enabled = Not ReadOnly
    'Me.grpDetalles.Enabled = Not ReadOnly
    'Me.gridDetalles.Enabled = Not ReadOnly
    Me.lblTipoFactura.Enabled = Not ReadOnly
    Me.Label1.Enabled = Not ReadOnly
    Me.Label3.Enabled = Not ReadOnly
    Me.Label2(0).Enabled = Not ReadOnly
    Me.Label4.Enabled = Not ReadOnly
    Me.Label5.Enabled = Not ReadOnly
    Me.Label12.Enabled = Not ReadOnly
    Me.Label7.Enabled = Not ReadOnly
    Me.lblCuit.Enabled = Not ReadOnly
    Me.lblIVA.Enabled = Not ReadOnly
    Me.lblDireccion.Enabled = Not ReadOnly
    Me.lblLocalidad.Enabled = Not ReadOnly
    Me.lblProvincia.Enabled = Not ReadOnly
    Me.lblCodPostal.Enabled = Not ReadOnly
    Me.lblNCND.Enabled = Not ReadOnly
    Me.lblAjuste.Enabled = Not ReadOnly
    Me.cboMonedaAjuste.Enabled = Not ReadOnly
    Me.grpTotales.Enabled = Not ReadOnly
    Me.btnItemsDescuentoAnticipo.Enabled = Not ReadOnly
    Me.Label9.Enabled = Not ReadOnly
    Me.lblSubTotal.Enabled = Not ReadOnly
    Me.lblPercepciones.Enabled = Not ReadOnly
    Me.lblIVATot.Enabled = Not ReadOnly
    Me.lblTotal.Enabled = Not ReadOnly
    Me.Label10.Enabled = Not ReadOnly
    Me.lblIva2.Enabled = Not ReadOnly
    Me.Label8.Enabled = Not ReadOnly

    Me.Label13.Enabled = Not ReadOnly
    Me.chkEsCredito.Enabled = Not ReadOnly
    Me.Label23.Enabled = Not ReadOnly

    Me.Frame1.Enabled = Not ReadOnly

    Me.btnSeleccionarOT.Enabled = Not ReadOnly
    Me.btnCrearItemConcepto.Enabled = Not ReadOnly

    ValidarEsCredito

    ''Me.caption = caption & " (" & Name & ")"

    'Me.cboCliente.ListIndex = "336"

End Sub

Private Sub LimpiarFactura()
    Me.txtNumero.text = vbNullString
    Me.lblTipoFactura.caption = vbNullString
    'Me.lblNCND.caption = vbNullString
    Me.txtReferencia.text = vbNullString
    Me.txtDiasVenc.text = vbNullString

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


    Me.txtNumero.Enabled = Not Factura.Tipo.PuntoVenta.EsElectronico Or Factura.Tipo.PuntoVenta.CaeManual And Not ReadOnly


    If Factura.estado <> EstadoFacturaCliente.EnProceso And Factura.Tipo.PuntoVenta.EsElectronico Then

        If LenB(Factura.CAE) > 0 Then
            Me.txtDetallesCAE.caption = "CAE " & Factura.CAE & " | CAE VTO " & Factura.CAEVto
            Me.txtNumero.Locked = True
        Else
            Me.txtDetallesCAE.caption = ""
        End If
    Else

        Me.txtDetallesCAE.caption = ""
    End If


    If IsSomething(Factura.Cliente) Then
        Me.cboCliente.ListIndex = funciones.PosIndexCbo(Factura.Cliente.Id, Me.cboCliente)
        MostrarCliente
    Else
        LimpiarCliente
    End If
    
    
    If Factura.moneda.Id = 0 Then
        Me.lblTC.caption = "Tipo Cambio: " & Factura.moneda.NombreCorto & " 1 "
    Else
        Me.lblTC.caption = "Tipo Cambio: " & Factura.moneda.NombreCorto & " " & Factura.CambioAPatron
    End If

    
    Me.cboMoneda.ListIndex = funciones.PosIndexCbo(Factura.moneda.Id, Me.cboMoneda)
    Me.cboConceptosAIncluir.ListIndex = funciones.PosIndexCbo(Factura.ConceptoIncluir, Me.cboConceptosAIncluir)

    Me.cboMonedaAjuste.ListIndex = funciones.PosIndexCbo(Factura.IdMonedaAjuste, Me.cboMonedaAjuste)

    If Factura.Id = 0 Then
        'creo que aaca no entra nunca
        '        Dim classA As New classAdministracion
        Me.txtNumero.text = Format(DAOFactura.proximaFactura(Factura))    'NuevoTipoDocumento, Factura.Tipo.TipoFactura.id), "0000")
    Else

        If Factura.estado = EstadoFacturaCliente.EnProceso Then

            If Factura.Tipo.PuntoVenta.CaeManual Then
                Me.txtNumero.text = Format(Factura.numero)
            Else

                Dim prox As Long
                prox = DAOFactura.proximaFactura(Factura)
                Factura.numero = prox
                Me.txtNumero.text = Format(prox)
            End If
        End If


        Set tipos = DAOTipoFacturaDiscriminado.FindAllByFilter("id_iva=" & Factura.TipoIVA.idIVA & " and tipo_Documento=" & Factura.TipoDocumento)    'acft.id IN (select TipoFactura FROM AdminConfigFacturas where idIVA = " & Factura.TipoIVA.idIVA & ")")

        Me.cboTiposFactura.Clear
        Dim T

        'lleno el combo de tipos de factura y dejo el primero x default

        For Each T In tipos
            cboTiposFactura.AddItem T.PuntoVenta.descripcion
            cboTiposFactura.ItemData(cboTiposFactura.NewIndex) = T.Id
        Next T

        Me.cboTiposFactura.ListIndex = funciones.PosIndexCbo(Factura.Tipo.Id, Me.cboTiposFactura)

        Me.txtNumero.text = Factura.numero
    End If

    Me.dtpFecha.value = Factura.FechaEmision
    Me.txtPercepcion.text = Round((Factura.AlicuotaPercepcionesIIBB - 1) * 100, 2)
    Me.txtDiasVenc.text = Factura.CantDiasPago
    Me.txtReferencia.text = Factura.OrdenCompra
    Me.txtCondObs.text = Factura.observaciones
    Me.txtTextoAdicional.text = Factura.TextoAdicional
    Me.lblTipoFactura.caption = Factura.Tipo.TipoFactura.Tipo

    Me.dtFechaPagoCredito = Factura.fechaPago

    'fce_nemer_28052020
    Me.dtFechaPagoCreditoDesde = Factura.FechaVtoDesde
    Me.dtFechaPagoCreditoHasta = Factura.FechaVtoHasta

    'fce_nemer_02062020_#113
    'Me.dtFechaServDesde = Factura.FechaServDesde
    'Me.dtFechaServHasta = Factura.FechaServHasta


    Me.txtTasaAjuste.text = Factura.TasaAjusteMensual
    ' Me.txtCbuCredito = Factura.CBU

    Dim c As CuentaBancaria

    If Factura.esCredito And LenB(Factura.CBU) > 0 Then

        Set c = DAOCuentaBancaria.FindByCBU(Factura.CBU)

        Me.chkEsCredito.value = Factura.esCredito


        If ReadOnly Then

            Me.cboCuentasCBU.Visible = IsSomething(c)

            Me.lblVerCbu.Visible = Not IsSomething(c)
            Me.chkEsCredito.Enabled = False
            Me.cboOpcional27.Enabled = False

            If IsSomething(c) Then
                Me.cboCuentasCBU.ListIndex = funciones.PosIndexCbo(c.Id, Me.cboCuentasCBU)
            Else
                Me.lblVerCbu = Factura.CBU
            End If
        Else
            Me.lblVerCbu.Visible = False
            If IsSomething(c) Then
                Me.cboCuentasCBU.ListIndex = funciones.PosIndexCbo(c.Id, Me.cboCuentasCBU)
            End If
            
        End If


    Else
        Me.cboCuentasCBU.Visible = False
        Me.lblVerCbu.Visible = True
        Me.lblVerCbu = "NO INFORMADO"
    End If
    
    ConceptosIncuir

    CargarDetalles

    Totalizar
    ValidarEsCredito

    Me.cboOpcional27.ListIndex = funciones.PosIndexCbo(Factura.Opcional27, Me.cboOpcional27)

End Sub


Private Sub ValidarEsCredito()

    Me.Frame1.Enabled = Factura.esCredito
    Me.frmFC.Enabled = Factura.esCredito

    Me.LblCBU.Enabled = Factura.esCredito
    Me.cboCuentasCBU.Enabled = Factura.esCredito
    Me.cboOpcional27.Enabled = Factura.esCredito
    Me.cboCuentasCBU.Visible = Factura.esCredito
    Me.txtCondObs.ListIndex = 1
    Me.Label22.Enabled = Factura.esCredito


    ConceptosIncuir
    If IsSomething(Factura) And IsSomething(Factura.Tipo) And IsSomething(Factura.Tipo.TipoFactura) Then
        If Factura.Tipo.TipoFactura.Tipo = "E" Then
            chkEsCredito.Enabled = False
        End If
    End If
End Sub


Private Sub Totalizar()

    Me.lblSubTotal.caption = Replace(FormatCurrency(funciones.FormatearDecimales(Factura.TotalSubTotal)), "$", "")
    Me.lblPercepciones.caption = Replace(FormatCurrency(funciones.FormatearDecimales(Factura.totalPercepciones)), "$", "")
    Me.lblIVATot.caption = Replace(FormatCurrency(funciones.FormatearDecimales(Factura.TotalIVA)), "$", "")
    Me.lblTotal.caption = Replace(FormatCurrency(funciones.FormatearDecimales(Factura.total)), "$", "")

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
    If Factura.Cliente Is Nothing Then Exit Sub
    Me.lblCuit.caption = Factura.Cliente.Cuit
    Me.lblIVA.caption = Factura.Cliente.TipoIVA.detalle
    Me.lblDireccion.caption = Factura.Cliente.Domicilio
    Me.lblLocalidad.caption = Factura.Cliente.localidad.nombre
    Me.lblCodPostal.caption = Factura.Cliente.CodigoPostal
    Me.lblCuitPais = Factura.Cliente.CuitPais
    Me.lblIdImpositivo = Factura.Cliente.IDImpositivo

    Me.lblProvincia = Factura.Cliente.provincia.nombre

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
    If Me.gridDetalles.row = -1 Then    'es nuevoF
        Me.gridDetalles.value(7) = True
        Me.gridDetalles.value(8) = True
    End If

    Cancel = Not IsNumeric(Me.gridDetalles.value(1)) Or Not IsNumeric(Me.gridDetalles.value(3)) Or Not IsNumeric(Me.gridDetalles.value(4))
End Sub

Private Sub gridDetalles_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 And ReadOnly And Me.gridDetalles.HitTest(x, y) = jgexHitTestConstants.jgexHTCell Then
        Dim row As Long: row = Me.gridDetalles.RowFromPoint(x, y)
        If row > 0 Then
            Set detalle = Factura.Detalles.item(Me.gridDetalles.rowIndex(row))
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
    it = Me.gridDetalles.rowIndex(gridDetalles.row)
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
    detalle.idFactura = Factura.Id
    detalle.Cantidad = Values(1)
    detalle.detalle = Values(2)
    detalle.PorcentajeDescuento = Values(3)
    detalle.Bruto = Values(4)
    detalle.IvaAplicado = Values(7)
    detalle.IBAplicado = Values(8)

    Factura.Detalles.Add detalle

    Totalizar
End Sub

Private Sub gridDetalles_UnboundDelete(ByVal rowIndex As Long, ByVal Bookmark As Variant)
    If rowIndex > 0 And Factura.Detalles.count > 0 Then
        Factura.Detalles.remove rowIndex
        Totalizar
    End If
End Sub

Private Sub gridDetalles_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If rowIndex <= Factura.Detalles.count Then
        Set detalle = Factura.Detalles.item(rowIndex)
        Values(1) = detalle.Cantidad
        Values(2) = detalle.detalle
        Values(3) = funciones.FormatearDecimales(detalle.PorcentajeDescuento)
        Values(4) = funciones.FormatearDecimales(detalle.Bruto)
        Values(5) = funciones.FormatearDecimales(detalle.SubTotal)
        Values(6) = funciones.FormatearDecimales(detalle.total)
        Values(7) = detalle.IvaAplicado
        Values(8) = detalle.IBAplicado
        Values(9) = detalle.VerOrigen

        Values(10) = detalle.idprovincia
    End If
End Sub

Private Sub gridDetalles_UnboundUpdate(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If rowIndex > 0 And Factura.Detalles.count > 0 Then
        Set detalle = Factura.Detalles.item(rowIndex)

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
                            detaFactRemito.DetalleRemitoId = redeta.Id
                            detaFactRemito.AplicadoARemito = True
                            redeta.Facturado = True
                            Dim transactionResult As Boolean: transactionResult = True
                            Dim q As String
                            conectar.BeginTransaction

                            q = "INSERT INTO AdminFacturasDetalleAplicacionRemitos (idFacturaDetalle, idRemitoDetalle, cantidadAplicada) VALUES (" & detaFactRemito.Id & ", " & redeta.Id & "  ,  " & redeta.Cantidad & ")"
                            transactionResult = transactionResult And conectar.execute(q)

                            Dim remi As Remito
                            Set remi = DAORemitoS.FindById(detaFactRemito.detalleRemito.Remito)
                            remi.EstadoFacturado = DAORemitoS.AnalizarEstadoFacturado(remi.Id)

                            transactionResult = transactionResult And DAORemitoS.Guardar(remi, False, False)
                            transactionResult = transactionResult And DAOFacturaDetalles.Guardar(detaFactRemito)
                            transactionResult = transactionResult And DAORemitoSDetalle.Guardar(redeta)
                            transactionResult = transactionResult And DAODetalleOrdenTrabajo.SaveCantidad(redeta.idDetallePedido, redeta.Cantidad, CantidadFacturada_, redeta.Valor, Factura.Id, Factura.moneda.Id, Factura.CambioAPatron, Factura.TipoCambioAjuste)

                            If transactionResult Then

                                conectar.CommitTransaction
                                remi.EstadoFacturado = DAORemitoS.AnalizarEstadoFacturado(remi.Id)
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

Private Sub Label23_Click()
    cboTiposFactura_Click

End Sub

Private Sub mnuAplicarDetalleRemito_Click()

    Set detaFactRemito = Nothing

    On Error Resume Next
    Dim f11 As New frmPlaneamientoRemitosListaProceso
    f11.idCliMostrar = Factura.Cliente.Id
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

        Set detaFactRemito = Factura.Detalles(Me.gridDetalles.rowIndex(Me.gridDetalles.row))

        frm.Show
    End If


End Sub

Private Sub btnSeleccionarOT_Click()

    If IsSomething(Factura.Cliente) Then
        Set Selecciones.OrdenTrabajo = Nothing

        Set frmPlaneamientoPedidosSeleccion.Cliente = Factura.Cliente

        frmPlaneamientoPedidosSeleccion.MostrarAnticipo = True

        frmPlaneamientoPedidosSeleccion.Show 1

        Dim Ot As OrdenTrabajo

        If IsSomething(Selecciones.OrdenTrabajo) Then

            If Not funciones.BuscarEnColeccion(Factura.OTsFacturadasAnticipo, CStr(Selecciones.OrdenTrabajo.Id)) Then
                Set Ot = DAOOrdenTrabajo.FindById(Selecciones.OrdenTrabajo.Id)
                Set Ot.Detalles = DAODetalleOrdenTrabajo.FindAllByOrdenTrabajo(Ot.Id, True, True, True)

                Factura.OTsFacturadasAnticipo.Add Ot, CStr(Ot.Id)

                Factura.Detalles = New Collection
                Factura.OrdenCompra = vbNullString
                Me.txtReferencia = "FACTURA POR ANTICIPO OT"
                Me.txtCondObs = vbNullString
                Me.txtDiasVenc = vbNullString

                Dim deta As FacturaDetalle

                For Each Ot In Factura.OTsFacturadasAnticipo
                    Me.txtReferencia.text = Me.txtReferencia.text & " " & Ot.IdFormateado

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

                    'bug #2 INICIO

                    If Ot.moneda.Id <> Factura.moneda.Id Then

                        If Factura.moneda.Patron Then

                            '                                        Set Monedas.MonedaConvertibles = Nothing
                            '                                        Set frmAdminComprobantesEmitidosCambioMoneda.Ot = Ot
                            '                                        Set frmAdminComprobantesEmitidosCambioMoneda.Factura = Factura
                            '                                        frmAdminComprobantesEmitidosCambioMoneda.Show 1

                            MsgBox ("La Moneda de la OT incluída es: " & Ot.moneda.NombreCorto & vbCrLf & "" _
                                  & "La Moneda del Comprobante que se está cargando es: " & Factura.moneda.NombreCorto & vbCrLf & "" _
                                  & "Se procede a realizar la conversión correspondiente." & vbCrLf & "" _
                                  & "Cálculo:" & vbCrLf & "Total de OT: " & Ot.total & vbCrLf & " * Valor de Moneda de OT: " & Ot.moneda.Cambio & vbCrLf & " * % Anticipo: " & Ot.Anticipo & vbCrLf & "/ 100")
                            deta.Bruto = deta.Bruto + funciones.RedondearDecimales((Ot.total * Ot.moneda.Cambio * Ot.Anticipo) / 100)

                        Else
                            MsgBox ("La Moneda de la OT incluída es: " & Ot.moneda.NombreCorto & vbCrLf & "" _
                                  & "La Moneda del Comprobante que se está cargando es: " & Factura.moneda.NombreCorto & vbCrLf & "" _
                                  & "Se procede a realizar la conversión correspondiente:" & vbCrLf & "" _
                                  & "Cálculo:" & vbCrLf & "Total de OT: " & Ot.total & vbCrLf & " * Valor de Moneda de Comprobante: " & Factura.moneda.Cambio & vbCrLf & " * % Anticipo: " & Ot.Anticipo & vbCrLf & "/ 100")
                            deta.Bruto = deta.Bruto + funciones.RedondearDecimales((Ot.total * Factura.moneda.Cambio * Ot.Anticipo) / 100)
                        End If

                    Else
                        MsgBox ("La Moneda de la OT es: " & Ot.moneda.NombreCorto & vbCrLf & "" _
                              & "La Moneda del Comprobante es: " & Factura.moneda.NombreCorto & vbCrLf & "" _
                              & "No se realiza conversión:" & vbCrLf & "" _
                              & "Cálculo:" & vbCrLf & " Total de OT: " & Ot.total & vbCrLf & " * % Anticipo: " & Ot.Anticipo & vbCrLf & "/ 100")
                        deta.Bruto = deta.Bruto + funciones.RedondearDecimales(((Ot.total * Ot.Anticipo) / 100) / Factura.moneda.Cambio)

                    End If

                Next Ot

                CargarDetalles

                Totalizar

            End If
        End If
    Else
        MsgBox "Debe seleccionar un cliente para poder operar.", vbExclamation
    End If
End Sub


'fce_nemer_09062020
Public Sub txtDiasVenc_LostFocus()
    If Me.txtDiasVenc = vbNullString Then
        Me.txtDiasVenc = 0
    End If

    Me.dtFechaPagoCredito.value = DateAdd("d", Me.txtDiasVenc, Me.dtpFecha)

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


Private Sub txtTextoAdicional_Change()
    Dim texto As String
    Dim caracteresRestantes As Integer
    
    ' Obtén el texto actual del TextBox
    texto = Me.txtTextoAdicional.text
    
    ' Calcula la cantidad de caracteres restantes
    caracteresRestantes = 300 - Len(texto)
    
    ' Actualiza el contenido del Label
    Me.lblCaracteresRestantes.caption = "Caracteres restantes: " & caracteresRestantes
End Sub

Private Sub CargarClientesEnCbo()
    DAOCliente.llenarComboXtremeSuite Me.cboCliente
    cboCliente.ListIndex = -1
End Sub
