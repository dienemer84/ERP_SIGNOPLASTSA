VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmAdminFacturasProformaEdicion 
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
   Icon            =   "frmAdminFacturasProformaEdicion.frx":0000
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
      TabIndex        =   47
      Top             =   1680
      Width           =   5535
      Begin VB.Label lblIdImpositivo 
         Height          =   255
         Left            =   2880
         TabIndex        =   62
         Top             =   2040
         Width           =   2415
      End
      Begin VB.Label lblCuitPais 
         Caption         =   "Label24"
         Height          =   255
         Left            =   1200
         TabIndex        =   61
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Provincia:"
         Height          =   195
         Left            =   360
         TabIndex        =   59
         Top             =   1470
         Width           =   705
      End
      Begin VB.Label lblProvincia 
         AutoSize        =   -1  'True
         Caption         =   "2343242"
         Height          =   195
         Left            =   1245
         TabIndex        =   58
         Top             =   1470
         Width           =   630
      End
      Begin VB.Label lblCodPostal 
         AutoSize        =   -1  'True
         Caption         =   "2343242"
         Height          =   195
         Left            =   1245
         TabIndex        =   57
         Top             =   1770
         Width           =   630
      End
      Begin VB.Label lblLocalidad 
         AutoSize        =   -1  'True
         Caption         =   "HHHHHH"
         Height          =   195
         Left            =   1245
         TabIndex        =   56
         Top             =   1155
         Width           =   630
      End
      Begin VB.Label lblDireccion 
         Caption         =   "RIVAD 3242"
         Height          =   195
         Left            =   1245
         TabIndex        =   55
         Top             =   840
         Width           =   4095
      End
      Begin VB.Label lblIVA 
         AutoSize        =   -1  'True
         Caption         =   "23"
         Height          =   195
         Left            =   1245
         TabIndex        =   54
         Top             =   555
         Width           =   180
      End
      Begin VB.Label lblCuit 
         AutoSize        =   -1  'True
         Caption         =   "23-30279550-9"
         Height          =   195
         Left            =   1245
         TabIndex        =   53
         Top             =   240
         Width           =   1110
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cod Postal:"
         Height          =   195
         Left            =   240
         TabIndex        =   52
         Top             =   1770
         Width           =   825
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Localidad:"
         Height          =   195
         Left            =   345
         TabIndex        =   51
         Top             =   1155
         Width           =   720
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "CUIT:"
         Height          =   195
         Left            =   645
         TabIndex        =   50
         Top             =   255
         Width           =   420
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Direccion:"
         Height          =   195
         Left            =   360
         TabIndex        =   49
         Top             =   840
         Width           =   705
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "IVA:"
         Height          =   195
         Index           =   0
         Left            =   750
         TabIndex        =   48
         Top             =   555
         Width           =   315
      End
   End
   Begin XtremeSuiteControls.PushButton btnExportarContenido 
      Height          =   615
      Left            =   15480
      TabIndex        =   46
      Top             =   6360
      Width           =   2055
      _Version        =   786432
      _ExtentX        =   3625
      _ExtentY        =   1085
      _StockProps     =   79
      Caption         =   "Exportar Detalle de Cbte a Excel"
      UseVisualStyle  =   -1  'True
   End
   Begin VB.Frame grpDatosCliente 
      Caption         =   "Cliente"
      Height          =   1575
      Left            =   120
      TabIndex        =   41
      Top             =   120
      Width           =   5535
      Begin XtremeSuiteControls.PushButton btnCrearCliente 
         Height          =   375
         Left            =   240
         TabIndex        =   44
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
         TabIndex        =   42
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
         TabIndex        =   43
         Top             =   240
         Width           =   1410
      End
   End
   Begin VB.Frame frmTextoAdicional 
      Caption         =   "Texto Adicional (Limite de 300 caracteres)"
      Height          =   2895
      Left            =   11640
      TabIndex        =   22
      Top             =   120
      Width           =   6015
      Begin VB.TextBox txtTextoAdicional 
         Height          =   2175
         Left            =   120
         MaxLength       =   300
         MultiLine       =   -1  'True
         TabIndex        =   23
         Top             =   240
         Width           =   5775
      End
      Begin VB.Label lblCaracteresRestantes 
         Caption         =   "Caracteres restantes: "
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   2520
         Width           =   3375
      End
   End
   Begin XtremeSuiteControls.GroupBox grpTotales 
      Height          =   1575
      Left            =   11640
      TabIndex        =   11
      Top             =   3000
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
         TabIndex        =   60
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   12
         Top             =   930
         Width           =   1080
      End
   End
   Begin XtremeSuiteControls.GroupBox grpDatos 
      Height          =   4335
      Left            =   5760
      TabIndex        =   4
      Top             =   120
      Width           =   5835
      _Version        =   786432
      _ExtentX        =   10292
      _ExtentY        =   7646
      _StockProps     =   79
      Caption         =   "Datos del Comprobante"
      Appearance      =   4
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
         TabIndex        =   0
         Text            =   "999999"
         Top             =   1800
         Width           =   3090
      End
      Begin XtremeSuiteControls.DateTimePicker dtpFecha 
         Height          =   405
         Left            =   2520
         TabIndex        =   1
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
      Begin XtremeSuiteControls.ComboBox cboMoneda 
         Height          =   405
         Left            =   2520
         TabIndex        =   2
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
         TabIndex        =   31
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
         Format          =   66977793
         CurrentDate     =   43967
      End
      Begin VB.Label Label2 
         Caption         =   "PROFORMA"
         Height          =   2295
         Index           =   1
         Left            =   120
         TabIndex        =   63
         Top             =   1560
         Width           =   855
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
         TabIndex        =   30
         Top             =   3720
         Width           =   2535
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFDBBF&
         DrawMode        =   9  'Not Mask Pen
         X1              =   14640
         X2              =   3120
         Y1              =   9000
         Y2              =   9000
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
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   7
         Top             =   1860
         Width           =   945
      End
      Begin VB.Label lblNCND 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "N/D"
         Height          =   195
         Left            =   240
         TabIndex        =   6
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
         TabIndex        =   5
         Top             =   270
         Width           =   645
      End
   End
   Begin XtremeSuiteControls.GroupBox grpDetalles 
      Height          =   3795
      Left            =   120
      TabIndex        =   10
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
         TabIndex        =   64
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
         Column(1)       =   "frmAdminFacturasProformaEdicion.frx":000C
         Column(2)       =   "frmAdminFacturasProformaEdicion.frx":0144
         Column(3)       =   "frmAdminFacturasProformaEdicion.frx":0238
         Column(4)       =   "frmAdminFacturasProformaEdicion.frx":0350
         Column(5)       =   "frmAdminFacturasProformaEdicion.frx":0470
         Column(6)       =   "frmAdminFacturasProformaEdicion.frx":05B0
         Column(7)       =   "frmAdminFacturasProformaEdicion.frx":06E4
         Column(8)       =   "frmAdminFacturasProformaEdicion.frx":080C
         Column(9)       =   "frmAdminFacturasProformaEdicion.frx":093C
         Column(10)      =   "frmAdminFacturasProformaEdicion.frx":0A4C
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmAdminFacturasProformaEdicion.frx":0B44
         FormatStyle(2)  =   "frmAdminFacturasProformaEdicion.frx":0C6C
         FormatStyle(3)  =   "frmAdminFacturasProformaEdicion.frx":0D1C
         FormatStyle(4)  =   "frmAdminFacturasProformaEdicion.frx":0DD0
         FormatStyle(5)  =   "frmAdminFacturasProformaEdicion.frx":0EA8
         FormatStyle(6)  =   "frmAdminFacturasProformaEdicion.frx":0F60
         ImageCount      =   0
         PrinterProperties=   "frmAdminFacturasProformaEdicion.frx":1040
      End
   End
   Begin XtremeSuiteControls.PushButton btnGuardar 
      Height          =   600
      Left            =   11880
      TabIndex        =   3
      Top             =   6360
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
   Begin XtremeSuiteControls.GroupBox grpPercep 
      Height          =   1695
      Left            =   6720
      TabIndex        =   24
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
         TabIndex        =   25
         Top             =   840
         Width           =   2715
      End
      Begin XtremeSuiteControls.ComboBox cboPadron 
         Height          =   315
         Left            =   1560
         TabIndex        =   26
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
         TabIndex        =   29
         Top             =   1320
         Visible         =   0   'False
         Width           =   2670
      End
      Begin XtremeSuiteControls.Label lblPadron 
         Height          =   195
         Left            =   240
         TabIndex        =   28
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
         TabIndex        =   27
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
      TabIndex        =   32
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
         Left            =   1035
         TabIndex        =   36
         Top             =   1200
         Width           =   5295
      End
      Begin VB.TextBox txtTasaAjuste 
         Height          =   300
         Left            =   5160
         TabIndex        =   35
         Top             =   720
         Width           =   1200
      End
      Begin VB.TextBox txtDiasVenc 
         Height          =   300
         Left            =   2160
         TabIndex        =   34
         Top             =   720
         Width           =   1080
      End
      Begin VB.TextBox txtReferencia 
         Height          =   300
         Left            =   1395
         TabIndex        =   33
         Top             =   240
         Width           =   4935
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "% Tasa ajuste mensual:"
         Height          =   195
         Left            =   3350
         TabIndex        =   40
         Top             =   780
         Width           =   1740
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Condicion:"
         Height          =   195
         Left            =   240
         TabIndex        =   39
         Top             =   1260
         Width           =   750
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cant Días Vencimiento FF:"
         Height          =   195
         Left            =   240
         TabIndex        =   38
         Top             =   780
         Width           =   1875
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "OC / Referencia:"
         Height          =   195
         Left            =   120
         TabIndex        =   37
         Top             =   300
         Width           =   1215
      End
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "23-30279550-9"
      Height          =   195
      Left            =   480
      TabIndex        =   21
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
      Left            =   11760
      TabIndex        =   20
      Top             =   10800
      Width           =   5385
   End
   Begin VB.Menu mnuDetalles 
      Caption         =   "mnuDetalles"
      Visible         =   0   'False
      Begin VB.Menu mnuAplicarDetalleRemito 
         Caption         =   "Aplicar detalle de remito"
      End
   End
End
Attribute VB_Name = "frmAdminFacturasProformaEdicion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tipos As Collection
Dim Tipo As clsTipoFactura
Implements ISuscriber
Private FacturaProforma As clsFacturaProforma
Private dataLoading As Boolean
Private detalle As clsFacturaProformaDetalle
Private suscId As String
'Private ErrorAfip As Boolean
Public NuevoTipoDocumento As tipoDocumentoContable
Public EsAnticipo As Boolean

Public ReadOnly As Boolean

Private detaFactRemito As clsFacturaProformaDetalle

Public Property Let idFactura(value As Long)
    Set FacturaProforma = DAOFacturaProforma.FindById(value, True, True)
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

    xlWorksheet.Cells(1, 1).value = "Detalle de Cbte " + FacturaProforma.GetShortDescription(False, False)

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

    Dim deta As clsFacturaProformaDetalle

    'DEFINE EL CONTADOR DEL PROGRESSBAR Y LO INICIA EN 0
    Dim d As Long
    d = 0

    For Each deta In FacturaProforma.Detalles
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
        MsgBox "Todavia esta editando algun detalle de la FacturaProforma.", vbExclamation
        Exit Sub
    End If



    If Not FacturaProforma.Cliente.CUITValido Or Not FacturaProforma.Cliente.ValidoRemitoFactura Then
        MsgBox "El cliente no es valido para poder facturar.", vbExclamation + vbOKOnly
        Exit Sub
    End If

    FacturaProforma.observaciones = Me.txtCondObs.text

    If LenB(FacturaProforma.numero) = 0 Or _
       LenB(FacturaProforma.OrdenCompra) = 0 Or _
       FacturaProforma.observaciones = "" Or _
       FacturaProforma.CantDiasPago = 0 Then
    MsgBox "El Comprobante debe poseer Nº, referencia, cantidad dias de vto. de FF y condición cargada.", vbExclamation + vbOKOnly

    End If


        Dim deta As clsFacturaProformaDetalle
        'Dim ot As OrdenTrabajo
        For Each deta In FacturaProforma.Detalles

        Next deta

        Dim c As CuentaBancaria

        If IsSomething(c) Then
            FacturaProforma.CBU = c.CBU
        End If

        FacturaProforma.observaciones = Me.txtCondObs.text
        FacturaProforma.TextoAdicional = Me.txtTextoAdicional

        FacturaProforma.fechaPago = Me.dtFechaPagoCredito.value

        If DAOFacturaProforma.Save(FacturaProforma, True) Then
            MsgBox "La Proforma ha sido guardada.", vbOKOnly + vbInformation
            Unload Me
        Else
            Err.Raise "9999", "Guardando Proforma", Err.Description
        End If

    Exit Sub
err1:
    MsgBox "Ocurrió un error al guardar." & Chr(10) & "Controle: " & Chr(10) & "- Que todos los datos estén cargados." & Chr(10) & "- Que el Nº de cbte. no esté ya asignado." & Chr(10) & "- Que se haya seleccionado OPCIÓN DE TRANSFERENCIA." & Chr(10) & "ERROR: " & Err.Description, vbCritical + vbOKOnly
End Sub


Private Sub cboCliente_Click()
    If IsSomething(FacturaProforma) And Me.cboCliente.ListIndex <> -1 And Not dataLoading Then

        Set FacturaProforma.Cliente = DAOCliente.BuscarPorID(Me.cboCliente.ItemData(Me.cboCliente.ListIndex))
        FacturaProforma.Detalles = New Collection

        Set FacturaProforma.TipoIVA = FacturaProforma.Cliente.TipoIVA

        Dim tipos As New Collection

'''        Set tipos = DAOTipoFacturaDiscriminado.FindAllByFilter("id_iva= " & FacturaProforma.TipoIVA.idIVA & " and tipo_documento=" & FacturaProforma.TipoDocumento)

        Dim Tipo As clsTipoFacturaDiscriminado

        Me.Label14.Enabled = True
        Me.Label15.Enabled = True
        Me.Label16.Enabled = True
        Me.txtNumero.Enabled = True    'Not factura.Tipo.PuntoVenta.EsElectronico
        Me.dtpFecha.Enabled = True
        Me.cboMoneda.Enabled = True

        Me.grpPercep.Enabled = True
        Me.grpInfo.Enabled = True

        Dim id_Default As Long
        id_Default = 0
        Dim nidx As Long
        'lleno el combo de tipos de factura y dejo el default marcado
        For Each Tipo In tipos
            If Tipo.PuntoVenta.default Then id_Default = nidx

        Next Tipo


        'pos on default pv

        FacturaProforma.AlicuotaAplicada = FacturaProforma.TipoIVA.alicuota
        
        Set FacturaProforma.Cliente = DAOCliente.BuscarPorID(FacturaProforma.Cliente.id)

''        If IsSomething(FacturaProforma.Tipo.TipoFactura) Then
''            FacturaProforma.EstaDiscriminada = FacturaProforma.Tipo.TipoFactura.Discrimina
''            Me.lblTipoFactura.caption = FacturaProforma.Tipo.TipoFactura.Tipo
''        Else
''            Me.lblTipoFactura.caption = vbNullString
''            Me.txtNumero.text = 0
''        End If


'''        Me.lblNCND.Visible = (FacturaProforma.TipoDocumento <> tipoDocumentoContable.Factura)
'''        Me.lblNCND.caption = FacturaProforma.GetShortDescription(True, True)


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

    If FacturaProforma.Cliente.CUITValido Then
        'Me.lblBuscandoPercepcion.Visible = True
        DoEvents
        Dim rs As Recordset
        Set rs = conectar.RSFactory("SELECT * FROM sp_permisos." & tabla & " WHERE cuit='" & FacturaProforma.Cliente.Cuit & "'")
        'Me.lblBuscandoPercepcion.Visible = False
        DoEvents
        If IsSomething(rs) Then
            If Not rs.EOF And Not rs.BOF Then
                'Me.lblVencido.Visible = (Now() > CDate(ConvertirAFechaAfip(rs!FechaHasta)))
                Me.lblVencido.Visible = Format(Now, "dd/mm/yyyy") > CDate(ConvertirAFechaAfip(rs!FechaHasta))
                'Me.lblVencido.Visible = False
                Me.txtPercepcion.text = rs!alicuota
                FacturaProforma.AlicuotaPercepcionesIIBB = (rs!alicuota / 100) + 1
            End If
        End If
    End If
End Sub


Private Sub cboMoneda_Click()
    If IsSomething(FacturaProforma) And Me.cboMoneda.ListIndex <> -1 And Not dataLoading Then
        Set FacturaProforma.moneda = DAOMoneda.GetById(Me.cboMoneda.ItemData(Me.cboMoneda.ListIndex))
    End If
End Sub


Private Sub cboPadron_Click()

    If IsSomething(FacturaProforma.Cliente) And Not dataLoading Then
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

    If Me.cboTiposFacturaProforma.ListIndex = -1 Then Exit Sub
 
    Dim id As Long

    id = Me.cboTiposFacturaProforma.ItemData(Me.cboTiposFacturaProforma.ListIndex)

    Set FacturaProforma.Tipo = DAOTipoFacturaDiscriminado.FindById(id)


    '1 11 19
    '    Me.lblCbuCredito.Visible = FacturaProforma.Tipo.PuntoVenta.EsCredito
    Me.Frame1.Enabled = FacturaProforma.esCredito
    Me.frmFC.Enabled = FacturaProforma.esCredito
    Me.Label22.Enabled = FacturaProforma.esCredito

    'Me.dtFechaPagoCredito.Enabled = FacturaProforma.EsCredito Or FacturaProforma.Tipo.PuntoVenta.CaeManual

    Me.dtFechaPagoCreditoDesde.Enabled = FacturaProforma.esCredito
    Me.dtFechaPagoCreditoHasta.Enabled = FacturaProforma.esCredito

    Me.cboCuentasCBU.Enabled = FacturaProforma.esCredito

    'Me.lblFechaPagoCredito.Enabled = FacturaProforma.EsCredito Or (FacturaProforma.ConceptoIncluir = ConceptoProductoServicio Or FacturaProforma.ConceptoIncluir = ConceptoServicio)
    'Me.dtFechaPagoCredito.Enabled = FacturaProforma.EsCredito Or (FacturaProforma.ConceptoIncluir = ConceptoProductoServicio Or FacturaProforma.ConceptoIncluir = ConceptoServicio)

    Me.LblCBU.Enabled = FacturaProforma.esCredito

    'fce_nemer_02062020_#113
    Me.lblPeriodoFacturadoT.Enabled = FacturaProforma.esCredito
    Me.lblPeriodoFacturadoD.Enabled = FacturaProforma.esCredito
    Me.lblPeriodoFacturadoH.Enabled = FacturaProforma.esCredito

    'fce_nemer_03062020_#133
    'Me.lblFechaPagoCredito.Enabled = FacturaProforma.Tipo.PuntoVenta.EsElectronico
    'Me.dtFechaPagoCredito.Enabled = FacturaProforma.Tipo.PuntoVenta.EsElectronico


    Me.lblEsCredito.caption = FacturaProforma.DescripcionCreditoAdicional

    Me.lblVerCbu.Visible = True
    If Not FacturaProforma.esCredito Then
        Me.lblVerCbu = "NO INFORMADO"
    End If


    If FacturaProforma.id = 0 Then    'agregado para q no cambie el nro de factura cuando estoy en edicion yu elijo otro cliente
        '       Me.txtNumero.Enabled = Not FacturaProforma.Tipo.PuntoVenta.EsElectronico
        '       If FacturaProforma.Tipo.PuntoVenta.EsElectronico Then


        '           Dim Ult As String
        '          Me.txtNumero.text = "0000"    'ERPHelper.GetUltimoAutorizado(FacturaProforma.Tipo.PuntoVenta.PuntoVenta, FacturaProforma.Tipo.id)
        'Else


        Me.txtNumero.text = Format(DAOFacturaProforma.proximaFactura(Factura), "00000000")    'NuevoTipoDocumento, FacturaProforma.Tipo.TipoFacturaProforma.id), "0000")
        Me.txtNumero.Enabled = Not FacturaProforma.Tipo.PuntoVenta.EsElectronico Or FacturaProforma.Tipo.PuntoVenta.CaeManual


        '        End If
    Else
        If FacturaProforma.estado <> EstadoFacturaCliente.EnProceso Then
            Me.txtNumero.text = Format(FacturaProforma.numero, "00000000")   'FacturaProforma.NumeroFormateado

        Else
            If FacturaProforma.Tipo.PuntoVenta.CaeManual Then
                Me.txtNumero.text = Format(FacturaProforma.numero, "00000000")
            Else
                Me.txtNumero.text = Format(DAOFacturaProforma.proximaFactura(Factura), "00000000")
            End If
        End If
        '        If FacturaProforma.Tipo.PuntoVenta.EsElectronico Then
        '           Me.txtNumero.text = "0000"
        '        Else
        'Me.txtNumero.text = Format(DAOFacturaProforma.proximaFactura(factura.Tipo.id), "00000000") 'NuevoTipoDocumento, FacturaProforma.Tipo.TipoFacturaProforma.id), "0000")
        '        End If
    End If

    Me.txtNumero.Enabled = Not FacturaProforma.Tipo.PuntoVenta.EsElectronico Or FacturaProforma.Tipo.PuntoVenta.CaeManual

    ValidarEsCredito
End Sub


Private Sub cmdNueva_Click()
    Dim frm2 As New frmAdminFacturasProformaEdicion
    frm2.Show
End Sub


Private Sub dtFechaPagoCredito_Change()
    If Not dataLoading Then
        FacturaProforma.fechaPago = Me.dtFechaPagoCredito.value
    End If

    Me.txtDiasVenc = DateDiff("d", Me.dtpFecha, Me.dtFechaPagoCredito)

End Sub


'fce_nemer_28052020
Private Sub dtFechaPagoCreditoDesde_Change()
    If Not dataLoading Then
        FacturaProforma.FechaVtoDesde = Me.dtFechaPagoCreditoDesde.value
    End If
End Sub


'fce_nemer_28052020
Private Sub dtFechaPagoCreditoHasta_Change()
    If Not dataLoading Then
        FacturaProforma.FechaVtoHasta = Me.dtFechaPagoCreditoHasta.value
    End If
End Sub

'fce_nemer_02062020_#113
'Private Sub dtFechaServDesde_Change()
'   If Not dataLoading Then
'        FacturaProforma.FechaServDesde = Me.dtFechaServDesde.value
'    End If
'End Sub

'fce_nemer_02062020_#113
'Private Sub dtFechaServHasta_Change()
'   If Not dataLoading Then
'        FacturaProforma.FechaServHasta = Me.dtFechaServHasta.value
'    End If
'End Sub


Private Sub dtpFecha_Change()
    If Not dataLoading Then

        FacturaProforma.FechaEmision = Me.dtpFecha.value

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
    
    If Not IsSomething(FacturaProforma) Then
        Set FacturaProforma = New clsFacturaProforma
        FacturaProforma.Detalles = New Collection
        Set FacturaProforma.Tipo = New clsTipoFacturaDiscriminado

        FacturaProforma.Tipo.TipoDoc = NuevoTipoDocumento
        Me.caption = "Nueva " & StrConv(FacturaProforma.TipoDocumentoDescription, vbProperCase)
        Me.dtpFecha.value = Now

        Me.dtFechaPagoCredito.value = Now

        If Me.cboMoneda.ListIndex <> -1 Then
            Set FacturaProforma.moneda = DAOMoneda.GetById(Me.cboMoneda.ItemData(Me.cboMoneda.ListIndex))
        End If
    Else
        Me.caption = FacturaProforma.GetShortDescription(False, True)
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

    If FacturaProforma.id = 0 Then
        FacturaProforma.FechaEmision = Now

        FacturaProforma.fechaPago = Now


        FacturaProforma.estado = EstadoFacturaCliente.EnProceso
        LimpiarFactura
        LimpiarCliente
        LimpiarTotales
        
    Else
        CargarFactura
    End If

    dataLoading = False

    Me.grpDatos.Enabled = Not ReadOnly

    If ReadOnly Then
        Me.gridDetalles.EditMode = jgexEditModeOff
        Me.gridDetalles.AllowAddNew = False
        Me.gridDetalles.ReadOnly = True
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Me.btnCrearCliente.Enabled = False
        Me.grpDatosCliente.Enabled = False
        Me.Frame.Enabled = False

        Dim mon_ajuste As clsMoneda
        Set mon_ajuste = DAOMoneda.GetById(FacturaProforma.IdMonedaAjuste)

        Dim colu As JSColumn
        For Each colu In Me.gridDetalles.Columns
            colu.EditType = jgexEditNone
        Next colu
    End If

    If EsAnticipo Or FacturaProforma.EsAnticipo Then

        Me.caption = "Anticipo " & Me.caption
        Me.gridDetalles.Columns(1).EditType = jgexEditNone
        Me.gridDetalles.AllowDelete = False
        FacturaProforma.origenFacturado = OrigenFacturadoAnticipoOT
    End If

    Me.btnGuardar.Enabled = Not ReadOnly Or EsAnticipo

    'fce_nemer_16062020
    Me.frmTextoAdicional.Enabled = Not ReadOnly
    Me.txtTextoAdicional.Enabled = Not ReadOnly
    Me.lblFechaPagoCredito.Enabled = Not ReadOnly
    Me.dtFechaPagoCredito.Enabled = Not ReadOnly
    Me.grpPercep.Enabled = Not ReadOnly
    'Me.frmFCE.Enabled = Not ReadOnly


'    If IsSomething(Factura) And IsSomething(FacturaProforma.Tipo) And IsSomething(FacturaProforma.Tipo.PuntoVenta) Then
'        Me.txtNumero.Enabled = Not ReadOnly And (Not FacturaProforma.Tipo.PuntoVenta.EsElectronico Or FacturaProforma.Tipo.PuntoVenta.CaeManual)
'    Else
'        Me.txtNumero.Enabled = Not ReadOnly    'And Not factura.Tipo.PuntoVenta.EsElectronico
'    End If

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
'    Me.Label2(0).Enabled = Not ReadOnly
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

    Me.grpTotales.Enabled = Not ReadOnly

    Me.Label9.Enabled = Not ReadOnly
    Me.lblSubTotal.Enabled = Not ReadOnly
    Me.lblPercepciones.Enabled = Not ReadOnly
    Me.lblIVATot.Enabled = Not ReadOnly
    Me.lblTotal.Enabled = Not ReadOnly
    Me.Label10.Enabled = Not ReadOnly
    Me.lblIva2.Enabled = Not ReadOnly
    Me.Label8.Enabled = Not ReadOnly

    ''Me.caption = caption & " (" & Name & ")"

'    Me.cboCliente.ListIndex = "336"
'    Me.txtCondObs = "CONDICION PRUEBA"
'    Me.txtDiasVenc = 1
'    Me.txtNumero = 100
'    Me.txtTasaAjuste = 2.5
'    Me.txtReferencia = "OC DE PRUEBA 1"
    

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

    If Not IsSomething(FacturaProforma) Then Exit Sub
'    Me.cboTiposFacturaProforma.Enabled = Not (FacturaProforma.estado = EstadoFacturaCliente.Aprobada)


'''    Me.txtNumero.Enabled = Not FacturaProforma.Tipo.PuntoVenta.EsElectronico Or FacturaProforma.Tipo.PuntoVenta.CaeManual And Not ReadOnly


    If FacturaProforma.estado <> EstadoFacturaCliente.EnProceso Then

        If LenB(FacturaProforma.CAE) > 0 Then
            Me.txtNumero.Locked = True
        Else
        
        End If
    Else

        Me.txtDetallesCAE.caption = ""
    End If


    If IsSomething(FacturaProforma.Cliente) Then
        Me.cboCliente.ListIndex = funciones.PosIndexCbo(FacturaProforma.Cliente.id, Me.cboCliente)
        MostrarCliente
    Else
        LimpiarCliente
    End If
    
    
    If FacturaProforma.moneda.id = 0 Then
        Me.lblTC.caption = "Tipo Cambio: " & FacturaProforma.moneda.NombreCorto & " 1 "
    Else
        Me.lblTC.caption = "Tipo Cambio: " & FacturaProforma.moneda.NombreCorto & " " & FacturaProforma.CambioAPatron
    End If

    
    Me.cboMoneda.ListIndex = funciones.PosIndexCbo(FacturaProforma.moneda.id, Me.cboMoneda)

    If FacturaProforma.id = 0 Then
        'creo que aaca no entra nunca
        '        Dim classA As New classAdministracion
        Me.txtNumero.text = Format(DAOFacturaProforma.proximaFactura(FacturaProforma))    'NuevoTipoDocumento, FacturaProforma.Tipo.TipoFacturaProforma.id), "0000")
    Else


'''        Set tipos = DAOTipoFacturaDiscriminado.FindAllByFilter("id_iva=" & FacturaProforma.TipoIVA.idIVA & " and tipo_Documento=" & FacturaProforma.TipoDocumento)    'acft.id IN (select TipoFactura FROM AdminConfigFacturas where idIVA = " & FacturaProforma.TipoIVA.idIVA & ")")

        Dim T

        'lleno el combo de tipos de factura y dejo el primero x default


'        Me.cboTiposFacturaProforma.ListIndex = funciones.PosIndexCbo(FacturaProforma.Tipo.id, Me.cboTiposFactura)

        Me.txtNumero.text = FacturaProforma.numero
    End If

    Me.dtpFecha.value = FacturaProforma.FechaEmision
    Me.txtPercepcion.text = Round((FacturaProforma.AlicuotaPercepcionesIIBB - 1) * 100, 2)
    Me.txtDiasVenc.text = FacturaProforma.CantDiasPago
    Me.txtReferencia.text = FacturaProforma.OrdenCompra
    Me.txtCondObs.text = FacturaProforma.observaciones
    Me.txtTextoAdicional.text = FacturaProforma.TextoAdicional
'    Me.lblTipoFacturaProforma.caption = FacturaProforma.Tipo.TipoFacturaProforma.Tipo

    Me.dtFechaPagoCredito = FacturaProforma.fechaPago

    'fce_nemer_02062020_#113
    'Me.dtFechaServDesde = FacturaProforma.FechaServDesde
    'Me.dtFechaServHasta = FacturaProforma.FechaServHasta


    Me.txtTasaAjuste.text = FacturaProforma.TasaAjusteMensual
    ' Me.txtCbuCredito = FacturaProforma.CBU

    Dim c As CuentaBancaria

    If FacturaProforma.esCredito And LenB(FacturaProforma.CBU) > 0 Then

        Set c = DAOCuentaBancaria.FindByCBU(FacturaProforma.CBU)


    Else

    End If
    
    CargarDetalles

    Totalizar

End Sub


Private Sub Totalizar()

    Me.lblSubTotal.caption = Replace(FormatCurrency(funciones.FormatearDecimales(FacturaProforma.TotalSubTotal)), "$", "")
    Me.lblPercepciones.caption = Replace(FormatCurrency(funciones.FormatearDecimales(FacturaProforma.totalPercepciones)), "$", "")
    Me.lblIVATot.caption = Replace(FormatCurrency(funciones.FormatearDecimales(FacturaProforma.TotalIVA)), "$", "")
    Me.lblTotal.caption = Replace(FormatCurrency(funciones.FormatearDecimales(FacturaProforma.total)), "$", "")

    GridEXHelper.AutoSizeColumns Me.gridDetalles
    
End Sub


Private Sub CargarDetalles()
    Me.gridDetalles.ItemCount = 0
    Me.gridDetalles.ItemCount = FacturaProforma.Detalles.count
    ActualizarCantDetalles

End Sub


Private Sub ActualizarCantDetalles()
    Me.grpDetalles.caption = "Detalles (Cant: " & Me.gridDetalles.ItemCount & ")"
End Sub


Private Sub MostrarCliente()
    On Error Resume Next
    If FacturaProforma Is Nothing Then Exit Sub
    If FacturaProforma.Cliente Is Nothing Then Exit Sub
    Me.lblCuit.caption = FacturaProforma.Cliente.Cuit
    Me.lblIVA.caption = FacturaProforma.Cliente.TipoIVA.detalle
    Me.lblDireccion.caption = FacturaProforma.Cliente.Domicilio
    Me.lblLocalidad.caption = FacturaProforma.Cliente.localidad.nombre
    Me.lblCodPostal.caption = FacturaProforma.Cliente.CodigoPostal
    Me.lblCuitPais = FacturaProforma.Cliente.CuitPais
    Me.lblIdImpositivo = FacturaProforma.Cliente.IDImpositivo

    Me.lblProvincia = FacturaProforma.Cliente.provincia.nombre

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
            Set detalle = FacturaProforma.Detalles.item(Me.gridDetalles.rowIndex(row))
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
        Set detalle = FacturaProforma.Detalles.item(it)

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
    Set detalle = New clsFacturaProformaDetalle
    Set detalle.Factura = FacturaProforma
    detalle.idFactura = FacturaProforma.id
    detalle.Cantidad = Values(1)
    detalle.detalle = Values(2)
    detalle.PorcentajeDescuento = Values(3)
    detalle.Bruto = Values(4)
    detalle.IvaAplicado = Values(7)
    detalle.IBAplicado = Values(8)

    FacturaProforma.Detalles.Add detalle

    Totalizar
End Sub


Private Sub gridDetalles_UnboundDelete(ByVal rowIndex As Long, ByVal Bookmark As Variant)
    If rowIndex > 0 And FacturaProforma.Detalles.count > 0 Then
        FacturaProforma.Detalles.remove rowIndex
        Totalizar
    End If
End Sub


Private Sub gridDetalles_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If rowIndex <= FacturaProforma.Detalles.count Then
        Set detalle = FacturaProforma.Detalles.item(rowIndex)
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
    If rowIndex > 0 And FacturaProforma.Detalles.count > 0 Then
        Set detalle = FacturaProforma.Detalles.item(rowIndex)

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
                            transactionResult = transactionResult And DAODetalleOrdenTrabajo.SaveCantidad(redeta.idDetallePedido, redeta.Cantidad, CantidadFacturada_, redeta.Valor, FacturaProforma.id, FacturaProforma.moneda.id, FacturaProforma.CambioAPatron, FacturaProforma.TipoCambioAjuste)

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


'fce_nemer_09062020
Public Sub txtDiasVenc_LostFocus()
    If Me.txtDiasVenc = vbNullString Then
        Me.txtDiasVenc = 0
    End If

    Me.dtFechaPagoCredito.value = DateAdd("d", Me.txtDiasVenc, Me.dtpFecha)

End Sub


Private Sub txtDiasVenc_Change()
    If Not dataLoading Then
        FacturaProforma.CantDiasPago = Val(Me.txtDiasVenc.text)
    End If
End Sub


Private Sub txtNumero_Change()
    If Not dataLoading Then
        FacturaProforma.numero = Me.txtNumero
    End If
End Sub


Private Sub txtPercepcion_Change()
    On Error GoTo E
    If Not dataLoading Then
        If LenB(Me.txtPercepcion.text) = 0 Then
            FacturaProforma.AlicuotaPercepcionesIIBB = 0
        Else
            FacturaProforma.AlicuotaPercepcionesIIBB = 1 + (CDbl(Me.txtPercepcion.text) / 100)
        End If
        Totalizar
    End If

    Exit Sub
E:
    FacturaProforma.AlicuotaPercepcionesIIBB = 0
    Me.txtPercepcion.text = 0
End Sub


Private Sub txtReferencia_Change()
    If Not dataLoading Then
        FacturaProforma.OrdenCompra = Me.txtReferencia.text
    End If
End Sub


Private Sub txtTasaAjuste_Change()
    If Not dataLoading Then
        FacturaProforma.TasaAjusteMensual = Val(Me.txtTasaAjuste.text)
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


