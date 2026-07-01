VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmAdminComprasNuevaFCProveedor 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comprobantes de Proveedores"
   ClientHeight    =   8130
   ClientLeft      =   45
   ClientTop       =   480
   ClientWidth     =   10515
   Icon            =   "frmAdminComprasNuevaFCProveedor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8130
   ScaleWidth      =   10515
   Begin XtremeSuiteControls.GroupBox GroupBox2 
      Height          =   855
      Left            =   120
      TabIndex        =   60
      Top             =   7200
      Width           =   10305
      _Version        =   786432
      _ExtentX        =   18177
      _ExtentY        =   1508
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.PushButton btnCtaCte 
         Height          =   495
         Left            =   120
         TabIndex        =   61
         Top             =   240
         Width           =   2535
         _Version        =   786432
         _ExtentX        =   4471
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Ver Cta. Cte."
         UseVisualStyle  =   -1  'True
      End
   End
   Begin XtremeSuiteControls.GroupBox fraDetalleCtaCte 
      Height          =   7935
      Left            =   10560
      TabIndex        =   58
      Top             =   120
      Width           =   9615
      _Version        =   786432
      _ExtentX        =   16960
      _ExtentY        =   13996
      _StockProps     =   79
      Caption         =   "Cuenta Corriente"
      UseVisualStyle  =   -1  'True
      Begin GridEX20.GridEX gridDetalleCtaCte 
         Height          =   7455
         Left            =   240
         TabIndex        =   59
         Top             =   360
         Width           =   9225
         _ExtentX        =   16272
         _ExtentY        =   13150
         Version         =   "2.0"
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         MethodHoldFields=   -1  'True
         GroupByBoxVisible=   0   'False
         DataMode        =   99
         ColumnHeaderHeight=   285
         IntProp1        =   0
         ColumnsCount    =   5
         Column(1)       =   "frmAdminComprasNuevaFCProveedor.frx":000C
         Column(2)       =   "frmAdminComprasNuevaFCProveedor.frx":011C
         Column(3)       =   "frmAdminComprasNuevaFCProveedor.frx":0220
         Column(4)       =   "frmAdminComprasNuevaFCProveedor.frx":030C
         Column(5)       =   "frmAdminComprasNuevaFCProveedor.frx":03F8
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmAdminComprasNuevaFCProveedor.frx":04E4
         FormatStyle(2)  =   "frmAdminComprasNuevaFCProveedor.frx":061C
         FormatStyle(3)  =   "frmAdminComprasNuevaFCProveedor.frx":06CC
         FormatStyle(4)  =   "frmAdminComprasNuevaFCProveedor.frx":0780
         FormatStyle(5)  =   "frmAdminComprasNuevaFCProveedor.frx":0858
         FormatStyle(6)  =   "frmAdminComprasNuevaFCProveedor.frx":0910
         ImageCount      =   0
         PrinterProperties=   "frmAdminComprasNuevaFCProveedor.frx":09F0
      End
   End
   Begin VB.CheckBox chkButton 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Cargada desde ARCA"
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
      Height          =   495
      Left            =   4080
      TabIndex        =   57
      Top             =   6600
      Width           =   2535
   End
   Begin XtremeSuiteControls.PushButton btnFormatoNumeroLIbre 
      Height          =   255
      Left            =   8700
      TabIndex        =   55
      Top             =   2160
      Width           =   1635
      _Version        =   786432
      _ExtentX        =   2884
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Quitar Formato"
      UseVisualStyle  =   -1  'True
   End
   Begin VB.TextBox txtNumeroCargado 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8700
      TabIndex        =   54
      Text            =   "txtNumeroCargado"
      Top             =   1800
      Width           =   1635
   End
   Begin VB.TextBox txtTipoCambio 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   8700
      TabIndex        =   51
      TabStop         =   0   'False
      Text            =   "1"
      Top             =   3615
      Width           =   1605
   End
   Begin VB.TextBox lblTotal 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   8760
      TabIndex        =   50
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   5685
      Width           =   1575
   End
   Begin XtremeSuiteControls.GroupBox fraFormaPago 
      Height          =   885
      Left            =   8220
      TabIndex        =   47
      Top             =   4665
      Width           =   2085
      _Version        =   786432
      _ExtentX        =   3678
      _ExtentY        =   1561
      _StockProps     =   79
      Caption         =   "Forma de Pago"
      UseVisualStyle  =   -1  'True
      Begin VB.OptionButton optContado 
         Caption         =   "Contado"
         Height          =   195
         Left            =   255
         TabIndex        =   49
         Top             =   540
         Width           =   1140
      End
      Begin VB.OptionButton optCtaCte 
         Caption         =   "Cuenta Corriente"
         Height          =   195
         Left            =   255
         TabIndex        =   48
         Top             =   285
         Width           =   1755
      End
   End
   Begin VB.TextBox txtIVA 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   8700
      TabIndex        =   10
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   4320
      Width           =   1605
   End
   Begin XtremeSuiteControls.ComboBox cboTiposFactura 
      Height          =   315
      Left            =   8700
      TabIndex        =   4
      Top             =   1050
      Width           =   1635
      _Version        =   786432
      _ExtentX        =   2884
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboMonedas 
      Height          =   315
      Left            =   8700
      TabIndex        =   2
      Top             =   300
      Width           =   1635
      _Version        =   786432
      _ExtentX        =   2884
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.PushButton btnNuevoProveedor 
      Height          =   300
      Left            =   4755
      TabIndex        =   14
      Top             =   915
      Width           =   1575
      _Version        =   786432
      _ExtentX        =   2778
      _ExtentY        =   529
      _StockProps     =   79
      Caption         =   "Nuevo Proveedor"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1500
      Left            =   120
      TabIndex        =   40
      Top             =   975
      Width           =   6435
      _Version        =   786432
      _ExtentX        =   11351
      _ExtentY        =   2646
      _StockProps     =   79
      Caption         =   "Datos del proveedor"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.PushButton btnDisponerProveedor 
         Height          =   375
         Left            =   5160
         TabIndex        =   19
         Top             =   990
         Width           =   1095
         _Version        =   786432
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Disponer"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cboTipoIva 
         Height          =   315
         Left            =   1440
         TabIndex        =   18
         Top             =   990
         Width           =   2895
         _Version        =   786432
         _ExtentX        =   5106
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Text            =   "ComboBox1"
      End
      Begin VB.TextBox txtIB 
         Height          =   285
         Left            =   3855
         TabIndex        =   17
         Top             =   630
         Width           =   2370
      End
      Begin VB.TextBox txtRazonSocial 
         Height          =   285
         Left            =   1440
         TabIndex        =   15
         Top             =   270
         Width           =   4785
      End
      Begin XtremeSuiteControls.FlatEdit txtCuit 
         Height          =   285
         Left            =   1440
         TabIndex        =   16
         Top             =   630
         Width           =   1590
         _Version        =   786432
         _ExtentX        =   2805
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
         MaxLength       =   13
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FF8080&
         Caption         =   "Tipo IVA"
         Height          =   255
         Left            =   135
         TabIndex        =   44
         Top             =   1035
         Width           =   1215
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FF8080&
         Caption         =   "IIBB"
         Height          =   255
         Left            =   3150
         TabIndex        =   43
         Top             =   660
         Width           =   615
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FF8080&
         Caption         =   "CUIT"
         Height          =   255
         Left            =   150
         TabIndex        =   42
         Top             =   660
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FF8080&
         Caption         =   "Razón Social"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   300
         Width           =   1215
      End
   End
   Begin XtremeSuiteControls.ComboBox cboProveedores 
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Top             =   165
      Width           =   4620
      _Version        =   786432
      _ExtentX        =   8149
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.GroupBox frame2 
      Height          =   2070
      Left            =   120
      TabIndex        =   37
      Top             =   5040
      Width           =   3630
      _Version        =   786432
      _ExtentX        =   6403
      _ExtentY        =   3651
      _StockProps     =   79
      Caption         =   "Percepciones"
      UseVisualStyle  =   -1  'True
      Begin GridEX20.GridEX grilla_percepciones 
         Height          =   1650
         Left            =   120
         TabIndex        =   13
         Top             =   255
         Width           =   3390
         _ExtentX        =   5980
         _ExtentY        =   2910
         Version         =   "2.0"
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         ColumnAutoResize=   -1  'True
         MethodHoldFields=   -1  'True
         AllowDelete     =   -1  'True
         RowHeaders      =   -1  'True
         DataMode        =   99
         AllowAddNew     =   -1  'True
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   2
         Column(1)       =   "frmAdminComprasNuevaFCProveedor.frx":0BC8
         Column(2)       =   "frmAdminComprasNuevaFCProveedor.frx":0D10
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmAdminComprasNuevaFCProveedor.frx":0E24
         FormatStyle(2)  =   "frmAdminComprasNuevaFCProveedor.frx":0F5C
         FormatStyle(3)  =   "frmAdminComprasNuevaFCProveedor.frx":100C
         FormatStyle(4)  =   "frmAdminComprasNuevaFCProveedor.frx":10C0
         FormatStyle(5)  =   "frmAdminComprasNuevaFCProveedor.frx":1198
         FormatStyle(6)  =   "frmAdminComprasNuevaFCProveedor.frx":1250
         ImageCount      =   0
         PrinterProperties=   "frmAdminComprasNuevaFCProveedor.frx":1330
      End
   End
   Begin XtremeSuiteControls.GroupBox fraAlicuotas 
      Height          =   2235
      Left            =   120
      TabIndex        =   36
      Tag             =   "Alicuotas IVA (Total: {VALUE})"
      Top             =   2520
      Width           =   3645
      _Version        =   786432
      _ExtentX        =   6429
      _ExtentY        =   3942
      _StockProps     =   79
      Caption         =   "Alicuotas IVA"
      UseVisualStyle  =   -1  'True
      Begin GridEX20.GridEX grilla_alicuotas 
         Height          =   1890
         Left            =   105
         TabIndex        =   11
         Top             =   255
         Width           =   3390
         _ExtentX        =   5980
         _ExtentY        =   3334
         Version         =   "2.0"
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         ColumnAutoResize=   -1  'True
         MethodHoldFields=   -1  'True
         AllowDelete     =   -1  'True
         RowHeaders      =   -1  'True
         DataMode        =   99
         AllowAddNew     =   -1  'True
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   2
         Column(1)       =   "frmAdminComprasNuevaFCProveedor.frx":1508
         Column(2)       =   "frmAdminComprasNuevaFCProveedor.frx":1674
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmAdminComprasNuevaFCProveedor.frx":1788
         FormatStyle(2)  =   "frmAdminComprasNuevaFCProveedor.frx":18C0
         FormatStyle(3)  =   "frmAdminComprasNuevaFCProveedor.frx":1970
         FormatStyle(4)  =   "frmAdminComprasNuevaFCProveedor.frx":1A24
         FormatStyle(5)  =   "frmAdminComprasNuevaFCProveedor.frx":1AFC
         FormatStyle(6)  =   "frmAdminComprasNuevaFCProveedor.frx":1BB4
         ImageCount      =   0
         PrinterProperties=   "frmAdminComprasNuevaFCProveedor.frx":1C94
      End
   End
   Begin GridEX20.GridEX grilla_alicuota 
      Height          =   2325
      Left            =   240
      TabIndex        =   34
      Top             =   8160
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   4101
      Version         =   "2.0"
      BoundColumnIndex=   "id"
      ReplaceColumnIndex=   "alicuota"
      ActAsDropDown   =   -1  'True
      ColumnAutoResize=   -1  'True
      HideSelection   =   2
      MethodHoldFields=   -1  'True
      ContScroll      =   -1  'True
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   2
      Column(1)       =   "frmAdminComprasNuevaFCProveedor.frx":1E6C
      Column(2)       =   "frmAdminComprasNuevaFCProveedor.frx":1F8C
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminComprasNuevaFCProveedor.frx":208C
      FormatStyle(2)  =   "frmAdminComprasNuevaFCProveedor.frx":21C4
      FormatStyle(3)  =   "frmAdminComprasNuevaFCProveedor.frx":2274
      FormatStyle(4)  =   "frmAdminComprasNuevaFCProveedor.frx":2328
      FormatStyle(5)  =   "frmAdminComprasNuevaFCProveedor.frx":2400
      FormatStyle(6)  =   "frmAdminComprasNuevaFCProveedor.frx":24B8
      ImageCount      =   0
      PrinterProperties=   "frmAdminComprasNuevaFCProveedor.frx":2598
   End
   Begin XtremeSuiteControls.PushButton btnGuardar 
      Height          =   495
      Left            =   8760
      TabIndex        =   20
      Top             =   6600
      Width           =   1635
      _Version        =   786432
      _ExtentX        =   2884
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Guardar"
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
   Begin VB.TextBox txtCodigoProveedor 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   540
      Width           =   1065
   End
   Begin VB.TextBox txtMontoManual 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   8760
      TabIndex        =   21
      Top             =   6060
      Width           =   1575
   End
   Begin VB.TextBox txtRedondeo 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   8700
      TabIndex        =   8
      Text            =   "0"
      Top             =   2835
      Width           =   1635
   End
   Begin VB.TextBox txtMontoNeto 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   8700
      TabIndex        =   9
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   3960
      Width           =   1605
   End
   Begin VB.TextBox txtImpuestos 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   8700
      TabIndex        =   7
      Text            =   "0"
      Top             =   2490
      Width           =   1635
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   300
      Left            =   8700
      TabIndex        =   5
      Top             =   1425
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   529
      _Version        =   393216
      Format          =   16842753
      CurrentDate     =   39897
   End
   Begin XtremeSuiteControls.GroupBox frame3 
      Height          =   2235
      Left            =   3840
      TabIndex        =   38
      Top             =   2520
      Width           =   3645
      _Version        =   786432
      _ExtentX        =   6429
      _ExtentY        =   3942
      _StockProps     =   79
      Caption         =   "Cuentas Contables"
      UseVisualStyle  =   -1  'True
      Begin GridEX20.GridEX grid_cuentascontables 
         Height          =   1890
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   3390
         _ExtentX        =   5980
         _ExtentY        =   3334
         Version         =   "2.0"
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         ColumnAutoResize=   -1  'True
         MethodHoldFields=   -1  'True
         AllowDelete     =   -1  'True
         RowHeaders      =   -1  'True
         DataMode        =   99
         AllowAddNew     =   -1  'True
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   2
         Column(1)       =   "frmAdminComprasNuevaFCProveedor.frx":2770
         Column(2)       =   "frmAdminComprasNuevaFCProveedor.frx":28AC
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmAdminComprasNuevaFCProveedor.frx":29C0
         FormatStyle(2)  =   "frmAdminComprasNuevaFCProveedor.frx":2AF8
         FormatStyle(3)  =   "frmAdminComprasNuevaFCProveedor.frx":2BA8
         FormatStyle(4)  =   "frmAdminComprasNuevaFCProveedor.frx":2C5C
         FormatStyle(5)  =   "frmAdminComprasNuevaFCProveedor.frx":2D34
         FormatStyle(6)  =   "frmAdminComprasNuevaFCProveedor.frx":2DEC
         ImageCount      =   0
         PrinterProperties=   "frmAdminComprasNuevaFCProveedor.frx":2ECC
      End
   End
   Begin GridEX20.GridEX grid_cuenta 
      Height          =   4725
      Left            =   2040
      TabIndex        =   39
      Top             =   8160
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   8334
      Version         =   "2.0"
      HoldSortSettings=   -1  'True
      BoundColumnIndex=   "id"
      ReplaceColumnIndex=   "cuenta"
      ActAsDropDown   =   -1  'True
      ColumnAutoResize=   -1  'True
      HideSelection   =   2
      MethodHoldFields=   -1  'True
      ContScroll      =   -1  'True
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      ColumnHeaders   =   0   'False
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   2
      Column(1)       =   "frmAdminComprasNuevaFCProveedor.frx":30A4
      Column(2)       =   "frmAdminComprasNuevaFCProveedor.frx":31BC
      SortKeysCount   =   1
      SortKey(1)      =   "frmAdminComprasNuevaFCProveedor.frx":32BC
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminComprasNuevaFCProveedor.frx":3324
      FormatStyle(2)  =   "frmAdminComprasNuevaFCProveedor.frx":345C
      FormatStyle(3)  =   "frmAdminComprasNuevaFCProveedor.frx":350C
      FormatStyle(4)  =   "frmAdminComprasNuevaFCProveedor.frx":35C0
      FormatStyle(5)  =   "frmAdminComprasNuevaFCProveedor.frx":3698
      FormatStyle(6)  =   "frmAdminComprasNuevaFCProveedor.frx":3750
      ImageCount      =   0
      PrinterProperties=   "frmAdminComprasNuevaFCProveedor.frx":3830
   End
   Begin GridEX20.GridEX grilla_percepcion 
      Height          =   4725
      Left            =   5640
      TabIndex        =   35
      Top             =   8160
      Width           =   2940
      _ExtentX        =   5186
      _ExtentY        =   8334
      Version         =   "2.0"
      BoundColumnIndex=   "id"
      ReplaceColumnIndex=   "percepcion"
      ActAsDropDown   =   -1  'True
      ColumnAutoResize=   -1  'True
      HideSelection   =   2
      MethodHoldFields=   -1  'True
      ContScroll      =   -1  'True
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   2
      Column(1)       =   "frmAdminComprasNuevaFCProveedor.frx":3A08
      Column(2)       =   "frmAdminComprasNuevaFCProveedor.frx":3B30
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminComprasNuevaFCProveedor.frx":3C30
      FormatStyle(2)  =   "frmAdminComprasNuevaFCProveedor.frx":3D68
      FormatStyle(3)  =   "frmAdminComprasNuevaFCProveedor.frx":3E18
      FormatStyle(4)  =   "frmAdminComprasNuevaFCProveedor.frx":3ECC
      FormatStyle(5)  =   "frmAdminComprasNuevaFCProveedor.frx":3FA4
      FormatStyle(6)  =   "frmAdminComprasNuevaFCProveedor.frx":405C
      ImageCount      =   0
      PrinterProperties=   "frmAdminComprasNuevaFCProveedor.frx":413C
   End
   Begin XtremeSuiteControls.ComboBox cboTipoDocContable 
      Height          =   315
      Left            =   8700
      TabIndex        =   3
      Top             =   675
      Width           =   1635
      _Version        =   786432
      _ExtentX        =   2884
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Sorted          =   -1  'True
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.PushButton btnNuevoCbte 
      Height          =   510
      Left            =   6840
      TabIndex        =   23
      Top             =   6600
      Width           =   1635
      _Version        =   786432
      _ExtentX        =   2884
      _ExtentY        =   900
      _StockProps     =   79
      Caption         =   "Nueva Factura"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtNumeroMask 
      Height          =   285
      Left            =   8700
      TabIndex        =   6
      Top             =   1785
      Width           =   1635
      _Version        =   786432
      _ExtentX        =   2884
      _ExtentY        =   503
      _StockProps     =   77
      BackColor       =   -2147483643
      Alignment       =   1
      MaxLength       =   13
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FF8080&
      Caption         =   "Nota: Unificar las lineas de alicuotas segun su indice. No cargar misma alicuota por duplicado."
      Height          =   495
      Index           =   1
      Left            =   240
      TabIndex        =   56
      Top             =   4800
      Width           =   7335
   End
   Begin VB.Label lblTipoCambioPago 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      Caption         =   "Tipo de Cambio pago: "
      Height          =   255
      Left            =   7005
      TabIndex        =   53
      Top             =   3240
      Width           =   3345
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      Caption         =   "Tipo de Cambio"
      Height          =   255
      Left            =   7440
      TabIndex        =   52
      Top             =   3645
      Width           =   1215
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      Caption         =   "IVA"
      Height          =   255
      Left            =   7080
      TabIndex        =   46
      Top             =   4350
      Width           =   1575
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      Caption         =   "Letra"
      Height          =   255
      Left            =   7455
      TabIndex        =   45
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label lblMoneda 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Moneda"
      Height          =   195
      Left            =   8040
      TabIndex        =   33
      Top             =   360
      Width           =   585
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      Caption         =   "Código"
      Height          =   255
      Left            =   240
      TabIndex        =   32
      Top             =   555
      Width           =   975
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      Caption         =   "Validar Factura"
      Height          =   255
      Left            =   7320
      TabIndex        =   31
      Top             =   6090
      Width           =   1335
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      Caption         =   "Redondeo IVA"
      Height          =   255
      Left            =   7335
      TabIndex        =   30
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      Caption         =   "Neto Gravado"
      Height          =   255
      Left            =   7080
      TabIndex        =   29
      Top             =   4005
      Width           =   1575
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      Caption         =   "Impuestos"
      Height          =   255
      Left            =   7215
      TabIndex        =   28
      Top             =   2505
      Width           =   1455
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      Caption         =   "Número"
      Height          =   195
      Left            =   7695
      TabIndex        =   26
      Top             =   1845
      Width           =   975
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      Caption         =   "Fecha"
      Height          =   255
      Left            =   7695
      TabIndex        =   25
      Top             =   1470
      Width           =   975
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      Caption         =   "Total Factura"
      Height          =   255
      Left            =   7320
      TabIndex        =   24
      Top             =   5715
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      Caption         =   "Proveedor"
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   225
      Width           =   975
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      Caption         =   "Tipo Documento"
      Height          =   255
      Index           =   0
      Left            =   7095
      TabIndex        =   27
      Top             =   720
      Width           =   1575
   End
End
Attribute VB_Name = "frmAdminComprasNuevaFCProveedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim moneda As clsMoneda
Dim loading As Boolean
Dim colAlicuotas As New Collection
Dim aliaplicada As clsAlicuotaAplicada

Dim colPercepciones As Collection
Dim colPercepcionesTMP As New Collection
Dim quitarFormato As Boolean

Dim perAplicada As clsPercepcionesAplicadas

Dim colCuentas As Collection

Dim ctaAplicada As clsCuentaFactura
Dim ctacontable As clsCuentaContable
Dim Percepcion As New clsPercepciones
Dim alicuota As clsAlicuotas
Dim idtipo As Long
Dim nroFacturaAnterior
Dim grabado As Boolean
Dim Proveedor As clsProveedor
Dim idProveedor As Long
Dim vFactura As clsFacturaProveedor
Dim VVer As Boolean

Private detallesCtaCte As Collection
Private detaCtaCte As DTODetalleCuentaCorriente

Private detalleCtaCteVisible As Boolean
Private anchoNormal As Long
Private anchoConDetalle As Long


Public Property Let ver(nVer As Boolean)
    VVer = nVer
End Property

Public Property Let Factura(nFactura As clsFacturaProveedor)
    If IsSomething(nFactura) Then
        Set vFactura = DAOFacturaProveedor.FindById(nFactura.Id)
    End If
End Property

Private Sub LlenarCuentasContables()
    Set colCuentas = DAOCuentaContable.GetAll()
    Me.grid_cuenta.ItemCount = 0
    Me.grid_cuenta.ItemCount = colCuentas.count
    Set Me.grid_cuentascontables.Columns("cuenta").DropDownControl = Me.grid_cuenta
End Sub

Private Sub llenarTiposFacturas()
    Dim i As Long
    Dim idIVA As Long
    Me.cboTiposFactura.Clear

    For i = 1 To Proveedor.TipoIVA.configFacturas.count
        idIVA = Proveedor.TipoIVA.Id
        Me.cboTiposFactura.AddItem Proveedor.TipoIVA.configFacturas(i).TipoFactura
        Me.cboTiposFactura.ItemData(Me.cboTiposFactura.NewIndex) = Proveedor.TipoIVA.configFacturas(i).Id
    Next i

    Dim idtipo As Long

    If Me.cboTiposFactura.ListCount > 0 Then
        Me.cboTiposFactura.ListIndex = 0
        idtipo = Me.cboTiposFactura.ItemData(Me.cboTiposFactura.ListIndex)
        llenarAlicuotas idtipo
    End If
End Sub

Private Sub btnCtaCte_Click()
    On Error GoTo err1

    If detalleCtaCteVisible Then
        
        OcultarDetalleCtaCte
        
    Else
        
        If Me.cboProveedores.ListIndex = -1 Then
            MsgBox "Debe seleccionar un proveedor.", vbExclamation, "Cuenta corriente"
            Exit Sub
        End If

        MostrarSaldoCtaCteProveedor
        
        CargarDetalleCtaCteProveedor

        Me.Width = anchoConDetalle
        Me.fraDetalleCtaCte.Visible = True
        Me.btnCtaCte.caption = "Ocultar Cta. Cte."
        detalleCtaCteVisible = True
        
    End If

    Exit Sub

err1:
    MsgBox "Error al mostrar/ocultar cuenta corriente: " & Err.Description, vbCritical, "Error"

End Sub

Private Sub btnFormatoNumeroLIbre_Click()
    If quitarFormato = True Then
        Me.txtNumeroMask.SetMask "", ""
        Me.txtNumeroMask.MaxLength = 16
        Me.btnFormatoNumeroLIbre.caption = "Reestablecer"
        quitarFormato = False

    Else

        Me.txtNumeroMask.SetMask "0000-00000000", "____-________"
        Me.txtNumeroMask.MaxLength = 13
        Me.btnFormatoNumeroLIbre.caption = "Quitar Formato"
        quitarFormato = True

    End If

End Sub


Private Sub cboMonedas_Click()
    Set moneda = DAOMoneda.GetById(Me.cboMonedas.ItemData(Me.cboMonedas.ListIndex))

    If IsSomething(vFactura) Then Set vFactura.moneda = moneda
    If IsSomething(moneda) Then
        If Not VVer Then
            Me.txtTipoCambio = moneda.Cambio
        End If
    End If
End Sub


Private Sub cboProveedores_Click()
    mostrar

    If Me.cboProveedores.ListIndex <> -1 Then
        Me.txtCodigoProveedor = Me.cboProveedores.ItemData(Me.cboProveedores.ListIndex)
    End If

    If Not loading Then
        OcultarDetalleCtaCte
'''        MostrarSaldoCtaCteProveedor
    End If
End Sub



'''Private Sub cboTiposFactura_Click()
'''
'''    grabado = False
'''    'vFactura.IvaAplicado = Nothing
'''    FacturaRequiereNumeroFormateado
'''
'''    If Me.cboTiposFactura.ListCount > 0 Then
'''        Dim idtipo As Long
'''        idtipo = Me.cboTiposFactura.ItemData(Me.cboTiposFactura.ListIndex)
'''
'''        llenarAlicuotas idtipo
'''
'''            If Not loading Then
'''                 If colAlicuotas.count > 0 Then
'''                    vFactura.IvaAplicado = Nothing
'''                    Me.grilla_alicuotas.ItemCount = 0
'''                    Me.grilla_alicuotas.Refresh
'''                    AddDefaultAlicuota colAlicuotas(1).Id
'''                End If
'''            End If
'''        End If
'''End Sub


Private Sub cboTiposFactura_Click()
    On Error GoTo ErrHandler
    
    Dim idx As Long
    Dim idtipoLocal As Long
    
    grabado = False
    
    If Me.cboTiposFactura.ListCount <= 0 Then Exit Sub
    
    idx = Me.cboTiposFactura.ListIndex
    If idx < 0 Then Exit Sub
    
    idtipoLocal = Me.cboTiposFactura.ItemData(idx)
    
    FacturaRequiereNumeroFormateado
    
    llenarAlicuotas idtipoLocal
    
    If Not loading Then
        If colAlicuotas.count > 0 Then
            vFactura.IvaAplicado = Nothing
            Me.grilla_alicuotas.ItemCount = 0
            Me.grilla_alicuotas.Refresh
            AddDefaultAlicuota colAlicuotas(1).Id
        End If
    End If
    
    Exit Sub

ErrHandler:
    MsgBox "Error al seleccionar el tipo de factura: " & Err.Description, vbExclamation, "Error"
End Sub
Private Sub mostrar()
    If Me.cboProveedores.ListIndex <> -1 Then
        idProveedor = CLng(Me.cboProveedores.ItemData(Me.cboProveedores.ListIndex))
        '        Dim lstRubros As ListView
        '        Dim accion As Integer
        Set Proveedor = DAOProveedor.FindById(idProveedor, False, False, False, False)
        If IsSomething(Proveedor) Then
            Me.txtCuit = Proveedor.Cuit
            Me.txtIB = Proveedor.IIBB
            Me.txtRazonSocial = Proveedor.RazonSocial
            Me.cboTipoIva.ListIndex = funciones.PosIndexCbo(Proveedor.TipoIVA.Id, Me.cboTipoIva)
            llenarTiposFacturas
            ProtegerProveedor
        End If
    End If
End Sub


Private Sub LimpiarProveedor()
    Me.txtCuit = vbNullString
    Me.txtIB = vbNullString
    Me.txtRazonSocial = vbNullString
    Me.cboTiposFactura.Clear
End Sub


Private Sub btnDisponerProveedor_Click()
    If Proveedor.Id = 0 Then
        Proveedor.RazonSocial = Me.txtRazonSocial
        Proveedor.Cuit = Replace(Me.txtCuit, "-", vbNullString)
        Proveedor.IIBB = Me.txtIB
        Proveedor.estado = EstadoProveedorContado
        Set Proveedor.TipoIVA = DAOTipoIvaProveedor.GetById(Me.cboTipoIva.ItemData(Me.cboTipoIva.ListIndex))
        llenarTiposFacturas
    End If
End Sub


'''Private Sub btnGuardar_Click()
'''    On Error GoTo err1
'''
'''    If Me.grilla_alicuotas.EditMode = jgexEditModeOn Then
'''        MsgBox "Todavia esta editando la grilla de Alicuotas de IVA." & vbNewLine & "Presione [ENTER] en la grilla para guardar los cambios.", vbExclamation + vbOKOnly
'''        Exit Sub
'''    End If
'''
'''
'''    If Me.grilla_percepciones.EditMode = jgexEditModeOn Then
'''        MsgBox "Todavia esta editando la grilla de Percepciones." & vbNewLine & "Presione [ENTER] en la grilla para guardar los cambios.", vbExclamation + vbOKOnly
'''        Exit Sub
'''    End If
'''
'''    If Me.grid_cuentascontables.EditMode = jgexEditModeOn Then
'''        MsgBox "Todavia esta editando la grilla de Cuentas Contables." & vbNewLine & "Presione [ENTER] en la grilla para guardar los cambios.", vbExclamation + vbOKOnly
'''        Exit Sub
'''    End If
'''
'''
'''    If Not Me.optContado.value And Not Me.optCtaCte.value Then
'''        MsgBox "Debe seleccionar la forma de pago.", vbExclamation
'''        Exit Sub
'''    End If
'''
'''    conectar.BeginTransaction
'''    '    Dim A As Boolean
'''    Dim montonero As Double
'''    '    Dim nroNuevo As Long
'''    Dim EVENTO As clsEventoObserver
'''    Dim nuevoproveedor As Boolean
'''
'''    If Not validarFactura Then
'''        Err.Raise 203
'''    End If
'''
''''''    ' Si se elije la letra B y hay cargado alicuotas
'''''''    If Me.cboTiposFactura.ListIndex = 1 Then
'''''''        MsgBox "Recuerdo que si el comprobante es letra B no debe tener alicuotas discriminadas", vbCritical, "Error"
'''''''    End If
'''
'''    'If MsgBox("¿Está seguro de guardar la factura?", vbYesNo, "Confirmación") = vbYes Then
'''
'''    armarFactura
'''
''''''    If vFactura.NetoGravado <= 0 Then
''''''        If vFactura.tipoDocumentoContable <> vFactura.tipoDocumentoContable.notaDebito Then Err.Raise 202
''''''        End If
''''''    End If
'''
'''
'''
'''    montonero = CDbl(Me.txtMontoNeto)
'''
'''
'''    If Me.txtNumeroMask.Text <> "______-________" And Len(Me.txtNumeroMask.Text) > 0 Then
'''        '    If Me.txtNumeroMask.text <> "" And Len(Me.txtNumeroMask.text) > 0 Then
'''
''''''        If vFactura.cuentasContables.count = 0 And vFactura.tipoDocumentoContable <> notaDebito Then Err.Raise 201
'''
'''        If funciones.RedondearDecimales(vFactura.TotalAplicadoACuentas) <> funciones.RedondearDecimales(vFactura.NetoGravado) Then Err.Raise 200
'''
'''        If Me.cboMonedas.ListIndex <> -1 Then
'''            Set vFactura.moneda = DAOMoneda.GetById(Me.cboMonedas.ItemData(Me.cboMonedas.ListIndex))
'''        Else
'''            Set vFactura.moneda = Nothing
'''        End If
'''
'''        'creo el proveedor si es contado
'''
'''        If vFactura.Proveedor.Id = 0 Then
'''            If Trim(Me.txtRazonSocial) = vbNullString Or Not IsNumeric(Replace(Me.txtCuit, "-", vbNullString)) Then
'''                If Not funciones.VerificarCUIT(Replace(Me.txtCuit, "-", vbNullString)) Then
'''                    Err.Raise 1000
'''                End If
'''            Else
'''                nuevoproveedor = True
'''                Set colPercepcionesTMP = colPercepciones
'''                If Not DAOProveedor.Guardar(vFactura.Proveedor) Then Err.Raise 300
'''            End If
'''
'''        End If
'''
'''        If DAOFacturaProveedor.existeFactura(vFactura) Then Err.Raise 101
'''
'''        Dim Nueva As Boolean
'''        Nueva = (vFactura.Id = 0)
'''
'''        If DAOFacturaProveedor.Guardar(vFactura) Then
'''            Set EVENTO = New clsEventoObserver
'''            Set EVENTO.Elemento = vFactura
'''            If Nueva Then
'''                EVENTO.EVENTO = agregar_
'''            Else
'''                EVENTO.EVENTO = modificar_
'''            End If
'''            Set EVENTO.Originador = Me
'''            EVENTO.Tipo = TipoSuscripcion.FacturaProveedor_
'''
'''            ' Desactivo este evento Notificar porque aparentemente da Error (dienemer 11.09.20)
'''            'Channel.Notificar EVENTO, TipoSuscripcion.FacturaProveedor_
'''
'''            MsgBox "Factura almacenada con éxito!", vbInformation, "Información"
'''            grabado = True
'''
'''            Me.cboProveedores.SetFocus
'''
'''        Else
'''            Err.Raise 100
'''        End If
'''    Else
'''        Err.Raise 101
'''    End If
'''    'End If
'''    conectar.CommitTransaction
'''    Exit Sub
'''
'''err1:
'''    conectar.RollBackTransaction
'''
'''    If Err.Number = 100 Then
'''        MsgBox "Se produjo algún error, no se guardarán los cambios!", vbCritical, "Error"
'''    ElseIf Err.Number = 101 Then
'''        MsgBox "La factura que intenta guardar ya existe!", vbCritical, "Error"
'''    ElseIf Err.Number = 200 Then
'''        MsgBox "Debe tener todo neto gravado aplicado a cuenta(s) contable(s)!", vbCritical, "Error"
'''    ElseIf Err.Number = 1000 Then
'''        MsgBox "Debe definir datos correctos para el proveedor que está creando!", vbCritical, "Error"
'''    ElseIf Err.Number = 201 Then
'''        MsgBox "Debe ingresar al menos una cuenta contable!", vbCritical, "Error"
'''    ElseIf Err.Number = 202 Then
'''        MsgBox "Debe ingresar montos válidos!", vbCritical, "Error"
'''    ElseIf Err.Number = 203 Then
'''        MsgBox "Los totales de la factura no coinciden." & vbNewLine & "Total esperado: " & funciones.RedondearDecimales(CDbl(Me.txtMontoManual)) & vbNewLine & "Total ingresado: " & vFactura.total, vbCritical, "Error"
'''    ElseIf Err.Number = 300 Or nuevoproveedor Then
'''        vFactura.Proveedor = Nothing
'''        nuevoproveedor = False
'''    Else
'''        MsgBox Err.Description, vbCritical
'''    End If
'''End Sub


Private Sub btnGuardar_Click()
    On Error GoTo err1

    Dim EVENTO As clsEventoObserver
    Dim nuevoproveedor As Boolean
    Dim Nueva As Boolean
    Dim totalManual As Double
    Dim transIniciada As Boolean
    Dim nroComprobante As String

    If Me.grilla_alicuotas.EditMode = jgexEditModeOn Then
        MsgBox "Todavia esta editando la grilla de Alicuotas de IVA." & vbNewLine & _
               "Presione [ENTER] en la grilla para guardar los cambios.", vbExclamation + vbOKOnly
        Exit Sub
    End If

    If Me.grilla_percepciones.EditMode = jgexEditModeOn Then
        MsgBox "Todavia esta editando la grilla de Percepciones." & vbNewLine & _
               "Presione [ENTER] en la grilla para guardar los cambios.", vbExclamation + vbOKOnly
        Exit Sub
    End If

    If Me.grid_cuentascontables.EditMode = jgexEditModeOn Then
        MsgBox "Todavia esta editando la grilla de Cuentas Contables." & vbNewLine & _
               "Presione [ENTER] en la grilla para guardar los cambios.", vbExclamation + vbOKOnly
        Exit Sub
    End If

    If Not Me.optContado.value And Not Me.optCtaCte.value Then
        MsgBox "Debe seleccionar la forma de pago.", vbExclamation
        Exit Sub
    End If

    If Not TryGetDouble(Me.txtMontoManual.Text, totalManual) Then
        MsgBox "Debe ingresar un total válido en el campo Validar Factura.", vbExclamation, "Validación"
        Me.txtMontoManual.SetFocus
        Exit Sub
    End If

    armarFactura

    If funciones.RedondearDecimales(vFactura.total) <> funciones.RedondearDecimales(totalManual) Then
        Err.Raise 203
    End If

    If Me.txtNumeroCargado.Visible Then
        nroComprobante = Trim$(Me.txtNumeroCargado.Text)
    Else
        nroComprobante = Trim$(Me.txtNumeroMask.Text)
    End If

    If nroComprobante = vbNullString Or nroComprobante = "____-________" Or nroComprobante = "______-________" Then
        MsgBox "Debe ingresar el número de comprobante.", vbExclamation, "Validación"
        Exit Sub
    End If

    If funciones.RedondearDecimales(vFactura.TotalAplicadoACuentas) <> funciones.RedondearDecimales(vFactura.NetoGravado) Then
        Err.Raise 200
    End If

    If Me.cboMonedas.ListIndex <> -1 Then
        Set vFactura.moneda = DAOMoneda.GetById(Me.cboMonedas.ItemData(Me.cboMonedas.ListIndex))
    Else
        Set vFactura.moneda = Nothing
    End If

    conectar.BeginTransaction
    transIniciada = True

    If vFactura.Proveedor.Id = 0 Then
        If Trim$(Me.txtRazonSocial.Text) = vbNullString Or Not IsNumeric(Replace(Me.txtCuit.Text, "-", vbNullString)) Then
            If Not funciones.VerificarCUIT(Replace(Me.txtCuit.Text, "-", vbNullString)) Then
                Err.Raise 1000
            End If
        Else
            nuevoproveedor = True
            Set colPercepcionesTMP = colPercepciones
            If Not DAOProveedor.Guardar(vFactura.Proveedor) Then Err.Raise 300
        End If
    End If

    If DAOFacturaProveedor.existeFactura(vFactura) Then Err.Raise 101

    Nueva = (vFactura.Id = 0)

    If DAOFacturaProveedor.Guardar(vFactura) Then
        Set EVENTO = New clsEventoObserver
        Set EVENTO.Elemento = vFactura

        If Nueva Then
            EVENTO.EVENTO = agregar_
        Else
            EVENTO.EVENTO = modificar_
        End If

        Set EVENTO.Originador = Me
        EVENTO.Tipo = TipoSuscripcion.FacturaProveedor_

        MsgBox "Factura almacenada con éxito!", vbInformation, "Información"
        grabado = True

        conectar.CommitTransaction
        transIniciada = False

        Me.cboProveedores.SetFocus
    Else
        Err.Raise 100
    End If

    Exit Sub

err1:
    If transIniciada Then conectar.RollBackTransaction

    If Err.Number = 100 Then
        MsgBox "Se produjo algún error, no se guardarán los cambios!", vbCritical, "Error"
    ElseIf Err.Number = 101 Then
        MsgBox "La factura que intenta guardar ya existe!", vbCritical, "Error"
    ElseIf Err.Number = 200 Then
        MsgBox "Debe tener todo neto gravado aplicado a cuenta(s) contable(s)!", vbCritical, "Error"
    ElseIf Err.Number = 1000 Then
        MsgBox "Debe definir datos correctos para el proveedor que está creando!", vbCritical, "Error"
    ElseIf Err.Number = 201 Then
        MsgBox "Debe ingresar al menos una cuenta contable!", vbCritical, "Error"
    ElseIf Err.Number = 202 Then
        MsgBox "Debe ingresar montos válidos!", vbCritical, "Error"
    ElseIf Err.Number = 203 Then
        MsgBox "Los totales de la factura no coinciden." & vbNewLine & _
               "Total esperado: " & funciones.RedondearDecimales(totalManual) & vbNewLine & _
               "Total ingresado: " & funciones.RedondearDecimales(vFactura.total), vbCritical, "Error"
    ElseIf Err.Number = 300 Or nuevoproveedor Then
        Set vFactura.Proveedor = Nothing
        nuevoproveedor = False
    Else
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
    End If
End Sub


Private Function TryGetDouble(ByVal Valor As String, ByRef Resultado As Double) As Boolean
    On Error GoTo err1

    Valor = Trim$(Valor)
    Valor = Replace(Valor, "$", "")
    Valor = Replace(Valor, " ", "")

    If Valor = vbNullString Then Exit Function
    If Not IsNumeric(Valor) Then Exit Function

    Resultado = CDbl(Valor)
    TryGetDouble = True
    Exit Function

err1:
    TryGetDouble = False
End Function


Private Sub DTPicker1_Click()
    grabado = False
End Sub

Private Sub Form_Activate()

    Me.cboProveedores.Enabled = True
    Me.cboProveedores.Visible = True
    Me.cboProveedores.SetFocus
    
End Sub

Private Sub Form_Load()
    loading = True

    anchoNormal = Me.Width
    anchoConDetalle = 20400

    detalleCtaCteVisible = False


    Me.txtCuit.SetMask "00-00000000-0", "__-________-_"

    Me.txtNumeroMask.Visible = True
    Me.txtNumeroMask.SetMask "0000-00000000", "____-________"
    Me.txtNumeroMask.MaxLength = 13

    quitarFormato = True

    Me.txtNumeroCargado.Visible = False

    FormHelper.Customize Me

    '    Set vFactura = DAOFacturaProveedor.FindById(vFactura.id)

    If Not IsSomething(vFactura) Then Set vFactura = New clsFacturaProveedor
    
    GridEXHelper.CustomizeGrid Me.grilla_alicuota, False, False
    GridEXHelper.CustomizeGrid Me.grilla_alicuotas, False, True
    GridEXHelper.CustomizeGrid Me.grilla_percepcion, False, False
    GridEXHelper.CustomizeGrid Me.grilla_percepciones, False, True
    GridEXHelper.CustomizeGrid Me.grid_cuenta, False, False
    GridEXHelper.CustomizeGrid Me.grid_cuentascontables, False, True

    DAOMoneda.llenarComboXtremeSuite Me.cboMonedas
    DAOTipoIvaProveedor.llenarComboXtremeSuite Me.cboTipoIva
    llenarComboProveedores

    Me.cboTipoDocContable.AddItem "Factura"
    Me.cboTipoDocContable.ItemData(Me.cboTipoDocContable.NewIndex) = tipoDocumentoContable.Factura
    Me.cboTipoDocContable.AddItem "Nota de crédito"
    Me.cboTipoDocContable.ItemData(Me.cboTipoDocContable.NewIndex) = tipoDocumentoContable.notaCredito
    Me.cboTipoDocContable.AddItem "Nota de débito"
    Me.cboTipoDocContable.ItemData(Me.cboTipoDocContable.NewIndex) = tipoDocumentoContable.notaDebito
    Me.cboTipoDocContable.AddItem "Despacho de Aduana"
    Me.cboTipoDocContable.ItemData(Me.cboTipoDocContable.NewIndex) = tipoDocumentoContable.DespachoAduana
    Me.cboTipoDocContable.AddItem "Liquidacion Bancaria"
    Me.cboTipoDocContable.ItemData(Me.cboTipoDocContable.NewIndex) = tipoDocumentoContable.LiquidacionBancaria

    'e5re52- SE AGREGA ESTE NUEVO TIPO DE COMPROBANTE COMPROBANTE DE COMPRA DE BIEN USADO
    Me.cboTipoDocContable.AddItem "Compra Bien Usado"
    Me.cboTipoDocContable.ItemData(Me.cboTipoDocContable.NewIndex) = tipoDocumentoContable.CompraBienesUsados

    ' MODIFICO EL VALOR DEL INDICE PARA QUE SE INICIE COMO "FACTURA"
    Me.cboTipoDocContable.ListIndex = 2

    Me.grilla_alicuotas.ItemCount = 0
    Me.grilla_percepciones.ItemCount = 0
    Me.grid_cuentascontables.ItemCount = 0

    FacturaRequiereNumeroFormateado

    llenarGrillaPercepciones
    
    LlenarCuentasContables

    Me.DTPicker1 = Now
    
    If vFactura.Id > 0 Then
        nroFacturaAnterior = vFactura.numero
        LlenarFactura
    End If
    

    If VVer Then
        LlenarFactura
        Me.txtTipoCambio.Enabled = False
        Me.cboTipoDocContable.Enabled = False
        Me.cboMonedas.Enabled = False
        Me.btnGuardar.Enabled = False
        Me.fraAlicuotas.Enabled = False
        Me.fraFormaPago.Enabled = False
        Me.Frame2.Enabled = False
        Me.Frame3.Enabled = False
        Me.cboProveedores.Enabled = False
        Me.cboTiposFactura.Enabled = False
        Me.txtImpuestos.Enabled = False
        Me.txtMontoNeto.Enabled = False
        Me.txtNumeroMask.Enabled = False
        Me.txtRedondeo.Enabled = False
        Me.DTPicker1.Enabled = False
        Me.lblTotal.Visible = True
        Me.Label10.Visible = True
        Me.Label17.Visible = False
        Me.txtMontoManual.Visible = False
        Me.btnNuevoCbte.Visible = False
        Me.txtTipoCambio = vFactura.TipoCambio
        Me.lblTipoCambioPago = "Tipo de cambio Pago: " & vFactura.TipoCambioPago

        grabado = True

    End If

    Me.lblTipoCambioPago.Visible = VVer
    
    TotalFactura
    
    Me.fraDetalleCtaCte.Visible = False
    Me.btnCtaCte.caption = "Ver Cta. Cte."
'''    Me.lblSaldoCtaCteProveedor.caption = "Saldo Cta. Cte.: 0.00"

    GridEXHelper.CustomizeGrid Me.gridDetalleCtaCte
    Me.gridDetalleCtaCte.ItemCount = 0
    

    loading = False
    
    'ESTA OPCION ES PARA ACTIVAR COMO CUENTA CORRIENTE SIEMPRE QUE SE ACTIVA EL FORM
    Me.optCtaCte.value = True
    Me.optContado.value = False
    
    Me.cboProveedores.Visible = True
    


    
End Sub

Private Sub FacturaRequiereNumeroFormateado()
    On Error GoTo ErrHandler
    
    Dim idx As Long
    Dim idtipoLocal As Long
    Dim cx As clsConfigFacturaProveedor
    
    If Me.cboTiposFactura.ListCount <= 0 Then Exit Sub
    
    idx = Me.cboTiposFactura.ListIndex
    If idx < 0 Then Exit Sub
    
    idtipoLocal = Me.cboTiposFactura.ItemData(idx)
    
    Set cx = DAOConfigFacturaProveedor.GetById(idtipoLocal)
    
    If Not IsSomething(cx) Then Exit Sub
    
    If cx.FormateaNumero Then
        Me.txtNumeroMask.SetMask "0000-00000000", "____-________"
        Me.txtNumeroMask.MaxLength = 13
    Else
        Me.txtNumeroMask.SetMask "", ""
        Me.txtNumeroMask.MaxLength = 16
    End If
    
    Exit Sub

ErrHandler:
    MsgBox "Error al configurar el formato del número: " & Err.Description, vbExclamation, "Error"
End Sub


Private Sub llenarGrillaPercepciones()
    Set colPercepciones = DAOPercepciones.GetAll
    Set colPercepcionesTMP = colPercepciones

    Me.grilla_percepcion.ItemCount = 0
    Me.grilla_percepcion.ItemCount = colPercepciones.count
    Set Me.grilla_percepciones.Columns("percepcion").DropDownControl = Me.grilla_percepcion
End Sub


Private Sub ProtegerProveedor()
    Me.GroupBox1.Enabled = (Proveedor.Id = 0)
    Me.btnDisponerProveedor.Visible = Me.GroupBox1.Enabled
End Sub


Private Sub TotalFactura()
    On Error GoTo er1
    Me.txtMontoNeto = funciones.FormatearDecimales(vFactura.NetoGravado)
    Me.lblTotal = funciones.FormatearDecimales(vFactura.total)
    Me.txtIVA.Text = funciones.FormatearDecimales(vFactura.TotalIVA)
    Me.fraAlicuotas.caption = Replace$(Me.fraAlicuotas.Tag, "{VALUE}", funciones.FormatearDecimales(vFactura.TotalIVA))
    Exit Sub
er1:
    Me.lblTotal = 0
End Sub

Private Sub Form_Terminate()
    Set colPercepciones = colPercepcionesTMP
End Sub

Private Sub grid_cuenta_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Set ctacontable = colCuentas.item(RowIndex)
    Values(1) = ctacontable.codigo & " - " & ctacontable.nombre
    Values(2) = ctacontable.Id
End Sub
Private Sub grid_cuentascontables_BeforeDelete(ByVal Cancel As GridEX20.JSRetBoolean)
    Cancel = Not (MsgBox("¿Está seguro de eliminar la cuenta contable seleccionada?", vbYesNo, "Confirmación") = vbYes)
End Sub
Private Sub grid_cuentascontables_BeforeUpdate(ByVal Cancel As GridEX20.JSRetBoolean)
    Cancel = (Not IsNumeric(Me.grid_cuentascontables.value(2))) Or (IsEmpty(Me.grid_cuentascontables.value(2)))
End Sub


Private Sub grid_cuentascontables_UnboundAddNew(ByVal NewRowBookmark As GridEX20.JSRetVariant, ByVal Values As GridEX20.JSRowData)
    Set ctaAplicada = New clsCuentaFactura
    ctaAplicada.Monto = funciones.FormatearDecimales(Values(2))
    ctaAplicada.cuentas = DAOCuentaContable.GetById(Values(1))
    vFactura.cuentasContables.Add ctaAplicada
    TotalFactura
    grabado = False
End Sub


Private Sub grid_cuentascontables_UnboundDelete(ByVal RowIndex As Long, ByVal Bookmark As Variant)
    vFactura.cuentasContables.remove RowIndex
    TotalFactura
    grabado = False
End Sub


Private Sub grid_cuentascontables_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    On Error GoTo ErrorHandler

    Set ctaAplicada = vFactura.cuentasContables.item(RowIndex)

    If Len(ctaAplicada.cuentas.codigo) > 0 Then
        Values(1) = ctaAplicada.cuentas.codigo & " - " & ctaAplicada.cuentas.nombre
        Values(2) = funciones.FormatearDecimales(ctaAplicada.Monto)
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Se produjo un error al cargar un valor vacío. Descripcion del Error:" & Err.Description, vbExclamation
End Sub


Private Sub grid_cuentascontables_UnboundUpdate(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If IsNumeric(Values(1)) And InStr(Values(1), ".") = 0 Then
        vFactura.cuentasContables(RowIndex).cuentas = DAOCuentaContable.GetById(Values(1))
    End If
    vFactura.cuentasContables(RowIndex).Monto = Values(2)
    TotalFactura
    grabado = False
End Sub


Private Sub gridDetalleCtaCte_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    On Error GoTo err1

    If detallesCtaCte Is Nothing Then Exit Sub
    If detallesCtaCte.count = 0 Then Exit Sub

    If RowIndex > 0 Then
        
        Set detaCtaCte = detallesCtaCte.item(RowIndex)

        Values(1) = detaCtaCte.FEcha
        Values(2) = detaCtaCte.Comprobante
        Values(3) = funciones.FormatearDecimales(detaCtaCte.Debe)
        Values(4) = funciones.FormatearDecimales(detaCtaCte.Haber)
        Values(5) = funciones.FormatearDecimales(detaCtaCte.saldo)
        
    End If

    Exit Sub

err1:
    MsgBox "Error al leer detalle de cuenta corriente: " & Err.Description, vbExclamation, "Error"
End Sub


Private Sub grilla_alicuota_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Set alicuota = colAlicuotas.item(RowIndex)
    Values(1) = funciones.FormatearDecimales(alicuota.alicuota)
    Values(2) = alicuota.Id
End Sub


Private Sub grilla_alicuotas_BeforeDelete(ByVal Cancel As GridEX20.JSRetBoolean)
    Cancel = Not (MsgBox("¿Está seguro de eliminar la alícuota seleccionada?", vbYesNo, "Confirmación") = vbYes)

End Sub


Private Sub AddDefaultAlicuota(id_alicuota As Long)
    Set aliaplicada = New clsAlicuotaAplicada
    aliaplicada.Monto = 0
    aliaplicada.alicuota = DAOAlicuotas.GetById(id_alicuota)
    vFactura.IvaAplicado.Add aliaplicada
    mostrarALicuotas
End Sub


Private Sub grilla_alicuotas_BeforeUpdate(ByVal Cancel As GridEX20.JSRetBoolean)
    Cancel = (Not IsNumeric(Me.grilla_alicuotas.value(2))) Or (Not IsNumeric(Me.grilla_alicuotas.value(1))) Or IsEmpty(Me.grilla_alicuotas.value(1))
End Sub

Private Sub grilla_alicuotas_GotFocus()
    grilla_alicuotas.SelStart = 0
    grilla_alicuotas.SelLength = -1
End Sub

Private Sub grilla_alicuotas_UnboundAddNew(ByVal NewRowBookmark As GridEX20.JSRetVariant, ByVal Values As GridEX20.JSRowData)
    Set aliaplicada = New clsAlicuotaAplicada

    aliaplicada.Monto = funciones.FormatearDecimales(Values(2))
    aliaplicada.alicuota = DAOAlicuotas.GetById(Values(1))
    vFactura.IvaAplicado.Add aliaplicada
    TotalFactura
    grabado = False

End Sub

Private Sub grilla_alicuotas_UnboundDelete(ByVal RowIndex As Long, ByVal Bookmark As Variant)
    vFactura.IvaAplicado.remove RowIndex
    TotalFactura
    grabado = False
End Sub

Private Sub grilla_alicuotas_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    On Error Resume Next
    Set aliaplicada = vFactura.IvaAplicado.item(RowIndex)
    Values(1) = funciones.FormatearDecimales(aliaplicada.alicuota.alicuota)
    Values(2) = funciones.FormatearDecimales(aliaplicada.Monto)
End Sub


Private Sub grilla_alicuotas_UnboundUpdate(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If IsNumeric(Values(1)) And InStr(Values(1), ".") = 0 Then
        vFactura.IvaAplicado(RowIndex).alicuota = DAOAlicuotas.GetById(Values(1))

    End If
    vFactura.IvaAplicado(RowIndex).Monto = Values(2)
    TotalFactura
    grabado = False
End Sub


Private Sub grilla_percepcion_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Set Percepcion = colPercepciones.item(RowIndex)
    Values(1) = Percepcion.Percepcion
    Values(2) = Percepcion.Id
End Sub


Private Sub grilla_percepciones_BeforeDelete(ByVal Cancel As GridEX20.JSRetBoolean)
    Cancel = Not (MsgBox("¿Está seguro de eliminar la percepción seleccionada?", vbYesNo, "Confirmación") = vbYes)
End Sub


Private Sub grilla_percepciones_BeforeUpdate(ByVal Cancel As GridEX20.JSRetBoolean)
    Cancel = (Not IsNumeric(Me.grilla_percepciones.value(2))) Or (IsEmpty(Me.grilla_percepciones.value(2)))
End Sub


Private Sub grilla_percepciones_UnboundAddNew(ByVal NewRowBookmark As GridEX20.JSRetVariant, ByVal Values As GridEX20.JSRowData)
    Set perAplicada = New clsPercepcionesAplicadas
    perAplicada.Monto = Values(2)
    perAplicada.Percepcion = DAOPercepciones.GetById(Values(1))
    vFactura.percepciones.Add perAplicada
    TotalFactura
    grabado = False
End Sub


Private Sub grilla_percepciones_UnboundDelete(ByVal RowIndex As Long, ByVal Bookmark As Variant)
    vFactura.percepciones.remove RowIndex
    TotalFactura
    grabado = False
End Sub


Private Sub grilla_percepciones_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    On Error GoTo ErrorHandler

    Set perAplicada = vFactura.percepciones.item(RowIndex)
    Values(1) = perAplicada.Percepcion.Percepcion
    Values(2) = funciones.FormatearDecimales(perAplicada.Monto)

    Exit Sub

ErrorHandler:
    MsgBox "Se produjo un error al cargar un valor vacío. Descripcion del Error: " & Err.Description, vbExclamation

End Sub


Private Sub grilla_percepciones_UnboundUpdate(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If IsNumeric(Values(1)) And InStr(Values(1), ".") = 0 Then
        vFactura.percepciones(RowIndex).Percepcion = DAOPercepciones.GetById(Values(1))
    End If
    vFactura.percepciones(RowIndex).Monto = Values(2)
    TotalFactura
    grabado = False
End Sub


Private Sub optContado_Click()
    vFactura.FormaPagoCuentaCorriente = False
End Sub


Private Sub optCtaCte_Click()
    vFactura.FormaPagoCuentaCorriente = True
End Sub


Private Sub btnNuevoProveedor_Click()
    Set Proveedor = New clsProveedor
    LimpiarProveedor
    ProtegerProveedor
End Sub


Private Sub btnNuevoCbte_Click()
    Dim frm1 As New frmAdminComprasNuevaFCProveedor
    frm1.Factura = Nothing
'''    frm1.Factura.total = 0
    frm1.cboProveedores.ListIndex = Me.cboProveedores.ListIndex
    frm1.Show
    
    frm1.Top = 100
    frm1.Left = 100
    Unload Me
End Sub


Private Sub txtCodigoProveedor_Change()
    On Error Resume Next
    Me.cboProveedores.ListIndex = funciones.PosIndexCbo(CLng(Me.txtCodigoProveedor), Me.cboProveedores)

End Sub


Private Sub txtCodigoProveedor_Validate(Cancel As Boolean)
    If Not IsNumeric(Me.txtCodigoProveedor) Then Cancel = True Else Cancel = False
End Sub


Private Sub txtCuit_Validate(Cancel As Boolean)
    On Error GoTo err1

    Dim cuitLimpio As String
    Dim F As String
    Dim duplicados As Collection

    cuitLimpio = Replace(Trim$(Me.txtCuit.Text), "-", vbNullString)

    If cuitLimpio = vbNullString Then Exit Sub

    'Si está editando un proveedor existente, excluyo el mismo proveedor
    F = "proveedores.cuit = " & Escape(cuitLimpio)

    If IsSomething(Proveedor) Then
        If Proveedor.Id > 0 Then
            F = F & " AND proveedores.id <> " & Proveedor.Id
        End If
    End If

    Set duplicados = DAOProveedor.FindAll(F)

    If duplicados.count > 0 Then

        'Solo permito CUIT duplicado si el proveedor que estoy cargando es del exterior
        If Not EsProveedorExteriorSeleccionado() Then
            Cancel = True
            MsgBox "Ya existe un proveedor con ese CUIT. Solo se permite CUIT duplicado para proveedores del exterior.", vbExclamation
            Exit Sub
        End If

    End If

    Exit Sub

err1:
    Cancel = True
    MsgBox "Error al validar CUIT del proveedor: " & Err.Description, vbCritical, "Error"
End Sub


Private Sub txtImpuestos_Change()
    On Error Resume Next
    vFactura.ImpuestoInterno = CDbl(Me.txtImpuestos)
    TotalFactura
    grabado = False
End Sub


Private Sub txtImpuestos_GotFocus()
    foco Me.txtImpuestos
End Sub


'''Private Function validarFactura() As Boolean
'''    validarFactura = (vFactura.total = funciones.RedondearDecimales(CDbl(Me.txtMontoManual)))
'''End Function

Private Function validarFactura() As Boolean
    Dim totalManual As Double

    If Not TryGetDouble(Me.txtMontoManual.Text, totalManual) Then
        validarFactura = False
        Exit Function
    End If

    validarFactura = funciones.RedondearDecimales(vFactura.total) = funciones.RedondearDecimales(totalManual)
End Function


Private Sub txtMontoManual_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then btnGuardar_Click
End Sub


Private Sub txtMontoNeto_Change()
    grabado = False
    TotalFactura
End Sub


Private Sub txtMontoNeto_GotFocus()
    foco Me.txtMontoNeto
End Sub


'Private Sub txtMontoPercep_Change()
'    grabado = False
'End Sub

'Private Sub txtNumero_Change()
'    On Error Resume Next
'    vFactura.Numero = Me.txtNumero
'    TotalFactura
'    grabado = False
'End Sub

'Private Sub txtNumero_GotFocus()
'    foco Me.txtNumero
'End Sub


Private Sub llenarAlicuotas(idtipo As Long)
    Me.grilla_alicuota.ItemCount = 0
    Set colAlicuotas = DAOAlicuotas.getByTipoFactura(idtipo)
    Me.grilla_alicuota.ItemCount = colAlicuotas.count
    Set Me.grilla_alicuotas.Columns("alicuota").DropDownControl = Me.grilla_alicuota
End Sub


'Private Sub limpiar()
'    If MsgBox("¿Desea limpiar la factura?", vbYesNo, "Confirmación") Then
'        Me.txtNumeroMask.text = vbNullString
'    End If
'End Sub

'Private Sub txtNoGravado_Change()
'    On Error Resume Next
'    vFactura.ConceptoNoGravado = CDbl(Me.txtNoGravado)
'    TotalFactura
'    grabado = False
'End Sub

'Private Sub txtNumeroMask_Change()
'    On Error Resume Next
'    If Me.txtNumeroMask.text <> "______-________" Then
'        '    If Me.txtNumeroMask.text <> "" Then
'        vFactura.numero = Me.txtNumeroMask.text
'        TotalFactura
'        grabado = False
'    End If
'
'
'End Sub

Private Sub txtNumeroMask_GotFocus()
    foco Me.txtNumeroMask
End Sub


Private Sub txtRedondeo_Change()
    On Error Resume Next
    vFactura.redondeo = CDbl(Me.txtRedondeo)
    TotalFactura
    grabado = False
End Sub


Private Sub txtRedondeo_GotFocus()
    foco Me.txtRedondeo
End Sub


Private Sub LlenarFactura()

    Me.cboTipoDocContable.ListIndex = funciones.PosIndexCbo(vFactura.tipoDocumentoContable, Me.cboTipoDocContable)
    Me.txtImpuestos = funciones.FormatearDecimales(vFactura.ImpuestoInterno)
    Me.DTPicker1 = vFactura.FEcha

    '    Me.Label12.Visible = False
    Me.txtNumeroMask.Visible = False

    Me.txtNumeroCargado.Visible = True
    Me.txtNumeroCargado.Text = vFactura.numero

    Me.txtRedondeo = vFactura.redondeo
    'Me.txtNoGravado = vFactura.ConceptoNoGravado
    Me.cboProveedores.ListIndex = funciones.PosIndexCbo(vFactura.Proveedor.Id, Me.cboProveedores)
    Me.cboTiposFactura.ListIndex = funciones.PosIndexCbo(vFactura.configFactura.Id, Me.cboTiposFactura)
    Me.cboMonedas.ListIndex = funciones.PosIndexCbo(vFactura.moneda.Id, Me.cboMonedas)
    Me.txtMontoNeto = vFactura.NetoGravado

    Me.optContado.value = Not vFactura.FormaPagoCuentaCorriente
    Me.optCtaCte.value = vFactura.FormaPagoCuentaCorriente

    Me.grid_cuentascontables.ItemCount = 0
    Me.grid_cuentascontables.ItemCount = vFactura.cuentasContables.count

    mostrarALicuotas

    Me.grilla_percepciones.ItemCount = 0
    Me.grilla_percepciones.ItemCount = vFactura.percepciones.count
    
    If vFactura.EsArca Then
        Me.chkButton.value = xtpChecked
    Else
        Me.chkButton.value = xtpUnchecked
    End If

    grabado = True
End Sub


Private Sub mostrarALicuotas()
    Me.grilla_alicuotas.ItemCount = 0
    Me.grilla_alicuotas.ItemCount = vFactura.IvaAplicado.count
End Sub


Private Sub armarFactura()
    vFactura.FEcha = (CDate(Format(Me.DTPicker1, "yyyy-mm-dd")))

    If Me.txtNumeroCargado.Text = "txtNumeroCargado" Then
        vFactura.numero = Me.txtNumeroMask.Text
    Else
        vFactura.numero = Me.txtNumeroCargado.Text
    End If

    vFactura.Proveedor = Proveedor
    
    vFactura.ImpuestoInterno = CDbl(Me.txtImpuestos)
    vFactura.Monto = CDbl(Me.txtMontoNeto)
    
    vFactura.estado = EstadoFacturaProveedor.EnProceso

    idtipo = Me.cboTiposFactura.ItemData(Me.cboTiposFactura.ListIndex)
    vFactura.tipoDocumentoContable = Me.cboTipoDocContable.ItemData(Me.cboTipoDocContable.ListIndex)
    vFactura.configFactura = DAOConfigFacturaProveedor.GetById(idtipo)
    
    vFactura.EsArca = (Me.chkButton.value = xtpChecked)

End Sub


Private Sub llenarComboProveedores()

'''    DAOProveedor.llenarComboXtremeSuite Me.cboProveedores, True, True, False

    DAOProveedor.llenarComboProveedores Me.cboProveedores

    
End Sub


Private Sub txtTipoCambio_Change()
    On Error Resume Next
    vFactura.TipoCambio = val(Me.txtTipoCambio)
    TotalFactura
    grabado = False
End Sub


Private Sub CargarDetalleCtaCteProveedor()
    On Error GoTo err1

    Dim Id As Long
    Dim condition As String

    Id = CLng(Me.cboProveedores.ItemData(Me.cboProveedores.ListIndex))

    condition = conectar.Escape(Format(Me.DTPicker1.value, "yyyy-mm-dd"))

    Set detallesCtaCte = DAOCuentaCorriente.FindAllDetallesProveedor2(Id, , condition, True)

    Me.gridDetalleCtaCte.ItemCount = 0

    If IsSomething(detallesCtaCte) Then
        
        If detallesCtaCte.count > 0 Then
            Me.gridDetalleCtaCte.ItemCount = detallesCtaCte.count
            GridEXHelper.AutoSizeColumns Me.gridDetalleCtaCte
        End If

    End If

    Me.gridDetalleCtaCte.Refresh

    Exit Sub

err1:
    MsgBox "Error al cargar el detalle de cuenta corriente: " & Err.Description, vbCritical, "Error"
End Sub


Private Sub OcultarDetalleCtaCte()
    On Error Resume Next

    If anchoNormal <= 0 Then anchoNormal = Me.Width

    Me.fraDetalleCtaCte.Visible = False
    Me.Width = anchoNormal
    Me.btnCtaCte.caption = "Ver Cta. Cte."
    detalleCtaCteVisible = False

    Me.gridDetalleCtaCte.ItemCount = 0
    Set detallesCtaCte = Nothing
End Sub


Private Sub MostrarSaldoCtaCteProveedor()
    On Error GoTo err1

    Dim Id As Long
    Dim condition As String
    Dim detallesSaldo As Collection
    Dim saldoProv As Double

'''    Me.lblSaldoCtaCteProveedor.caption = "Saldo Cta. Cte.: 0.00"

    If Me.cboProveedores.ListIndex = -1 Then Exit Sub

    Id = CLng(Me.cboProveedores.ItemData(Me.cboProveedores.ListIndex))
    If Id <= 0 Then Exit Sub

    condition = conectar.Escape(Format(Me.DTPicker1.value, "yyyy-mm-dd"))

    Set detallesSaldo = DAOCuentaCorriente.FindAllDetallesProveedor2(Id, , condition, True)

'''    If IsSomething(detallesSaldo) Then
'''        saldoProv = DAOCuentaCorriente.GetSaldo(detallesSaldo)
'''
'''        Me.lblSaldoCtaCteProveedor.caption = "Saldo Cta. Cte.: " & _
'''            Replace(FormatCurrency(funciones.FormatearDecimales(saldoProv)), "$", "")
'''    End If

    Exit Sub

err1:
'''    Me.lblSaldoCtaCteProveedor.caption = "Saldo Cta. Cte.: -"
End Sub


Private Function EsProveedorExteriorSeleccionado() As Boolean
    On Error GoTo err1

    Dim textoTipoIva As String

    If Me.cboTipoIva.ListIndex = -1 Then Exit Function

    textoTipoIva = UCase$(Trim$(Me.cboTipoIva.list(Me.cboTipoIva.ListIndex)))

    EsProveedorExteriorSeleccionado = _
        (InStr(1, textoTipoIva, "EXTERIOR", vbTextCompare) > 0) Or _
        (InStr(1, textoTipoIva, "EXTRANJERO", vbTextCompare) > 0)

    Exit Function

err1:
    EsProveedorExteriorSeleccionado = False
End Function
