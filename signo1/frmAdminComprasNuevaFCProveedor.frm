VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmAdminComprasNuevaFCProveedor 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Factura de Proveedor"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   480
   ClientWidth     =   10515
   Icon            =   "frmAdminComprasNuevaFCProveedor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   10515
   Begin VB.TextBox txtTipoCambio 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   8700
      TabIndex        =   51
      TabStop         =   0   'False
      Text            =   "0"
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
   Begin XtremeSuiteControls.PushButton PushButton1 
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
      Left            =   75
      TabIndex        =   40
      Top             =   975
      Width           =   6435
      _Version        =   786432
      _ExtentX        =   11351
      _ExtentY        =   2646
      _StockProps     =   79
      Caption         =   "Datos del proveedor"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.PushButton cmdDisponer 
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
      Left            =   105
      TabIndex        =   37
      Top             =   4770
      Width           =   3765
      _Version        =   786432
      _ExtentX        =   6641
      _ExtentY        =   3651
      _StockProps     =   79
      Caption         =   "Percepciones"
      UseVisualStyle  =   -1  'True
      Begin GridEX20.GridEX grilla_percepciones 
         Height          =   1650
         Left            =   120
         TabIndex        =   13
         Top             =   255
         Width           =   3495
         _ExtentX        =   6165
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
         Column(1)       =   "frmAdminComprasNuevaFCProveedor.frx":000C
         Column(2)       =   "frmAdminComprasNuevaFCProveedor.frx":0154
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmAdminComprasNuevaFCProveedor.frx":0268
         FormatStyle(2)  =   "frmAdminComprasNuevaFCProveedor.frx":03A0
         FormatStyle(3)  =   "frmAdminComprasNuevaFCProveedor.frx":0450
         FormatStyle(4)  =   "frmAdminComprasNuevaFCProveedor.frx":0504
         FormatStyle(5)  =   "frmAdminComprasNuevaFCProveedor.frx":05DC
         FormatStyle(6)  =   "frmAdminComprasNuevaFCProveedor.frx":0694
         ImageCount      =   0
         PrinterProperties=   "frmAdminComprasNuevaFCProveedor.frx":0774
      End
   End
   Begin XtremeSuiteControls.GroupBox fraAlicuotas 
      Height          =   2115
      Left            =   120
      TabIndex        =   36
      Tag             =   "Alicuotas IVA (Total: {VALUE})"
      Top             =   2610
      Width           =   2445
      _Version        =   786432
      _ExtentX        =   4313
      _ExtentY        =   3731
      _StockProps     =   79
      Caption         =   "Alicuotas IVA"
      UseVisualStyle  =   -1  'True
      Begin GridEX20.GridEX grilla_alicuotas 
         Height          =   1680
         Left            =   105
         TabIndex        =   11
         Top             =   255
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   2963
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
         Column(1)       =   "frmAdminComprasNuevaFCProveedor.frx":094C
         Column(2)       =   "frmAdminComprasNuevaFCProveedor.frx":0AB8
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmAdminComprasNuevaFCProveedor.frx":0BCC
         FormatStyle(2)  =   "frmAdminComprasNuevaFCProveedor.frx":0D04
         FormatStyle(3)  =   "frmAdminComprasNuevaFCProveedor.frx":0DB4
         FormatStyle(4)  =   "frmAdminComprasNuevaFCProveedor.frx":0E68
         FormatStyle(5)  =   "frmAdminComprasNuevaFCProveedor.frx":0F40
         FormatStyle(6)  =   "frmAdminComprasNuevaFCProveedor.frx":0FF8
         ImageCount      =   0
         PrinterProperties=   "frmAdminComprasNuevaFCProveedor.frx":10D8
      End
   End
   Begin GridEX20.GridEX grilla_alicuota 
      Height          =   1485
      Left            =   555
      TabIndex        =   34
      Top             =   3060
      Visible         =   0   'False
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   2619
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
      Column(1)       =   "frmAdminComprasNuevaFCProveedor.frx":12B0
      Column(2)       =   "frmAdminComprasNuevaFCProveedor.frx":13D0
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminComprasNuevaFCProveedor.frx":14D0
      FormatStyle(2)  =   "frmAdminComprasNuevaFCProveedor.frx":1608
      FormatStyle(3)  =   "frmAdminComprasNuevaFCProveedor.frx":16B8
      FormatStyle(4)  =   "frmAdminComprasNuevaFCProveedor.frx":176C
      FormatStyle(5)  =   "frmAdminComprasNuevaFCProveedor.frx":1844
      FormatStyle(6)  =   "frmAdminComprasNuevaFCProveedor.frx":18FC
      ImageCount      =   0
      PrinterProperties=   "frmAdminComprasNuevaFCProveedor.frx":19DC
   End
   Begin XtremeSuiteControls.PushButton cmdGuardar 
      Height          =   390
      Left            =   8745
      TabIndex        =   20
      Top             =   6480
      Width           =   1635
      _Version        =   786432
      _ExtentX        =   2884
      _ExtentY        =   688
      _StockProps     =   79
      Caption         =   "Guardar"
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
      Text            =   "0"
      Top             =   6060
      Width           =   1575
   End
   Begin VB.TextBox txtRedondeo 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   8700
      TabIndex        =   8
      Text            =   "0"
      Top             =   2475
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
      Top             =   2130
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
      Format          =   59244545
      CurrentDate     =   39897
   End
   Begin XtremeSuiteControls.GroupBox frame3 
      Height          =   2115
      Left            =   2655
      TabIndex        =   38
      Top             =   2610
      Width           =   3840
      _Version        =   786432
      _ExtentX        =   6773
      _ExtentY        =   3731
      _StockProps     =   79
      Caption         =   "Cuentas Contables"
      UseVisualStyle  =   -1  'True
      Begin GridEX20.GridEX grid_cuentascontables 
         Height          =   1650
         Left            =   135
         TabIndex        =   12
         Top             =   285
         Width           =   3555
         _ExtentX        =   6271
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
         Column(1)       =   "frmAdminComprasNuevaFCProveedor.frx":1BB4
         Column(2)       =   "frmAdminComprasNuevaFCProveedor.frx":1CF0
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmAdminComprasNuevaFCProveedor.frx":1E04
         FormatStyle(2)  =   "frmAdminComprasNuevaFCProveedor.frx":1F3C
         FormatStyle(3)  =   "frmAdminComprasNuevaFCProveedor.frx":1FEC
         FormatStyle(4)  =   "frmAdminComprasNuevaFCProveedor.frx":20A0
         FormatStyle(5)  =   "frmAdminComprasNuevaFCProveedor.frx":2178
         FormatStyle(6)  =   "frmAdminComprasNuevaFCProveedor.frx":2230
         ImageCount      =   0
         PrinterProperties=   "frmAdminComprasNuevaFCProveedor.frx":2310
      End
   End
   Begin GridEX20.GridEX grid_cuenta 
      Height          =   1485
      Left            =   3195
      TabIndex        =   39
      Top             =   2970
      Visible         =   0   'False
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   2619
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
      Column(1)       =   "frmAdminComprasNuevaFCProveedor.frx":24E8
      Column(2)       =   "frmAdminComprasNuevaFCProveedor.frx":2600
      SortKeysCount   =   1
      SortKey(1)      =   "frmAdminComprasNuevaFCProveedor.frx":2700
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminComprasNuevaFCProveedor.frx":2768
      FormatStyle(2)  =   "frmAdminComprasNuevaFCProveedor.frx":28A0
      FormatStyle(3)  =   "frmAdminComprasNuevaFCProveedor.frx":2950
      FormatStyle(4)  =   "frmAdminComprasNuevaFCProveedor.frx":2A04
      FormatStyle(5)  =   "frmAdminComprasNuevaFCProveedor.frx":2ADC
      FormatStyle(6)  =   "frmAdminComprasNuevaFCProveedor.frx":2B94
      ImageCount      =   0
      PrinterProperties=   "frmAdminComprasNuevaFCProveedor.frx":2C74
   End
   Begin GridEX20.GridEX grilla_percepcion 
      Height          =   1485
      Left            =   885
      TabIndex        =   35
      Top             =   5085
      Visible         =   0   'False
      Width           =   2340
      _ExtentX        =   4128
      _ExtentY        =   2619
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
      Column(1)       =   "frmAdminComprasNuevaFCProveedor.frx":2E4C
      Column(2)       =   "frmAdminComprasNuevaFCProveedor.frx":2F74
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminComprasNuevaFCProveedor.frx":3074
      FormatStyle(2)  =   "frmAdminComprasNuevaFCProveedor.frx":31AC
      FormatStyle(3)  =   "frmAdminComprasNuevaFCProveedor.frx":325C
      FormatStyle(4)  =   "frmAdminComprasNuevaFCProveedor.frx":3310
      FormatStyle(5)  =   "frmAdminComprasNuevaFCProveedor.frx":33E8
      FormatStyle(6)  =   "frmAdminComprasNuevaFCProveedor.frx":34A0
      ImageCount      =   0
      PrinterProperties=   "frmAdminComprasNuevaFCProveedor.frx":3580
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
   Begin XtremeSuiteControls.PushButton PushButton2 
      Height          =   390
      Left            =   7035
      TabIndex        =   23
      Top             =   6480
      Width           =   1635
      _Version        =   786432
      _ExtentX        =   2884
      _ExtentY        =   688
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
   Begin VB.Label lblTipoCambioPago 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      Caption         =   "Tipo de Cambio pago: "
      Height          =   255
      Left            =   7005
      TabIndex        =   53
      Top             =   3210
      Width           =   3345
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      Caption         =   "Tipo de Cambio"
      Height          =   255
      Left            =   6990
      TabIndex        =   52
      Top             =   3645
      Width           =   1575
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      Caption         =   "IVA"
      Height          =   255
      Left            =   7035
      TabIndex        =   46
      Top             =   4350
      Width           =   1575
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      Caption         =   "Letra"
      Height          =   255
      Left            =   7320
      TabIndex        =   45
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label lblMoneda 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Moneda"
      Height          =   195
      Left            =   8010
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
      Left            =   7335
      TabIndex        =   31
      Top             =   6090
      Width           =   1335
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      Caption         =   "Redondeo IVA"
      Height          =   255
      Left            =   7260
      TabIndex        =   30
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      Caption         =   "Neto Gravado"
      Height          =   255
      Left            =   7020
      TabIndex        =   29
      Top             =   4005
      Width           =   1575
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      Caption         =   "Impuestos"
      Height          =   255
      Left            =   7140
      TabIndex        =   28
      Top             =   2145
      Width           =   1455
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      Caption         =   "Número"
      Height          =   285
      Left            =   7635
      TabIndex        =   26
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      Caption         =   "Fecha"
      Height          =   255
      Left            =   7635
      TabIndex        =   25
      Top             =   1470
      Width           =   975
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      Caption         =   "Total Factura"
      Height          =   255
      Left            =   7275
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
      Left            =   7080
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



Dim perAplicada As clsPercepcionesAplicadas

Dim colCuentas As Collection


Dim ctaAplicada As clsCuentaFactura
Dim ctaContable As clsCuentaContable
Dim Percepcion As New clsPercepciones
Dim alicuota As clsAlicuotas
Dim idtipo As Long
Dim nroFacturaAnterior
Dim grabado As Boolean
Dim Proveedor As clsProveedor
Dim idProveedor As Long
Dim vFactura As clsFacturaProveedor
Dim VVer As Boolean

Public Property Let ver(nVer As Boolean)
    VVer = nVer
End Property
Public Property Let Factura(nFactura As clsFacturaProveedor)
    If IsSomething(nFactura) Then
        Set vFactura = DAOFacturaProveedor.FindById(nFactura.id)
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
        idIVA = Proveedor.TipoIVA.id
        Me.cboTiposFactura.AddItem Proveedor.TipoIVA.configFacturas(i).TipoFactura
        Me.cboTiposFactura.ItemData(Me.cboTiposFactura.NewIndex) = Proveedor.TipoIVA.configFacturas(i).id
    Next i



    Dim idtipo As Long

    If Me.cboTiposFactura.ListCount > 0 Then
        Me.cboTiposFactura.ListIndex = 0
        idtipo = Me.cboTiposFactura.ItemData(Me.cboTiposFactura.ListIndex)
        llenarAlicuotas idtipo
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
     

End Sub






Private Sub cboTiposFactura_Click()

    grabado = False
    'vFactura.IvaAplicado = Nothing
FacturaRequiereNumeroFormateado
    If Me.cboTiposFactura.ListCount > 0 Then
        Dim idtipo As Long
        idtipo = Me.cboTiposFactura.ItemData(Me.cboTiposFactura.ListIndex)
        llenarAlicuotas idtipo

        Dim id_ali As Long

        If Not loading Then

            If colAlicuotas.count > 0 Then
                vFactura.IvaAplicado = Nothing
                Me.grilla_alicuotas.ItemCount = 0
                Me.grilla_alicuotas.Refresh
                AddDefaultAlicuota colAlicuotas(1).id

            End If
        End If


    End If
End Sub
Private Sub mostrar()
    If Me.cboProveedores.ListIndex <> -1 Then
        idProveedor = CLng(Me.cboProveedores.ItemData(Me.cboProveedores.ListIndex))
        Dim lstRubros As ListView
        Dim accion As Integer
        Set Proveedor = DAOProveedor.FindById(idProveedor, False, False, False, False)
        If IsSomething(Proveedor) Then
            Me.txtCuit = Proveedor.Cuit
            Me.txtIB = Proveedor.IIBB
            Me.txtRazonSocial = Proveedor.RazonSocial
            Me.cboTipoIva.ListIndex = funciones.PosIndexCbo(Proveedor.TipoIVA.id, Me.cboTipoIva)
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


Private Sub cmdDisponer_Click()


    If Proveedor.id = 0 Then
        Proveedor.RazonSocial = Me.txtRazonSocial
        Proveedor.Cuit = Replace(Me.txtCuit, "-", vbNullString)
        Proveedor.IIBB = Me.txtIB
        Proveedor.estado = EstadoProveedorContado
        Set Proveedor.TipoIVA = DAOTipoIvaProveedor.GetById(Me.cboTipoIva.ItemData(Me.cboTipoIva.ListIndex))
        llenarTiposFacturas
    End If
End Sub

Private Sub cmdGuardar_Click()
    On Error GoTo err1

    If Me.grilla_alicuotas.EditMode = jgexEditModeOn Then
        MsgBox "Todavia esta editando la grilla de Alicuotas de IVA." & vbNewLine & "Presione [ENTER] en la grilla para guardar los cambios.", vbExclamation + vbOKOnly
        Exit Sub
    End If


    If Me.grilla_percepciones.EditMode = jgexEditModeOn Then
        MsgBox "Todavia esta editando la grilla de Percepciones." & vbNewLine & "Presione [ENTER] en la grilla para guardar los cambios.", vbExclamation + vbOKOnly
        Exit Sub
    End If

    If Me.grid_cuentascontables.EditMode = jgexEditModeOn Then
        MsgBox "Todavia esta editando la grilla de Cuentas Contables." & vbNewLine & "Presione [ENTER] en la grilla para guardar los cambios.", vbExclamation + vbOKOnly
        Exit Sub
    End If


    If Not Me.optContado.value And Not Me.optCtaCte.value Then
        MsgBox "Debe seleccionar la forma de pago.", vbExclamation
        Exit Sub
    End If


    conectar.BeginTransaction
    Dim A As Boolean
    Dim montonero As Double
    Dim nroNuevo As Long
    Dim EVENTO As clsEventoObserver
    Dim nuevoproveedor As Boolean

    If Not validarFactura Then
        Err.Raise 203
    End If

    'If MsgBox("¿Está seguro de guardar la factura?", vbYesNo, "Confirmación") = vbYes Then
    armarFactura
    If vFactura.NetoGravado <= 0 Then
        If vFactura.tipoDocumentoContable <> notaDebito Then
            Err.Raise 202
        End If
    End If


    montonero = CDbl(Me.txtMontoNeto)
    If Me.txtNumeroMask.text <> "____-________" And Len(Me.txtNumeroMask.text) > 0 Then


        If vFactura.cuentasContables.count = 0 And vFactura.tipoDocumentoContable <> notaDebito Then Err.Raise 201

        If funciones.RedondearDecimales(vFactura.TotalAplicadoACuentas) <> funciones.RedondearDecimales(vFactura.NetoGravado) Then Err.Raise 200

        If Me.cboMonedas.ListIndex <> -1 Then
            Set vFactura.moneda = DAOMoneda.GetById(Me.cboMonedas.ItemData(Me.cboMonedas.ListIndex))
        Else
            Set vFactura.moneda = Nothing
        End If

        'creo el proveedor si es contado

        If vFactura.Proveedor.id = 0 Then
            If Trim(Me.txtRazonSocial) = vbNullString Or Not IsNumeric(Replace(Me.txtCuit, "-", vbNullString)) Then
                If Not funciones.VerificarCUIT(Replace(Me.txtCuit, "-", vbNullString)) Then
                    Err.Raise 1000
                End If
            Else
                nuevoproveedor = True
                Set colPercepcionesTMP = colPercepciones
                If Not DAOProveedor.Guardar(vFactura.Proveedor) Then Err.Raise 300
            End If

        End If


        If DAOFacturaProveedor.existeFactura(vFactura) Then Err.Raise 101

        Dim NUEVA As Boolean
        NUEVA = (vFactura.id = 0)

        If DAOFacturaProveedor.Guardar(vFactura) Then
            Set EVENTO = New clsEventoObserver
            Set EVENTO.Elemento = vFactura
            If NUEVA Then
                EVENTO.EVENTO = agregar_
            Else
                EVENTO.EVENTO = modificar_
            End If
            Set EVENTO.Originador = Me
            EVENTO.Tipo = TipoSuscripcion.FacturaProveedor_
            
            ' Desactivo este evento Notificar porque aparentemente da Error (dienemer 11.09.20)
                   'Channel.Notificar EVENTO, TipoSuscripcion.FacturaProveedor_
                   
            MsgBox "Factura almacenada con éxito!", vbInformation, "Información"
            grabado = True
        Else
            Err.Raise 100
        End If
    Else
        Err.Raise 101
    End If
    'End If
    conectar.CommitTransaction
    Exit Sub

err1:
    conectar.RollBackTransaction
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
        MsgBox "Los totales de la factura no coinciden." & vbNewLine & "Total esperado: " & funciones.RedondearDecimales(CDbl(Me.txtMontoManual)) & vbNewLine & "Total ingresado: " & vFactura.Total, vbCritical, "Error"
    ElseIf Err.Number = 300 Or nuevoproveedor Then
        vFactura.Proveedor = Nothing
        nuevoproveedor = False
    Else
        MsgBox Err.Description, vbCritical
    End If
End Sub
Private Sub Command2_Click()
    grabado = False
    TotalFactura
End Sub
Private Sub DTPicker1_Click()
    grabado = False
End Sub

Private Sub Form_Load()
    loading = True
    
  
    Me.txtCuit.SetMask "00-00000000-0", "__-________-_"

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


    Me.cboTipoDocContable.ListIndex = 1

    Me.grilla_alicuotas.ItemCount = 0
    Me.grilla_percepciones.ItemCount = 0
    Me.grid_cuentascontables.ItemCount = 0
'    If vFactura.configFactura.FormateaNumero Then
'       Me.txtNumeroMask.SetMask "0000-00000000", "____-________"
'       Me.txtNumeroMask.MaxLength = 0
'  End If
FacturaRequiereNumeroFormateado
    llenarGrillaPercepciones
    LlenarCuentasContables


    Me.DTPicker1 = Now
    If vFactura.id > 0 Then
        nroFacturaAnterior = vFactura.numero
        LlenarFactura
    End If


    If VVer Then
        LlenarFactura
        Me.txtTipoCambio.Enabled = False
        Me.cboTipoDocContable.Enabled = False
        Me.cboMonedas.Enabled = False
        Me.cmdGuardar.Enabled = False
        Me.fraAlicuotas.Enabled = False
        Me.fraFormaPago.Enabled = False
        Me.frame2.Enabled = False
        Me.frame3.Enabled = False
        Me.cboProveedores.Enabled = False
        Me.cboTiposFactura.Enabled = False
        Me.txtImpuestos.Enabled = False
        Me.txtMontoNeto.Enabled = False
        Me.txtNumeroMask.Enabled = False
        Me.txtRedondeo.Enabled = False
        'Me.txtNoGravado.Enabled = False
        Me.DTPicker1.Enabled = False
        Me.lblTotal.Visible = True
        Me.Label10.Visible = True
        Me.Label17.Visible = False
        Me.txtMontoManual.Visible = False
        Me.PushButton2.Visible = False
        Me.txtTipoCambio = vFactura.TipoCambio


        Me.lblTipoCambioPago = "Tipo de cambio Pago: " & vFactura.TipoCambioPago
        grabado = True
    End If
    Me.lblTipoCambioPago.Visible = VVer
    TotalFactura
    loading = False



End Sub

Private Sub FacturaRequiereNumeroFormateado()
'dnemer
    Dim idtipo As Integer
    idtipo = Me.cboTiposFactura.ItemData(Me.cboTiposFactura.ListIndex)
 Dim cx As clsConfigFacturaProveedor
 
 Set cx = DAOConfigFacturaProveedor.GetById(idtipo)

If IsSomething(cx) Then
   If cx.FormateaNumero Then
       Me.txtNumeroMask.SetMask "0000-00000000", "____-________"
       Me.txtNumeroMask.MaxLength = 13
  Else
    Me.txtNumeroMask.SetMask "", ""
       Me.txtNumeroMask.MaxLength = 0
  End If
End If
End Sub

Private Sub llenarGrillaPercepciones()
    Set colPercepciones = DAOPercepciones.GetAll
    Set colPercepcionesTMP = colPercepciones

    Me.grilla_percepcion.ItemCount = 0
    Me.grilla_percepcion.ItemCount = colPercepciones.count
    Set Me.grilla_percepciones.Columns("percepcion").DropDownControl = Me.grilla_percepcion
End Sub
Private Sub ProtegerProveedor()
    Me.GroupBox1.Enabled = (Proveedor.id = 0)
    Me.cmdDisponer.Visible = Me.GroupBox1.Enabled
End Sub

Private Sub TotalFactura()
    On Error GoTo er1
    Me.txtMontoNeto = funciones.FormatearDecimales(vFactura.NetoGravado)
    Me.lblTotal = funciones.FormatearDecimales(vFactura.Total)
    Me.txtIVA.text = funciones.FormatearDecimales(vFactura.TotalIVA)

    Me.fraAlicuotas.caption = Replace$(Me.fraAlicuotas.Tag, "{VALUE}", funciones.FormatearDecimales(vFactura.TotalIVA))
    Exit Sub
er1:
    Me.lblTotal = 0
End Sub

Private Sub Form_Terminate()
    Set colPercepciones = colPercepcionesTMP
End Sub

Private Sub grid_cuenta_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Set ctaContable = colCuentas.item(RowIndex)
    Values(1) = ctaContable.codigo & " - " & ctaContable.nombre
    Values(2) = ctaContable.id
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
    Set ctaAplicada = vFactura.cuentasContables.item(RowIndex)
    Values(1) = ctaAplicada.cuentas.codigo & " - " & ctaAplicada.cuentas.nombre
    Values(2) = funciones.FormatearDecimales(ctaAplicada.Monto)
End Sub
Private Sub grid_cuentascontables_UnboundUpdate(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If IsNumeric(Values(1)) And InStr(Values(1), ".") = 0 Then
        vFactura.cuentasContables(RowIndex).cuentas = DAOCuentaContable.GetById(Values(1))
    End If
    vFactura.cuentasContables(RowIndex).Monto = Values(2)
    TotalFactura
    grabado = False
End Sub

Private Sub grilla_alicuota_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Set alicuota = colAlicuotas.item(RowIndex)
    Values(1) = funciones.FormatearDecimales(alicuota.alicuota)
    Values(2) = alicuota.id
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
    Values(2) = Percepcion.id
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
    Set perAplicada = vFactura.percepciones.item(RowIndex)
    Values(1) = perAplicada.Percepcion.Percepcion
    Values(2) = funciones.FormatearDecimales(perAplicada.Monto)
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

Private Sub PushButton1_Click()
    Set Proveedor = New clsProveedor
    LimpiarProveedor
    ProtegerProveedor
End Sub

Private Sub PushButton2_Click()
    Dim frm1 As New frmAdminComprasNuevaFCProveedor
    frm1.Factura = Nothing
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
    Dim F As String
    F = "proveedores.cuit = " & Escape(Replace(Me.txtCuit.text, "-", vbNullString))
    Cancel = DAOProveedor.FindAll(F).count > 0
    If Cancel Then MsgBox "Ya existe un proveedor con ese CUIT.", vbExclamation
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

Private Function validarFactura() As Boolean
    validarFactura = (vFactura.Total = funciones.RedondearDecimales(CDbl(Me.txtMontoManual)))
End Function




Private Sub txtMontoManual_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdGuardar_Click
End Sub


Private Sub txtMontoNeto_Change()
    grabado = False
    TotalFactura
End Sub
Private Sub txtMontoNeto_GotFocus()
    foco Me.txtMontoNeto
End Sub
Private Sub txtMontoPercep_Change()
    grabado = False
End Sub

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
Private Sub limpiar()
    If MsgBox("¿Desea limpiar la factura?", vbYesNo, "Confirmación") Then
        Me.txtNumeroMask.text = vbNullString
    End If
End Sub

'Private Sub txtNoGravado_Change()
'    On Error Resume Next
'    vFactura.ConceptoNoGravado = CDbl(Me.txtNoGravado)
'    TotalFactura
'    grabado = False
'End Sub

Private Sub txtNumeroMask_Change()
    On Error Resume Next
    If Me.txtNumeroMask.text <> "____-________" Then
        vFactura.numero = Me.txtNumeroMask.text
        TotalFactura
        grabado = False
    End If


End Sub

Private Sub txtNumeroMask_GotFocus()
    foco Me.txtNumeroMask
End Sub

Private Sub txtRedondeo_Change()
    On Error Resume Next
    vFactura.Redondeo = CDbl(Me.txtRedondeo)
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


If vFactura.configFactura.FormateaNumero Then
    If InStr(1, vFactura.numero, "-") = 0 Then
      Me.txtNumeroMask.text = vFactura.numero
        'Me.txtNumeroMask.text = Mid(vFactura.numero, 1, 4) & "-" & String(8 - Len(Mid(vFactura.numero, 5)), "0") & Mid(vFactura.numero, 5)
    Else
        Me.txtNumeroMask.text = vFactura.numero
    End If
Else
      Me.txtNumeroMask.text = vFactura.numero
End If
    Me.txtRedondeo = vFactura.Redondeo
    'Me.txtNoGravado = vFactura.ConceptoNoGravado
    Me.cboProveedores.ListIndex = funciones.PosIndexCbo(vFactura.Proveedor.id, Me.cboProveedores)
    Me.cboTiposFactura.ListIndex = funciones.PosIndexCbo(vFactura.configFactura.id, Me.cboTiposFactura)
    Me.cboMonedas.ListIndex = funciones.PosIndexCbo(vFactura.moneda.id, Me.cboMonedas)
    Me.txtMontoNeto = vFactura.NetoGravado

    Me.optContado.value = Not vFactura.FormaPagoCuentaCorriente
    Me.optCtaCte.value = vFactura.FormaPagoCuentaCorriente

    Me.grid_cuentascontables.ItemCount = 0
    Me.grid_cuentascontables.ItemCount = vFactura.cuentasContables.count

    mostrarALicuotas

    Me.grilla_percepciones.ItemCount = 0
    Me.grilla_percepciones.ItemCount = vFactura.percepciones.count
    grabado = True
End Sub
Private Sub mostrarALicuotas()
    Me.grilla_alicuotas.ItemCount = 0
    Me.grilla_alicuotas.ItemCount = vFactura.IvaAplicado.count
End Sub

Private Sub armarFactura()
    vFactura.FEcha = (CDate(Format(Me.DTPicker1, "yyyy-mm-dd")))
    vFactura.numero = Me.txtNumeroMask.text

    vFactura.Proveedor = Proveedor
    vFactura.ImpuestoInterno = CDbl(Me.txtImpuestos)
    vFactura.Monto = CDbl(Me.txtMontoNeto)
    vFactura.estado = EstadoFacturaProveedor.EnProceso

    idtipo = Me.cboTiposFactura.ItemData(Me.cboTiposFactura.ListIndex)
    vFactura.tipoDocumentoContable = Me.cboTipoDocContable.ItemData(Me.cboTipoDocContable.ListIndex)
    vFactura.configFactura = DAOConfigFacturaProveedor.GetById(idtipo)
End Sub
Private Sub llenarComboProveedores()
    DAOProveedor.llenarComboXtremeSuite Me.cboProveedores, True, True, False
End Sub

Private Sub txtTipoCambio_Change()
    On Error Resume Next
    vFactura.TipoCambio = Val(Me.txtTipoCambio)
    TotalFactura
    grabado = False
End Sub
