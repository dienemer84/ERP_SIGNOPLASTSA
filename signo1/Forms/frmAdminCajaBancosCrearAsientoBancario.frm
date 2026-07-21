VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmAdminCajaBancosCrearAsientoBancario 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Crear Movimiento de Caja y Bancos"
   ClientHeight    =   10110
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10110
   ScaleWidth      =   10005
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.PushButton cmdCrear 
      Height          =   495
      Left            =   7800
      TabIndex        =   19
      Top             =   9480
      Width           =   2055
      _Version        =   786432
      _ExtentX        =   3625
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Guardar"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1095
      Left            =   120
      TabIndex        =   12
      Top             =   8280
      Width           =   9735
      _Version        =   786432
      _ExtentX        =   17171
      _ExtentY        =   1931
      _StockProps     =   79
      Caption         =   "Observaciones"
      Appearance      =   4
      Begin XtremeSuiteControls.FlatEdit FlatEdit1 
         Height          =   735
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   9255
         _Version        =   786432
         _ExtentX        =   16325
         _ExtentY        =   1296
         _StockProps     =   77
         BackColor       =   -2147483643
      End
   End
   Begin XtremeSuiteControls.GroupBox grpTipo 
      Height          =   735
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   9735
      _Version        =   786432
      _ExtentX        =   17171
      _ExtentY        =   1296
      _StockProps     =   79
      Caption         =   "Tipo de Movimiento"
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
      Begin XtremeSuiteControls.RadioButton RadioButton3 
         Height          =   375
         Left            =   6720
         TabIndex        =   24
         Top             =   240
         Width           =   2655
         _Version        =   786432
         _ExtentX        =   4683
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   " TRANSFERENCIAS"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton RadioButton2 
         Height          =   375
         Left            =   3720
         TabIndex        =   10
         Top             =   240
         Width           =   1575
         _Version        =   786432
         _ExtentX        =   2778
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "  EGRESO"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton RadioButton1 
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   1575
         _Version        =   786432
         _ExtentX        =   2778
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "  INGRESO"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox 
      Height          =   7215
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   9735
      _Version        =   786432
      _ExtentX        =   17171
      _ExtentY        =   12726
      _StockProps     =   79
      Caption         =   "Detalles"
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
      Begin XtremeSuiteControls.PushButton btnClearCtaBancariaDestino 
         Height          =   375
         Left            =   7680
         TabIndex        =   26
         Top             =   1895
         Width           =   375
         _Version        =   786432
         _ExtentX        =   661
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "X"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cboCuentaBancariaDestino 
         Height          =   315
         Left            =   2160
         TabIndex        =   25
         Top             =   1920
         Width           =   5415
         _Version        =   786432
         _ExtentX        =   9551
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.PushButton btnClearCtaBancaria 
         Height          =   375
         Left            =   7680
         TabIndex        =   20
         Top             =   1440
         Width           =   375
         _Version        =   786432
         _ExtentX        =   661
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "X"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cboCuentasBancarias 
         Height          =   315
         Left            =   2160
         TabIndex        =   18
         Top             =   1470
         Width           =   5415
         _Version        =   786432
         _ExtentX        =   9551
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.PushButton btnClearCtaContable 
         Height          =   375
         Left            =   7680
         TabIndex        =   23
         Top             =   2760
         Width           =   375
         _Version        =   786432
         _ExtentX        =   661
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "X"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cboCuentasContables 
         Height          =   315
         Left            =   2160
         TabIndex        =   21
         Top             =   2790
         Width           =   5415
         _Version        =   786432
         _ExtentX        =   9551
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Text            =   "cboCuentas"
      End
      Begin XtremeSuiteControls.DateTimePicker dtpFecha 
         Height          =   330
         Left            =   2160
         TabIndex        =   15
         Top             =   900
         Width           =   1245
         _Version        =   786432
         _ExtentX        =   2196
         _ExtentY        =   582
         _StockProps     =   68
         Format          =   1
         CurrentDate     =   40183.7263657407
      End
      Begin XtremeSuiteControls.ComboBox cboMonedas 
         Height          =   315
         Left            =   2160
         TabIndex        =   17
         Top             =   360
         Width           =   1245
         _Version        =   786432
         _ExtentX        =   2196
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Style           =   2
         Text            =   "cboMonedas"
      End
      Begin XtremeSuiteControls.GroupBox grpOrigen 
         Height          =   3855
         Left            =   120
         TabIndex        =   22
         Top             =   3240
         Width           =   9495
         _Version        =   786432
         _ExtentX        =   16748
         _ExtentY        =   6800
         _StockProps     =   79
         Caption         =   "Valores"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   3
         Begin XtremeSuiteControls.GroupBox GroupBox3 
            Height          =   2055
            Left            =   120
            TabIndex        =   33
            Top             =   1680
            Width           =   9255
            _Version        =   786432
            _ExtentX        =   16325
            _ExtentY        =   3625
            _StockProps     =   79
            Caption         =   "de Cheques Disponibles"
            UseVisualStyle  =   -1  'True
            Begin GridEX20.GridEX gridChequesPropios 
               Height          =   1575
               Left            =   120
               TabIndex        =   34
               Top             =   360
               Width           =   8970
               _ExtentX        =   15822
               _ExtentY        =   2778
               Version         =   "2.0"
               BoundColumnIndex=   ""
               ReplaceColumnIndex=   ""
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
               ColumnsCount    =   5
               Column(1)       =   "frmAdminCajaBancosCrearAsientoBancario.frx":0000
               Column(2)       =   "frmAdminCajaBancosCrearAsientoBancario.frx":0168
               Column(3)       =   "frmAdminCajaBancosCrearAsientoBancario.frx":029C
               Column(4)       =   "frmAdminCajaBancosCrearAsientoBancario.frx":03D8
               Column(5)       =   "frmAdminCajaBancosCrearAsientoBancario.frx":0540
               FormatStylesCount=   6
               FormatStyle(1)  =   "frmAdminCajaBancosCrearAsientoBancario.frx":0638
               FormatStyle(2)  =   "frmAdminCajaBancosCrearAsientoBancario.frx":0770
               FormatStyle(3)  =   "frmAdminCajaBancosCrearAsientoBancario.frx":0820
               FormatStyle(4)  =   "frmAdminCajaBancosCrearAsientoBancario.frx":08D4
               FormatStyle(5)  =   "frmAdminCajaBancosCrearAsientoBancario.frx":09AC
               FormatStyle(6)  =   "frmAdminCajaBancosCrearAsientoBancario.frx":0A64
               ImageCount      =   0
               PrinterProperties=   "frmAdminCajaBancosCrearAsientoBancario.frx":0B44
            End
         End
         Begin XtremeSuiteControls.GroupBox GroupBox2 
            Height          =   1335
            Left            =   120
            TabIndex        =   28
            Top             =   240
            Width           =   9255
            _Version        =   786432
            _ExtentX        =   16325
            _ExtentY        =   2355
            _StockProps     =   79
            Caption         =   "de Bancos y Caja"
            UseVisualStyle  =   -1  'True
            Begin VB.TextBox txtMonto 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   1920
               TabIndex        =   30
               Top             =   360
               Width           =   1815
            End
            Begin VB.TextBox txtComprobante 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   1920
               TabIndex        =   29
               Top             =   840
               Width           =   3135
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               Caption         =   "Monto"
               Height          =   255
               Left            =   120
               TabIndex        =   32
               Top             =   390
               Width           =   1695
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               Caption         =   "Comprobante"
               Height          =   255
               Left            =   0
               TabIndex        =   31
               Top             =   855
               Width           =   1815
            End
         End
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   375
         Index           =   1
         Left            =   720
         TabIndex        =   27
         Top             =   1860
         Width           =   1335
         _Version        =   786432
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Cuenta destino"
         Alignment       =   1
      End
      Begin VB.Line Line1 
         X1              =   480
         X2              =   9120
         Y1              =   2520
         Y2              =   2520
      End
      Begin XtremeSuiteControls.Label Label 
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   16
         Top             =   390
         Width           =   975
         _Version        =   786432
         _ExtentX        =   1720
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Moneda"
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label Label 
         Height          =   375
         Index           =   2
         Left            =   1200
         TabIndex        =   14
         Top             =   885
         Width           =   855
         _Version        =   786432
         _ExtentX        =   1508
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Fecha"
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   375
         Index           =   0
         Left            =   720
         TabIndex        =   11
         Top             =   1420
         Width           =   1335
         _Version        =   786432
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Cuenta destino"
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label Label 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   2760
         Width           =   1935
         _Version        =   786432
         _ExtentX        =   3413
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Cuenta contable destino"
         Alignment       =   1
      End
   End
   Begin GridEX20.GridEX gridBancos 
      Height          =   1845
      Left            =   10080
      TabIndex        =   2
      Top             =   5880
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
      Column(1)       =   "frmAdminCajaBancosCrearAsientoBancario.frx":0D1C
      Column(2)       =   "frmAdminCajaBancosCrearAsientoBancario.frx":0E1C
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminCajaBancosCrearAsientoBancario.frx":0F0C
      FormatStyle(2)  =   "frmAdminCajaBancosCrearAsientoBancario.frx":1044
      FormatStyle(3)  =   "frmAdminCajaBancosCrearAsientoBancario.frx":10F4
      FormatStyle(4)  =   "frmAdminCajaBancosCrearAsientoBancario.frx":11A8
      FormatStyle(5)  =   "frmAdminCajaBancosCrearAsientoBancario.frx":1280
      FormatStyle(6)  =   "frmAdminCajaBancosCrearAsientoBancario.frx":1338
      ImageCount      =   0
      PrinterProperties=   "frmAdminCajaBancosCrearAsientoBancario.frx":1418
   End
   Begin GridEX20.GridEX gridCuentasBancarias 
      Height          =   1695
      Left            =   10080
      TabIndex        =   3
      Top             =   4080
      Width           =   4185
      _ExtentX        =   7382
      _ExtentY        =   2990
      Version         =   "2.0"
      BoundColumnIndex=   "id"
      ReplaceColumnIndex=   "cuenta"
      ActAsDropDown   =   -1  'True
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
      Column(1)       =   "frmAdminCajaBancosCrearAsientoBancario.frx":15F0
      Column(2)       =   "frmAdminCajaBancosCrearAsientoBancario.frx":1714
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminCajaBancosCrearAsientoBancario.frx":1808
      FormatStyle(2)  =   "frmAdminCajaBancosCrearAsientoBancario.frx":1940
      FormatStyle(3)  =   "frmAdminCajaBancosCrearAsientoBancario.frx":19F0
      FormatStyle(4)  =   "frmAdminCajaBancosCrearAsientoBancario.frx":1AA4
      FormatStyle(5)  =   "frmAdminCajaBancosCrearAsientoBancario.frx":1B7C
      FormatStyle(6)  =   "frmAdminCajaBancosCrearAsientoBancario.frx":1C34
      ImageCount      =   0
      PrinterProperties=   "frmAdminCajaBancosCrearAsientoBancario.frx":1D14
   End
   Begin GridEX20.GridEX gridMonedas 
      Height          =   1815
      Left            =   10080
      TabIndex        =   4
      Top             =   7800
      Width           =   4260
      _ExtentX        =   7514
      _ExtentY        =   3201
      Version         =   "2.0"
      BoundColumnIndex=   "id"
      ReplaceColumnIndex=   "moneda"
      ActAsDropDown   =   -1  'True
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
      Column(1)       =   "frmAdminCajaBancosCrearAsientoBancario.frx":1EEC
      Column(2)       =   "frmAdminCajaBancosCrearAsientoBancario.frx":2010
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminCajaBancosCrearAsientoBancario.frx":2104
      FormatStyle(2)  =   "frmAdminCajaBancosCrearAsientoBancario.frx":223C
      FormatStyle(3)  =   "frmAdminCajaBancosCrearAsientoBancario.frx":22EC
      FormatStyle(4)  =   "frmAdminCajaBancosCrearAsientoBancario.frx":23A0
      FormatStyle(5)  =   "frmAdminCajaBancosCrearAsientoBancario.frx":2478
      FormatStyle(6)  =   "frmAdminCajaBancosCrearAsientoBancario.frx":2530
      ImageCount      =   0
      PrinterProperties=   "frmAdminCajaBancosCrearAsientoBancario.frx":2610
   End
   Begin GridEX20.GridEX gridChequesDisponibles 
      Height          =   1905
      Left            =   10080
      TabIndex        =   5
      Top             =   2040
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   3360
      Version         =   "2.0"
      BoundColumnIndex=   "id"
      ReplaceColumnIndex=   "numero"
      ActAsDropDown   =   -1  'True
      ColumnAutoResize=   -1  'True
      HideSelection   =   2
      MethodHoldFields=   -1  'True
      ContScroll      =   -1  'True
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      RowHeaders      =   -1  'True
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   8
      Column(1)       =   "frmAdminCajaBancosCrearAsientoBancario.frx":27E8
      Column(2)       =   "frmAdminCajaBancosCrearAsientoBancario.frx":2968
      Column(3)       =   "frmAdminCajaBancosCrearAsientoBancario.frx":2B08
      Column(4)       =   "frmAdminCajaBancosCrearAsientoBancario.frx":2C00
      Column(5)       =   "frmAdminCajaBancosCrearAsientoBancario.frx":2D3C
      Column(6)       =   "frmAdminCajaBancosCrearAsientoBancario.frx":2E48
      Column(7)       =   "frmAdminCajaBancosCrearAsientoBancario.frx":2F68
      Column(8)       =   "frmAdminCajaBancosCrearAsientoBancario.frx":3074
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminCajaBancosCrearAsientoBancario.frx":3168
      FormatStyle(2)  =   "frmAdminCajaBancosCrearAsientoBancario.frx":32A0
      FormatStyle(3)  =   "frmAdminCajaBancosCrearAsientoBancario.frx":3350
      FormatStyle(4)  =   "frmAdminCajaBancosCrearAsientoBancario.frx":3404
      FormatStyle(5)  =   "frmAdminCajaBancosCrearAsientoBancario.frx":34DC
      FormatStyle(6)  =   "frmAdminCajaBancosCrearAsientoBancario.frx":3594
      ImageCount      =   0
      PrinterProperties=   "frmAdminCajaBancosCrearAsientoBancario.frx":3674
   End
   Begin GridEX20.GridEX gridChequeras 
      Height          =   1815
      Left            =   360
      TabIndex        =   6
      Top             =   10200
      Width           =   4035
      _ExtentX        =   7117
      _ExtentY        =   3201
      Version         =   "2.0"
      BoundColumnIndex=   "id"
      ReplaceColumnIndex=   "chequera"
      ActAsDropDown   =   -1  'True
      ColumnAutoResize=   -1  'True
      HideSelection   =   2
      MethodHoldFields=   -1  'True
      ContScroll      =   -1  'True
      AllowColumnDrag =   0   'False
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      ColumnHeaders   =   0   'False
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   2
      Column(1)       =   "frmAdminCajaBancosCrearAsientoBancario.frx":384C
      Column(2)       =   "frmAdminCajaBancosCrearAsientoBancario.frx":396C
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminCajaBancosCrearAsientoBancario.frx":3A6C
      FormatStyle(2)  =   "frmAdminCajaBancosCrearAsientoBancario.frx":3BA4
      FormatStyle(3)  =   "frmAdminCajaBancosCrearAsientoBancario.frx":3C54
      FormatStyle(4)  =   "frmAdminCajaBancosCrearAsientoBancario.frx":3D08
      FormatStyle(5)  =   "frmAdminCajaBancosCrearAsientoBancario.frx":3DE0
      FormatStyle(6)  =   "frmAdminCajaBancosCrearAsientoBancario.frx":3E98
      ImageCount      =   0
      PrinterProperties=   "frmAdminCajaBancosCrearAsientoBancario.frx":3F78
   End
   Begin GridEX20.GridEX gridChequesChequera 
      Height          =   1710
      Left            =   10080
      TabIndex        =   7
      Top             =   240
      Width           =   1980
      _ExtentX        =   3493
      _ExtentY        =   3016
      Version         =   "2.0"
      HoldSortSettings=   -1  'True
      BoundColumnIndex=   "id"
      ReplaceColumnIndex=   "nro"
      ActAsDropDown   =   -1  'True
      ColumnAutoResize=   -1  'True
      HideSelection   =   2
      MethodHoldFields=   -1  'True
      ContScroll      =   -1  'True
      AllowColumnDrag =   0   'False
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      ColumnHeaders   =   0   'False
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   2
      Column(1)       =   "frmAdminCajaBancosCrearAsientoBancario.frx":4150
      Column(2)       =   "frmAdminCajaBancosCrearAsientoBancario.frx":4280
      SortKeysCount   =   1
      SortKey(1)      =   "frmAdminCajaBancosCrearAsientoBancario.frx":4380
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminCajaBancosCrearAsientoBancario.frx":43E8
      FormatStyle(2)  =   "frmAdminCajaBancosCrearAsientoBancario.frx":4520
      FormatStyle(3)  =   "frmAdminCajaBancosCrearAsientoBancario.frx":45D0
      FormatStyle(4)  =   "frmAdminCajaBancosCrearAsientoBancario.frx":4684
      FormatStyle(5)  =   "frmAdminCajaBancosCrearAsientoBancario.frx":475C
      FormatStyle(6)  =   "frmAdminCajaBancosCrearAsientoBancario.frx":4814
      ImageCount      =   0
      PrinterProperties=   "frmAdminCajaBancosCrearAsientoBancario.frx":48F4
   End
End
Attribute VB_Name = "frmAdminCajaBancosCrearAsientoBancario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim colCuentas As New Collection
Dim CuentaContable As clsCuentaContable
Dim formLoading As Boolean
Dim formLoaded As Boolean

Private operacion As operacion
Private AsientoContable As New clsAsientoContable
Private Banco As Banco
Private caja As caja
Private CuentaBancaria As CuentaBancaria
Private moneda As clsMoneda
Private alicuotaRetencion As DTORetencionAlicuota
Private CuentasBancarias As New Collection
Private retenciones As New Collection
Private Monedas As New Collection
Private Cajas As New Collection
Private bancos As New Collection
Private chequesDisponibles As New Collection
Private chequeras As New Collection
Private cheque As cheque
Private tmpChequera As chequera
Private chequesChequeraSeleccionada As New Collection
Public ReadOnly As Boolean


Private Sub btnClearCtaContable_Click()
        Me.cboCuentasContables.ListIndex = -1
End Sub

Private Sub btnClearCtaBancaria_Click()
        Me.cboCuentasBancarias.ListIndex = -1
End Sub

Private Sub cboMonedas_Click()
    If Me.cboMonedas.ListIndex = -1 Then
        Set AsientoContable.moneda = Nothing
    Else
        Set AsientoContable.moneda = DAOMoneda.GetById(Me.cboMonedas.ItemData(Me.cboMonedas.ListIndex))
    End If
'    Totalizar
End Sub


Private Sub cboCuentas_Click()
    If IsSomething(AsientoContable) And Me.cboCuentasContables.ListIndex <> -1 Then
        Set AsientoContable.CuentaContable = DAOCuentaContable.GetById(Me.cboCuentasContables.ItemData(Me.cboCuentasContables.ListIndex))
    End If

End Sub


Private Sub cmdCrear_Click()
    On Error GoTo err1

    Dim esIngreso As Boolean
    Dim esEgreso As Boolean
    Dim esTransferencia As Boolean

    Dim cuentaOrigen As CuentaBancaria
    Dim cuentaDestino As CuentaBancaria

    Dim monto As Double
    Dim comprobante As String
    Dim entradaSalida As Integer

    Dim fechaMantener As Date
    Dim idCuentaMantener As Long

    esIngreso = Me.RadioButton1.value
    esEgreso = Me.RadioButton2.value
    esTransferencia = Me.RadioButton3.value

    fechaMantener = Me.dtpFecha.value

    If Me.cboCuentasBancarias.ListIndex <> -1 Then
        idCuentaMantener = Me.cboCuentasBancarias.ItemData(Me.cboCuentasBancarias.ListIndex)
    Else
        idCuentaMantener = 0
    End If

    '-----------------------------
    ' Validaciones generales
    '-----------------------------
    If Not esIngreso And Not esEgreso And Not esTransferencia Then
        MsgBox "Debe seleccionar el tipo de movimiento.", vbExclamation
        Exit Sub
    End If

    If Me.cboMonedas.ListIndex = -1 Then
        MsgBox "Debe seleccionar una moneda.", vbExclamation
        Exit Sub
    End If

    Set AsientoContable.moneda = _
        DAOMoneda.GetById(Me.cboMonedas.ItemData(Me.cboMonedas.ListIndex))

    AsientoContable.FEcha = Me.dtpFecha.value
    AsientoContable.Observaciones = Trim$(Me.FlatEdit1.Text)
    AsientoContable.idUsuario = funciones.GetUserObj.Id

    comprobante = Trim$(Me.txtComprobante.Text)

    If LenB(Trim$(Me.txtMonto.Text)) > 0 Then
        If Not IsNumeric(Me.txtMonto.Text) Then
            MsgBox "Debe ingresar un monto válido.", vbExclamation
            Exit Sub
        End If

        monto = CDbl(Me.txtMonto.Text)
    Else
        monto = 0
    End If

    ' Limpio operaciones para reconstruirlas según el tipo de movimiento
    LimpiarOperacionesMovimiento

    '-----------------------------
    ' INGRESO
    '-----------------------------
    If esIngreso Then

        AsientoContable.TipoMovimiento = "INGRESO"
        Set AsientoContable.CuentaContable = Nothing

        If Me.cboCuentasBancarias.ListIndex = -1 Then
            MsgBox "Debe seleccionar una cuenta destino.", vbExclamation
            Exit Sub
        End If

        Set cuentaDestino = ObtenerCuentaSeleccionada(Me.cboCuentasBancarias)
        Set AsientoContable.CuentaBancaria = cuentaDestino

        If Not CuentaCoincideConMoneda(cuentaDestino, AsientoContable.moneda) Then
            MsgBox "La moneda de la cuenta destino no coincide con la moneda seleccionada.", vbExclamation
            Exit Sub
        End If

        If monto > 0 Then
            If LenB(comprobante) = 0 Then
                MsgBox "Debe ingresar un comprobante.", vbExclamation
                Exit Sub
            End If

            AsientoContable.operacionesBanco.Add _
                CrearOperacionBancaria(cuentaDestino, monto, comprobante, OPEntrada)
        End If

        If monto <= 0 And AsientoContable.ChequesPropios.count = 0 Then
            MsgBox "Debe cargar un monto o al menos un cheque propio.", vbExclamation
            Exit Sub
        End If

    '-----------------------------
    ' EGRESO
    '-----------------------------
    ElseIf esEgreso Then

        AsientoContable.TipoMovimiento = "EGRESO"

        If Me.cboCuentasBancarias.ListIndex = -1 Then
            MsgBox "Debe seleccionar una cuenta origen.", vbExclamation
            Exit Sub
        End If

        Set cuentaOrigen = ObtenerCuentaSeleccionada(Me.cboCuentasBancarias)
        Set AsientoContable.CuentaBancaria = cuentaOrigen

        If Not CuentaCoincideConMoneda(cuentaOrigen, AsientoContable.moneda) Then
            MsgBox "La moneda de la cuenta origen no coincide con la moneda seleccionada.", vbExclamation
            Exit Sub
        End If

        If Me.cboCuentasContables.ListIndex = -1 Then
            MsgBox "Debe seleccionar una cuenta contable destino para el egreso.", vbExclamation
            Exit Sub
        End If

        Set AsientoContable.CuentaContable = _
            DAOCuentaContable.GetById(Me.cboCuentasContables.ItemData(Me.cboCuentasContables.ListIndex))

        If monto > 0 Then
            If LenB(comprobante) = 0 Then
                MsgBox "Debe ingresar un comprobante.", vbExclamation
                Exit Sub
            End If

            AsientoContable.operacionesBanco.Add _
                CrearOperacionBancaria(cuentaOrigen, monto, comprobante, OPSalida)
        End If

        If monto <= 0 And AsientoContable.ChequesPropios.count = 0 Then
            MsgBox "Debe cargar un monto o al menos un cheque propio.", vbExclamation
            Exit Sub
        End If

    '-----------------------------
    ' TRANSFERENCIA
    '-----------------------------
    ElseIf esTransferencia Then

        AsientoContable.TipoMovimiento = "TRANSFERENCIA"
        Set AsientoContable.CuentaContable = Nothing

        If AsientoContable.ChequesPropios.count > 0 Then
            MsgBox "Las transferencias bancarias no deben tener cheques propios.", vbExclamation
            Exit Sub
        End If

        If Me.cboCuentasBancarias.ListIndex = -1 Then
            MsgBox "Debe seleccionar una cuenta origen.", vbExclamation
            Exit Sub
        End If

        If Me.cboCuentaBancariaDestino.ListIndex = -1 Then
            MsgBox "Debe seleccionar una cuenta destino.", vbExclamation
            Exit Sub
        End If

        Set cuentaOrigen = ObtenerCuentaSeleccionada(Me.cboCuentasBancarias)
        Set cuentaDestino = ObtenerCuentaSeleccionada(Me.cboCuentaBancariaDestino)

        If cuentaOrigen.Id = cuentaDestino.Id Then
            MsgBox "La cuenta origen y la cuenta destino no pueden ser la misma.", vbExclamation
            Exit Sub
        End If

        If Not CuentaCoincideConMoneda(cuentaOrigen, AsientoContable.moneda) Then
            MsgBox "La moneda de la cuenta origen no coincide con la moneda seleccionada.", vbExclamation
            Exit Sub
        End If

        If Not CuentaCoincideConMoneda(cuentaDestino, AsientoContable.moneda) Then
            MsgBox "La moneda de la cuenta destino no coincide con la moneda seleccionada.", vbExclamation
            Exit Sub
        End If

        If monto <= 0 Then
            MsgBox "Debe ingresar un monto mayor a cero para la transferencia.", vbExclamation
            Exit Sub
        End If

        If LenB(comprobante) = 0 Then
            MsgBox "Debe ingresar un comprobante.", vbExclamation
            Exit Sub
        End If

        Set AsientoContable.CuentaBancaria = cuentaOrigen

        ' Salida de la cuenta origen
        AsientoContable.operacionesBanco.Add _
            CrearOperacionBancaria(cuentaOrigen, monto, comprobante, OPSalida)

        ' Entrada en la cuenta destino
        AsientoContable.operacionesBanco.Add _
            CrearOperacionBancaria(cuentaDestino, monto, comprobante, OPEntrada)

    End If

    '-----------------------------
    ' Totales y guardado
    '-----------------------------
    If AsientoContable.TipoMovimiento = "TRANSFERENCIA" Then
    AsientoContable.StaticTotalOrigenes = monto
    Else
        AsientoContable.StaticTotalOrigenes = AsientoContable.TotalOrigenes
    End If

    If AsientoContable.IsValid Then

        Dim n As Boolean
        n = (AsientoContable.Id = 0)

        If DAOAsientoContable.Save(AsientoContable, True) Then

            If n Then
                MsgBox "Movimiento Nro " & AsientoContable.Id & " creado con éxito.", vbInformation
            Else
                MsgBox "Movimiento modificado con éxito.", vbInformation
            End If

            If n Then
                If MsgBox("Desea registrar un nuevo movimiento?", vbQuestion + vbYesNo) = vbYes Then
                    Dim f12 As New frmAdminCajaBancosCrearAsientoBancario
                    f12.Show
                    f12.CargarValoresIniciales fechaMantener, idCuentaMantener
                End If
            End If

            Unload Me

        Else
            MsgBox "Hubo un problema al guardar el movimiento.", vbCritical
        End If

    Else
        MsgBox AsientoContable.ValidationMessages, vbCritical, "Error"
    End If

    Exit Sub

err1:
    MsgBox "Error al guardar el movimiento: " & Err.Description, vbCritical

End Sub


Private Sub Form_Load()

    FormHelper.Customize Me

    formLoading = True

    Me.Left = _
    frmPrincipal.ScaleWidth / 6
    Me.Top = frmPrincipal.ScaleHeight / 22
    
    Me.gridChequeras.Visible = False
    Me.gridChequesChequera.Visible = False
    
    GridEXHelper.CustomizeGrid Me.gridChequesDisponibles, False, False

    llenarComboCuentas

    DAOCuentaBancaria.llenarComboXtremeSuite _
        Me.cboCuentasBancarias

    DAOCuentaBancaria.llenarComboXtremeSuite _
        Me.cboCuentaBancariaDestino

    Me.cboCuentasBancarias.ListIndex = -1
    Me.cboCuentaBancariaDestino.ListIndex = -1

    DAOMoneda.llenarComboXtremeSuite Me.cboMonedas
    Me.cboMonedas.ListIndex = -1
    
    Set chequeras = DAOChequeras.FindAllWithChequesDisponibles()
    
    Me.gridChequeras.ItemCount = chequeras.count
    
    Set Me.gridChequesPropios.Columns("chequera").DropDownControl = Me.gridChequeras

    Set Me.gridChequesPropios.Columns("numero").DropDownControl = Me.gridChequesChequera

    gridChequesChequera.ItemCount = 0
    
    GridEXHelper.AutoSizeColumns Me.gridChequeras

    Me.dtpFecha.value = Date

    Me.RadioButton1.value = False
    Me.RadioButton2.value = False
    Me.RadioButton3.value = False

    ActualizarModoMovimiento

    formLoaded = True
    formLoading = False

End Sub

Private Sub llenarComboCuentas()
  
    DAOCuentaContable.llenarComboXtremeSuiteConCodigo Me.cboCuentasContables, False, False, False
    
    Me.cboCuentasContables.ListIndex = -1
    
End Sub


Private Sub gridBancos_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex <= bancos.count Then
        Set Banco = bancos.item(RowIndex)
        Values(1) = Banco.Id
        Values(2) = Banco.nombre
    End If
End Sub


'''Private Sub gridChequeras_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
'''    If RowIndex <= chequeras.count Then
'''        Set tmpChequera = chequeras.item(RowIndex)
'''        Values(1) = tmpChequera.Description
'''        Values(2) = tmpChequera.Id
'''    End If
'''End Sub


'''Private Sub gridCheques_BeforeUpdate(ByVal Cancel As GridEX20.JSRetBoolean)
'''    Dim msg As New Collection
'''
'''    ' REVISA QUE EN LA COLECCION DE CHEQUES DE TERCEROS QUE SE ESTAN CARGANDO NO EST? INGRESADO EL MISMO CHEQUE, SI LO DETECTA GENERA MSG DE ERROR
'''    If funciones.BuscarEnColeccion(AsientoContable.ChequesTerceros, CStr(Me.gridCheques.value(1))) Then
'''        msg.Add "El cheque seleccionado ya fue ingresado anteriormente."
'''    End If
'''
'''    Cancel = (msg.count > 0)
'''    If Cancel Then MsgBox funciones.JoinCollectionValues(msg, vbNewLine), vbExclamation
'''
'''End Sub


Private Sub gridCheques_UnboundAddNew(ByVal NewRowBookmark As GridEX20.JSRetVariant, ByVal Values As GridEX20.JSRowData)
    Set cheque = Nothing
    If IsNumeric(Values(1)) Then Set cheque = DAOCheques.FindById(Values(1))
    If IsSomething(cheque) Then
        AsientoContable.ChequesTerceros.Add cheque, CStr(cheque.Id)

    End If


End Sub

Private Sub gridCheques_UnboundDelete(ByVal RowIndex As Long, ByVal Bookmark As Variant)
    If RowIndex > 0 Then
        AsientoContable.ChequesTerceros.remove RowIndex

    End If
End Sub

Private Sub gridCheques_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex <= AsientoContable.ChequesTerceros.count Then
        Set cheque = AsientoContable.ChequesTerceros.item(RowIndex)

        Values(1) = cheque.numero & " "

        'FORMATCURRENCY
        Values(2) = FormatCurrency(cheque.monto)
        Values(3) = cheque.FechaVencimiento
        If IsSomething(cheque.moneda) Then Values(4) = cheque.moneda.NombreCorto
        If IsSomething(cheque.Banco) Then Values(5) = cheque.Banco.nombre
        Values(6) = cheque.OrigenDestino
        Values(7) = cheque.OrigenCheque
    
    End If
End Sub

Private Sub gridCheques_UnboundUpdate(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex > 0 And AsientoContable.ChequesTerceros.count >= RowIndex Then
        Set cheque = Nothing
        If IsNumeric(Values(1)) Then Set cheque = DAOCheques.FindById(Values(1))
        If IsSomething(cheque) Then
            AsientoContable.ChequesTerceros.Add cheque, , , RowIndex
            AsientoContable.ChequesTerceros.remove RowIndex
        End If
'        Totalizar
    End If
End Sub

Private Sub gridChequesChequera_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex > 0 And chequesChequeraSeleccionada.count > 0 Then
        Values(1) = chequesChequeraSeleccionada(RowIndex).numero
        Values(2) = chequesChequeraSeleccionada(RowIndex).Id
    End If
End Sub


Private Sub gridChequesDisponibles_ColumnHeaderClick(ByVal Column As GridEX20.JSColumn)
    GridEXHelper.ColumnHeaderClick Me.gridChequesDisponibles, Column
End Sub


Private Sub gridChequesDisponibles_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex <= chequesDisponibles.count Then
        Set cheque = chequesDisponibles.item(RowIndex)
        Values(1) = cheque.numero
        'FORMATCURRENCY
        Values(2) = FormatCurrency(cheque.monto)
        Values(3) = cheque.FechaVencimiento
        If IsSomething(cheque.moneda) Then Values(4) = cheque.moneda.NombreCorto
        If IsSomething(cheque.Banco) Then Values(5) = cheque.Banco.nombre
        Values(6) = cheque.Id
        Values(7) = cheque.OrigenCheque
        Values(8) = cheque.OrigenDestino

    End If

End Sub

Private Sub gridChequesPropios_BeforeUpdate(ByVal Cancel As GridEX20.JSRetBoolean)
    Dim msg As New Collection

    If LenB(Me.gridChequesPropios.value(1)) = 0 Then
        msg.Add "Debe especificar una chequera."
    End If

    If LenB(Me.gridChequesPropios.value(2)) = 0 Then
        msg.Add "Debe especificar un cheque."
    End If

    ' REVISA QUE EN LA COLECCION DE CHEQUES PROPIOS QUE SE ESTAN CARGANDO NO EST? INGRESADO EL MISMO CHEQUE, SI LO DETECTA GENERA MSG DE ERROR
    If funciones.BuscarEnColeccion(AsientoContable.ChequesPropios, CStr(Me.gridChequesPropios.value(2))) Then
        msg.Add "El cheque seleccionado ya fue ingresado anteriormente."
    End If

    If Not IsNumeric(Me.gridChequesPropios.value(3)) Then
        msg.Add "Debe especificar un monto válido."
    End If
    ' REVISA QUE SE HAYA CARGADO UN MONTO DEL CHEQUE INGRESADO, SI NO SE CARGA GENERA MSG DE ERROR

    If LenB(Me.gridChequesPropios.value(3)) = 0 Then
        msg.Add "Debe especificar un monto mayor a 0."
    End If

    If Not IsDate(Me.gridChequesPropios.value(4)) Then
        msg.Add "Debe especificar una fecha valida."
    End If

    Cancel = (msg.count > 0)
    If Cancel Then MsgBox funciones.JoinCollectionValues(msg, vbNewLine), vbExclamation

End Sub


Private Sub gridChequesPropios_ListSelected(ByVal ColIndex As Integer, ByVal ValueListIndex As Long, ByVal value As Variant)
    If ColIndex = 1 Then
        'If Not IsNumeric(Me.gridChequesPropios.Value(1)) Or LenB(Me.gridChequesPropios.Value(1)) = 0 Then
        If Not IsNumeric(value) Or LenB(value) = 0 Then
            Set chequesChequeraSeleccionada = New Collection
        Else
            Set chequesChequeraSeleccionada = DAOCheques.FindAllDisponiblesByChequera(val(value))  ' Me.gridChequesPropios.Value(1))
        End If

        Me.gridChequesChequera.ItemCount = chequesChequeraSeleccionada.count
    End If
End Sub


Private Sub gridChequesPropios_UnboundAddNew(ByVal NewRowBookmark As GridEX20.JSRetVariant, ByVal Values As GridEX20.JSRowData)

    If Me.RadioButton3.value = True Then
        MsgBox "No se pueden cargar cheques propios en una transferencia bancaria.", vbExclamation
        Exit Sub
    End If

    Set cheque = Nothing

    If IsNumeric(Values(2)) Then
        Set cheque = DAOCheques.FindById(Values(2))
    End If

    If IsSomething(cheque) Then
        cheque.monto = CDbl(Values(3))
        cheque.FechaVencimiento = Values(4)
        cheque.Propio = True
        cheque.EnCartera = False

        If Not funciones.BuscarEnColeccion(AsientoContable.ChequesPropios, CStr(cheque.Id)) Then
            AsientoContable.ChequesPropios.Add cheque, CStr(cheque.Id)
        End If
    End If

End Sub


Private Sub gridChequesPropios_UnboundDelete(ByVal RowIndex As Long, ByVal Bookmark As Variant)
    If RowIndex > 0 And AsientoContable.ChequesPropios.count >= RowIndex Then
        AsientoContable.ChequesPropios.remove RowIndex
    End If
End Sub


Private Sub gridChequesPropios_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If AsientoContable.ChequesPropios.count >= RowIndex Then
        Set cheque = AsientoContable.ChequesPropios.item(RowIndex)
        Values(1) = cheque.chequera.Description
        Values(2) = vbNullString
        'FORMATCURRENCY
        Values(3) = FormatCurrency(cheque.monto)
        Values(4) = cheque.FechaVencimiento
        Values(5) = cheque.numero
    End If
End Sub


Private Sub gridChequesPropios_UnboundUpdate(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If AsientoContable.ChequesPropios.count >= RowIndex Then
        Set cheque = AsientoContable.ChequesPropios.item(RowIndex)
        cheque.monto = Values(3)
        cheque.FechaVencimiento = Values(4)
    End If
    
End Sub


Private Sub gridCuentasBancarias_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If CuentasBancarias.count >= RowIndex Then
        Set CuentaBancaria = CuentasBancarias.item(RowIndex)
        Values(1) = CuentaBancaria.Id
        Values(2) = CuentaBancaria.DescripcionFormateada
    End If
End Sub


Private Sub gridMonedas_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex > 0 And Monedas.count > 0 Then
        Set moneda = Monedas.item(RowIndex)
        Values(1) = moneda.Id
        Values(2) = moneda.NombreCorto
    End If
End Sub


Public Sub Cargar(aContable As clsAsientoContable)

    Set AsientoContable = DAOAsientoContable.FindById(aContable.Id)
    
    Me.caption = "Movimiento de Caja y Bancos Nro " & AsientoContable.Id
    
    With AsientoContable
    
        '------------------------------
        ' Tipo de movimiento
        '------------------------------
        If .TipoMovimiento = "INGRESO" Then
            Me.RadioButton1.value = True
            Me.RadioButton2.value = False
            
            Me.Label1(0).caption = "Cuenta destino"
            
            Me.cboCuentasContables.Enabled = False
            Me.Label(0).Enabled = False
            Me.btnClearCtaContable.Enabled = False
            
        ElseIf .TipoMovimiento = "SALIDA" Or .TipoMovimiento = "EGRESO" Then
            Me.RadioButton1.value = False
            Me.RadioButton2.value = True
            
            Me.Label1(0).caption = "Cuenta origen"
            
            Me.cboCuentasContables.Enabled = True
            Me.Label(0).Enabled = True
            Me.btnClearCtaContable.Enabled = True
            
        ElseIf .TipoMovimiento = "TRANSFERENCIA" Then
        
            Me.RadioButton1.value = False
            Me.RadioButton2.value = False
            Me.RadioButton3.value = True
        
            Me.Label1(0).caption = "Cuenta origen"
            Me.Label1(1).caption = "Cuenta destino"
        
            If IsSomething(.CuentaBancaria) Then
                Me.cboCuentasBancarias.ListIndex = funciones.PosIndexCbo(.CuentaBancaria.Id, Me.cboCuentasBancarias)
            Else
                Me.cboCuentasBancarias.ListIndex = -1
            End If
        
            If IsSomething(.CuentaBancariaDestino) Then
                Me.cboCuentaBancariaDestino.ListIndex = funciones.PosIndexCbo(.CuentaBancariaDestino.Id, Me.cboCuentaBancariaDestino)
            Else
                Me.cboCuentaBancariaDestino.ListIndex = -1
            End If
        
        End If
                
        Me.cboCuentasBancarias.Enabled = True
        Me.Label1(0).Enabled = True
        Me.btnClearCtaBancaria.Enabled = True
        
        '------------------------------
        ' Cuenta bancaria principal
        '------------------------------
        If IsSomething(.CuentaBancaria) Then
            Me.cboCuentasBancarias.ListIndex = funciones.PosIndexCbo(.CuentaBancaria.Id, Me.cboCuentasBancarias)
        Else
            Me.cboCuentasBancarias.ListIndex = -1
        End If
        
        '------------------------------
        ' Cuenta contable
        ' Puede venir Nothing en INGRESO
        '------------------------------
        If IsSomething(.CuentaContable) Then
            Me.cboCuentasContables.ListIndex = funciones.PosIndexCbo(.CuentaContable.Id, Me.cboCuentasContables)
        Else
            Me.cboCuentasContables.ListIndex = -1
        End If
        
        '------------------------------
        ' Moneda
        ' Si ya no la usás en cabecera, puede venir Nothing
        '------------------------------
        If IsSomething(.moneda) Then
            Me.cboMonedas.ListIndex = funciones.PosIndexCbo(.moneda.Id, Me.cboMonedas)
        Else
            Me.cboMonedas.ListIndex = -1
        End If
        
        '------------------------------
        ' Fecha / observaciones
        '------------------------------
        Me.dtpFecha.value = .FEcha
        Me.FlatEdit1.Text = .Observaciones
        
        '------------------------------
        ' Grillas
        '------------------------------
        
    End With

    Me.cboCuentasBancarias.Enabled = Not ReadOnly
    Me.gridBancos.AllowEdit = Not ReadOnly

    Me.cboMonedas.Enabled = Not ReadOnly
    Me.dtpFecha.Enabled = Not ReadOnly
    Me.btnClearCtaBancaria.Enabled = Not ReadOnly
    Me.btnClearCtaContable.Enabled = Not ReadOnly
    Me.FlatEdit1.Enabled = Not ReadOnly
    Me.RadioButton1.Enabled = Not ReadOnly
    Me.RadioButton2.Enabled = Not ReadOnly
    Me.cmdCrear.Enabled = Not ReadOnly
    Me.Label(0).Enabled = Not ReadOnly
    Me.Label1(0).Enabled = Not ReadOnly
    Me.Label(1).Enabled = Not ReadOnly
    Me.Label(2).Enabled = Not ReadOnly
    

End Sub


Private Sub RadioButton1_Click()

    If Me.RadioButton1.value Then
        ActualizarModoMovimiento
    End If

End Sub


Private Sub RadioButton2_Click()

    If Me.RadioButton2.value Then
        ActualizarModoMovimiento
    End If

End Sub


Private Sub RadioButton3_Click()

    If Me.RadioButton3.value Then
        ActualizarModoMovimiento
    End If

End Sub


Private Sub ActualizarModoMovimiento()

    Dim esIngreso As Boolean
    Dim esEgreso As Boolean
    Dim esTransferencia As Boolean

    esIngreso = Me.RadioButton1.value
    esEgreso = Me.RadioButton2.value
    esTransferencia = Me.RadioButton3.value

    ' Cuenta bancaria principal
    Me.cboCuentasBancarias.Enabled = _
        esIngreso Or esEgreso Or esTransferencia

    Me.btnClearCtaBancaria.Enabled = _
        Me.cboCuentasBancarias.Enabled

    Me.Label1(0).Enabled = _
        Me.cboCuentasBancarias.Enabled

    If esIngreso Then
        Me.Label1(0).caption = "Cuenta destino"

    ElseIf esEgreso Or esTransferencia Then
        Me.Label1(0).caption = "Cuenta origen"

    Else
        Me.Label1(0).caption = "Cuenta"
    End If

    ' Segunda cuenta: solo transferencia
    Me.Label1(1).Visible = esTransferencia
    Me.cboCuentaBancariaDestino.Visible = esTransferencia
    Me.btnClearCtaBancariaDestino.Visible = esTransferencia

    Me.Label1(1).Enabled = esTransferencia
    Me.cboCuentaBancariaDestino.Enabled = esTransferencia
    Me.btnClearCtaBancariaDestino.Enabled = esTransferencia

    Me.Label1(1).caption = "Cuenta destino"

    If Not esTransferencia Then
        Me.cboCuentaBancariaDestino.ListIndex = -1
    End If

    ' Cuenta contable: solo egreso
    Me.cboCuentasContables.Enabled = esEgreso
    Me.Label(0).Enabled = esEgreso
    Me.btnClearCtaContable.Enabled = esEgreso

    If Not esEgreso Then
        Me.cboCuentasContables.ListIndex = -1
        Set AsientoContable.CuentaContable = Nothing
    End If
    
    ' Cheques propios: solo ingreso/egreso, nunca transferencia
    Me.gridChequesPropios.Enabled = esIngreso Or esEgreso
    Me.gridChequesPropios.Visible = esIngreso Or esEgreso
    
    Me.gridChequeras.Enabled = esIngreso Or esEgreso
    Me.gridChequesChequera.Enabled = esIngreso Or esEgreso
    
    Me.gridChequeras.Visible = esIngreso Or esEgreso
    Me.gridChequesChequera.Visible = esIngreso Or esEgreso
    
    If esTransferencia Then
        Do While AsientoContable.ChequesPropios.count > 0
            AsientoContable.ChequesPropios.remove 1
        Loop
    
        Me.gridChequesPropios.ItemCount = 0
        Me.gridChequesPropios.Refresh
    End If


End Sub


Public Sub CargarValoresIniciales(ByVal pFecha As Date, ByVal pIdCuentaBancaria As Long)

    Me.dtpFecha.value = pFecha

    If pIdCuentaBancaria > 0 Then
        Me.cboCuentasBancarias.ListIndex = funciones.PosIndexCbo(pIdCuentaBancaria, Me.cboCuentasBancarias)

        If Me.cboCuentasBancarias.ListIndex <> -1 Then
            Set AsientoContable.CuentaBancaria = DAOCuentaBancaria.FindById(pIdCuentaBancaria)
        End If
    Else
        Me.cboCuentasBancarias.ListIndex = -1
    End If

End Sub


Private Sub btnClearCtaBancariaDestino_Click()

    Me.cboCuentaBancariaDestino.ListIndex = -1

End Sub


Private Function ObtenerCuentaSeleccionada( _
    ByVal cbo As Object _
) As CuentaBancaria

    If cbo.ListIndex = -1 Then
        Set ObtenerCuentaSeleccionada = Nothing
        Exit Function
    End If

    Set ObtenerCuentaSeleccionada = _
        DAOCuentaBancaria.FindById( _
            cbo.ItemData(cbo.ListIndex))

End Function


Private Sub LimpiarOperacionesMovimiento()

    Do While AsientoContable.operacionesBanco.count > 0
        AsientoContable.operacionesBanco.remove 1
    Loop

    Do While AsientoContable.OperacionesCaja.count > 0
        AsientoContable.OperacionesCaja.remove 1
    Loop

End Sub


Private Function CrearOperacionBancaria( _
    ByVal cuenta As CuentaBancaria, _
    ByVal monto As Double, _
    ByVal comprobante As String, _
    ByVal entradaSalida As Integer _
) As operacion

    Dim op As New operacion

    op.Pertenencia = OrigenOperacion.Banco
    op.monto = monto
    op.comprobante = comprobante
    op.FechaOperacion = Me.dtpFecha.value
    op.entradaSalida = entradaSalida

    Set op.CuentaBancaria = cuenta
    Set op.moneda = AsientoContable.moneda

    If IsSomething(AsientoContable.CuentaContable) Then
        Set op.CuentaContable = AsientoContable.CuentaContable
    Else
        Set op.CuentaContable = Nothing
    End If

    Set CrearOperacionBancaria = op

End Function


Private Function CuentaCoincideConMoneda( _
    ByVal cuenta As CuentaBancaria, _
    ByVal monedaSeleccionada As clsMoneda _
) As Boolean

    CuentaCoincideConMoneda = False

    If Not IsSomething(cuenta) Then Exit Function
    If Not IsSomething(monedaSeleccionada) Then Exit Function
    If Not IsSomething(cuenta.moneda) Then Exit Function

    CuentaCoincideConMoneda = _
        (cuenta.moneda.Id = monedaSeleccionada.Id)

End Function


Private Sub gridChequeras_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex <= chequeras.count Then
        Set tmpChequera = chequeras.item(RowIndex)
        Values(1) = tmpChequera.Description
        Values(2) = tmpChequera.Id
    End If
End Sub

