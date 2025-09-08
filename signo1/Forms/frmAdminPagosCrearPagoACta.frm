VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmAdminPagosCrearPagoACta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Crear Pago a Cuenta"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   9960
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.GroupBox GroupBox 
      Height          =   855
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   6000
      Width           =   9735
      _Version        =   786432
      _ExtentX        =   17171
      _ExtentY        =   1508
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.PushButton cmdCrear 
         Height          =   495
         Left            =   7440
         TabIndex        =   9
         Top             =   240
         Width           =   2055
         _Version        =   786432
         _ExtentX        =   3625
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Guardar"
         UseVisualStyle  =   -1  'True
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox 
      Height          =   1695
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9735
      _Version        =   786432
      _ExtentX        =   17171
      _ExtentY        =   2990
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
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.DateTimePicker dtpFecha 
         Height          =   330
         Left            =   1080
         TabIndex        =   17
         Top             =   1200
         Width           =   1245
         _Version        =   786432
         _ExtentX        =   2196
         _ExtentY        =   582
         _StockProps     =   68
         Format          =   1
         CurrentDate     =   40183.7263657407
      End
      Begin XtremeSuiteControls.PushButton btnClearProveedor 
         Height          =   375
         Left            =   5520
         TabIndex        =   18
         Top             =   330
         Width           =   375
         _Version        =   786432
         _ExtentX        =   661
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "X"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cboProveedores 
         Height          =   315
         Left            =   1080
         TabIndex        =   19
         Top             =   360
         Width           =   4335
         _Version        =   786432
         _ExtentX        =   7646
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Text            =   "cboProveedores"
      End
      Begin XtremeSuiteControls.ComboBox cboMonedas 
         Height          =   315
         Left            =   1080
         TabIndex        =   23
         Top             =   775
         Width           =   1245
         _Version        =   786432
         _ExtentX        =   2196
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Style           =   2
         Text            =   "cboMonedas"
      End
      Begin XtremeSuiteControls.Label Label 
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   22
         Top             =   1178
         Width           =   855
         _Version        =   786432
         _ExtentX        =   1508
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Fecha"
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label Label 
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   21
         Top             =   810
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
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   330
         Width           =   855
         _Version        =   786432
         _ExtentX        =   1508
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Proveedor"
         Alignment       =   1
      End
   End
   Begin XtremeSuiteControls.GroupBox grpOrigen 
      Height          =   3975
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   9735
      _Version        =   786432
      _ExtentX        =   17171
      _ExtentY        =   7011
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
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.TabControl TabControl 
         Height          =   3540
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   9540
         _Version        =   786432
         _ExtentX        =   16828
         _ExtentY        =   6244
         _StockProps     =   68
         Appearance      =   10
         Color           =   32
         PaintManager.ShowIcons=   -1  'True
         ItemCount       =   4
         SelectedItem    =   3
         Item(0).Caption =   "Cheques Propios"
         Item(0).ControlCount=   1
         Item(0).Control(0)=   "gridChequesPropios"
         Item(1).Caption =   "Banco"
         Item(1).ControlCount=   1
         Item(1).Control(0)=   "gridDepositosOperaciones"
         Item(2).Caption =   "Cheques 3ros"
         Item(2).ControlCount=   1
         Item(2).Control(0)=   "gridCheques"
         Item(3).Caption =   "Caja"
         Item(3).ControlCount=   2
         Item(3).Control(0)=   "gridCajaOperaciones"
         Item(3).Control(1)=   "gridCompensatorios"
         Begin GridEX20.GridEX gridDepositosOperaciones 
            Height          =   2910
            Left            =   -69895
            TabIndex        =   3
            Top             =   435
            Visible         =   0   'False
            Width           =   9330
            _ExtentX        =   16457
            _ExtentY        =   5133
            Version         =   "2.0"
            BoundColumnIndex=   ""
            ReplaceColumnIndex=   ""
            ColumnAutoResize=   -1  'True
            MethodHoldFields=   -1  'True
            ContScroll      =   -1  'True
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
            Column(1)       =   "frmAdminPagosCrearPagoACta.frx":0000
            Column(2)       =   "frmAdminPagosCrearPagoACta.frx":0160
            Column(3)       =   "frmAdminPagosCrearPagoACta.frx":029C
            Column(4)       =   "frmAdminPagosCrearPagoACta.frx":03D0
            Column(5)       =   "frmAdminPagosCrearPagoACta.frx":0514
            FormatStylesCount=   6
            FormatStyle(1)  =   "frmAdminPagosCrearPagoACta.frx":0618
            FormatStyle(2)  =   "frmAdminPagosCrearPagoACta.frx":0750
            FormatStyle(3)  =   "frmAdminPagosCrearPagoACta.frx":0800
            FormatStyle(4)  =   "frmAdminPagosCrearPagoACta.frx":08B4
            FormatStyle(5)  =   "frmAdminPagosCrearPagoACta.frx":098C
            FormatStyle(6)  =   "frmAdminPagosCrearPagoACta.frx":0A44
            ImageCount      =   0
            PrinterProperties=   "frmAdminPagosCrearPagoACta.frx":0B24
         End
         Begin GridEX20.GridEX gridCajaOperaciones 
            Height          =   2910
            Left            =   105
            TabIndex        =   4
            Top             =   435
            Width           =   9330
            _ExtentX        =   16457
            _ExtentY        =   5133
            Version         =   "2.0"
            BoundColumnIndex=   ""
            ReplaceColumnIndex=   ""
            ColumnAutoResize=   -1  'True
            MethodHoldFields=   -1  'True
            ContScroll      =   -1  'True
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
            Column(1)       =   "frmAdminPagosCrearPagoACta.frx":0CFC
            Column(2)       =   "frmAdminPagosCrearPagoACta.frx":0E5C
            Column(3)       =   "frmAdminPagosCrearPagoACta.frx":0F98
            Column(4)       =   "frmAdminPagosCrearPagoACta.frx":10CC
            Column(5)       =   "frmAdminPagosCrearPagoACta.frx":1200
            FormatStylesCount=   6
            FormatStyle(1)  =   "frmAdminPagosCrearPagoACta.frx":1304
            FormatStyle(2)  =   "frmAdminPagosCrearPagoACta.frx":143C
            FormatStyle(3)  =   "frmAdminPagosCrearPagoACta.frx":14EC
            FormatStyle(4)  =   "frmAdminPagosCrearPagoACta.frx":15A0
            FormatStyle(5)  =   "frmAdminPagosCrearPagoACta.frx":1678
            FormatStyle(6)  =   "frmAdminPagosCrearPagoACta.frx":1730
            ImageCount      =   0
            PrinterProperties=   "frmAdminPagosCrearPagoACta.frx":1810
         End
         Begin GridEX20.GridEX gridChequesPropios 
            Height          =   2910
            Left            =   -69895
            TabIndex        =   5
            Top             =   435
            Visible         =   0   'False
            Width           =   9330
            _ExtentX        =   16457
            _ExtentY        =   5133
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
            Column(1)       =   "frmAdminPagosCrearPagoACta.frx":19E8
            Column(2)       =   "frmAdminPagosCrearPagoACta.frx":1B50
            Column(3)       =   "frmAdminPagosCrearPagoACta.frx":1C84
            Column(4)       =   "frmAdminPagosCrearPagoACta.frx":1DC0
            Column(5)       =   "frmAdminPagosCrearPagoACta.frx":1F28
            FormatStylesCount=   6
            FormatStyle(1)  =   "frmAdminPagosCrearPagoACta.frx":2020
            FormatStyle(2)  =   "frmAdminPagosCrearPagoACta.frx":2158
            FormatStyle(3)  =   "frmAdminPagosCrearPagoACta.frx":2208
            FormatStyle(4)  =   "frmAdminPagosCrearPagoACta.frx":22BC
            FormatStyle(5)  =   "frmAdminPagosCrearPagoACta.frx":2394
            FormatStyle(6)  =   "frmAdminPagosCrearPagoACta.frx":244C
            ImageCount      =   0
            PrinterProperties=   "frmAdminPagosCrearPagoACta.frx":252C
         End
         Begin GridEX20.GridEX gridCheques 
            Height          =   2910
            Left            =   -69895
            TabIndex        =   6
            Top             =   435
            Visible         =   0   'False
            Width           =   9330
            _ExtentX        =   16457
            _ExtentY        =   5133
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
            ColumnsCount    =   7
            Column(1)       =   "frmAdminPagosCrearPagoACta.frx":2704
            Column(2)       =   "frmAdminPagosCrearPagoACta.frx":2884
            Column(3)       =   "frmAdminPagosCrearPagoACta.frx":2A24
            Column(4)       =   "frmAdminPagosCrearPagoACta.frx":2B1C
            Column(5)       =   "frmAdminPagosCrearPagoACta.frx":2C58
            Column(6)       =   "frmAdminPagosCrearPagoACta.frx":2D64
            Column(7)       =   "frmAdminPagosCrearPagoACta.frx":2E34
            FormatStylesCount=   6
            FormatStyle(1)  =   "frmAdminPagosCrearPagoACta.frx":2F20
            FormatStyle(2)  =   "frmAdminPagosCrearPagoACta.frx":3058
            FormatStyle(3)  =   "frmAdminPagosCrearPagoACta.frx":3108
            FormatStyle(4)  =   "frmAdminPagosCrearPagoACta.frx":31BC
            FormatStyle(5)  =   "frmAdminPagosCrearPagoACta.frx":3294
            FormatStyle(6)  =   "frmAdminPagosCrearPagoACta.frx":334C
            ImageCount      =   0
            PrinterProperties=   "frmAdminPagosCrearPagoACta.frx":342C
         End
         Begin GridEX20.GridEX gridCompensatorios 
            Height          =   2910
            Left            =   -69895
            TabIndex        =   7
            Top             =   435
            Visible         =   0   'False
            Width           =   9330
            _ExtentX        =   16457
            _ExtentY        =   5133
            Version         =   "2.0"
            BoundColumnIndex=   ""
            ReplaceColumnIndex=   ""
            PreviewColumn   =   "observacion"
            PreviewRowLines =   1
            ColumnAutoResize=   -1  'True
            MethodHoldFields=   -1  'True
            ContScroll      =   -1  'True
            AllowColumnDrag =   0   'False
            AllowDelete     =   -1  'True
            GroupByBoxVisible=   0   'False
            RowHeaders      =   -1  'True
            DataMode        =   99
            ColumnHeaderHeight=   285
            IntProp1        =   0
            IntProp2        =   0
            IntProp7        =   0
            ColumnsCount    =   5
            Column(1)       =   "frmAdminPagosCrearPagoACta.frx":3604
            Column(2)       =   "frmAdminPagosCrearPagoACta.frx":374C
            Column(3)       =   "frmAdminPagosCrearPagoACta.frx":3858
            Column(4)       =   "frmAdminPagosCrearPagoACta.frx":3944
            Column(5)       =   "frmAdminPagosCrearPagoACta.frx":3A48
            FormatStylesCount=   6
            FormatStyle(1)  =   "frmAdminPagosCrearPagoACta.frx":3B88
            FormatStyle(2)  =   "frmAdminPagosCrearPagoACta.frx":3CC0
            FormatStyle(3)  =   "frmAdminPagosCrearPagoACta.frx":3D70
            FormatStyle(4)  =   "frmAdminPagosCrearPagoACta.frx":3E24
            FormatStyle(5)  =   "frmAdminPagosCrearPagoACta.frx":3EFC
            FormatStyle(6)  =   "frmAdminPagosCrearPagoACta.frx":3FB4
            ImageCount      =   0
            PrinterProperties=   "frmAdminPagosCrearPagoACta.frx":4094
         End
      End
   End
   Begin GridEX20.GridEX gridBancos 
      Height          =   1845
      Left            =   4440
      TabIndex        =   10
      Top             =   7560
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
      Column(1)       =   "frmAdminPagosCrearPagoACta.frx":426C
      Column(2)       =   "frmAdminPagosCrearPagoACta.frx":436C
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminPagosCrearPagoACta.frx":445C
      FormatStyle(2)  =   "frmAdminPagosCrearPagoACta.frx":4594
      FormatStyle(3)  =   "frmAdminPagosCrearPagoACta.frx":4644
      FormatStyle(4)  =   "frmAdminPagosCrearPagoACta.frx":46F8
      FormatStyle(5)  =   "frmAdminPagosCrearPagoACta.frx":47D0
      FormatStyle(6)  =   "frmAdminPagosCrearPagoACta.frx":4888
      ImageCount      =   0
      PrinterProperties=   "frmAdminPagosCrearPagoACta.frx":4968
   End
   Begin GridEX20.GridEX gridCuentasBancarias 
      Height          =   1695
      Left            =   8040
      TabIndex        =   11
      Top             =   7560
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
      Column(1)       =   "frmAdminPagosCrearPagoACta.frx":4B40
      Column(2)       =   "frmAdminPagosCrearPagoACta.frx":4C64
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminPagosCrearPagoACta.frx":4D58
      FormatStyle(2)  =   "frmAdminPagosCrearPagoACta.frx":4E90
      FormatStyle(3)  =   "frmAdminPagosCrearPagoACta.frx":4F40
      FormatStyle(4)  =   "frmAdminPagosCrearPagoACta.frx":4FF4
      FormatStyle(5)  =   "frmAdminPagosCrearPagoACta.frx":50CC
      FormatStyle(6)  =   "frmAdminPagosCrearPagoACta.frx":5184
      ImageCount      =   0
      PrinterProperties=   "frmAdminPagosCrearPagoACta.frx":5264
   End
   Begin GridEX20.GridEX gridMonedas 
      Height          =   1815
      Left            =   120
      TabIndex        =   12
      Top             =   7560
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
      Column(1)       =   "frmAdminPagosCrearPagoACta.frx":543C
      Column(2)       =   "frmAdminPagosCrearPagoACta.frx":5560
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminPagosCrearPagoACta.frx":5654
      FormatStyle(2)  =   "frmAdminPagosCrearPagoACta.frx":578C
      FormatStyle(3)  =   "frmAdminPagosCrearPagoACta.frx":583C
      FormatStyle(4)  =   "frmAdminPagosCrearPagoACta.frx":58F0
      FormatStyle(5)  =   "frmAdminPagosCrearPagoACta.frx":59C8
      FormatStyle(6)  =   "frmAdminPagosCrearPagoACta.frx":5A80
      ImageCount      =   0
      PrinterProperties=   "frmAdminPagosCrearPagoACta.frx":5B60
   End
   Begin GridEX20.GridEX gridCajas 
      Height          =   1695
      Left            =   120
      TabIndex        =   13
      Top             =   9480
      Width           =   3420
      _ExtentX        =   6033
      _ExtentY        =   2990
      Version         =   "2.0"
      BoundColumnIndex=   "id"
      ReplaceColumnIndex=   "caja"
      ActAsDropDown   =   -1  'True
      HideSelection   =   2
      MethodHoldFields=   -1  'True
      ContScroll      =   -1  'True
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      ColumnHeaders   =   0   'False
      NewRowPos       =   1
      RowHeaders      =   -1  'True
      DataMode        =   99
      ColumnHeaderHeight=   285
      ColumnsCount    =   2
      Column(1)       =   "frmAdminPagosCrearPagoACta.frx":5D38
      Column(2)       =   "frmAdminPagosCrearPagoACta.frx":5E38
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminPagosCrearPagoACta.frx":5F24
      FormatStyle(2)  =   "frmAdminPagosCrearPagoACta.frx":605C
      FormatStyle(3)  =   "frmAdminPagosCrearPagoACta.frx":610C
      FormatStyle(4)  =   "frmAdminPagosCrearPagoACta.frx":61C0
      FormatStyle(5)  =   "frmAdminPagosCrearPagoACta.frx":6298
      FormatStyle(6)  =   "frmAdminPagosCrearPagoACta.frx":6350
      ImageCount      =   0
      PrinterProperties=   "frmAdminPagosCrearPagoACta.frx":6430
   End
   Begin GridEX20.GridEX gridChequesDisponibles 
      Height          =   1905
      Left            =   4440
      TabIndex        =   14
      Top             =   9480
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
      Column(1)       =   "frmAdminPagosCrearPagoACta.frx":6608
      Column(2)       =   "frmAdminPagosCrearPagoACta.frx":6788
      Column(3)       =   "frmAdminPagosCrearPagoACta.frx":6928
      Column(4)       =   "frmAdminPagosCrearPagoACta.frx":6A20
      Column(5)       =   "frmAdminPagosCrearPagoACta.frx":6B5C
      Column(6)       =   "frmAdminPagosCrearPagoACta.frx":6C68
      Column(7)       =   "frmAdminPagosCrearPagoACta.frx":6D88
      Column(8)       =   "frmAdminPagosCrearPagoACta.frx":6E94
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminPagosCrearPagoACta.frx":6F88
      FormatStyle(2)  =   "frmAdminPagosCrearPagoACta.frx":70C0
      FormatStyle(3)  =   "frmAdminPagosCrearPagoACta.frx":7170
      FormatStyle(4)  =   "frmAdminPagosCrearPagoACta.frx":7224
      FormatStyle(5)  =   "frmAdminPagosCrearPagoACta.frx":72FC
      FormatStyle(6)  =   "frmAdminPagosCrearPagoACta.frx":73B4
      ImageCount      =   0
      PrinterProperties=   "frmAdminPagosCrearPagoACta.frx":7494
   End
   Begin GridEX20.GridEX gridChequeras 
      Height          =   1815
      Left            =   10440
      TabIndex        =   15
      Top             =   9480
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
      Column(1)       =   "frmAdminPagosCrearPagoACta.frx":766C
      Column(2)       =   "frmAdminPagosCrearPagoACta.frx":778C
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminPagosCrearPagoACta.frx":788C
      FormatStyle(2)  =   "frmAdminPagosCrearPagoACta.frx":79C4
      FormatStyle(3)  =   "frmAdminPagosCrearPagoACta.frx":7A74
      FormatStyle(4)  =   "frmAdminPagosCrearPagoACta.frx":7B28
      FormatStyle(5)  =   "frmAdminPagosCrearPagoACta.frx":7C00
      FormatStyle(6)  =   "frmAdminPagosCrearPagoACta.frx":7CB8
      ImageCount      =   0
      PrinterProperties=   "frmAdminPagosCrearPagoACta.frx":7D98
   End
   Begin GridEX20.GridEX gridChequesChequera 
      Height          =   1710
      Left            =   12360
      TabIndex        =   16
      Top             =   7560
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
      Column(1)       =   "frmAdminPagosCrearPagoACta.frx":7F70
      Column(2)       =   "frmAdminPagosCrearPagoACta.frx":80A0
      SortKeysCount   =   1
      SortKey(1)      =   "frmAdminPagosCrearPagoACta.frx":81A0
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminPagosCrearPagoACta.frx":8208
      FormatStyle(2)  =   "frmAdminPagosCrearPagoACta.frx":8340
      FormatStyle(3)  =   "frmAdminPagosCrearPagoACta.frx":83F0
      FormatStyle(4)  =   "frmAdminPagosCrearPagoACta.frx":84A4
      FormatStyle(5)  =   "frmAdminPagosCrearPagoACta.frx":857C
      FormatStyle(6)  =   "frmAdminPagosCrearPagoACta.frx":8634
      ImageCount      =   0
      PrinterProperties=   "frmAdminPagosCrearPagoACta.frx":8714
   End
End
Attribute VB_Name = "frmAdminPagosCrearPagoACta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim colProveedores As New Collection
Dim prov As clsProveedor
Dim formLoading As Boolean
Dim formLoaded As Boolean

Private operacion As operacion
Private PagoACta As New clsPagoACta
Private Banco As Banco
Private caja As caja
Private CuentaBancaria As CuentaBancaria
Private moneda As clsMoneda
Private alicuotaRetencion As DTORetencionAlicuota
Private cuentasBancarias As New Collection
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

Private Sub btnClearProveedor_Click()
        Me.cboProveedores.ListIndex = -1
End Sub


Private Sub cboMonedas_Click()
    If Me.cboMonedas.ListIndex = -1 Then
        Set PagoACta.moneda = Nothing
    Else
        Set PagoACta.moneda = DAOMoneda.GetById(Me.cboMonedas.ItemData(Me.cboMonedas.ListIndex))
    End If
'    Totalizar
End Sub


Private Sub cboProveedores_Click()
    If IsSomething(PagoACta) And Me.cboProveedores.ListIndex <> -1 Then

        Set PagoACta.Proveedor = DAOProveedor.FindById(Me.cboProveedores.ItemData(Me.cboProveedores.ListIndex))
    
    End If
        
End Sub


Private Sub cmdCrear_Click()
    If Me.gridChequesPropios.EditMode = jgexEditModeOn Then
        MsgBox "Todavia esta editando la grilla de cheques propios.", vbExclamation
        Exit Sub
    End If

    If Me.gridCheques.EditMode = jgexEditModeOn Then
        MsgBox "Todavia esta editando la grilla de cheques de 3ros.", vbExclamation
        Exit Sub
    End If

    If Me.gridCajaOperaciones.EditMode = jgexEditModeOn Then
        MsgBox "Todavia esta editando la grilla de caja.", vbExclamation
        Exit Sub
    End If

    If Me.gridDepositosOperaciones.EditMode = jgexEditModeOn Then
        MsgBox "Todavia esta editando la grilla de banco.", vbExclamation
        Exit Sub
    End If

    PagoACta.FEcha = Me.dtpFecha.value
    
    PagoACta.StaticTotalOrigenes = PagoACta.TotalOrigenes
    
    
    If PagoACta.IsValid Then

        Dim n As Boolean: n = (PagoACta.Id = 0)

        If DAOPagoACta.Save(PagoACta, True) Then

            'Me.btnGuardar.Enabled = False

            If n Then
                MsgBox "Pago a cuenta Nro " & PagoACta.Id & " creado con éxito.", vbInformation
            Else

                MsgBox "Pago a cuenta modificado con éxito.", vbInformation
            End If

            If n Then
                If MsgBox("Desea crear un nuevo Pago a cuenta?", vbQuestion + vbYesNo) = vbYes Then
                    Dim f12 As New frmAdminPagosCrearPagoACta
                    f12.Show
                End If
            End If

            Unload Me
        Else
            MsgBox "Hubo un problema al guardar el Pago a cuenta.", vbCritical
        End If
    Else
        MsgBox PagoACta.ValidationMessages, vbCritical, "Error"
    End If


End Sub

Private Sub Form_Load()

    FormHelper.Customize Me
    
    LlenarComboProveedores
    
    
    formLoading = True
    
    Me.Left = frmPrincipal.ScaleWidth / 6
    Me.Top = frmPrincipal.ScaleHeight / 22
    
    Me.gridChequeras.Visible = False
    Me.gridChequesChequera.Visible = False
    Me.gridCompensatorios.ItemCount = 0
    
    GridEXHelper.CustomizeGrid Me.gridCajaOperaciones, False, True
    GridEXHelper.CustomizeGrid Me.gridDepositosOperaciones, False, True
    GridEXHelper.CustomizeGrid Me.gridCheques, False, True
    GridEXHelper.CustomizeGrid Me.gridChequesDisponibles, False, False
    GridEXHelper.CustomizeGrid Me.gridBancos, False, False
    GridEXHelper.CustomizeGrid Me.gridCuentasBancarias, False, False
    GridEXHelper.CustomizeGrid Me.gridMonedas, False, False
    GridEXHelper.CustomizeGrid Me.gridCajas, False, False
    GridEXHelper.CustomizeGrid Me.gridChequeras, False, False
    GridEXHelper.CustomizeGrid Me.gridChequesPropios, False, True
    GridEXHelper.CustomizeGrid Me.gridCompensatorios, False, True
    GridEXHelper.CustomizeGrid Me.gridChequesChequera
      
    Set Cajas = DAOCaja.FindAll()
    Me.gridCajas.ItemCount = Cajas.count

    Set Monedas = DAOMoneda.GetAll()
    Me.gridMonedas.ItemCount = Monedas.count

    Set cuentasBancarias = DAOCuentaBancaria.FindAll()
    Me.gridCuentasBancarias.ItemCount = cuentasBancarias.count

    Set bancos = DAOBancos.GetAll()
    Me.gridBancos.ItemCount = bancos.count

    Set chequeras = DAOChequeras.FindAllWithChequesDisponibles()
    Me.gridChequeras.ItemCount = chequeras.count

    CargarChequesDisponibles

    Me.gridCajaOperaciones.ItemCount = PagoACta.OperacionesCaja.count
    Me.gridDepositosOperaciones.ItemCount = PagoACta.operacionesBanco.count
    Me.gridCheques.ItemCount = PagoACta.ChequesTerceros.count
    Me.gridChequesPropios.ItemCount = PagoACta.ChequesPropios.count



    Set Me.gridCheques.Columns("numero").DropDownControl = Me.gridChequesDisponibles

    Set Me.gridDepositosOperaciones.Columns("moneda").DropDownControl = Me.gridMonedas
    Set Me.gridDepositosOperaciones.Columns("cuenta").DropDownControl = Me.gridCuentasBancarias

    Set Me.gridCajaOperaciones.Columns("moneda").DropDownControl = Me.gridMonedas
    Set Me.gridCajaOperaciones.Columns("caja").DropDownControl = Me.gridCajas

    Set Me.gridChequesPropios.Columns("chequera").DropDownControl = Me.gridChequeras
    Set Me.gridChequesPropios.Columns("numero").DropDownControl = Me.gridChequesChequera
    gridChequesChequera.ItemCount = 0
    GridEXHelper.AutoSizeColumns Me.gridChequeras


    DAOMoneda.llenarComboXtremeSuite Me.cboMonedas

    Me.dtpFecha.value = PagoACta.FEcha
    
    'lstFacturas_Click
'    Totalizar

    formLoaded = True
    formLoading = False
End Sub


Private Sub LlenarComboProveedores()
    
'''    DAOProveedor.llenarComboXtremeSuite Me.cboProveedores, True, True, True

    Call DAOProveedor.LlenarComboProveedores(cboProveedores)
    Me.cboProveedores.ListIndex = -1
    
End Sub


Private Sub gridBancos_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex <= bancos.count Then
        Set Banco = bancos.item(RowIndex)
        Values(1) = Banco.Id
        Values(2) = Banco.nombre
    End If
End Sub


Private Sub gridCajaOperaciones_BeforeUpdate(ByVal Cancel As GridEX20.JSRetBoolean)
    Dim cond1 As Boolean
    Dim cond2 As Boolean
    Dim cond3 As Boolean
    Dim cond4 As Boolean


    cond1 = Not IsNumeric(Me.gridCajaOperaciones.value(1))
    cond2 = Not IsNumeric(Me.gridCajaOperaciones.value(2)) And LenB(Me.gridCajaOperaciones.value(2)) = 0
    cond3 = Not IsDate(Me.gridCajaOperaciones.value(3))
    cond4 = LenB(Me.gridCajaOperaciones.value(4)) = 0 Or IsEmpty(Me.gridCajaOperaciones.value(4))    'or Not IsNumeric(Me.gridCajaOperaciones.value(4))

    Cancel = cond1 Or cond2 Or cond3 Or cond4
    
End Sub


Private Sub gridCajaOperaciones_UnboundAddNew(ByVal NewRowBookmark As GridEX20.JSRetVariant, ByVal Values As GridEX20.JSRowData)
    Set operacion = New operacion
    'operacion.IdPertenencia = recibo.Id
    operacion.Pertenencia = OrigenOperacion.caja
    operacion.Monto = Values(1)
    operacion.Comprobante = Values(5)
    If IsNumeric(Values(2)) Then
        Set operacion.moneda = DAOMoneda.GetById(Values(2))
    End If
    operacion.FechaOperacion = Values(3)
    If IsNumeric(Values(4)) Then
        Set operacion.caja = DAOCaja.FindById(Values(4))
    End If
    operacion.EntradaSalida = OPSalida
    PagoACta.OperacionesCaja.Add operacion
'    Totalizar
End Sub


Private Sub gridCajaOperaciones_UnboundDelete(ByVal RowIndex As Long, ByVal Bookmark As Variant)
    If RowIndex > 0 And PagoACta.OperacionesCaja.count >= RowIndex Then
        PagoACta.OperacionesCaja.remove RowIndex
'        Totalizar
    End If
End Sub


Private Sub gridCajaOperaciones_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex <= PagoACta.OperacionesCaja.count Then
        Set operacion = PagoACta.OperacionesCaja.item(RowIndex)
        'FORMATCURRENCY
        Values(1) = FormatCurrency(funciones.FormatearDecimales(operacion.Monto))
        If IsSomething(operacion.moneda) Then
            Values(2) = operacion.moneda.NombreCorto
        End If
        Values(3) = operacion.FechaOperacion
        If IsSomething(operacion.caja) Then
            Values(4) = operacion.caja.nombre
        End If
        If IsSomething(operacion) Then
            Values(5) = operacion.Comprobante
        End If
    End If
End Sub

Private Sub gridCajaOperaciones_UnboundUpdate(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex > 0 And PagoACta.OperacionesCaja.count > 0 Then
        Set operacion = PagoACta.OperacionesCaja.item(RowIndex)
        'operacion.IdPertenencia = recibo.id
        'operacion.Pertenencia = Banco
        operacion.Monto = Values(1)
        operacion.Comprobante = Values(5)
        If IsNumeric(Values(2)) Then
            Set operacion.moneda = DAOMoneda.GetById(Values(2))
        End If
        operacion.FechaOperacion = Values(3)
        If IsNumeric(Values(4)) Then
            Set operacion.caja = DAOCaja.FindById(Values(4))
        End If
        operacion.EntradaSalida = OPSalida
'        Totalizar
    End If
End Sub

Private Sub gridCajas_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex > 0 And Cajas.count > 0 Then
        Set caja = Cajas.item(RowIndex)
        Values(1) = caja.Id
        Values(2) = caja.nombre
    End If
End Sub


Private Sub gridChequeras_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex <= chequeras.count Then
        Set tmpChequera = chequeras.item(RowIndex)
        Values(1) = tmpChequera.Description
        Values(2) = tmpChequera.Id
    End If
End Sub

Private Sub gridCheques_BeforeUpdate(ByVal Cancel As GridEX20.JSRetBoolean)
    Dim msg As New Collection

    ' REVISA QUE EN LA COLECCION DE CHEQUES DE TERCEROS QUE SE ESTAN CARGANDO NO EST? INGRESADO EL MISMO CHEQUE, SI LO DETECTA GENERA MSG DE ERROR
    If funciones.BuscarEnColeccion(PagoACta.ChequesTerceros, CStr(Me.gridCheques.value(1))) Then
        msg.Add "El cheque seleccionado ya fue ingresado anteriormente."
    End If

    Cancel = (msg.count > 0)
    If Cancel Then MsgBox funciones.JoinCollectionValues(msg, vbNewLine), vbExclamation

End Sub

Private Sub gridCheques_UnboundAddNew(ByVal NewRowBookmark As GridEX20.JSRetVariant, ByVal Values As GridEX20.JSRowData)
    Set cheque = Nothing
    If IsNumeric(Values(1)) Then Set cheque = DAOCheques.FindById(Values(1))
    If IsSomething(cheque) Then
        PagoACta.ChequesTerceros.Add cheque, CStr(cheque.Id)

    End If
'    Totalizar

End Sub

Private Sub gridCheques_UnboundDelete(ByVal RowIndex As Long, ByVal Bookmark As Variant)
    If RowIndex > 0 Then
        PagoACta.ChequesTerceros.remove RowIndex
'        Totalizar
    End If
End Sub

Private Sub gridCheques_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex <= PagoACta.ChequesTerceros.count Then
        Set cheque = PagoACta.ChequesTerceros.item(RowIndex)

        Values(1) = cheque.numero & " "

        'FORMATCURRENCY
        Values(2) = FormatCurrency(cheque.Monto)
        Values(3) = cheque.FechaVencimiento
        If IsSomething(cheque.moneda) Then Values(4) = cheque.moneda.NombreCorto
        If IsSomething(cheque.Banco) Then Values(5) = cheque.Banco.nombre
        Values(6) = cheque.OrigenDestino
        Values(7) = cheque.OrigenCheque
    
    End If
End Sub

Private Sub gridCheques_UnboundUpdate(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex > 0 And PagoACta.ChequesTerceros.count >= RowIndex Then
        Set cheque = Nothing
        If IsNumeric(Values(1)) Then Set cheque = DAOCheques.FindById(Values(1))
        If IsSomething(cheque) Then
            PagoACta.ChequesTerceros.Add cheque, , , RowIndex
            PagoACta.ChequesTerceros.remove RowIndex
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
        Values(2) = FormatCurrency(cheque.Monto)
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
    If funciones.BuscarEnColeccion(PagoACta.ChequesPropios, CStr(Me.gridChequesPropios.value(2))) Then
        msg.Add "El cheque seleccionado ya fue ingresado anteriormente."
    End If

    If Not IsNumeric(Me.gridChequesPropios.value(3)) Then
        msg.Add "Debe especificar un monto vÃ¡lido."
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
            Set chequesChequeraSeleccionada = DAOCheques.FindAllDisponiblesByChequera(Val(value))  ' Me.gridChequesPropios.Value(1))
        End If

        Me.gridChequesChequera.ItemCount = chequesChequeraSeleccionada.count
    End If
End Sub


Private Sub gridChequesPropios_UnboundAddNew(ByVal NewRowBookmark As GridEX20.JSRetVariant, ByVal Values As GridEX20.JSRowData)
    Set cheque = Nothing
    If IsNumeric(Values(2)) Then Set cheque = DAOCheques.FindById(Values(2))
    If IsSomething(cheque) Then
        cheque.Monto = Values(3)
        cheque.FechaVencimiento = Values(4)

        PagoACta.ChequesPropios.Add cheque, CStr(cheque.Id)


    End If
'    Totalizar
End Sub

Private Sub gridChequesPropios_UnboundDelete(ByVal RowIndex As Long, ByVal Bookmark As Variant)
    If RowIndex > 0 Then
        PagoACta.ChequesPropios.remove RowIndex
'        Totalizar
    End If
End Sub

Private Sub gridChequesPropios_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If PagoACta.ChequesPropios.count >= RowIndex Then
        Set cheque = PagoACta.ChequesPropios.item(RowIndex)
        Values(1) = cheque.chequera.Description
        Values(2) = vbNullString
        'FORMATCURRENCY
        Values(3) = FormatCurrency(cheque.Monto)
        Values(4) = cheque.FechaVencimiento
        Values(5) = cheque.numero


'        Totalizar
    End If
End Sub

Private Sub gridChequesPropios_UnboundUpdate(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If PagoACta.ChequesPropios.count >= RowIndex Then
        Set cheque = PagoACta.ChequesPropios.item(RowIndex)

        '        If Values(2) <> Cheque.Id Then
        '            ordenPago.ChequesPropios.remove CStr(Cheque.Id)
        '            Set Cheque = DAOCheques.FindById(Values(2))
        '            ordenPago.ChequesPropios.Add Cheque, CStr(Cheque.Id)
        '        End If

        cheque.Monto = Values(3)
        cheque.FechaVencimiento = Values(4)
    End If

'    Totalizar
End Sub

Private Sub gridCuentasBancarias_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If cuentasBancarias.count >= RowIndex Then
        Set CuentaBancaria = cuentasBancarias.item(RowIndex)
        Values(1) = CuentaBancaria.Id
        Values(2) = CuentaBancaria.DescripcionFormateada
    End If
End Sub

Private Sub gridDepositosOperaciones_BeforeUpdate(ByVal Cancel As GridEX20.JSRetBoolean)
    
    Dim cond1 As Boolean
    Dim cond2 As Boolean
    Dim cond3 As Boolean
    Dim cond4 As Boolean


    cond1 = Not IsNumeric(Me.gridDepositosOperaciones.value(1))
    cond2 = Not IsNumeric(Me.gridDepositosOperaciones.value(2)) And LenB(Me.gridDepositosOperaciones.value(2)) = 0
    cond3 = Not IsDate(Me.gridDepositosOperaciones.value(3))
    cond4 = Not IsNumeric(Me.gridDepositosOperaciones.value(4)) And LenB(Me.gridDepositosOperaciones.value(4)) = 0

    Cancel = cond1 Or cond2 Or cond3 Or cond4
    
End Sub

Private Sub gridDepositosOperaciones_UnboundAddNew(ByVal NewRowBookmark As GridEX20.JSRetVariant, ByVal Values As GridEX20.JSRowData)
    Set operacion = New operacion
    'operacion.IdPertenencia = recibo.Id
    operacion.Pertenencia = OrigenOperacion.Banco
    operacion.Monto = Values(1)
    operacion.Comprobante = Values(5)
    If IsNumeric(Values(2)) Then
        Set operacion.moneda = DAOMoneda.GetById(Values(2))
    End If
    operacion.FechaOperacion = Values(3)
    If IsNumeric(Values(4)) Then
        Set operacion.CuentaBancaria = DAOCuentaBancaria.FindById(Values(4))
    End If
    operacion.EntradaSalida = OPSalida
    PagoACta.operacionesBanco.Add operacion
''    Totalizar
End Sub

Private Sub gridDepositosOperaciones_UnboundDelete(ByVal RowIndex As Long, ByVal Bookmark As Variant)
    If RowIndex > 0 And PagoACta.operacionesBanco.count >= RowIndex Then
        PagoACta.operacionesBanco.remove RowIndex
'        Totalizar
    End If
End Sub

Private Sub gridDepositosOperaciones_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex <= PagoACta.operacionesBanco.count Then
        Set operacion = PagoACta.operacionesBanco.item(RowIndex)
        'FORMATCURRENCY
        Values(1) = FormatCurrency(funciones.FormatearDecimales(operacion.Monto))
        If IsSomething(operacion.moneda) Then
            Values(2) = operacion.moneda.NombreCorto
        End If
        Values(3) = operacion.FechaOperacion
        If IsSomething(operacion.CuentaBancaria) Then
            Values(4) = operacion.CuentaBancaria.DescripcionFormateada
        End If
        If IsSomething(operacion) Then
            Values(5) = operacion.Comprobante
        End If
    End If
End Sub

Private Sub gridDepositosOperaciones_UnboundUpdate(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex > 0 And PagoACta.operacionesBanco.count > 0 Then
        Set operacion = PagoACta.operacionesBanco.item(RowIndex)
        'operacion.IdPertenencia = recibo.id
        'operacion.Pertenencia = Banco
        operacion.Monto = Values(1)
        operacion.Comprobante = Values(5)
        If IsNumeric(Values(2)) Then
            Set operacion.moneda = DAOMoneda.GetById(Values(2))
        End If
        operacion.FechaOperacion = Values(3)
        If IsNumeric(Values(4)) Then
            Set operacion.CuentaBancaria = DAOCuentaBancaria.FindById(Values(4))
        End If
        operacion.EntradaSalida = OPSalida
'        Totalizar
    End If
End Sub


Private Sub CargarChequesDisponibles()
    Set chequesDisponibles = DAOCheques.FindAllEnCarteraDeTerceros
    Me.gridChequesDisponibles.ItemCount = chequesDisponibles.count
End Sub


Private Sub gridMonedas_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex > 0 And Monedas.count > 0 Then
        Set moneda = Monedas.item(RowIndex)
        Values(1) = moneda.Id
        Values(2) = moneda.NombreCorto
    End If
End Sub


Public Sub Cargar(pcta As clsPagoACta)

    Me.caption = "Pago a Cuenta Nro " & pcta.Id

    If Not IsSomething(pcta) Then
        MsgBox "La OP que está intentando visualizar est? en estado PENDIENTE. " & vbNewLine & "Por lo tanto no puede ser mostrada porque puede estar siendo editada." & vbNewLine & "Verifiquelo por favor.", vbCritical, "OP Pendiente"
        Unload Me
        Exit Sub

    End If

    Set PagoACta = DAOPagoACta.FindById(pcta.Id)
    
    Me.caption = "Pago a Cuenta Nro " & pcta.Id
   
    With PagoACta

        Me.cboProveedores.ListIndex = funciones.PosIndexCbo(.Proveedor.Id, Me.cboProveedores)
        
        Me.gridCajaOperaciones.ItemCount = .OperacionesCaja.count
        Me.gridDepositosOperaciones.ItemCount = .operacionesBanco.count
        Me.gridCheques.ItemCount = .ChequesTerceros.count
        Me.gridChequesPropios.ItemCount = .ChequesPropios.count

        Me.cboMonedas.ListIndex = funciones.PosIndexCbo(.moneda.Id, Me.cboMonedas)
        Me.dtpFecha.value = .FEcha

    End With

    Me.cboProveedores.Enabled = Not ReadOnly
    Me.btnClearProveedor.Enabled = Not ReadOnly
    Me.gridDepositosOperaciones.AllowEdit = Not ReadOnly
    Me.gridDepositosOperaciones.AllowDelete = Not ReadOnly
    Me.gridBancos.AllowEdit = Not ReadOnly
    Me.gridCajaOperaciones.AllowEdit = Not ReadOnly
    Me.gridCajaOperaciones.AllowDelete = Not ReadOnly
    Me.gridCajas.AllowEdit = Not ReadOnly
    Me.gridChequeras.AllowEdit = Not ReadOnly
    Me.gridCheques.AllowEdit = Not ReadOnly
    Me.gridCheques.AllowDelete = Not ReadOnly
    Me.gridChequesChequera.AllowEdit = Not ReadOnly
    Me.gridChequesDisponibles.AllowEdit = Not ReadOnly
    Me.gridChequesPropios.AllowEdit = Not ReadOnly
    Me.gridChequesPropios.AllowDelete = Not ReadOnly
    Me.cboMonedas.Enabled = Not ReadOnly
    Me.dtpFecha.Enabled = Not ReadOnly

End Sub

