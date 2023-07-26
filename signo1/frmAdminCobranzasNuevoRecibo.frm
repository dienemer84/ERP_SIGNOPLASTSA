VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmAdminCobranzasNuevoRecibo 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recibo"
   ClientHeight    =   8010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   18000
   ClipControls    =   0   'False
   Icon            =   "frmAdminCobranzasNuevoRecibo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   18000
   Begin XtremeSuiteControls.PushButton cmdCerrar 
      Height          =   405
      Left            =   10200
      TabIndex        =   38
      Top             =   240
      Width           =   1425
      _Version        =   786432
      _ExtentX        =   2514
      _ExtentY        =   714
      _StockProps     =   79
      Caption         =   "Cerrar"
      BackColor       =   -2147483633
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton cmdActualizar 
      Height          =   405
      Left            =   15960
      TabIndex        =   37
      Top             =   240
      Width           =   1425
      _Version        =   786432
      _ExtentX        =   2514
      _ExtentY        =   714
      _StockProps     =   79
      Caption         =   "Actualizar"
      BackColor       =   -2147483633
      Enabled         =   0   'False
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton cmdGuardar 
      Height          =   405
      Left            =   12960
      TabIndex        =   10
      Top             =   240
      Width           =   1425
      _Version        =   786432
      _ExtentX        =   2514
      _ExtentY        =   714
      _StockProps     =   79
      Caption         =   "Guardar"
      BackColor       =   -2147483633
      Enabled         =   0   'False
      UseVisualStyle  =   -1  'True
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Otros datos del Recibo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7095
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   17805
      Begin XtremeSuiteControls.TabControl TabFacturasRetenciones 
         Height          =   4335
         Left            =   150
         TabIndex        =   28
         Top             =   300
         Width           =   6585
         _Version        =   786432
         _ExtentX        =   11615
         _ExtentY        =   7646
         _StockProps     =   68
         Appearance      =   10
         Color           =   32
         ItemCount       =   2
         Item(0).Caption =   "Comprobantes"
         Item(0).ControlCount=   2
         Item(0).Control(0)=   "gridFacturas"
         Item(0).Control(1)=   "gridFacturasCombo"
         Item(1).Caption =   "Retenciones"
         Item(1).ControlCount=   2
         Item(1).Control(0)=   "gridRetenciones"
         Item(1).Control(1)=   "gridTipoRetenciones"
         Begin GridEX20.GridEX gridFacturasCombo 
            Height          =   3180
            Left            =   5400
            TabIndex        =   30
            Top             =   4365
            Width           =   3210
            _ExtentX        =   5662
            _ExtentY        =   5609
            Version         =   "2.0"
            BoundColumnIndex=   "id"
            ReplaceColumnIndex=   "factura"
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
            ColumnsCount    =   3
            Column(1)       =   "frmAdminCobranzasNuevoRecibo.frx":000C
            Column(2)       =   "frmAdminCobranzasNuevoRecibo.frx":0130
            Column(3)       =   "frmAdminCobranzasNuevoRecibo.frx":0224
            FormatStylesCount=   6
            FormatStyle(1)  =   "frmAdminCobranzasNuevoRecibo.frx":0338
            FormatStyle(2)  =   "frmAdminCobranzasNuevoRecibo.frx":0470
            FormatStyle(3)  =   "frmAdminCobranzasNuevoRecibo.frx":0520
            FormatStyle(4)  =   "frmAdminCobranzasNuevoRecibo.frx":05D4
            FormatStyle(5)  =   "frmAdminCobranzasNuevoRecibo.frx":06AC
            FormatStyle(6)  =   "frmAdminCobranzasNuevoRecibo.frx":0764
            ImageCount      =   0
            PrinterProperties=   "frmAdminCobranzasNuevoRecibo.frx":0844
         End
         Begin GridEX20.GridEX gridTipoRetenciones 
            Height          =   2175
            Left            =   -62725
            TabIndex        =   32
            Top             =   4365
            Visible         =   0   'False
            Width           =   4110
            _ExtentX        =   7250
            _ExtentY        =   3836
            Version         =   "2.0"
            BoundColumnIndex=   "id"
            ReplaceColumnIndex=   "retencion"
            ActAsDropDown   =   -1  'True
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
            ColumnsCount    =   4
            Column(1)       =   "frmAdminCobranzasNuevoRecibo.frx":0A1C
            Column(2)       =   "frmAdminCobranzasNuevoRecibo.frx":0B3C
            Column(3)       =   "frmAdminCobranzasNuevoRecibo.frx":0C3C
            Column(4)       =   "frmAdminCobranzasNuevoRecibo.frx":0D30
            FormatStylesCount=   6
            FormatStyle(1)  =   "frmAdminCobranzasNuevoRecibo.frx":0E34
            FormatStyle(2)  =   "frmAdminCobranzasNuevoRecibo.frx":0F6C
            FormatStyle(3)  =   "frmAdminCobranzasNuevoRecibo.frx":101C
            FormatStyle(4)  =   "frmAdminCobranzasNuevoRecibo.frx":10D0
            FormatStyle(5)  =   "frmAdminCobranzasNuevoRecibo.frx":11A8
            FormatStyle(6)  =   "frmAdminCobranzasNuevoRecibo.frx":1260
            ImageCount      =   0
            PrinterProperties=   "frmAdminCobranzasNuevoRecibo.frx":1340
         End
         Begin GridEX20.GridEX gridFacturas 
            Height          =   3870
            Left            =   135
            TabIndex        =   29
            Top             =   345
            Width           =   6375
            _ExtentX        =   11245
            _ExtentY        =   6826
            Version         =   "2.0"
            BoundColumnIndex=   ""
            ReplaceColumnIndex=   ""
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
            ColumnsCount    =   4
            Column(1)       =   "frmAdminCobranzasNuevoRecibo.frx":1518
            Column(2)       =   "frmAdminCobranzasNuevoRecibo.frx":1660
            Column(3)       =   "frmAdminCobranzasNuevoRecibo.frx":17A8
            Column(4)       =   "frmAdminCobranzasNuevoRecibo.frx":1900
            FormatStylesCount=   6
            FormatStyle(1)  =   "frmAdminCobranzasNuevoRecibo.frx":1A30
            FormatStyle(2)  =   "frmAdminCobranzasNuevoRecibo.frx":1B68
            FormatStyle(3)  =   "frmAdminCobranzasNuevoRecibo.frx":1C18
            FormatStyle(4)  =   "frmAdminCobranzasNuevoRecibo.frx":1CCC
            FormatStyle(5)  =   "frmAdminCobranzasNuevoRecibo.frx":1DA4
            FormatStyle(6)  =   "frmAdminCobranzasNuevoRecibo.frx":1E5C
            ImageCount      =   0
            PrinterProperties=   "frmAdminCobranzasNuevoRecibo.frx":1F3C
         End
         Begin GridEX20.GridEX gridRetenciones 
            Height          =   3870
            Left            =   -69865
            TabIndex        =   31
            Top             =   345
            Visible         =   0   'False
            Width           =   6390
            _ExtentX        =   11271
            _ExtentY        =   6826
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
            ColumnsCount    =   4
            Column(1)       =   "frmAdminCobranzasNuevoRecibo.frx":2114
            Column(2)       =   "frmAdminCobranzasNuevoRecibo.frx":224C
            Column(3)       =   "frmAdminCobranzasNuevoRecibo.frx":2380
            Column(4)       =   "frmAdminCobranzasNuevoRecibo.frx":24BC
            FormatStylesCount=   6
            FormatStyle(1)  =   "frmAdminCobranzasNuevoRecibo.frx":25C8
            FormatStyle(2)  =   "frmAdminCobranzasNuevoRecibo.frx":2700
            FormatStyle(3)  =   "frmAdminCobranzasNuevoRecibo.frx":27B0
            FormatStyle(4)  =   "frmAdminCobranzasNuevoRecibo.frx":2864
            FormatStyle(5)  =   "frmAdminCobranzasNuevoRecibo.frx":293C
            FormatStyle(6)  =   "frmAdminCobranzasNuevoRecibo.frx":29F4
            ImageCount      =   0
            PrinterProperties=   "frmAdminCobranzasNuevoRecibo.frx":2AD4
         End
      End
      Begin VB.TextBox txtRedondeo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1275
         TabIndex        =   24
         Text            =   "0"
         Top             =   6480
         Width           =   900
      End
      Begin VB.ComboBox cboMonedas 
         Height          =   315
         Left            =   4200
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   4755
         Width           =   855
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Valores"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6705
         Left            =   7440
         TabIndex        =   7
         Top             =   240
         Width           =   10560
         Begin XtremeSuiteControls.GroupBox grpCheques 
            Height          =   2625
            Left            =   165
            TabIndex        =   8
            Top             =   3825
            Width           =   10290
            _Version        =   786432
            _ExtentX        =   18150
            _ExtentY        =   4630
            _StockProps     =   79
            Caption         =   "Cheques Recibidos"
            UseVisualStyle  =   -1  'True
            Begin GridEX20.GridEX gridCheques 
               Height          =   2280
               Left            =   75
               TabIndex        =   9
               Top             =   225
               Width           =   10020
               _ExtentX        =   17674
               _ExtentY        =   4022
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
               ColumnsCount    =   7
               Column(1)       =   "frmAdminCobranzasNuevoRecibo.frx":2CAC
               Column(2)       =   "frmAdminCobranzasNuevoRecibo.frx":2E38
               Column(3)       =   "frmAdminCobranzasNuevoRecibo.frx":2FA0
               Column(4)       =   "frmAdminCobranzasNuevoRecibo.frx":3108
               Column(5)       =   "frmAdminCobranzasNuevoRecibo.frx":3240
               Column(6)       =   "frmAdminCobranzasNuevoRecibo.frx":3384
               Column(7)       =   "frmAdminCobranzasNuevoRecibo.frx":34EC
               FormatStylesCount=   6
               FormatStyle(1)  =   "frmAdminCobranzasNuevoRecibo.frx":35E8
               FormatStyle(2)  =   "frmAdminCobranzasNuevoRecibo.frx":3720
               FormatStyle(3)  =   "frmAdminCobranzasNuevoRecibo.frx":37D0
               FormatStyle(4)  =   "frmAdminCobranzasNuevoRecibo.frx":3884
               FormatStyle(5)  =   "frmAdminCobranzasNuevoRecibo.frx":395C
               FormatStyle(6)  =   "frmAdminCobranzasNuevoRecibo.frx":3A14
               ImageCount      =   0
               PrinterProperties=   "frmAdminCobranzasNuevoRecibo.frx":3AF4
            End
            Begin GridEX20.GridEX gridBancos 
               Height          =   1845
               Left            =   150
               TabIndex        =   11
               Top             =   2640
               Visible         =   0   'False
               Width           =   3705
               _ExtentX        =   6535
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
               Column(1)       =   "frmAdminCobranzasNuevoRecibo.frx":3CCC
               Column(2)       =   "frmAdminCobranzasNuevoRecibo.frx":3DCC
               FormatStylesCount=   6
               FormatStyle(1)  =   "frmAdminCobranzasNuevoRecibo.frx":3EC0
               FormatStyle(2)  =   "frmAdminCobranzasNuevoRecibo.frx":3FF8
               FormatStyle(3)  =   "frmAdminCobranzasNuevoRecibo.frx":40A8
               FormatStyle(4)  =   "frmAdminCobranzasNuevoRecibo.frx":415C
               FormatStyle(5)  =   "frmAdminCobranzasNuevoRecibo.frx":4234
               FormatStyle(6)  =   "frmAdminCobranzasNuevoRecibo.frx":42EC
               ImageCount      =   0
               PrinterProperties=   "frmAdminCobranzasNuevoRecibo.frx":43CC
            End
         End
         Begin XtremeSuiteControls.GroupBox grpBanco 
            Height          =   1920
            Left            =   135
            TabIndex        =   12
            Top             =   1860
            Width           =   10290
            _Version        =   786432
            _ExtentX        =   18150
            _ExtentY        =   3387
            _StockProps     =   79
            Caption         =   "Banco"
            UseVisualStyle  =   -1  'True
            Begin GridEX20.GridEX gridDepositosOperaciones 
               Height          =   1545
               Left            =   90
               TabIndex        =   13
               Top             =   225
               Width           =   10065
               _ExtentX        =   17754
               _ExtentY        =   2725
               Version         =   "2.0"
               BoundColumnIndex=   ""
               ReplaceColumnIndex=   ""
               ColumnAutoResize=   -1  'True
               MethodHoldFields=   -1  'True
               ContScroll      =   -1  'True
               AllowDelete     =   -1  'True
               GroupByBoxVisible=   0   'False
               RowHeaders      =   -1  'True
               ItemCount       =   3
               DataMode        =   99
               AllowAddNew     =   -1  'True
               ColumnHeaderHeight=   285
               IntProp1        =   0
               IntProp2        =   0
               IntProp7        =   0
               ColumnsCount    =   4
               Column(1)       =   "frmAdminCobranzasNuevoRecibo.frx":45A4
               Column(2)       =   "frmAdminCobranzasNuevoRecibo.frx":4730
               Column(3)       =   "frmAdminCobranzasNuevoRecibo.frx":48C0
               Column(4)       =   "frmAdminCobranzasNuevoRecibo.frx":4A48
               FormatStylesCount=   6
               FormatStyle(1)  =   "frmAdminCobranzasNuevoRecibo.frx":4BB8
               FormatStyle(2)  =   "frmAdminCobranzasNuevoRecibo.frx":4CF0
               FormatStyle(3)  =   "frmAdminCobranzasNuevoRecibo.frx":4DA0
               FormatStyle(4)  =   "frmAdminCobranzasNuevoRecibo.frx":4E54
               FormatStyle(5)  =   "frmAdminCobranzasNuevoRecibo.frx":4F2C
               FormatStyle(6)  =   "frmAdminCobranzasNuevoRecibo.frx":4FE4
               ImageCount      =   0
               PrinterProperties=   "frmAdminCobranzasNuevoRecibo.frx":50C4
            End
         End
         Begin XtremeSuiteControls.GroupBox grpCaja 
            Height          =   1635
            Left            =   135
            TabIndex        =   14
            Top             =   180
            Width           =   10290
            _Version        =   786432
            _ExtentX        =   18150
            _ExtentY        =   2884
            _StockProps     =   79
            Caption         =   "Caja"
            UseVisualStyle  =   -1  'True
            Begin GridEX20.GridEX gridCajaOperaciones 
               Height          =   1260
               Left            =   90
               TabIndex        =   15
               Top             =   225
               Width           =   10050
               _ExtentX        =   17727
               _ExtentY        =   2223
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
               ColumnsCount    =   4
               Column(1)       =   "frmAdminCobranzasNuevoRecibo.frx":529C
               Column(2)       =   "frmAdminCobranzasNuevoRecibo.frx":5428
               Column(3)       =   "frmAdminCobranzasNuevoRecibo.frx":5590
               Column(4)       =   "frmAdminCobranzasNuevoRecibo.frx":5718
               FormatStylesCount=   6
               FormatStyle(1)  =   "frmAdminCobranzasNuevoRecibo.frx":584C
               FormatStyle(2)  =   "frmAdminCobranzasNuevoRecibo.frx":5984
               FormatStyle(3)  =   "frmAdminCobranzasNuevoRecibo.frx":5A34
               FormatStyle(4)  =   "frmAdminCobranzasNuevoRecibo.frx":5AE8
               FormatStyle(5)  =   "frmAdminCobranzasNuevoRecibo.frx":5BC0
               FormatStyle(6)  =   "frmAdminCobranzasNuevoRecibo.frx":5C78
               ImageCount      =   0
               PrinterProperties=   "frmAdminCobranzasNuevoRecibo.frx":5D58
            End
         End
      End
      Begin VB.Label lblTotalCTACTE 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Total Cta Cte:"
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
         Left            =   3000
         TabIndex        =   33
         Top             =   6045
         Width           =   1200
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFDBBF&
         DrawMode        =   9  'Not Mask Pen
         X1              =   7710
         X2              =   135
         Y1              =   6360
         Y2              =   6360
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFDBBF&
         DrawMode        =   9  'Not Mask Pen
         X1              =   7710
         X2              =   120
         Y1              =   5355
         Y2              =   5355
      End
      Begin VB.Label lblDiferencia 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Diferencia entre Recibido y Recibo:"
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
         Left            =   3000
         TabIndex        =   27
         Tag             =   "Diferencia entre Recibido y Recibo: "
         Top             =   6495
         Width           =   3060
      End
      Begin VB.Label lblTotalRecibido 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Total Recibido:"
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
         Left            =   3000
         TabIndex        =   26
         Top             =   5460
         Width           =   1320
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Redondeo:"
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
         Left            =   270
         TabIndex        =   25
         Top             =   6510
         Width           =   945
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Total Recibo:"
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
         Left            =   3000
         TabIndex        =   23
         Top             =   4815
         Width           =   1170
      End
      Begin VB.Label lblTotalRecibo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5175
         TabIndex        =   22
         Top             =   4770
         Width           =   1410
      End
      Begin VB.Label lblTotalCaja 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Total Caja:"
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
         Left            =   255
         TabIndex        =   20
         Top             =   6060
         Width           =   945
      End
      Begin VB.Label lblTotalBanco 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Total Banco:"
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
         Left            =   255
         TabIndex        =   19
         Top             =   5760
         Width           =   1110
      End
      Begin VB.Label lblTotalCheques 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Total Cheques:"
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
         TabIndex        =   18
         Top             =   5460
         Width           =   1305
      End
      Begin VB.Label lblTotalFactura 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Total Cbtes:"
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
         Left            =   270
         TabIndex        =   17
         Top             =   4815
         Width           =   1050
      End
      Begin VB.Label lblTotalRetenciones 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Total Retenciones:"
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
         Left            =   270
         TabIndex        =   16
         Top             =   5055
         Width           =   1635
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos del Recibo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   9735
      Begin VB.ComboBox cboClientes 
         Height          =   315
         Left            =   2550
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   255
         Width           =   4260
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   300
         Left            =   8085
         TabIndex        =   4
         Top             =   270
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         Format          =   62521345
         CurrentDate     =   39199
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   5
         Top             =   315
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7245
         TabIndex        =   3
         Top             =   315
         Width           =   735
      End
      Begin VB.Label lblNumeroRecibo 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Número "
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
         Left            =   270
         TabIndex        =   2
         Top             =   315
         Width           =   720
      End
   End
   Begin GridEX20.GridEX gridCajas 
      Height          =   975
      Left            =   120
      TabIndex        =   34
      Top             =   8640
      Visible         =   0   'False
      Width           =   6180
      _ExtentX        =   10901
      _ExtentY        =   1720
      Version         =   "2.0"
      BoundColumnIndex=   "id"
      ReplaceColumnIndex=   "caja"
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
      Column(1)       =   "frmAdminCobranzasNuevoRecibo.frx":5F30
      Column(2)       =   "frmAdminCobranzasNuevoRecibo.frx":6054
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminCobranzasNuevoRecibo.frx":6140
      FormatStyle(2)  =   "frmAdminCobranzasNuevoRecibo.frx":6278
      FormatStyle(3)  =   "frmAdminCobranzasNuevoRecibo.frx":6328
      FormatStyle(4)  =   "frmAdminCobranzasNuevoRecibo.frx":63DC
      FormatStyle(5)  =   "frmAdminCobranzasNuevoRecibo.frx":64B4
      FormatStyle(6)  =   "frmAdminCobranzasNuevoRecibo.frx":656C
      ImageCount      =   0
      PrinterProperties=   "frmAdminCobranzasNuevoRecibo.frx":664C
   End
   Begin GridEX20.GridEX gridCuentasBancarias 
      Height          =   975
      Left            =   120
      TabIndex        =   35
      Top             =   9720
      Visible         =   0   'False
      Width           =   6225
      _ExtentX        =   10980
      _ExtentY        =   1720
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
      Column(1)       =   "frmAdminCobranzasNuevoRecibo.frx":6824
      Column(2)       =   "frmAdminCobranzasNuevoRecibo.frx":6948
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminCobranzasNuevoRecibo.frx":6A3C
      FormatStyle(2)  =   "frmAdminCobranzasNuevoRecibo.frx":6B74
      FormatStyle(3)  =   "frmAdminCobranzasNuevoRecibo.frx":6C24
      FormatStyle(4)  =   "frmAdminCobranzasNuevoRecibo.frx":6CD8
      FormatStyle(5)  =   "frmAdminCobranzasNuevoRecibo.frx":6DB0
      FormatStyle(6)  =   "frmAdminCobranzasNuevoRecibo.frx":6E68
      ImageCount      =   0
      PrinterProperties=   "frmAdminCobranzasNuevoRecibo.frx":6F48
   End
   Begin GridEX20.GridEX gridMonedas 
      Height          =   1215
      Left            =   6600
      TabIndex        =   36
      Top             =   8640
      Visible         =   0   'False
      Width           =   4020
      _ExtentX        =   7091
      _ExtentY        =   2143
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
      Column(1)       =   "frmAdminCobranzasNuevoRecibo.frx":7120
      Column(2)       =   "frmAdminCobranzasNuevoRecibo.frx":7244
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminCobranzasNuevoRecibo.frx":7338
      FormatStyle(2)  =   "frmAdminCobranzasNuevoRecibo.frx":7470
      FormatStyle(3)  =   "frmAdminCobranzasNuevoRecibo.frx":7520
      FormatStyle(4)  =   "frmAdminCobranzasNuevoRecibo.frx":75D4
      FormatStyle(5)  =   "frmAdminCobranzasNuevoRecibo.frx":76AC
      FormatStyle(6)  =   "frmAdminCobranzasNuevoRecibo.frx":7764
      ImageCount      =   0
      PrinterProperties=   "frmAdminCobranzasNuevoRecibo.frx":7844
   End
End
Attribute VB_Name = "frmAdminCobranzasNuevoRecibo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private dataLoaded As Boolean
Private clienteChange As Boolean
Private cheque As cheque
Private recibo As recibo
Private retencionRecibo As retencionRecibo
Private retenciones As New Collection
Private Retencion As Retencion
Private facturasCliente As New Collection
Private Factura As Factura
Private bancos As New Collection
Private Banco As Banco
Private CuentaBancaria As CuentaBancaria
Private cuentasBancarias As New Collection
Private Monedas As New Collection
Private moneda As clsMoneda
Private operacion As operacion
Private Cajas As New Collection
Private caja As caja
'Private RecibosACuenta As New Collection

Private Editar_ As Boolean

Public Property Let editar(nvalue As Boolean)
    Editar_ = nvalue
End Property


Public Property Let reciboId(nIdRecibo As Long)
    Set recibo = DAORecibo.FindById(nIdRecibo, True, True, True, True, True)

    If recibo Is Nothing Then
        MsgBox "recibo no encontrado, cierre pantalla", vbCritical
    End If


    Me.dtpFecha.value = recibo.FEcha
    lblNumeroRecibo.caption = "Número: " & recibo.Id
    Me.txtRedondeo.text = recibo.Redondeo

    Set Cajas = DAOCaja.FindAll()
    Me.gridCajas.ItemCount = Cajas.count

    Set Monedas = DAOMoneda.GetAll()
    Me.gridMonedas.ItemCount = Monedas.count

    Set cuentasBancarias = DAOCuentaBancaria.FindAll()
    Me.gridCuentasBancarias.ItemCount = cuentasBancarias.count

    Set bancos = DAOBancos.GetAll()
    Me.gridBancos.ItemCount = bancos.count

    'Me.chkACuenta.value = CInt(recibo.PagoACuenta) * -1

    Set gridRetenciones.Columns("tipo").DropDownControl = Me.gridTipoRetenciones
    Set Me.gridFacturas.Columns("factura").DropDownControl = Me.gridFacturasCombo
    Set Me.gridCheques.Columns("banco").DropDownControl = Me.gridBancos
    Set Me.gridCheques.Columns("moneda").DropDownControl = Me.gridMonedas

    Set Me.gridDepositosOperaciones.Columns("moneda").DropDownControl = Me.gridMonedas
    Set Me.gridDepositosOperaciones.Columns("cuenta").DropDownControl = Me.gridCuentasBancarias

    Set Me.gridCajaOperaciones.Columns("caja").DropDownControl = Me.gridCajas
    Set Me.gridCajaOperaciones.Columns("moneda").DropDownControl = Me.gridMonedas

    Set retenciones = DAORetenciones.FindAll()
    Me.gridTipoRetenciones.ItemCount = retenciones.count

    Me.gridRetenciones.ItemCount = recibo.retenciones.count
    Me.gridFacturas.ItemCount = recibo.facturas.count

    DAOMoneda.LlenarCombo Me.cboMonedas
    Me.cboMonedas.ListIndex = PosIndexCbo(recibo.moneda.Id, Me.cboMonedas)
    DAOCliente.LlenarCombo Me.cboClientes, True, True
    Me.cboClientes.ListIndex = PosIndexCbo(recibo.cliente.Id, Me.cboClientes)

    Me.gridCajaOperaciones.ItemCount = recibo.OperacionesCaja.count
    Me.gridDepositosOperaciones.ItemCount = recibo.OperacionesBanco.count
    Me.gridCheques.ItemCount = recibo.Cheques.count

    Totalizar

    'Me.gridFacturas.Enabled = Not recibo.PagoACuenta
    'Me.gridRetenciones.Enabled = Not recibo.PagoACuenta

    CargarFacturasCliente    'para el combo
    Me.cboMonedas.Enabled = Editar_
    Me.txtRedondeo.Enabled = Editar_
    Me.gridRetenciones.AllowEdit = Editar_
    Me.gridFacturas.AllowEdit = Editar_
    Me.gridCajaOperaciones.AllowEdit = Editar_
    Me.gridBancos.AllowEdit = Editar_
    Me.gridCheques.AllowEdit = Editar_
    Me.gridDepositosOperaciones.AllowEdit = Editar_
    Me.gridRetenciones.AllowDelete = Editar_
    Me.gridFacturas.AllowDelete = Editar_
    Me.gridCajaOperaciones.AllowDelete = Editar_
    Me.gridCheques.AllowDelete = Editar_
    Me.gridDepositosOperaciones.AllowDelete = Editar_
    Me.gridRetenciones.AllowAddNew = Editar_
    Me.gridFacturas.AllowAddNew = Editar_
    Me.gridCajaOperaciones.AllowAddNew = Editar_
    Me.gridCheques.AllowAddNew = Editar_
    gridDepositosOperaciones.AllowAddNew = Editar_
    Me.Frame1.Enabled = Editar_

    '    Me.cmdGuardar.Enabled = Editar_
    '    Me.cmdActualizar.Enabled = Editar_

    dataLoaded = True

End Property

Private Sub CargarFacturasCliente()
    If Me.cboClientes.ListIndex = -1 Then
        Set facturasCliente = New Collection
    Else
        Set facturasCliente = DAOFactura.FindAllNoSaldadasTotalByCliente(Me.cboClientes.ItemData(Me.cboClientes.ListIndex), True)
    End If
    Me.gridFacturasCombo.ItemCount = 0
    Me.gridFacturasCombo.ItemCount = facturasCliente.count
End Sub

Private Sub cboClientes_Click()
    If clienteChange Then Exit Sub

    If Me.cboClientes.ListIndex <> -1 Then
        If dataLoaded Then
            If recibo.facturas.count > 0 Then
                If MsgBox("Va a cambiar de cliente y perder las facturas." & vbNewLine & "¿Desea continuar?", vbQuestion + vbYesNo) = vbNo Then
                    clienteChange = True
                    Me.cboClientes.ListIndex = funciones.PosIndexCbo(recibo.cliente.Id, Me.cboClientes)
                    clienteChange = False
                    Exit Sub

                End If
            End If
            VaciarFacturasRetenciones
            CargarFacturasCliente
            Set recibo.cliente = DAOCliente.BuscarPorID(Me.cboClientes.ItemData(Me.cboClientes.ListIndex))
            'Else
            '    VaciarFacturasRetenciones
            '    CargarFacturasCliente
        End If
    End If

    clienteChange = False
End Sub



Private Sub cboMonedas_Click()
    If Me.cboMonedas.ListIndex <> -1 And dataLoaded Then
        Set recibo.moneda = DAOMoneda.GetById(Me.cboMonedas.ItemData(Me.cboMonedas.ListIndex))
        Totalizar
    End If
End Sub


Private Sub VaciarFacturasRetenciones(Optional ByVal cleanRetenciones As Boolean = True)
    Set recibo.facturas = New Collection
    Me.gridFacturas.ItemCount = 0

    If cleanRetenciones Then
        Set recibo.retenciones = New Collection
        Me.gridRetenciones.ItemCount = 0
    End If
End Sub

Private Sub cmdActualizar_Click()
    If Not recibo.IsValid Then
        MsgBox recibo.ValidationMessages, vbExclamation
        Exit Sub
    End If

    If DAORecibo.Save(recibo) And DAORecibo.aprobar(recibo) Then

        MsgBox "Aprobación actualizada!", vbInformation, "Información"

        Unload Me
    Else
        MsgBox "Hubo un error al intentar actualizar el recibo.", vbCritical

    End If

    'Dim idRecibo As Long
    '   'If MsgBox("¿Está seguro de aprobar este recibo?", vbYesNo, "Confirmación") = vbYes Then
    '
    '        Set recibo = DAORecibo.FindById(recibo.id, True, True, True, True, True)
    '
    '        If DAORecibo.aprobar(recibo) Then
    '            MsgBox "Aprobación actualizada!", vbInformation, "Información"
    '        Else
    '            MsgBox "Error, no se pudo actualizar el recibo nuevamente!", vbCritical, "Error"
    '        End If
    'End If

End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub cmdGuardar_Click()
    If Not recibo.IsValid Then
        MsgBox recibo.ValidationMessages, vbExclamation
        Exit Sub
    End If

    If DAORecibo.Save(recibo) Then
        MsgBox "Recibo guardado.", vbInformation
        Unload Me
    Else
        MsgBox "Hubo un error al intentar guardar el recibo.", vbCritical
    End If


End Sub


Private Sub Totalizar()

    Me.lblTotalFactura.caption = "Total Facturas: " & funciones.FormatearDecimales(recibo.TotalFacturas)
    Me.lblTotalRetenciones.caption = "Total Retenciones: " & funciones.FormatearDecimales(recibo.TotalRetenciones)
    Me.lblTotalCheques.caption = "Total Cheques: " & funciones.FormatearDecimales(recibo.TotalCheques)
    Me.lblTotalBanco.caption = "Total Banco: " & funciones.FormatearDecimales(recibo.TotalOperacionesBanco)
    Me.lblTotalCaja.caption = "Total Caja: " & funciones.FormatearDecimales(recibo.TotalOperacionesCaja)

    Dim totalRecibo As Double
    totalRecibo = funciones.FormatearDecimales(recibo.total)
    Dim totalCancelado As Double
    totalCancelado = funciones.FormatearDecimales(recibo.TotalRecibido)

    Me.lblTotalRecibo.caption = funciones.FormatearDecimales(totalRecibo)
    Me.lblTotalRecibido.caption = "Total Recibido: " & funciones.FormatearDecimales(totalCancelado)

    If totalCancelado < totalRecibo Then
        lblTotalRecibo.backColor = vbRed
    ElseIf totalCancelado = totalRecibo Then
        lblTotalRecibo.backColor = vbYellow
    ElseIf totalCancelado > totalRecibo Then
        lblTotalRecibo.backColor = vbGreen
    End If

    Me.lblDiferencia.caption = Me.lblDiferencia.Tag & funciones.FormatearDecimales(totalCancelado - MonedaConverter.Convertir(totalRecibo, recibo.moneda.Id, DAOMoneda.MONEDA_PESO_ID))
    'Debug.Print MonedaConverter.Convertir(totalRecibo, recibo.moneda.id, DAOMoneda.MONEDA_PESO_ID)
    recibo.ACuenta = (totalCancelado - totalRecibo)
End Sub


'Private Sub Command1_Click()
'    Dim ret As retencionRecibo
'    Dim colrec As New Collection
'    Dim recibo As recibo
'    Set colrec = DAORecibo.FindAll
'    For Each recibo In colrec
'
'    Set recibo.Retenciones = DAOReciboRetencion.FindAllByRecibo(recibo.Id)
'        For Each ret In recibo.Retenciones
'            conectar.execute "update AdminRecibosDetalleRetenciones set fecha=" & conectar.Escape(recibo.Fecha) & " where id=" & ret.Id
'        Next ret
'
'    Next recibo
'End Sub

Private Sub dtpFecha_LostFocus()
    recibo.FEcha = Me.dtpFecha.value
End Sub

Private Sub Form_Load()
    dataLoaded = False

    FormHelper.Customize Me

    GridEXHelper.CustomizeGrid Me.gridRetenciones, False, True
    GridEXHelper.CustomizeGrid Me.gridTipoRetenciones, False, False

    GridEXHelper.CustomizeGrid Me.gridFacturas, False, True
    GridEXHelper.CustomizeGrid Me.gridFacturasCombo, False, False

    GridEXHelper.CustomizeGrid Me.gridCheques, False, True
    GridEXHelper.CustomizeGrid Me.gridBancos, False, False

    GridEXHelper.CustomizeGrid Me.gridCuentasBancarias, False, False
    GridEXHelper.CustomizeGrid Me.gridMonedas, False, False
    GridEXHelper.CustomizeGrid Me.gridCajas, False, False

    GridEXHelper.CustomizeGrid Me.gridDepositosOperaciones, False, True
    GridEXHelper.CustomizeGrid Me.gridCajaOperaciones, False, True

End Sub

'Private Sub VerRecibosConSaldoACuenta()
'
'    Set RecibosACuenta = New Collection
'
'    Set RecibosACuenta = DAORecibo.FindAll("(a_cuenta-a_cuenta_usado) >0.02")
''
'End Sub

Private Sub gridBancos_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If rowIndex <= bancos.count Then
        Set Banco = bancos.item(rowIndex)
        Values(1) = Banco.Id
        Values(2) = Banco.nombre
    End If

End Sub

Private Sub gridCajaOperaciones_BeforeUpdate(ByVal Cancel As GridEX20.JSRetBoolean)
    Cancel = _
    Not IsNumeric(Me.gridCajaOperaciones.value(1)) Or _
             (Not IsNumeric(Me.gridCajaOperaciones.value(2)) And LenB(Me.gridCajaOperaciones.value(2)) = 0) Or _
             Not IsDate(Me.gridCajaOperaciones.value(3)) Or _
             (Not IsNumeric(Me.gridCajaOperaciones.value(4)) And LenB(Me.gridCajaOperaciones.value(4)) = 0)
End Sub

Private Sub gridCajaOperaciones_UnboundAddNew(ByVal NewRowBookmark As GridEX20.JSRetVariant, ByVal Values As GridEX20.JSRowData)
    Set operacion = New operacion
    operacion.IdPertenencia = recibo.Id
    operacion.Pertenencia = OrigenOperacion.caja
    operacion.Monto = Values(1)
    If IsNumeric(Values(2)) Then
        Set operacion.moneda = DAOMoneda.GetById(Values(2))
    End If
    operacion.FechaOperacion = Values(3)
    If IsNumeric(Values(4)) Then
        Set operacion.caja = DAOCaja.FindById(Values(4))
    End If
    operacion.EntradaSalida = OPEntrada
    recibo.OperacionesCaja.Add operacion

    Totalizar
End Sub

Private Sub gridCajaOperaciones_UnboundDelete(ByVal rowIndex As Long, ByVal Bookmark As Variant)
    If rowIndex > 0 And recibo.OperacionesCaja.count >= rowIndex Then
        recibo.OperacionesCaja.remove rowIndex
        Totalizar
    End If
End Sub

Private Sub gridCajaOperaciones_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If rowIndex <= recibo.OperacionesCaja.count Then
        Set operacion = recibo.OperacionesCaja.item(rowIndex)
        Values(1) = funciones.FormatearDecimales(operacion.Monto)
        If IsSomething(operacion.moneda) Then
            Values(2) = operacion.moneda.NombreCorto
        End If
        Values(3) = operacion.FechaOperacion
        If IsSomething(operacion.caja) Then
            Values(4) = operacion.caja.nombre
        End If
        Totalizar
    End If
End Sub

Private Sub gridCajaOperaciones_UnboundUpdate(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If rowIndex > 0 And recibo.OperacionesCaja.count > 0 Then
        Set operacion = recibo.OperacionesCaja.item(rowIndex)
        'operacion.IdPertenencia = recibo.id
        'operacion.Pertenencia = Banco
        operacion.Monto = Values(1)
        If IsNumeric(Values(2)) Then
            Set operacion.moneda = DAOMoneda.GetById(Values(2))
        End If
        operacion.FechaOperacion = Values(3)
        If IsNumeric(Values(4)) Then
            Set operacion.caja = DAOCaja.FindById(Values(4))
        End If
        operacion.EntradaSalida = OPEntrada
        Totalizar
    End If
End Sub

Private Sub gridCajas_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If rowIndex > 0 And Cajas.count > 0 Then
        Set caja = Cajas.item(rowIndex)
        Values(1) = caja.Id
        Values(2) = caja.nombre
    End If
End Sub

Private Sub gridCheques_BeforeUpdate(ByVal Cancel As GridEX20.JSRetBoolean)
    If Me.gridCheques.row = -1 Then    'es nuevo
        If Val(Me.gridCheques.value(6)) = -1 Then
            Me.gridCheques.value(7) = recibo.cliente.razon
        End If
    End If

    Cancel = _
    LenB(Me.gridCheques.value(1)) = 0 Or _
             Not IsNumeric(Me.gridCheques.value(2)) Or _
             (Not IsNumeric(Me.gridCheques.value(3)) And LenB(Me.gridCheques.value(3)) = 0) Or _
             (Not IsNumeric(Me.gridCheques.value(4)) And LenB(Me.gridCheques.value(4)) = 0) Or _
             Not IsDate(Me.gridCheques.value(5))
End Sub

Private Sub gridCheques_UnboundAddNew(ByVal NewRowBookmark As GridEX20.JSRetVariant, ByVal Values As GridEX20.JSRowData)
    Set cheque = New cheque
    If IsNumeric(Values(4)) Then Set cheque.Banco = DAOBancos.GetById(Values(4))
    cheque.EnCartera = True
    cheque.FechaRecibido = recibo.FEcha
    If IsNumeric(Values(3)) Then Set cheque.moneda = DAOMoneda.GetById(Values(3))
    cheque.Monto = funciones.RedondearDecimales(Val(Values(2)))
    cheque.numero = Values(1)
    cheque.Propio = False
    cheque.FechaVencimiento = Values(5)

    cheque.TercerosPropio = (Values(6) = -1)

    If cheque.TercerosPropio Then
        cheque.OrigenDestino = recibo.cliente.razon
    Else
        cheque.OrigenDestino = Values(7)
    End If
    recibo.Cheques.Add cheque

    Totalizar
End Sub

Private Sub gridCheques_UnboundDelete(ByVal rowIndex As Long, ByVal Bookmark As Variant)
    If rowIndex > 0 Then
        recibo.Cheques.remove rowIndex
        Totalizar
    End If
End Sub

Private Sub gridCheques_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If rowIndex <= recibo.Cheques.count Then
        Set cheque = recibo.Cheques.item(rowIndex)
        Values(1) = cheque.numero
        Values(2) = funciones.FormatearDecimales(cheque.Monto)
        Values(6) = cheque.TercerosPropio
        Values(7) = cheque.OrigenDestino
        If IsSomething(cheque.moneda) Then Values(3) = cheque.moneda.NombreCorto
        If IsSomething(cheque.Banco) Then Values(4) = cheque.Banco.nombre
        If CDbl(cheque.FechaVencimiento) > 0 Then Values(5) = cheque.FechaVencimiento
        Totalizar
    End If
End Sub

Private Sub gridCheques_UnboundUpdate(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If rowIndex > 0 And recibo.Cheques.count >= rowIndex Then
        Set cheque = recibo.Cheques.item(rowIndex)
        If IsNumeric(Values(4)) Then Set cheque.Banco = DAOBancos.GetById(Values(4))
        cheque.EnCartera = True
        cheque.FechaRecibido = recibo.FEcha
        If IsNumeric(Values(3)) Then Set cheque.moneda = DAOMoneda.GetById(Values(3))
        cheque.Monto = Val(Values(2))
        cheque.numero = Values(1)
        cheque.Propio = False
        cheque.FechaVencimiento = Values(5)
        cheque.OrigenDestino = Values(7)

        cheque.TercerosPropio = (Values(6) = -1)

        'recibo.Cheques.Add cheque, , , RowIndex
        'recibo.Cheques.Remove RowIndex
        Totalizar
    End If
End Sub



Private Sub gridCuentasBancarias_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If cuentasBancarias.count >= rowIndex Then
        Set CuentaBancaria = cuentasBancarias.item(rowIndex)
        Values(1) = CuentaBancaria.Id
        Values(2) = CuentaBancaria.DescripcionFormateada
    End If
End Sub

Private Sub gridDepositosOperaciones_BeforeUpdate(ByVal Cancel As GridEX20.JSRetBoolean)
    Cancel = _
    Not IsNumeric(Me.gridDepositosOperaciones.value(1)) Or _
             (Not IsNumeric(Me.gridDepositosOperaciones.value(2)) And LenB(Me.gridDepositosOperaciones.value(2)) = 0) Or _
             Not IsDate(Me.gridDepositosOperaciones.value(3)) Or _
             (Not IsNumeric(Me.gridDepositosOperaciones.value(4)) And LenB(Me.gridDepositosOperaciones.value(4)) = 0)
End Sub

Private Sub gridDepositosOperaciones_UnboundAddNew(ByVal NewRowBookmark As GridEX20.JSRetVariant, ByVal Values As GridEX20.JSRowData)
    Set operacion = New operacion
    operacion.IdPertenencia = recibo.Id
    operacion.Pertenencia = OrigenOperacion.Banco
    operacion.Monto = Values(1)
    If IsNumeric(Values(2)) Then
        Set operacion.moneda = DAOMoneda.GetById(Values(2))
    End If
    operacion.FechaOperacion = Values(3)
    If IsNumeric(Values(4)) Then
        Set operacion.CuentaBancaria = DAOCuentaBancaria.FindById(Values(4))
    End If
    operacion.EntradaSalida = OPEntrada


    recibo.OperacionesBanco.Add operacion

    Totalizar
End Sub

Private Sub gridDepositosOperaciones_UnboundDelete(ByVal rowIndex As Long, ByVal Bookmark As Variant)
    If rowIndex > 0 And recibo.OperacionesBanco.count >= rowIndex Then
        recibo.OperacionesBanco.remove rowIndex
        Totalizar
    End If
End Sub

Private Sub gridDepositosOperaciones_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If rowIndex <= recibo.OperacionesBanco.count Then
        Set operacion = recibo.OperacionesBanco.item(rowIndex)
        Values(1) = funciones.FormatearDecimales(operacion.Monto)
        If IsSomething(operacion.moneda) Then
            Values(2) = operacion.moneda.NombreCorto
        End If
        Values(3) = operacion.FechaOperacion
        If IsSomething(operacion.CuentaBancaria) Then
            Values(4) = operacion.CuentaBancaria.DescripcionFormateada
        End If
        Totalizar
    End If
End Sub

Private Sub gridDepositosOperaciones_UnboundUpdate(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If rowIndex > 0 And recibo.OperacionesBanco.count > 0 Then
        Set operacion = recibo.OperacionesBanco.item(rowIndex)
        'operacion.IdPertenencia = recibo.id
        'operacion.Pertenencia = Banco
        operacion.Monto = Values(1)
        If IsNumeric(Values(2)) Then
            Set operacion.moneda = DAOMoneda.GetById(Values(2))
        End If
        operacion.FechaOperacion = Values(3)
        If IsNumeric(Values(4)) Then
            Set operacion.CuentaBancaria = DAOCuentaBancaria.FindById(Values(4))
        End If
        operacion.EntradaSalida = OPEntrada
        Totalizar
    End If
End Sub

Private Sub gridFacturas_BeforeUpdate(ByVal Cancel As GridEX20.JSRetBoolean)
    Cancel = LenB(Me.gridFacturas.value(1)) = 0 Or Val(Me.gridFacturas.value(4)) > Val(Me.gridFacturas.value(3))    ' no selecciono factura
End Sub

Private Sub gridFacturas_UnboundAddNew(ByVal NewRowBookmark As GridEX20.JSRetVariant, ByVal Values As GridEX20.JSRowData)
    If IsNumeric(Values(1)) Then
        Set Factura = DAOFactura.FindById(Values(1), True)
        recibo.facturas.Add Factura

        If recibo.PagosDeFacturas.Exists(CStr(Factura.Id)) Then
            recibo.PagosDeFacturas.remove CStr(Factura.Id)
        End If



        recibo.PagosDeFacturas.Add CStr(Factura.Id), Factura.total - DAOFactura.PagosRealizados(Factura.Id)




        Totalizar
    End If

End Sub

Private Sub gridFacturas_UnboundDelete(ByVal rowIndex As Long, ByVal Bookmark As Variant)
    If rowIndex > 0 And recibo.facturas.count >= rowIndex Then
        recibo.facturas.remove rowIndex
        If recibo.PagosDeFacturas.Exists(CStr(Factura.Id)) Then
            recibo.PagosDeFacturas.remove CStr(Factura.Id)
        End If
        Totalizar
    End If
End Sub

Private Sub gridFacturas_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If rowIndex > 0 And recibo.facturas.count >= rowIndex Then
        Set Factura = recibo.facturas.item(rowIndex)
        Values(1) = Factura.GetShortDescription(False, True)
        Values(2) = funciones.FormatearDecimales(Factura.total)
        Values(3) = funciones.FormatearDecimales(Factura.total - DAOFactura.PagosRealizados(Factura.Id))

        If recibo.PagosDeFacturas.Exists(CStr(Factura.Id)) Then
            Values(4) = funciones.FormatearDecimales(recibo.PagosDeFacturas.item(CStr(Factura.Id)))
        Else
            Values(4) = funciones.FormatearDecimales(0)
        End If

    End If
End Sub

Private Sub gridFacturas_UnboundUpdate(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If rowIndex > 0 And recibo.facturas.count >= rowIndex Then
        Set Factura = recibo.facturas.item(rowIndex)
        'Set Factura = DAOFactura.FindById(Values(1), True)
        'recibo.facturas.Add Factura, , , RowIndex
        'recibo.facturas.Remove RowIndex

        If recibo.PagosDeFacturas.Exists(CStr(Factura.Id)) Then
            recibo.PagosDeFacturas.remove CStr(Factura.Id)
        End If
        recibo.PagosDeFacturas.Add CStr(Factura.Id), Val(Values(4))

        Totalizar    ' no se si hay que totalizar
    End If
End Sub

Private Sub gridFacturasCombo_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If rowIndex > 0 And facturasCliente.count >= rowIndex Then
        Set Factura = facturasCliente.item(rowIndex)
        Values(1) = Factura.Id
        Values(2) = Factura.GetShortDescription(False, True)
        Values(3) = funciones.FormatearDecimales(Factura.total)
    End If
End Sub

Private Sub gridMonedas_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If rowIndex > 0 And Monedas.count > 0 Then
        Set moneda = Monedas.item(rowIndex)
        Values(1) = moneda.Id
        Values(2) = moneda.NombreCorto
    End If
End Sub

Private Sub gridRetenciones_BeforeUpdate(ByVal Cancel As GridEX20.JSRetBoolean)
    Cancel = LenB(Me.gridRetenciones.value(1)) = 0 Or _
             Not IsNumeric(Me.gridRetenciones.value(2)) Or _
             Not IsNumeric(Me.gridRetenciones.value(3))
End Sub

Private Sub gridRetenciones_UnboundDelete(ByVal rowIndex As Long, ByVal Bookmark As Variant)
    If rowIndex > 0 Then
        recibo.retenciones.remove rowIndex
        Totalizar
    End If
End Sub

Private Sub gridRetenciones_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If rowIndex > 0 And rowIndex <= recibo.retenciones.count Then

        Set retencionRecibo = recibo.retenciones.item(rowIndex)
        Values(1) = retencionRecibo.NroRetencion
        If IsSomething(retencionRecibo.Retencion) Then
            Values(2) = retencionRecibo.Retencion.nombre
        End If
        Values(3) = funciones.FormatearDecimales(retencionRecibo.Valor)
        Values(4) = Format(retencionRecibo.FEcha, "dd-mm-yyyy")

    End If
End Sub

Private Sub UpdateAddRetencion(ByRef retencionRecibo As retencionRecibo, Values As GridEX20.JSRowData)
    retencionRecibo.idRecibo = recibo.Id
    retencionRecibo.NroRetencion = Values(1)
    retencionRecibo.Valor = Val(Values(3))

    If LenB(Values(4)) = 0 Or Not IsDate(Values(4)) Then
        retencionRecibo.FEcha = Me.dtpFecha.value    'Now
    Else
        retencionRecibo.FEcha = CDate(Values(4))
    End If

    If Not IsEmpty(Values(2)) And IsNumeric(Values(2)) Then
        Set retencionRecibo.Retencion = retenciones.item(CStr(Values(2)))    'DAORetenciones.
    End If
End Sub

Private Sub gridRetenciones_UnboundAddNew(ByVal NewRowBookmark As GridEX20.JSRetVariant, ByVal Values As GridEX20.JSRowData)

    Set retencionRecibo = New retencionRecibo
    UpdateAddRetencion retencionRecibo, Values
    recibo.retenciones.Add retencionRecibo
    Totalizar
End Sub

Private Sub gridRetenciones_UnboundUpdate(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If rowIndex > 0 Then
        Set retencionRecibo = recibo.retenciones(rowIndex)
        UpdateAddRetencion retencionRecibo, Values
        Totalizar
    End If
End Sub

Private Sub gridTipoRetenciones_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If retenciones.count >= rowIndex Then
        Set Retencion = retenciones.item(rowIndex)
        Values(1) = Retencion.nombre
        Values(2) = Retencion.Id
        Values(3) = Retencion.codigo
        Values(4) = Retencion.Porcentaje

    End If
End Sub

Private Sub txtRedondeo_Change()
    recibo.Redondeo = Val(Me.txtRedondeo.text)
    Totalizar
End Sub
