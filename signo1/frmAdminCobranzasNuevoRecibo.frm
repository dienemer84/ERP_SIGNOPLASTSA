VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmAdminCobranzasNuevoRecibo 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recibo"
   ClientHeight    =   10485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   18735
   ClipControls    =   0   'False
   Icon            =   "frmAdminCobranzasNuevoRecibo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10485
   ScaleWidth      =   18735
   Begin XtremeSuiteControls.PushButton cmdCerrar 
      Height          =   405
      Left            =   10200
      TabIndex        =   37
      Top             =   360
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
      Left            =   16080
      TabIndex        =   36
      Top             =   360
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
      Top             =   360
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
      Top             =   960
      Width           =   17925
      Begin XtremeSuiteControls.TabControl TabFacturasRetenciones 
         Height          =   4335
         Left            =   120
         TabIndex        =   28
         Top             =   300
         Width           =   7065
         _Version        =   786432
         _ExtentX        =   12462
         _ExtentY        =   7646
         _StockProps     =   68
         Appearance      =   10
         Color           =   32
         ItemCount       =   2
         Item(0).Caption =   "Comprobantes"
         Item(0).ControlCount=   1
         Item(0).Control(0)=   "gridFacturas"
         Item(1).Caption =   "Retenciones"
         Item(1).ControlCount=   2
         Item(1).Control(0)=   "gridRetenciones"
         Item(1).Control(1)=   "gridTipoRetenciones"
         Begin GridEX20.GridEX gridTipoRetenciones 
            Height          =   2175
            Left            =   -62725
            TabIndex        =   31
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
            Column(1)       =   "frmAdminCobranzasNuevoRecibo.frx":000C
            Column(2)       =   "frmAdminCobranzasNuevoRecibo.frx":012C
            Column(3)       =   "frmAdminCobranzasNuevoRecibo.frx":022C
            Column(4)       =   "frmAdminCobranzasNuevoRecibo.frx":0320
            FormatStylesCount=   6
            FormatStyle(1)  =   "frmAdminCobranzasNuevoRecibo.frx":0424
            FormatStyle(2)  =   "frmAdminCobranzasNuevoRecibo.frx":055C
            FormatStyle(3)  =   "frmAdminCobranzasNuevoRecibo.frx":060C
            FormatStyle(4)  =   "frmAdminCobranzasNuevoRecibo.frx":06C0
            FormatStyle(5)  =   "frmAdminCobranzasNuevoRecibo.frx":0798
            FormatStyle(6)  =   "frmAdminCobranzasNuevoRecibo.frx":0850
            ImageCount      =   0
            PrinterProperties=   "frmAdminCobranzasNuevoRecibo.frx":0930
         End
         Begin GridEX20.GridEX gridFacturas 
            Height          =   3870
            Left            =   135
            TabIndex        =   29
            Top             =   345
            Width           =   6840
            _ExtentX        =   12065
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
            ColumnsCount    =   5
            Column(1)       =   "frmAdminCobranzasNuevoRecibo.frx":0B08
            Column(2)       =   "frmAdminCobranzasNuevoRecibo.frx":0C7C
            Column(3)       =   "frmAdminCobranzasNuevoRecibo.frx":0DBC
            Column(4)       =   "frmAdminCobranzasNuevoRecibo.frx":0F30
            Column(5)       =   "frmAdminCobranzasNuevoRecibo.frx":10B4
            FormatStylesCount=   6
            FormatStyle(1)  =   "frmAdminCobranzasNuevoRecibo.frx":1210
            FormatStyle(2)  =   "frmAdminCobranzasNuevoRecibo.frx":1348
            FormatStyle(3)  =   "frmAdminCobranzasNuevoRecibo.frx":13F8
            FormatStyle(4)  =   "frmAdminCobranzasNuevoRecibo.frx":14AC
            FormatStyle(5)  =   "frmAdminCobranzasNuevoRecibo.frx":1584
            FormatStyle(6)  =   "frmAdminCobranzasNuevoRecibo.frx":163C
            ImageCount      =   0
            PrinterProperties=   "frmAdminCobranzasNuevoRecibo.frx":171C
         End
         Begin GridEX20.GridEX gridRetenciones 
            Height          =   3870
            Left            =   -69865
            TabIndex        =   30
            Top             =   345
            Visible         =   0   'False
            Width           =   6870
            _ExtentX        =   12118
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
            Column(1)       =   "frmAdminCobranzasNuevoRecibo.frx":18F4
            Column(2)       =   "frmAdminCobranzasNuevoRecibo.frx":1A10
            Column(3)       =   "frmAdminCobranzasNuevoRecibo.frx":1B44
            Column(4)       =   "frmAdminCobranzasNuevoRecibo.frx":1C80
            FormatStylesCount=   6
            FormatStyle(1)  =   "frmAdminCobranzasNuevoRecibo.frx":1D8C
            FormatStyle(2)  =   "frmAdminCobranzasNuevoRecibo.frx":1EC4
            FormatStyle(3)  =   "frmAdminCobranzasNuevoRecibo.frx":1F74
            FormatStyle(4)  =   "frmAdminCobranzasNuevoRecibo.frx":2028
            FormatStyle(5)  =   "frmAdminCobranzasNuevoRecibo.frx":2100
            FormatStyle(6)  =   "frmAdminCobranzasNuevoRecibo.frx":21B8
            ImageCount      =   0
            PrinterProperties=   "frmAdminCobranzasNuevoRecibo.frx":2298
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
         Width           =   10320
         Begin XtremeSuiteControls.GroupBox grpCheques 
            Height          =   2625
            Left            =   120
            TabIndex        =   8
            Top             =   3840
            Width           =   10050
            _Version        =   786432
            _ExtentX        =   17727
            _ExtentY        =   4630
            _StockProps     =   79
            Caption         =   "Cheques Recibidos"
            UseVisualStyle  =   -1  'True
            Begin GridEX20.GridEX gridCheques 
               Height          =   2280
               Left            =   75
               TabIndex        =   9
               Top             =   225
               Width           =   9780
               _ExtentX        =   17251
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
               Column(1)       =   "frmAdminCobranzasNuevoRecibo.frx":2470
               Column(2)       =   "frmAdminCobranzasNuevoRecibo.frx":25FC
               Column(3)       =   "frmAdminCobranzasNuevoRecibo.frx":2764
               Column(4)       =   "frmAdminCobranzasNuevoRecibo.frx":28CC
               Column(5)       =   "frmAdminCobranzasNuevoRecibo.frx":2A04
               Column(6)       =   "frmAdminCobranzasNuevoRecibo.frx":2B48
               Column(7)       =   "frmAdminCobranzasNuevoRecibo.frx":2CB0
               FormatStylesCount=   6
               FormatStyle(1)  =   "frmAdminCobranzasNuevoRecibo.frx":2DAC
               FormatStyle(2)  =   "frmAdminCobranzasNuevoRecibo.frx":2EE4
               FormatStyle(3)  =   "frmAdminCobranzasNuevoRecibo.frx":2F94
               FormatStyle(4)  =   "frmAdminCobranzasNuevoRecibo.frx":3048
               FormatStyle(5)  =   "frmAdminCobranzasNuevoRecibo.frx":3120
               FormatStyle(6)  =   "frmAdminCobranzasNuevoRecibo.frx":31D8
               ImageCount      =   0
               PrinterProperties=   "frmAdminCobranzasNuevoRecibo.frx":32B8
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
               Column(1)       =   "frmAdminCobranzasNuevoRecibo.frx":3490
               Column(2)       =   "frmAdminCobranzasNuevoRecibo.frx":3590
               FormatStylesCount=   6
               FormatStyle(1)  =   "frmAdminCobranzasNuevoRecibo.frx":3684
               FormatStyle(2)  =   "frmAdminCobranzasNuevoRecibo.frx":37BC
               FormatStyle(3)  =   "frmAdminCobranzasNuevoRecibo.frx":386C
               FormatStyle(4)  =   "frmAdminCobranzasNuevoRecibo.frx":3920
               FormatStyle(5)  =   "frmAdminCobranzasNuevoRecibo.frx":39F8
               FormatStyle(6)  =   "frmAdminCobranzasNuevoRecibo.frx":3AB0
               ImageCount      =   0
               PrinterProperties=   "frmAdminCobranzasNuevoRecibo.frx":3B90
            End
         End
         Begin XtremeSuiteControls.GroupBox grpBanco 
            Height          =   1920
            Left            =   120
            TabIndex        =   12
            Top             =   1860
            Width           =   10050
            _Version        =   786432
            _ExtentX        =   17727
            _ExtentY        =   3387
            _StockProps     =   79
            Caption         =   "Banco"
            UseVisualStyle  =   -1  'True
            Begin GridEX20.GridEX gridDepositosOperaciones 
               Height          =   1545
               Left            =   120
               TabIndex        =   13
               Top             =   225
               Width           =   9795
               _ExtentX        =   17277
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
               Column(1)       =   "frmAdminCobranzasNuevoRecibo.frx":3D68
               Column(2)       =   "frmAdminCobranzasNuevoRecibo.frx":3EF4
               Column(3)       =   "frmAdminCobranzasNuevoRecibo.frx":4084
               Column(4)       =   "frmAdminCobranzasNuevoRecibo.frx":420C
               FormatStylesCount=   6
               FormatStyle(1)  =   "frmAdminCobranzasNuevoRecibo.frx":437C
               FormatStyle(2)  =   "frmAdminCobranzasNuevoRecibo.frx":44B4
               FormatStyle(3)  =   "frmAdminCobranzasNuevoRecibo.frx":4564
               FormatStyle(4)  =   "frmAdminCobranzasNuevoRecibo.frx":4618
               FormatStyle(5)  =   "frmAdminCobranzasNuevoRecibo.frx":46F0
               FormatStyle(6)  =   "frmAdminCobranzasNuevoRecibo.frx":47A8
               ImageCount      =   0
               PrinterProperties=   "frmAdminCobranzasNuevoRecibo.frx":4888
            End
         End
         Begin XtremeSuiteControls.GroupBox grpCaja 
            Height          =   1635
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   10050
            _Version        =   786432
            _ExtentX        =   17727
            _ExtentY        =   2884
            _StockProps     =   79
            Caption         =   "Caja"
            UseVisualStyle  =   -1  'True
            Begin GridEX20.GridEX gridCajaOperaciones 
               Height          =   1260
               Left            =   90
               TabIndex        =   15
               Top             =   225
               Width           =   9810
               _ExtentX        =   17304
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
               Column(1)       =   "frmAdminCobranzasNuevoRecibo.frx":4A60
               Column(2)       =   "frmAdminCobranzasNuevoRecibo.frx":4BEC
               Column(3)       =   "frmAdminCobranzasNuevoRecibo.frx":4D54
               Column(4)       =   "frmAdminCobranzasNuevoRecibo.frx":4EDC
               FormatStylesCount=   6
               FormatStyle(1)  =   "frmAdminCobranzasNuevoRecibo.frx":5010
               FormatStyle(2)  =   "frmAdminCobranzasNuevoRecibo.frx":5148
               FormatStyle(3)  =   "frmAdminCobranzasNuevoRecibo.frx":51F8
               FormatStyle(4)  =   "frmAdminCobranzasNuevoRecibo.frx":52AC
               FormatStyle(5)  =   "frmAdminCobranzasNuevoRecibo.frx":5384
               FormatStyle(6)  =   "frmAdminCobranzasNuevoRecibo.frx":543C
               ImageCount      =   0
               PrinterProperties=   "frmAdminCobranzasNuevoRecibo.frx":551C
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
         TabIndex        =   32
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
      Top             =   120
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
         Format          =   66519041
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
      Height          =   2655
      Left            =   120
      TabIndex        =   33
      Top             =   8160
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   4683
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
      Column(1)       =   "frmAdminCobranzasNuevoRecibo.frx":56F4
      Column(2)       =   "frmAdminCobranzasNuevoRecibo.frx":5818
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminCobranzasNuevoRecibo.frx":5904
      FormatStyle(2)  =   "frmAdminCobranzasNuevoRecibo.frx":5A3C
      FormatStyle(3)  =   "frmAdminCobranzasNuevoRecibo.frx":5AEC
      FormatStyle(4)  =   "frmAdminCobranzasNuevoRecibo.frx":5BA0
      FormatStyle(5)  =   "frmAdminCobranzasNuevoRecibo.frx":5C78
      FormatStyle(6)  =   "frmAdminCobranzasNuevoRecibo.frx":5D30
      ImageCount      =   0
      PrinterProperties=   "frmAdminCobranzasNuevoRecibo.frx":5E10
   End
   Begin GridEX20.GridEX gridCuentasBancarias 
      Height          =   4215
      Left            =   2040
      TabIndex        =   34
      Top             =   8160
      Visible         =   0   'False
      Width           =   5265
      _ExtentX        =   9287
      _ExtentY        =   7435
      Version         =   "2.0"
      BoundColumnIndex=   "id"
      ReplaceColumnIndex=   "cuenta"
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
      Column(1)       =   "frmAdminCobranzasNuevoRecibo.frx":5FE8
      Column(2)       =   "frmAdminCobranzasNuevoRecibo.frx":610C
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminCobranzasNuevoRecibo.frx":6200
      FormatStyle(2)  =   "frmAdminCobranzasNuevoRecibo.frx":6338
      FormatStyle(3)  =   "frmAdminCobranzasNuevoRecibo.frx":63E8
      FormatStyle(4)  =   "frmAdminCobranzasNuevoRecibo.frx":649C
      FormatStyle(5)  =   "frmAdminCobranzasNuevoRecibo.frx":6574
      FormatStyle(6)  =   "frmAdminCobranzasNuevoRecibo.frx":662C
      ImageCount      =   0
      PrinterProperties=   "frmAdminCobranzasNuevoRecibo.frx":670C
   End
   Begin GridEX20.GridEX gridMonedas 
      Height          =   2655
      Left            =   8040
      TabIndex        =   35
      Top             =   8160
      Width           =   4980
      _ExtentX        =   8784
      _ExtentY        =   4683
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
      Column(1)       =   "frmAdminCobranzasNuevoRecibo.frx":68E4
      Column(2)       =   "frmAdminCobranzasNuevoRecibo.frx":6A08
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminCobranzasNuevoRecibo.frx":6AFC
      FormatStyle(2)  =   "frmAdminCobranzasNuevoRecibo.frx":6C34
      FormatStyle(3)  =   "frmAdminCobranzasNuevoRecibo.frx":6CE4
      FormatStyle(4)  =   "frmAdminCobranzasNuevoRecibo.frx":6D98
      FormatStyle(5)  =   "frmAdminCobranzasNuevoRecibo.frx":6E70
      FormatStyle(6)  =   "frmAdminCobranzasNuevoRecibo.frx":6F28
      ImageCount      =   0
      PrinterProperties=   "frmAdminCobranzasNuevoRecibo.frx":7008
   End
   Begin GridEX20.GridEX gridFacturasCombo 
      Height          =   2700
      Left            =   13440
      TabIndex        =   38
      Top             =   8160
      Width           =   5490
      _ExtentX        =   9684
      _ExtentY        =   4763
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
      Column(1)       =   "frmAdminCobranzasNuevoRecibo.frx":71E0
      Column(2)       =   "frmAdminCobranzasNuevoRecibo.frx":7304
      Column(3)       =   "frmAdminCobranzasNuevoRecibo.frx":73F8
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminCobranzasNuevoRecibo.frx":750C
      FormatStyle(2)  =   "frmAdminCobranzasNuevoRecibo.frx":7644
      FormatStyle(3)  =   "frmAdminCobranzasNuevoRecibo.frx":76F4
      FormatStyle(4)  =   "frmAdminCobranzasNuevoRecibo.frx":77A8
      FormatStyle(5)  =   "frmAdminCobranzasNuevoRecibo.frx":7880
      FormatStyle(6)  =   "frmAdminCobranzasNuevoRecibo.frx":7938
      ImageCount      =   0
      PrinterProperties=   "frmAdminCobranzasNuevoRecibo.frx":7A18
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
Private Recibo As Recibo
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
    Set Recibo = DAORecibo.FindById(nIdRecibo, True, True, True, True, True)

    If Recibo Is Nothing Then
        MsgBox "recibo no encontrado, cierre pantalla", vbCritical
    End If


    Me.dtpFecha.value = Recibo.FEcha
    lblNumeroRecibo.caption = "Número: " & Recibo.Id
    Me.txtRedondeo.Text = Recibo.Redondeo

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

    Me.gridRetenciones.ItemCount = Recibo.retenciones.count
    Me.gridFacturas.ItemCount = Recibo.facturas.count

    DAOMoneda.LlenarCombo Me.cboMonedas
    Me.cboMonedas.ListIndex = PosIndexCbo(Recibo.moneda.Id, Me.cboMonedas)
    DAOCliente.LlenarCombo Me.cboClientes, True, True
    Me.cboClientes.ListIndex = PosIndexCbo(Recibo.cliente.Id, Me.cboClientes)

    Me.gridCajaOperaciones.ItemCount = Recibo.operacionesCaja.count
    Me.gridDepositosOperaciones.ItemCount = Recibo.operacionesBanco.count
    Me.gridCheques.ItemCount = Recibo.cheques.count

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
            If Recibo.facturas.count > 0 Then
                If MsgBox("Va a cambiar de cliente y perder las facturas." & vbNewLine & "¿Desea continuar?", vbQuestion + vbYesNo) = vbNo Then
                    clienteChange = True
                    Me.cboClientes.ListIndex = funciones.PosIndexCbo(Recibo.cliente.Id, Me.cboClientes)
                    clienteChange = False
                    Exit Sub

                End If
            End If
            VaciarFacturasRetenciones
            CargarFacturasCliente
            Set Recibo.cliente = DAOCliente.BuscarPorID(Me.cboClientes.ItemData(Me.cboClientes.ListIndex))
            'Else
            '    VaciarFacturasRetenciones
            '    CargarFacturasCliente
        End If
    End If

    clienteChange = False
End Sub



Private Sub cboMonedas_Click()
    If Me.cboMonedas.ListIndex <> -1 And dataLoaded Then
        Set Recibo.moneda = DAOMoneda.GetById(Me.cboMonedas.ItemData(Me.cboMonedas.ListIndex))
        Totalizar
    End If
End Sub


Private Sub VaciarFacturasRetenciones(Optional ByVal cleanRetenciones As Boolean = True)
    Set Recibo.facturas = New Collection
    Me.gridFacturas.ItemCount = 0

    If cleanRetenciones Then
        Set Recibo.retenciones = New Collection
        Me.gridRetenciones.ItemCount = 0
    End If
End Sub


Private Sub cmdActualizar_Click()
    If Not Recibo.IsValid Then
        MsgBox Recibo.ValidationMessages, vbExclamation
        Exit Sub
    End If

    If DAORecibo.Save(Recibo) And DAORecibo.aprobar(Recibo) Then

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
    If Not Recibo.IsValid Then
        MsgBox Recibo.ValidationMessages, vbExclamation
        Exit Sub
    End If

    If DAORecibo.Save(Recibo) Then
        MsgBox "Recibo guardado.", vbInformation
        Unload Me
    Else
        MsgBox "Hubo un error al intentar guardar el recibo.", vbCritical
    End If

End Sub


Private Sub Totalizar()

    Me.lblTotalFactura.caption = "Total Facturas: " & Replace(FormatCurrency(funciones.FormatearDecimales(Recibo.TotalFacturas)), "$", "")
    Me.lblTotalRetenciones.caption = "Total Retenciones: " & Replace(FormatCurrency(funciones.FormatearDecimales(Recibo.TotalRetenciones)), "$", "")
    Me.lblTotalCheques.caption = "Total Cheques: " & Replace(FormatCurrency(funciones.FormatearDecimales(Recibo.TotalCheques)), "$", "")
    Me.lblTotalBanco.caption = "Total Banco: " & Replace(FormatCurrency(funciones.FormatearDecimales(Recibo.TotalOperacionesBanco)), "$", "")
    Me.lblTotalCaja.caption = "Total Caja: " & Replace(FormatCurrency(funciones.FormatearDecimales(Recibo.TotalOperacionesCaja)), "$", "")
    
    'Replace(FormatCurrency(funciones.FormatearDecimales(saldoComprobante) * i), "$", "")

    Dim totalRecibo As Double
    totalRecibo = funciones.FormatearDecimales(Recibo.total)
    Dim totalCancelado As Double
    totalCancelado = funciones.FormatearDecimales(Recibo.TotalRecibido)

    Me.lblTotalRecibo.caption = Replace(FormatCurrency(funciones.FormatearDecimales(totalRecibo)), "$", "")
    Me.lblTotalRecibido.caption = "Total Recibido: " & Replace(FormatCurrency(funciones.FormatearDecimales(totalCancelado)), "$", "")

    If totalCancelado < totalRecibo Then
        lblTotalRecibo.backColor = vbRed
    ElseIf totalCancelado = totalRecibo Then
        lblTotalRecibo.backColor = vbYellow
    ElseIf totalCancelado > totalRecibo Then
        lblTotalRecibo.backColor = vbGreen
    End If

    Me.lblDiferencia.caption = Me.lblDiferencia.Tag & Replace(FormatCurrency(funciones.FormatearDecimales(totalCancelado - MonedaConverter.Convertir(totalRecibo, Recibo.moneda.Id, DAOMoneda.MONEDA_PESO_ID))), "$", "")
    'Debug.Print MonedaConverter.Convertir(totalRecibo, recibo.moneda.id, DAOMoneda.MONEDA_PESO_ID)
    Recibo.ACuenta = (totalCancelado - totalRecibo)
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
    Recibo.FEcha = Me.dtpFecha.value
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

Private Sub gridBancos_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex <= bancos.count Then
        Set Banco = bancos.item(RowIndex)
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
    operacion.IdPertenencia = Recibo.Id
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
    Recibo.operacionesCaja.Add operacion

    Totalizar
End Sub

Private Sub gridCajaOperaciones_UnboundDelete(ByVal RowIndex As Long, ByVal Bookmark As Variant)
    If RowIndex > 0 And Recibo.operacionesCaja.count >= RowIndex Then
        Recibo.operacionesCaja.remove RowIndex
        Totalizar
    End If
End Sub

Private Sub gridCajaOperaciones_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex <= Recibo.operacionesCaja.count Then
        Set operacion = Recibo.operacionesCaja.item(RowIndex)
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

Private Sub gridCajaOperaciones_UnboundUpdate(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex > 0 And Recibo.operacionesCaja.count > 0 Then
        Set operacion = Recibo.operacionesCaja.item(RowIndex)
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

Private Sub gridCajas_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex > 0 And Cajas.count > 0 Then
        Set caja = Cajas.item(RowIndex)
        Values(1) = caja.Id
        Values(2) = caja.nombre
    End If
End Sub

Private Sub gridCheques_BeforeUpdate(ByVal Cancel As GridEX20.JSRetBoolean)
    If Me.gridCheques.row = -1 Then    'es nuevo
        If Val(Me.gridCheques.value(6)) = -1 Then
            Me.gridCheques.value(7) = Recibo.cliente.razon
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
    cheque.FechaRecibido = Recibo.FEcha
    If IsNumeric(Values(3)) Then Set cheque.moneda = DAOMoneda.GetById(Values(3))
    cheque.Monto = funciones.RedondearDecimales(Val(Values(2)))
    cheque.numero = Values(1)
    cheque.Propio = False
    cheque.FechaVencimiento = Values(5)

    cheque.TercerosPropio = (Values(6) = -1)

    If cheque.TercerosPropio Then
        cheque.OrigenDestino = Recibo.cliente.razon
    Else
        cheque.OrigenDestino = Values(7)
    End If
    Recibo.cheques.Add cheque

    Totalizar
End Sub

Private Sub gridCheques_UnboundDelete(ByVal RowIndex As Long, ByVal Bookmark As Variant)
    If RowIndex > 0 Then
        Recibo.cheques.remove RowIndex
        Totalizar
    End If
End Sub


Private Sub gridCheques_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex <= Recibo.cheques.count Then
        Set cheque = Recibo.cheques.item(RowIndex)
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


Private Sub gridCheques_UnboundUpdate(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex > 0 And Recibo.cheques.count >= RowIndex Then
        Set cheque = Recibo.cheques.item(RowIndex)
        If IsNumeric(Values(4)) Then Set cheque.Banco = DAOBancos.GetById(Values(4))
        cheque.EnCartera = True
        cheque.FechaRecibido = Recibo.FEcha
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


Private Sub gridCuentasBancarias_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If cuentasBancarias.count >= RowIndex Then
        Set CuentaBancaria = cuentasBancarias.item(RowIndex)
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
    operacion.IdPertenencia = Recibo.Id
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

    Recibo.operacionesBanco.Add operacion

    Totalizar
End Sub


Private Sub gridDepositosOperaciones_UnboundDelete(ByVal RowIndex As Long, ByVal Bookmark As Variant)
    If RowIndex > 0 And Recibo.operacionesBanco.count >= RowIndex Then
        Recibo.operacionesBanco.remove RowIndex
        Totalizar
    End If
End Sub


Private Sub gridDepositosOperaciones_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex <= Recibo.operacionesBanco.count Then
        Set operacion = Recibo.operacionesBanco.item(RowIndex)
        Values(1) = funciones.FormatearDecimales(operacion.Monto)
        If IsSomething(operacion.moneda) Then
            Values(2) = operacion.moneda.NombreCorto
        End If
        Values(3) = operacion.FechaOperacion
        If IsSomething(operacion.CuentaBancaria) Then
        '''Values(4) = operacion.CuentaBancaria.DescripcionFormateada & " | " & operacion.CuentaBancaria.BancoNombre & " | " & operacion.CuentaBancaria.moneda.NombreCorto
            Values(4) = operacion.CuentaBancaria.DescripcionFormateada
        End If
        Totalizar
    End If
End Sub


Private Sub gridDepositosOperaciones_UnboundUpdate(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex > 0 And Recibo.operacionesBanco.count > 0 Then
        Set operacion = Recibo.operacionesBanco.item(RowIndex)
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
        Recibo.facturas.Add Factura

        If Recibo.PagosDeFacturas.Exists(CStr(Factura.Id)) Then
            Recibo.PagosDeFacturas.remove CStr(Factura.Id)
        End If

        Recibo.PagosDeFacturas.Add CStr(Factura.Id), Factura.total - DAOFactura.PagosRealizados(Factura.Id)

        Totalizar
        
    End If

End Sub


Private Sub gridFacturas_UnboundDelete(ByVal RowIndex As Long, ByVal Bookmark As Variant)
    If RowIndex > 0 And Recibo.facturas.count >= RowIndex Then
        Recibo.facturas.remove RowIndex
        If Recibo.PagosDeFacturas.Exists(CStr(Factura.Id)) Then
            Recibo.PagosDeFacturas.remove CStr(Factura.Id)
        End If
        Totalizar
    End If
End Sub


Private Sub gridFacturas_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex > 0 And Recibo.facturas.count >= RowIndex Then
        Set Factura = Recibo.facturas.item(RowIndex)
        Values(1) = Factura.GetShortDescription(False, True)
        
        Values(2) = Factura.FechaEmision
        
        Values(3) = funciones.FormatearDecimales(Factura.total)
        
        Values(4) = funciones.FormatearDecimales(Factura.total - DAOFactura.PagosRealizados(Factura.Id))

        If Recibo.PagosDeFacturas.Exists(CStr(Factura.Id)) Then
            Values(5) = funciones.FormatearDecimales(Recibo.PagosDeFacturas.item(CStr(Factura.Id)))
        Else
            Values(5) = funciones.FormatearDecimales(0)
        End If

    End If
End Sub


Private Sub gridFacturas_UnboundUpdate(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex > 0 And Recibo.facturas.count >= RowIndex Then
        Set Factura = Recibo.facturas.item(RowIndex)
        'Set Factura = DAOFactura.FindById(Values(1), True)
        'recibo.facturas.Add Factura, , , RowIndex
        'recibo.facturas.Remove RowIndex

        If Recibo.PagosDeFacturas.Exists(CStr(Factura.Id)) Then
            Recibo.PagosDeFacturas.remove CStr(Factura.Id)
        End If
        
        'ANTERIORMENTE ESTABA EL VALOR 4, PERO CUANDO AGREGUÉ LA COLUMNA DE FECHA SE MODIFICÓ AL 5
        'recibo.PagosDeFacturas.Add CStr(Factura.Id), Val(Values(4))
        Recibo.PagosDeFacturas.Add CStr(Factura.Id), Val(Values(5))

        Totalizar    ' no se si hay que totalizar
    End If
End Sub


Private Sub gridFacturasCombo_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex > 0 And facturasCliente.count >= RowIndex Then
        Set Factura = facturasCliente.item(RowIndex)
        Values(1) = Factura.Id
        Values(2) = Factura.GetShortDescription(False, True)
        Values(3) = funciones.FormatearDecimales(Factura.total)
    End If
End Sub


Private Sub gridMonedas_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex > 0 And Monedas.count > 0 Then
        Set moneda = Monedas.item(RowIndex)
        Values(1) = moneda.Id
        Values(2) = moneda.NombreCorto
    End If
End Sub


Private Sub gridRetenciones_BeforeUpdate(ByVal Cancel As GridEX20.JSRetBoolean)
    Cancel = LenB(Me.gridRetenciones.value(1)) = 0 Or _
             Not IsNumeric(Me.gridRetenciones.value(2)) Or _
             Not IsNumeric(Me.gridRetenciones.value(3))
End Sub


Private Sub gridRetenciones_UnboundDelete(ByVal RowIndex As Long, ByVal Bookmark As Variant)
    If RowIndex > 0 Then
        Recibo.retenciones.remove RowIndex
        Totalizar
    End If
End Sub


Private Sub gridRetenciones_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex > 0 And RowIndex <= Recibo.retenciones.count Then

        Set retencionRecibo = Recibo.retenciones.item(RowIndex)
        Values(1) = retencionRecibo.NroRetencion
        If IsSomething(retencionRecibo.Retencion) Then
            Values(2) = retencionRecibo.Retencion.nombre
        End If
        Values(3) = funciones.FormatearDecimales(retencionRecibo.Valor)
        Values(4) = Format(retencionRecibo.FEcha, "dd-mm-yyyy")

    End If
End Sub


Private Sub UpdateAddRetencion(ByRef retencionRecibo As retencionRecibo, Values As GridEX20.JSRowData)
    retencionRecibo.idRecibo = Recibo.Id
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
    Recibo.retenciones.Add retencionRecibo
    Totalizar
End Sub

Private Sub gridRetenciones_UnboundUpdate(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex > 0 Then
        Set retencionRecibo = Recibo.retenciones(RowIndex)
        UpdateAddRetencion retencionRecibo, Values
        Totalizar
    End If
End Sub


Private Sub gridTipoRetenciones_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If retenciones.count >= RowIndex Then
        Set Retencion = retenciones.item(RowIndex)
        Values(1) = Retencion.nombre
        Values(2) = Retencion.Id
        Values(3) = Retencion.codigo
        Values(4) = Retencion.Porcentaje

    End If
End Sub


Private Sub txtRedondeo_Change()
    Recibo.Redondeo = Val(Me.txtRedondeo.Text)
    Totalizar
End Sub
