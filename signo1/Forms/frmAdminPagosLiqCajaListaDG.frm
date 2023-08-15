VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmAdminPagosLiqCajaListaDG 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   10575
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15870
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10575
   ScaleWidth      =   15870
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.PushButton btnConfirmar 
      Height          =   495
      Left            =   5520
      TabIndex        =   31
      Top             =   6600
      Width           =   2295
      _Version        =   786432
      _ExtentX        =   4048
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Confirmar"
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
   Begin XtremeSuiteControls.PushButton btnExportarComprobantes 
      Height          =   495
      Left            =   360
      TabIndex        =   30
      Top             =   6600
      Width           =   2295
      _Version        =   786432
      _ExtentX        =   4048
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Exportar"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton btnExportarConfirmados 
      Height          =   495
      Left            =   13200
      TabIndex        =   29
      Top             =   6600
      Width           =   2415
      _Version        =   786432
      _ExtentX        =   4260
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Exportar"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton PusGuardar 
      Height          =   495
      Left            =   13320
      TabIndex        =   25
      Top             =   240
      Width           =   2295
      _Version        =   786432
      _ExtentX        =   4048
      _ExtentY        =   873
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
   Begin XtremeSuiteControls.PushButton btnQuitarSeleccionado 
      Height          =   495
      Left            =   8160
      TabIndex        =   22
      Top             =   6600
      Width           =   2295
      _Version        =   786432
      _ExtentX        =   4048
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Quitar Seleccionado"
      UseVisualStyle  =   -1  'True
   End
   Begin VB.TextBox txtNumerodeLiquidacion 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      TabIndex        =   20
      Top             =   360
      Width           =   1815
   End
   Begin VB.TextBox txtOtrosDescuentos 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   17880
      TabIndex        =   5
      Top             =   11760
      Visible         =   0   'False
      Width           =   960
   End
   Begin XtremeSuiteControls.PushButton btnLimpiarNúmero 
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   900
      Width           =   495
      _Version        =   786432
      _ExtentX        =   873
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "X"
      UseVisualStyle  =   -1  'True
   End
   Begin VB.TextBox txtFiltroNumero 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   900
      Width           =   2775
   End
   Begin GridEX20.GridEX grillaConfirmados 
      Height          =   4455
      Left            =   8160
      TabIndex        =   2
      Top             =   2040
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   7858
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      HeaderStyle     =   2
      MethodHoldFields=   -1  'True
      GroupByBoxVisible=   0   'False
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   7
      Column(1)       =   "frmAdminPagosLiqCajaListaDG.frx":0000
      Column(2)       =   "frmAdminPagosLiqCajaListaDG.frx":0188
      Column(3)       =   "frmAdminPagosLiqCajaListaDG.frx":02EC
      Column(4)       =   "frmAdminPagosLiqCajaListaDG.frx":0434
      Column(5)       =   "frmAdminPagosLiqCajaListaDG.frx":0574
      Column(6)       =   "frmAdminPagosLiqCajaListaDG.frx":06BC
      Column(7)       =   "frmAdminPagosLiqCajaListaDG.frx":07FC
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminPagosLiqCajaListaDG.frx":0900
      FormatStyle(2)  =   "frmAdminPagosLiqCajaListaDG.frx":0A28
      FormatStyle(3)  =   "frmAdminPagosLiqCajaListaDG.frx":0AD8
      FormatStyle(4)  =   "frmAdminPagosLiqCajaListaDG.frx":0B8C
      FormatStyle(5)  =   "frmAdminPagosLiqCajaListaDG.frx":0C64
      FormatStyle(6)  =   "frmAdminPagosLiqCajaListaDG.frx":0D1C
      ImageCount      =   0
      PrinterProperties=   "frmAdminPagosLiqCajaListaDG.frx":0DFC
   End
   Begin GridEX20.GridEX grilla 
      Height          =   4455
      Left            =   360
      TabIndex        =   1
      Top             =   2040
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   7858
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      MethodHoldFields=   -1  'True
      GroupByBoxVisible=   0   'False
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   7
      Column(1)       =   "frmAdminPagosLiqCajaListaDG.frx":0FCC
      Column(2)       =   "frmAdminPagosLiqCajaListaDG.frx":110C
      Column(3)       =   "frmAdminPagosLiqCajaListaDG.frx":1244
      Column(4)       =   "frmAdminPagosLiqCajaListaDG.frx":138C
      Column(5)       =   "frmAdminPagosLiqCajaListaDG.frx":14CC
      Column(6)       =   "frmAdminPagosLiqCajaListaDG.frx":1614
      Column(7)       =   "frmAdminPagosLiqCajaListaDG.frx":1754
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminPagosLiqCajaListaDG.frx":1850
      FormatStyle(2)  =   "frmAdminPagosLiqCajaListaDG.frx":1978
      FormatStyle(3)  =   "frmAdminPagosLiqCajaListaDG.frx":1A28
      FormatStyle(4)  =   "frmAdminPagosLiqCajaListaDG.frx":1ADC
      FormatStyle(5)  =   "frmAdminPagosLiqCajaListaDG.frx":1BB4
      FormatStyle(6)  =   "frmAdminPagosLiqCajaListaDG.frx":1C6C
      ImageCount      =   0
      PrinterProperties=   "frmAdminPagosLiqCajaListaDG.frx":1D4C
   End
   Begin XtremeSuiteControls.PushButton btnCargarCbtes 
      Height          =   495
      Left            =   5520
      TabIndex        =   0
      Top             =   240
      Width           =   2295
      _Version        =   786432
      _ExtentX        =   4048
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Cargar Cbtes"
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
   Begin GridEX20.GridEX gridBancos 
      Height          =   1845
      Left            =   360
      TabIndex        =   6
      Top             =   10800
      Visible         =   0   'False
      Width           =   6705
      _ExtentX        =   11827
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
      Column(1)       =   "frmAdminPagosLiqCajaListaDG.frx":1F1C
      Column(2)       =   "frmAdminPagosLiqCajaListaDG.frx":201C
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminPagosLiqCajaListaDG.frx":210C
      FormatStyle(2)  =   "frmAdminPagosLiqCajaListaDG.frx":2244
      FormatStyle(3)  =   "frmAdminPagosLiqCajaListaDG.frx":22F4
      FormatStyle(4)  =   "frmAdminPagosLiqCajaListaDG.frx":23A8
      FormatStyle(5)  =   "frmAdminPagosLiqCajaListaDG.frx":2480
      FormatStyle(6)  =   "frmAdminPagosLiqCajaListaDG.frx":2538
      ImageCount      =   0
      PrinterProperties=   "frmAdminPagosLiqCajaListaDG.frx":2618
   End
   Begin GridEX20.GridEX gridCuentasBancarias 
      Height          =   1935
      Left            =   6720
      TabIndex        =   7
      Top             =   10800
      Visible         =   0   'False
      Width           =   6225
      _ExtentX        =   10980
      _ExtentY        =   3413
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
      Column(1)       =   "frmAdminPagosLiqCajaListaDG.frx":27F0
      Column(2)       =   "frmAdminPagosLiqCajaListaDG.frx":2914
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminPagosLiqCajaListaDG.frx":2A08
      FormatStyle(2)  =   "frmAdminPagosLiqCajaListaDG.frx":2B40
      FormatStyle(3)  =   "frmAdminPagosLiqCajaListaDG.frx":2BF0
      FormatStyle(4)  =   "frmAdminPagosLiqCajaListaDG.frx":2CA4
      FormatStyle(5)  =   "frmAdminPagosLiqCajaListaDG.frx":2D7C
      FormatStyle(6)  =   "frmAdminPagosLiqCajaListaDG.frx":2E34
      ImageCount      =   0
      PrinterProperties=   "frmAdminPagosLiqCajaListaDG.frx":2F14
   End
   Begin GridEX20.GridEX gridMonedas 
      Height          =   1815
      Left            =   240
      TabIndex        =   8
      Top             =   11040
      Visible         =   0   'False
      Width           =   1380
      _ExtentX        =   2434
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
      Column(1)       =   "frmAdminPagosLiqCajaListaDG.frx":30EC
      Column(2)       =   "frmAdminPagosLiqCajaListaDG.frx":3210
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminPagosLiqCajaListaDG.frx":3304
      FormatStyle(2)  =   "frmAdminPagosLiqCajaListaDG.frx":343C
      FormatStyle(3)  =   "frmAdminPagosLiqCajaListaDG.frx":34EC
      FormatStyle(4)  =   "frmAdminPagosLiqCajaListaDG.frx":35A0
      FormatStyle(5)  =   "frmAdminPagosLiqCajaListaDG.frx":3678
      FormatStyle(6)  =   "frmAdminPagosLiqCajaListaDG.frx":3730
      ImageCount      =   0
      PrinterProperties=   "frmAdminPagosLiqCajaListaDG.frx":3810
   End
   Begin GridEX20.GridEX gridCajas 
      Height          =   1935
      Left            =   16200
      TabIndex        =   9
      Top             =   9240
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   3413
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
      Column(1)       =   "frmAdminPagosLiqCajaListaDG.frx":39E8
      Column(2)       =   "frmAdminPagosLiqCajaListaDG.frx":3AE8
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminPagosLiqCajaListaDG.frx":3BD4
      FormatStyle(2)  =   "frmAdminPagosLiqCajaListaDG.frx":3D0C
      FormatStyle(3)  =   "frmAdminPagosLiqCajaListaDG.frx":3DBC
      FormatStyle(4)  =   "frmAdminPagosLiqCajaListaDG.frx":3E70
      FormatStyle(5)  =   "frmAdminPagosLiqCajaListaDG.frx":3F48
      FormatStyle(6)  =   "frmAdminPagosLiqCajaListaDG.frx":4000
      ImageCount      =   0
      PrinterProperties=   "frmAdminPagosLiqCajaListaDG.frx":40E0
   End
   Begin GridEX20.GridEX gridChequesDisponibles 
      Height          =   1920
      Left            =   5040
      TabIndex        =   10
      Top             =   11040
      Visible         =   0   'False
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   3387
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
      ColumnsCount    =   7
      Column(1)       =   "frmAdminPagosLiqCajaListaDG.frx":42B8
      Column(2)       =   "frmAdminPagosLiqCajaListaDG.frx":4438
      Column(3)       =   "frmAdminPagosLiqCajaListaDG.frx":45D8
      Column(4)       =   "frmAdminPagosLiqCajaListaDG.frx":4714
      Column(5)       =   "frmAdminPagosLiqCajaListaDG.frx":4820
      Column(6)       =   "frmAdminPagosLiqCajaListaDG.frx":4940
      Column(7)       =   "frmAdminPagosLiqCajaListaDG.frx":4A4C
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminPagosLiqCajaListaDG.frx":4B40
      FormatStyle(2)  =   "frmAdminPagosLiqCajaListaDG.frx":4C78
      FormatStyle(3)  =   "frmAdminPagosLiqCajaListaDG.frx":4D28
      FormatStyle(4)  =   "frmAdminPagosLiqCajaListaDG.frx":4DDC
      FormatStyle(5)  =   "frmAdminPagosLiqCajaListaDG.frx":4EB4
      FormatStyle(6)  =   "frmAdminPagosLiqCajaListaDG.frx":4F6C
      ImageCount      =   0
      PrinterProperties=   "frmAdminPagosLiqCajaListaDG.frx":504C
   End
   Begin GridEX20.GridEX gridChequeras 
      Height          =   1935
      Left            =   8760
      TabIndex        =   11
      Top             =   10920
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   3413
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
      Column(1)       =   "frmAdminPagosLiqCajaListaDG.frx":5224
      Column(2)       =   "frmAdminPagosLiqCajaListaDG.frx":5344
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminPagosLiqCajaListaDG.frx":5444
      FormatStyle(2)  =   "frmAdminPagosLiqCajaListaDG.frx":557C
      FormatStyle(3)  =   "frmAdminPagosLiqCajaListaDG.frx":562C
      FormatStyle(4)  =   "frmAdminPagosLiqCajaListaDG.frx":56E0
      FormatStyle(5)  =   "frmAdminPagosLiqCajaListaDG.frx":57B8
      FormatStyle(6)  =   "frmAdminPagosLiqCajaListaDG.frx":5870
      ImageCount      =   0
      PrinterProperties=   "frmAdminPagosLiqCajaListaDG.frx":5950
   End
   Begin GridEX20.GridEX gridChequesChequera 
      Height          =   1935
      Left            =   14160
      TabIndex        =   12
      Top             =   11160
      Width           =   1980
      _ExtentX        =   3493
      _ExtentY        =   3413
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
      Column(1)       =   "frmAdminPagosLiqCajaListaDG.frx":5B28
      Column(2)       =   "frmAdminPagosLiqCajaListaDG.frx":5C58
      SortKeysCount   =   1
      SortKey(1)      =   "frmAdminPagosLiqCajaListaDG.frx":5D58
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminPagosLiqCajaListaDG.frx":5DC0
      FormatStyle(2)  =   "frmAdminPagosLiqCajaListaDG.frx":5EF8
      FormatStyle(3)  =   "frmAdminPagosLiqCajaListaDG.frx":5FA8
      FormatStyle(4)  =   "frmAdminPagosLiqCajaListaDG.frx":605C
      FormatStyle(5)  =   "frmAdminPagosLiqCajaListaDG.frx":6134
      FormatStyle(6)  =   "frmAdminPagosLiqCajaListaDG.frx":61EC
      ImageCount      =   0
      PrinterProperties=   "frmAdminPagosLiqCajaListaDG.frx":62CC
   End
   Begin XtremeSuiteControls.RadioButton RadSeleccioneProveedor 
      Height          =   210
      Left            =   17880
      TabIndex        =   13
      Top             =   9600
      Visible         =   0   'False
      Width           =   2760
      _Version        =   786432
      _ExtentX        =   4868
      _ExtentY        =   370
      _StockProps     =   79
      Caption         =   "Seleccione Proveedor"
      Appearance      =   6
   End
   Begin XtremeSuiteControls.GroupBox GroValores 
      Height          =   3255
      Left            =   240
      TabIndex        =   15
      Top             =   7200
      Width           =   15495
      _Version        =   786432
      _ExtentX        =   27331
      _ExtentY        =   5741
      _StockProps     =   79
      Caption         =   "3- Valores de pago"
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
      Begin XtremeSuiteControls.TabControl TabControl 
         Height          =   2820
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   15180
         _Version        =   786432
         _ExtentX        =   26776
         _ExtentY        =   4974
         _StockProps     =   68
         Appearance      =   10
         Color           =   32
         PaintManager.ShowIcons=   -1  'True
         ItemCount       =   2
         Item(0).Caption =   "Banco"
         Item(0).ControlCount=   2
         Item(0).Control(0)=   "gridCompensatorios"
         Item(0).Control(1)=   "gridDepositosOperaciones"
         Item(1).Caption =   "Caja"
         Item(1).ControlCount=   1
         Item(1).Control(0)=   "gridCajaOperaciones"
         Begin GridEX20.GridEX gridCompensatorios 
            Height          =   4710
            Left            =   -69895
            TabIndex        =   17
            Top             =   435
            Width           =   9330
            _ExtentX        =   16457
            _ExtentY        =   8308
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
            Column(1)       =   "frmAdminPagosLiqCajaListaDG.frx":64A4
            Column(2)       =   "frmAdminPagosLiqCajaListaDG.frx":65EC
            Column(3)       =   "frmAdminPagosLiqCajaListaDG.frx":66F8
            Column(4)       =   "frmAdminPagosLiqCajaListaDG.frx":67E4
            Column(5)       =   "frmAdminPagosLiqCajaListaDG.frx":68E8
            FormatStylesCount=   6
            FormatStyle(1)  =   "frmAdminPagosLiqCajaListaDG.frx":6A28
            FormatStyle(2)  =   "frmAdminPagosLiqCajaListaDG.frx":6B60
            FormatStyle(3)  =   "frmAdminPagosLiqCajaListaDG.frx":6C10
            FormatStyle(4)  =   "frmAdminPagosLiqCajaListaDG.frx":6CC4
            FormatStyle(5)  =   "frmAdminPagosLiqCajaListaDG.frx":6D9C
            FormatStyle(6)  =   "frmAdminPagosLiqCajaListaDG.frx":6E54
            ImageCount      =   0
            PrinterProperties=   "frmAdminPagosLiqCajaListaDG.frx":6F34
         End
         Begin GridEX20.GridEX gridDepositosOperaciones 
            Height          =   2055
            Left            =   120
            TabIndex        =   18
            Top             =   480
            Width           =   9930
            _ExtentX        =   17515
            _ExtentY        =   3625
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
            Column(1)       =   "frmAdminPagosLiqCajaListaDG.frx":710C
            Column(2)       =   "frmAdminPagosLiqCajaListaDG.frx":726C
            Column(3)       =   "frmAdminPagosLiqCajaListaDG.frx":73A8
            Column(4)       =   "frmAdminPagosLiqCajaListaDG.frx":74DC
            Column(5)       =   "frmAdminPagosLiqCajaListaDG.frx":7620
            FormatStylesCount=   6
            FormatStyle(1)  =   "frmAdminPagosLiqCajaListaDG.frx":7724
            FormatStyle(2)  =   "frmAdminPagosLiqCajaListaDG.frx":785C
            FormatStyle(3)  =   "frmAdminPagosLiqCajaListaDG.frx":790C
            FormatStyle(4)  =   "frmAdminPagosLiqCajaListaDG.frx":79C0
            FormatStyle(5)  =   "frmAdminPagosLiqCajaListaDG.frx":7A98
            FormatStyle(6)  =   "frmAdminPagosLiqCajaListaDG.frx":7B50
            ImageCount      =   0
            PrinterProperties=   "frmAdminPagosLiqCajaListaDG.frx":7C30
         End
         Begin GridEX20.GridEX gridCajaOperaciones 
            Height          =   2055
            Left            =   -69880
            TabIndex        =   19
            Top             =   480
            Visible         =   0   'False
            Width           =   9930
            _ExtentX        =   17515
            _ExtentY        =   3625
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
            Column(1)       =   "frmAdminPagosLiqCajaListaDG.frx":7E08
            Column(2)       =   "frmAdminPagosLiqCajaListaDG.frx":7F68
            Column(3)       =   "frmAdminPagosLiqCajaListaDG.frx":80A4
            Column(4)       =   "frmAdminPagosLiqCajaListaDG.frx":81D8
            Column(5)       =   "frmAdminPagosLiqCajaListaDG.frx":830C
            FormatStylesCount=   6
            FormatStyle(1)  =   "frmAdminPagosLiqCajaListaDG.frx":8410
            FormatStyle(2)  =   "frmAdminPagosLiqCajaListaDG.frx":8548
            FormatStyle(3)  =   "frmAdminPagosLiqCajaListaDG.frx":85F8
            FormatStyle(4)  =   "frmAdminPagosLiqCajaListaDG.frx":86AC
            FormatStyle(5)  =   "frmAdminPagosLiqCajaListaDG.frx":8784
            FormatStyle(6)  =   "frmAdminPagosLiqCajaListaDG.frx":883C
            ImageCount      =   0
            PrinterProperties=   "frmAdminPagosLiqCajaListaDG.frx":891C
         End
      End
   End
   Begin XtremeSuiteControls.DateTimePicker dtpFecha 
      Height          =   375
      Left            =   10440
      TabIndex        =   21
      Top             =   360
      Width           =   1575
      _Version        =   786432
      _ExtentX        =   2778
      _ExtentY        =   661
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   1
      CurrentDate     =   45105.6793981481
   End
   Begin XtremeSuiteControls.PushButton btnConfirmarCbte 
      Height          =   495
      Index           =   0
      Left            =   5520
      TabIndex        =   32
      Top             =   6600
      Width           =   2295
      _Version        =   786432
      _ExtentX        =   4048
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Confirmar"
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
   Begin VB.Line Line 
      BorderColor     =   &H8000000B&
      X1              =   360
      X2              =   15600
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label lblCbtesConfirmados 
      Caption         =   "Comprobantes: "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8160
      TabIndex        =   34
      Top             =   1750
      Width           =   1815
   End
   Begin VB.Label lblComprobantesMostrados 
      Caption         =   "Comprobantes: "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   33
      Top             =   1750
      Width           =   1935
   End
   Begin VB.Label Label 
      Caption         =   "Fecha de Liquidación"
      Height          =   255
      Index           =   2
      Left            =   10440
      TabIndex        =   28
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label 
      Caption         =   "Nro Liquidación"
      Height          =   255
      Index           =   1
      Left            =   8160
      TabIndex        =   27
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label 
      Caption         =   "Filtrar por Nro de Comprobante"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   26
      Top             =   600
      Width           =   2775
   End
   Begin XtremeSuiteControls.Label lblLabel2 
      Height          =   375
      Left            =   8160
      TabIndex        =   24
      Top             =   840
      Width           =   3735
      _Version        =   786432
      _ExtentX        =   6588
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Total Valores Cargados: $ 0"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.Label lblLabel1 
      Height          =   375
      Left            =   8160
      TabIndex        =   23
      Top             =   1200
      Width           =   3615
      _Version        =   786432
      _ExtentX        =   6376
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Total Comprobantes: $ 0"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblMoneda 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Moneda"
      Height          =   195
      Left            =   17880
      TabIndex        =   14
      Tag             =   "Total: "
      Top             =   10020
      Visible         =   0   'False
      Width           =   570
   End
End
Attribute VB_Name = "frmAdminPagosLiqCajaListaDG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vId As String
Private Factura As clsFacturaProveedor
Private facturaConfirmada As clsFacturaProveedor
Private facturas As Collection
Private facturasConfirmadas As Collection
Dim m_Archivos As Dictionary

'Dim compe As Compensatorio

Private Banco As Banco
Private caja As caja
Private CuentaBancaria As CuentaBancaria
Private moneda As clsMoneda
Private cuentasBancarias As New Collection
Private Monedas As New Collection
Private Cajas As New Collection
Private bancos As New Collection
Private chequesDisponibles As New Collection
Private chequeras As New Collection

Private LiquidacionCaja As New clsLiquidacionCaja
Private operacion As operacion
Private cheque As cheque
Private tmpChequera As chequera

Private chequesChequeraSeleccionada As New Collection

Public ReadOnly As Boolean
Public esNueva As Boolean

Private Sub btnCargarCbtes_Click()
    llenarGrilla
End Sub


Private Sub btnConfirmar_Click()
    ' Verificar si se ha seleccionado una fila en el DataGrid
    If grilla.row < 0 Then
        MsgBox "No se ha seleccionado ninguna fila."
        Exit Sub
    End If

    ' Seleccionar la factura
    SeleccionarFactura

    ' Verificar si se asignó correctamente la factura
    If Not Factura Is Nothing Then
        ' Verificar si la colección facturasConfirmadas ya ha sido inicializada
        If facturasConfirmadas Is Nothing Then
            Set facturasConfirmadas = New Collection
        End If

        ' Verificar si la factura ya está en la colección facturasConfirmadas
        Dim facturaExistente As Boolean
        facturaExistente = False

        For Each fac In facturasConfirmadas
            If fac.numero = Factura.numero Then
                facturaExistente = True
                Exit For
            End If
        Next fac

        If facturaExistente Then
            MsgBox ("El comprobante " & Factura.NumeroFormateado & " ya existe en el listado de Cbtes. Confirmados")
        Else
            ' Agregar la factura seleccionada a la colección facturasConfirmadas
            facturasConfirmadas.Add Factura

            Me.txtFiltroNumero.SetFocus
            Me.txtFiltroNumero = ""

            grillaConfirmados.ItemCount = facturasConfirmadas.count

            llenarGrilla

            TotalizarComprobantes
        End If
    Else
        MsgBox "No se pudo seleccionar la factura."
    End If
End Sub


Private Sub btnExportarComprobantes_Click()

    If Me.grilla.ItemCount = 0 Then
    MsgBox ("No hay comprobantes para exportar!")
   Else
    ExportToXslComprobantes
        End If
    
End Sub

Private Sub btnExportarConfirmados_Click()
    
    If Me.grillaConfirmados.ItemCount = 0 Then
    MsgBox ("No hay comprobantes para exportar!")
   Else
       
    ExportToXslConfirmados
    End If
    

    
End Sub

Public Function ExportToXslConfirmados()

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

    xlWorksheet.Range("A1:G1").Merge
    xlWorksheet.Range("A2:G2").Merge
    xlWorksheet.Range("A1:G3").Font.Bold = True
    xlWorksheet.Cells(1, 1).value = "Detalle de Comprobante de Liquidación"
    xlWorksheet.Cells(3, 1).value = "Tipo"
    xlWorksheet.Cells(3, 2).value = "Letra"
    xlWorksheet.Cells(3, 3).value = "Nro Cbte"
    xlWorksheet.Cells(3, 4).value = "Fecha"
    xlWorksheet.Cells(3, 5).value = "Moneda"
    xlWorksheet.Cells(3, 6).value = "Monto"
    xlWorksheet.Cells(3, 7).value = "Proveedor"


    Dim idx As Integer
    idx = 4


    For Each facturaConfirmada In facturasConfirmadas


        xlWorksheet.Cells(idx, 1).value = enums.EnumTipoDocumentoContableShort(facturaConfirmada.tipoDocumentoContable)
        xlWorksheet.Cells(idx, 2).value = facturaConfirmada.configFactura.TipoFactura
        xlWorksheet.Cells(idx, 3).value = facturaConfirmada.numero
        xlWorksheet.Cells(idx, 4).value = facturaConfirmada.FEcha
        xlWorksheet.Cells(idx, 5).value = facturaConfirmada.moneda.NombreCorto
        
        Dim c As Integer
        
        If facturaConfirmada.tipoDocumentoContable = tipoDocumentoContable.notaCredito Then c = -1 Else c = 1
        
        xlWorksheet.Cells(idx, 6).value = facturaConfirmada.total * c
       
        xlWorksheet.Cells(idx, 7).value = UCase(funciones.RazonSocialFormateada(facturaConfirmada.Proveedor.RazonSocial))
        
        idx = idx + 1

    Next
    


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


End Function


Public Function ExportToXslComprobantes()

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

    xlWorksheet.Range("A1:G1").Merge
    xlWorksheet.Range("A2:G2").Merge
    xlWorksheet.Range("A1:G3").Font.Bold = True
    xlWorksheet.Cells(1, 1).value = "Detalle de Comprobante de Liquidación"
    xlWorksheet.Cells(3, 1).value = "Tipo"
    xlWorksheet.Cells(3, 2).value = "Letra"
    xlWorksheet.Cells(3, 3).value = "Nro Cbte"
    xlWorksheet.Cells(3, 4).value = "Fecha"
    xlWorksheet.Cells(3, 5).value = "Moneda"
    xlWorksheet.Cells(3, 6).value = "Monto"
    xlWorksheet.Cells(3, 7).value = "Proveedor"


    Dim idx As Integer
    idx = 4

    For Each Factura In facturas


        xlWorksheet.Cells(idx, 1).value = enums.EnumTipoDocumentoContableShort(Factura.tipoDocumentoContable)
        xlWorksheet.Cells(idx, 2).value = Factura.configFactura.TipoFactura
        xlWorksheet.Cells(idx, 3).value = Factura.numero
        xlWorksheet.Cells(idx, 4).value = Factura.FEcha
        xlWorksheet.Cells(idx, 5).value = Factura.moneda.NombreCorto
        
        Dim c As Integer
        
        If Factura.tipoDocumentoContable = tipoDocumentoContable.notaCredito Then c = -1 Else c = 1
        
        xlWorksheet.Cells(idx, 6).value = Factura.total * c
        
        xlWorksheet.Cells(idx, 7).value = UCase(funciones.RazonSocialFormateada(Factura.Proveedor.RazonSocial))
        
        idx = idx + 1

    Next

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


End Function

Private Sub btnLimpiarNúmero_Click()
    Me.txtFiltroNumero = ""
    
    llenarGrilla
    
End Sub
Private Sub SeleccionarFactura()
    On Error Resume Next
    Set Factura = facturas.item(grilla.rowIndex(grilla.row))
End Sub



Private Sub btnQuitarSeleccionado_Click()
    ' Verificar si se ha seleccionado una factura confirmada
    If facturaConfirmada Is Nothing Then
        MsgBox "No se ha seleccionado ninguna factura."
        Exit Sub
    End If
    
                    
    ' Preguntar al usuario para confirmar la eliminación
    Dim respuesta As Integer
    respuesta = MsgBox("¿Estás seguro de que deseas eliminar la factura " & facturaConfirmada.NumeroFormateado & "?", vbQuestion + vbYesNo, "Confirmar eliminación")
    
    If respuesta = vbYes Then
    
        Dim facturaAEliminar As clsFacturaProveedor
        Dim i As Integer
    
        ' Buscar la factura correspondiente en la colección
        For i = facturasConfirmadas.count To 1 Step -1
            If facturasConfirmadas(i).Id = facturaConfirmada.Id Then
                ' Encontrar la factura a eliminar
                Set facturaAEliminar = facturasConfirmadas(i)
                
                Dim q As String
                If facturaConfirmada.estado = Saldada Then
                    q = "UPDATE AdminComprasFacturasProveedores SET estado = 2 WHERE id = " & facturaConfirmada.Id
                    If Not conectar.execute(q) Then GoTo err1
                 End If
                
                Exit For
            End If
        Next i
    
        ' Verificar si se encontró la factura a eliminar
        If Not facturaAEliminar Is Nothing Then
            ' Eliminar la factura de la colección facturasConfirmadas
            facturasConfirmadas.remove i

            grillaConfirmados.ItemCount = 0
            grillaConfirmados.ItemCount = facturasConfirmadas.count
            
            TotalizarComprobantes
            
        Else
            MsgBox "No se encontró la factura a eliminar."
            grillaConfirmados.ItemCount = 0
            grillaConfirmados.ItemCount = facturasConfirmadas.count
        End If
    Else
err1:
    End If
End Sub






Private Sub Form_Load()
    Set m_Archivos = DAOArchivo.GetCantidadArchivosPorReferencia(OA_FacturaProveedor)
    vId = funciones.CreateGUID
    FormHelper.Customize Me

    GridEXHelper.CustomizeGrid Me.grilla, False
    GridEXHelper.CustomizeGrid Me.grillaConfirmados, False
    

    Me.grilla.ItemCount = 0
    Me.grillaConfirmados.ItemCount = 0

    Me.grilla.Refresh
    Me.grillaConfirmados.Refresh

    GridEXHelper.CustomizeGrid Me.gridCajaOperaciones, False, True
    GridEXHelper.CustomizeGrid Me.gridDepositosOperaciones, False, True
    '    GridEXHelper.CustomizeGrid Me.gridCheques, False, True
    GridEXHelper.CustomizeGrid Me.gridChequesDisponibles, False, False
    GridEXHelper.CustomizeGrid Me.gridBancos, False, False
    GridEXHelper.CustomizeGrid Me.gridCuentasBancarias, False, False
    GridEXHelper.CustomizeGrid Me.gridMonedas, False, False
    GridEXHelper.CustomizeGrid Me.gridCajas, False, False
    GridEXHelper.CustomizeGrid Me.gridChequeras, False, False
    '    GridEXHelper.CustomizeGrid Me.gridChequesPropios, False, True
    GridEXHelper.CustomizeGrid Me.gridCompensatorios, False, True
    GridEXHelper.CustomizeGrid Me.gridChequesChequera
    '    GridEXHelper.CustomizeGrid Me.gridRetenciones, False, True

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

    Me.gridCajaOperaciones.ItemCount = LiquidacionCaja.OperacionesCaja.count

    Me.gridDepositosOperaciones.ItemCount = LiquidacionCaja.OperacionesBanco.count

    Set Me.gridDepositosOperaciones.Columns("moneda").DropDownControl = Me.gridMonedas
    Set Me.gridDepositosOperaciones.Columns("cuenta").DropDownControl = Me.gridCuentasBancarias

    Set Me.gridCajaOperaciones.Columns("moneda").DropDownControl = Me.gridMonedas
    Set Me.gridCajaOperaciones.Columns("caja").DropDownControl = Me.gridCajas

    gridChequesChequera.ItemCount = 0
    GridEXHelper.AutoSizeColumns Me.gridChequeras
    
    llenarGrilla
    
    Me.caption = "Creación de Liquidación de Caja"

End Sub


Public Sub llenarGrilla()

    grilla.ItemCount = 0

    Dim condition As String

    condition = " AdminComprasFacturasProveedores.estado = 2 "

    Set facturas = DAOFacturaProveedor.FindAll(condition, , , Permisos.AdminFaPVerSoloPropias)

    Dim F As clsFacturaProveedor
    Dim total As Double
    Dim totalneto As Double
    Dim totIva As Double
    Dim totalno As Double
    Dim totalpercep As Double
    Dim totalsaldo As Double

    Dim c As Integer

    total = 0

    For Each F In facturas

        If F.tipoDocumentoContable = tipoDocumentoContable.notaCredito Then c = -1 Else c = 1
        total = total + MonedaConverter.Convertir(F.total * c, F.moneda.Id, MonedaConverter.Patron.Id)
        totalneto = totalneto + MonedaConverter.Convertir(F.Monto * c - F.TotalNetoGravadoDiscriminado(0) * c, F.moneda.Id, MonedaConverter.Patron.Id)
        totalno = totalno + MonedaConverter.Convertir(F.TotalNetoGravadoDiscriminado(0) * c, F.moneda.Id, MonedaConverter.Patron.Id)
        totIva = totIva + MonedaConverter.Convertir(F.TotalIVA * c, F.moneda.Id, MonedaConverter.Patron.Id)
        totalpercep = totalpercep + F.totalPercepciones * c
        totalsaldo = totalsaldo + ((F.total - (F.NetoGravadoAbonadoGlobal + F.OtrosAbonadoGlobal)) * c)

    Next

    grilla.ItemCount = facturas.count

'    Me.caption = "Cbtes. filtrados [Cantidad: " & facturas.count & "]"
    
    lblComprobantesMostrados.caption = "Comprobantes: " & facturas.count
    
End Sub


Private Sub grilla_ColumnHeaderClick(ByVal Column As GridEX20.JSColumn)
    GridEXHelper.ColumnHeaderClick Me.grilla, Column

End Sub



Private Sub grilla_FetchIcon(ByVal rowIndex As Long, ByVal ColIndex As Integer, ByVal RowBookmark As Variant, ByVal IconIndex As GridEX20.JSRetInteger)
    If ColIndex = 15 And m_Archivos.item(Factura.Id) > 0 Then IconIndex = 1

End Sub


' LLENADO DE GRILLA DE COMPROBANTES APROBADOS PARA PAGAR
Private Sub grilla_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)

    Set Factura = facturas.item(rowIndex)

    Dim i As Integer

    If Factura.tipoDocumentoContable = tipoDocumentoContable.notaCredito Then i = -1 Else i = 1

    With Factura

        Values(1) = enums.EnumTipoDocumentoContableShort(Factura.tipoDocumentoContable)
        Values(2) = Factura.configFactura.TipoFactura
        Values(3) = Factura.numero
        Values(4) = Factura.FEcha
        Values(5) = Factura.moneda.NombreCorto
        Values(6) = Replace(FormatCurrency(funciones.FormatearDecimales(Factura.total) * i), "$", "")
        Values(7) = UCase(funciones.RazonSocialFormateada(Factura.Proveedor.RazonSocial))

    End With

End Sub



Private Sub grillaConfirmados_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    SeleccionarFactura

End Sub

' LLENADO DE GRILLA DE COMPROBANTES CONFIRMADOS
Private Sub grillaConfirmados_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)

    Set facturaConfirmada = facturasConfirmadas.item(rowIndex)

    Dim i As Integer

    If facturaConfirmada.tipoDocumentoContable = tipoDocumentoContable.notaCredito Then i = -1 Else i = 1

    With facturaConfirmada

        Values(1) = enums.EnumTipoDocumentoContableShort(facturaConfirmada.tipoDocumentoContable)
        Values(2) = facturaConfirmada.configFactura.TipoFactura
        Values(3) = facturaConfirmada.numero
        Values(4) = facturaConfirmada.FEcha
        Values(5) = facturaConfirmada.moneda.NombreCorto
        Values(6) = Replace(FormatCurrency(funciones.FormatearDecimales(facturaConfirmada.total) * i), "$", "")
        Values(7) = UCase(funciones.RazonSocialFormateada(facturaConfirmada.Proveedor.RazonSocial))

    End With
    ' La variable facturasConfirmadas no ha sido inicializada correctamente
    ' Realiza las acciones necesarias para manejar este caso de error.

End Sub

'ESTA FUNCION SE ENCARGA DE GUARDAR LA LIQUIDACIÓN QUE SE ESTÁ CREANDO

Private Sub PusGuardar_Click()

    If Me.gridCajaOperaciones.EditMode = jgexEditModeOn Then
        MsgBox "Todavia esta editando la grilla de caja.", vbExclamation
        Exit Sub
    End If

    If Me.gridDepositosOperaciones.EditMode = jgexEditModeOn Then
        MsgBox "Todavia esta editando la grilla de banco.", vbExclamation
        Exit Sub
    End If

    LiquidacionCaja.FEcha = Me.dtpFecha.value

    If Me.txtNumerodeLiquidacion.text = "" Then
        MsgBox ("El número de Liquidación no puede estar vacío.")
        Exit Sub
    Else
        LiquidacionCaja.NumeroLiq = Me.txtNumerodeLiquidacion.text

    End If

    Set LiquidacionCaja.FacturasProveedor = New Collection


    For Each Factura In facturasConfirmadas
        LiquidacionCaja.FacturasProveedor.Add Factura
    Next


    If LiquidacionCaja.IsValid Then

        Dim n As Boolean: n = (LiquidacionCaja.Id = 0)

        If DAOLiquidacionCaja.Save(LiquidacionCaja, True) Then

            If n Then
                MsgBox "Liquidación de Caja Nº " & Me.txtNumerodeLiquidacion & " creada con exito.", vbInformation
            Else
                MsgBox "Liquidación de Caja modificada con exito.", vbInformation
            End If

            If n Then
                If MsgBox("¿Desea crear una Liquidación de Caja nueva", vbQuestion + vbYesNo) = vbYes Then
                    Dim f12 As New frmAdminPagosLiqCajaListaDG
                    f12.Show
                End If
            End If

            Unload Me
        Else
            MsgBox "Hubo un problema al guardar la Liquidación.", vbCritical
        End If
    Else
        MsgBox LiquidacionCaja.ValidationMessages, vbCritical, "Error"
    End If
End Sub


'FUNCION PARA CALCULAR LOS TOTALES DE LOS COMPROBANTES DEL DATAGRID CONFIRMADOS
Public Sub TotalizarComprobantes()
    Dim total As Double
    Dim i As Integer
    Dim Factura As clsFacturaProveedor

    For i = 1 To facturasConfirmadas.count
        Set Factura = facturasConfirmadas.item(i)

        If Factura.tipoDocumentoContable = tipoDocumentoContable.notaCredito Then
            total = total - Factura.total    ' Resta el total de las facturas tipo nota de crédito
        Else
            total = total + Factura.total    ' Suma el total de las demás facturas
        End If
    Next i

    LiquidacionCaja.StaticTotalFacturas = funciones.RedondearDecimales(total)

    lblLabel1.caption = "Total Comprobantes: " & FormatCurrency(funciones.FormatearDecimales(total))
    
    lblCbtesConfirmados.caption = "Comprobantes: " & facturasConfirmadas.count

End Sub


'TODAS LAS FUNCIONES QUE VIENEN DE LA LIQUIDACION DE CAJA ANTERIOR
'Private Sub CargarChequesDisponibles()
'    Set chequesDisponibles = DAOCheques.FindAllEnCarteraDeTerceros
'    Me.gridChequesDisponibles.ItemCount = chequesDisponibles.count
'End Sub


Private Sub gridBancos_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If rowIndex <= bancos.count Then
        Set Banco = bancos.item(rowIndex)
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

Private Sub gridCajas_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If rowIndex > 0 And Cajas.count > 0 Then
        Set caja = Cajas.item(rowIndex)
        Values(1) = caja.Id
        Values(2) = caja.nombre
    End If
End Sub


Private Sub gridChequeras_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If rowIndex <= chequeras.count Then
        Set tmpChequera = chequeras.item(rowIndex)
        Values(1) = tmpChequera.Description
        Values(2) = tmpChequera.Id
    End If
End Sub


Private Sub gridChequesChequera_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If rowIndex > 0 And chequesChequeraSeleccionada.count > 0 Then
        Values(1) = chequesChequeraSeleccionada(rowIndex).numero
        Values(2) = chequesChequeraSeleccionada(rowIndex).Id
    End If
End Sub

'Private Sub gridChequesDisponibles_ColumnHeaderClick(ByVal Column As GridEX20.JSColumn)
'    GridEXHelper.ColumnHeaderClick Me.gridChequesDisponibles, Column
'End Sub

Private Sub gridChequesDisponibles_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If rowIndex <= chequesDisponibles.count Then
        Set cheque = chequesDisponibles.item(rowIndex)
        Values(1) = cheque.numero
        'FORMATCURRENCY
        Values(2) = FormatCurrency(cheque.Monto)
        If IsSomething(cheque.moneda) Then Values(3) = cheque.moneda.NombreCorto
        If IsSomething(cheque.Banco) Then Values(4) = cheque.Banco.nombre
        Values(5) = cheque.Id
        Values(6) = cheque.OrigenCheque
        Values(7) = cheque.OrigenDestino

    End If

End Sub


Private Sub gridCuentasBancarias_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If cuentasBancarias.count >= rowIndex Then
        Set CuentaBancaria = cuentasBancarias.item(rowIndex)
        Values(1) = CuentaBancaria.Id
        Values(2) = CuentaBancaria.DescripcionFormateada
    End If
End Sub

Private Sub gridMonedas_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If rowIndex > 0 And Monedas.count > 0 Then
        Set moneda = Monedas.item(rowIndex)
        Values(1) = moneda.Id
        Values(2) = moneda.NombreCorto
    End If
End Sub


'ESTA FUNCION ES LA QUE CARGA LOS REGISTROS DE VALORES CAJA EN UNA OPERACION DENTRO DE UNA LIQUIDACION DE CAJA
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
    LiquidacionCaja.OperacionesCaja.Add operacion

    Totalizar

End Sub

Private Sub gridCajaOperaciones_UnboundDelete(ByVal rowIndex As Long, ByVal Bookmark As Variant)
    If rowIndex > 0 And LiquidacionCaja.OperacionesCaja.count >= rowIndex Then
        LiquidacionCaja.OperacionesCaja.remove rowIndex

        Totalizar

    End If
End Sub

Private Sub gridCajaOperaciones_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If rowIndex <= LiquidacionCaja.OperacionesCaja.count Then
        Set operacion = LiquidacionCaja.OperacionesCaja.item(rowIndex)
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


Private Sub gridCajaOperaciones_UnboundUpdate(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If rowIndex > 0 And LiquidacionCaja.OperacionesCaja.count > 0 Then
        Set operacion = LiquidacionCaja.OperacionesCaja.item(rowIndex)
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

        Totalizar

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
    LiquidacionCaja.OperacionesBanco.Add operacion

    Totalizar

End Sub

Private Sub gridDepositosOperaciones_UnboundDelete(ByVal rowIndex As Long, ByVal Bookmark As Variant)
    If rowIndex > 0 And LiquidacionCaja.OperacionesBanco.count >= rowIndex Then
        LiquidacionCaja.OperacionesBanco.remove rowIndex

        Totalizar

    End If
End Sub

Private Sub gridDepositosOperaciones_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If rowIndex <= LiquidacionCaja.OperacionesBanco.count Then
        Set operacion = LiquidacionCaja.OperacionesBanco.item(rowIndex)
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

Private Sub gridDepositosOperaciones_UnboundUpdate(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If rowIndex > 0 And LiquidacionCaja.OperacionesBanco.count > 0 Then
        Set operacion = LiquidacionCaja.OperacionesBanco.item(rowIndex)
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

        Totalizar

    End If
End Sub




Private Sub TabControl_SelectedChanged(ByVal item As Xtremesuitecontrols.ITabControlItem)
    Me.TabControl.TabIndex = 0
    
End Sub


Private Sub Totalizar()
    LiquidacionCaja.StaticTotalOrigenes = LiquidacionCaja.TotalOrigenes
    Me.lblLabel2.caption = "Total Valores Cargados: " & FormatCurrency(funciones.FormatearDecimales(LiquidacionCaja.StaticTotalOrigenes + LiquidacionCaja.StaticTotalRetenido))
    GridEXHelper.AutoSizeColumns Me.gridCajaOperaciones
    GridEXHelper.AutoSizeColumns Me.gridDepositosOperaciones

End Sub

Private Sub txtFiltroNumero_Change()
    Dim filterText As String
    filterText = Trim(Me.txtFiltroNumero.text)
    
        ' Limpiar la grilla
    grilla.ItemCount = 0
    
    ' Filtrar las facturas en base al texto ingresado
    Dim facturasFiltradas As New Collection
    Dim Factura As clsFacturaProveedor
    
    For Each Factura In facturas
        If InStr(1, Factura.NumeroFormateado, filterText, vbTextCompare) > 0 Then
            facturasFiltradas.Add Factura
        End If
    Next
    
    grilla.ItemCount = 0
    
    Set facturas = facturasFiltradas
        
    grilla.ItemCount = facturasFiltradas.count
    
End Sub


'''''''''''''''''''''''''''''''''''

Public Sub Cargar(liq As clsLiquidacionCaja)

    Me.gridDepositosOperaciones.AllowEdit = Not ReadOnly
    Me.gridDepositosOperaciones.AllowDelete = Not ReadOnly
    Me.gridBancos.AllowEdit = Not ReadOnly
    Me.gridCajaOperaciones.AllowEdit = Not ReadOnly
    Me.gridCajaOperaciones.AllowDelete = Not ReadOnly
    Me.gridCajas.AllowEdit = Not ReadOnly
    Me.gridChequeras.AllowEdit = Not ReadOnly
    Me.gridChequesChequera.AllowEdit = Not ReadOnly
    Me.gridChequesDisponibles.AllowEdit = Not ReadOnly

    Me.PusGuardar.Enabled = Not ReadOnly
    Me.btnQuitarSeleccionado.Enabled = Not ReadOnly
    Me.btnCargarCbtes.Enabled = Not ReadOny
    Me.btnLimpiarNúmero.Enabled = Not ReadOnly
    Me.txtFiltroNumero.Enabled = Not ReadOnly
    Me.txtNumerodeLiquidacion.Enabled = Not ReadOnly
    Me.dtpFecha.Enabled = Not ReadOnly
    Me.btnExportarComprobantes.Enabled = Not ReadOnly
    Me.btnConfirmar.Enabled = Not ReadOnly
    Me.btnCargarCbtes.Enabled = Not ReadOnly
    

    If Not IsSomething(liq) Then
        MsgBox "La Liquidación que está intentando visualizar está en estado PENDIENTE. " & vbNewLine & "Por lo tanto no puede ser mostrada porque puede estar siendo editada." & vbNewLine & "Verifiquelo por favor.", vbCritical, "OP Pendiente"
        Unload Me
        Exit Sub
    End If

    Set LiquidacionCaja = DAOLiquidacionCaja.FindById(liq.Id)

    Set facturasConfirmadas = DAOFacturaProveedor.FindAllByLiquidacionCaja(liq.Id)

    Me.grillaConfirmados.ItemCount = 0
'    Me.grillaConfirmados.ItemCount = liq.FacturasProveedor.count
    
    Me.grillaConfirmados.ItemCount = facturasConfirmadas.count
    
    Me.gridCajaOperaciones.ItemCount = LiquidacionCaja.OperacionesCaja.count
    Me.gridDepositosOperaciones.ItemCount = LiquidacionCaja.OperacionesBanco.count

    Me.dtpFecha.value = LiquidacionCaja.FEcha
    Me.txtOtrosDescuentos.text = LiquidacionCaja.OtrosDescuentos

    Me.caption = "Liquidación Nº " & LiquidacionCaja.NumeroLiq

    Me.txtNumerodeLiquidacion = LiquidacionCaja.NumeroLiq

    Totalizar
    TotalizarComprobantes

    esNueva = False

End Sub





