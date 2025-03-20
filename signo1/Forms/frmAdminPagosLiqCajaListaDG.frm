VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmAdminPagosLiqCajaListaDG 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   10695
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15780
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
   ScaleHeight     =   10695
   ScaleWidth      =   15780
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
      Left            =   17760
      TabIndex        =   5
      Top             =   12000
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
      ColumnsCount    =   8
      Column(1)       =   "frmAdminPagosLiqCajaListaDG.frx":0000
      Column(2)       =   "frmAdminPagosLiqCajaListaDG.frx":0188
      Column(3)       =   "frmAdminPagosLiqCajaListaDG.frx":02EC
      Column(4)       =   "frmAdminPagosLiqCajaListaDG.frx":0434
      Column(5)       =   "frmAdminPagosLiqCajaListaDG.frx":0574
      Column(6)       =   "frmAdminPagosLiqCajaListaDG.frx":06BC
      Column(7)       =   "frmAdminPagosLiqCajaListaDG.frx":07FC
      Column(8)       =   "frmAdminPagosLiqCajaListaDG.frx":093C
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminPagosLiqCajaListaDG.frx":0A40
      FormatStyle(2)  =   "frmAdminPagosLiqCajaListaDG.frx":0B68
      FormatStyle(3)  =   "frmAdminPagosLiqCajaListaDG.frx":0C18
      FormatStyle(4)  =   "frmAdminPagosLiqCajaListaDG.frx":0CCC
      FormatStyle(5)  =   "frmAdminPagosLiqCajaListaDG.frx":0DA4
      FormatStyle(6)  =   "frmAdminPagosLiqCajaListaDG.frx":0E5C
      ImageCount      =   0
      PrinterProperties=   "frmAdminPagosLiqCajaListaDG.frx":0F3C
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
      Column(1)       =   "frmAdminPagosLiqCajaListaDG.frx":110C
      Column(2)       =   "frmAdminPagosLiqCajaListaDG.frx":124C
      Column(3)       =   "frmAdminPagosLiqCajaListaDG.frx":1384
      Column(4)       =   "frmAdminPagosLiqCajaListaDG.frx":14CC
      Column(5)       =   "frmAdminPagosLiqCajaListaDG.frx":160C
      Column(6)       =   "frmAdminPagosLiqCajaListaDG.frx":1754
      Column(7)       =   "frmAdminPagosLiqCajaListaDG.frx":1894
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminPagosLiqCajaListaDG.frx":1990
      FormatStyle(2)  =   "frmAdminPagosLiqCajaListaDG.frx":1AB8
      FormatStyle(3)  =   "frmAdminPagosLiqCajaListaDG.frx":1B68
      FormatStyle(4)  =   "frmAdminPagosLiqCajaListaDG.frx":1C1C
      FormatStyle(5)  =   "frmAdminPagosLiqCajaListaDG.frx":1CF4
      FormatStyle(6)  =   "frmAdminPagosLiqCajaListaDG.frx":1DAC
      ImageCount      =   0
      PrinterProperties=   "frmAdminPagosLiqCajaListaDG.frx":1E8C
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
      Height          =   1245
      Left            =   10080
      TabIndex        =   6
      Top             =   10560
      Visible         =   0   'False
      Width           =   6705
      _ExtentX        =   11827
      _ExtentY        =   2196
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
      Column(1)       =   "frmAdminPagosLiqCajaListaDG.frx":205C
      Column(2)       =   "frmAdminPagosLiqCajaListaDG.frx":215C
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminPagosLiqCajaListaDG.frx":224C
      FormatStyle(2)  =   "frmAdminPagosLiqCajaListaDG.frx":2384
      FormatStyle(3)  =   "frmAdminPagosLiqCajaListaDG.frx":2434
      FormatStyle(4)  =   "frmAdminPagosLiqCajaListaDG.frx":24E8
      FormatStyle(5)  =   "frmAdminPagosLiqCajaListaDG.frx":25C0
      FormatStyle(6)  =   "frmAdminPagosLiqCajaListaDG.frx":2678
      ImageCount      =   0
      PrinterProperties=   "frmAdminPagosLiqCajaListaDG.frx":2758
   End
   Begin GridEX20.GridEX gridCuentasBancarias 
      Height          =   1335
      Left            =   10680
      TabIndex        =   7
      Top             =   11520
      Visible         =   0   'False
      Width           =   6225
      _ExtentX        =   10980
      _ExtentY        =   2355
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
      Column(1)       =   "frmAdminPagosLiqCajaListaDG.frx":2930
      Column(2)       =   "frmAdminPagosLiqCajaListaDG.frx":2A54
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminPagosLiqCajaListaDG.frx":2B48
      FormatStyle(2)  =   "frmAdminPagosLiqCajaListaDG.frx":2C80
      FormatStyle(3)  =   "frmAdminPagosLiqCajaListaDG.frx":2D30
      FormatStyle(4)  =   "frmAdminPagosLiqCajaListaDG.frx":2DE4
      FormatStyle(5)  =   "frmAdminPagosLiqCajaListaDG.frx":2EBC
      FormatStyle(6)  =   "frmAdminPagosLiqCajaListaDG.frx":2F74
      ImageCount      =   0
      PrinterProperties=   "frmAdminPagosLiqCajaListaDG.frx":3054
   End
   Begin GridEX20.GridEX gridMonedas 
      Height          =   1215
      Left            =   15120
      TabIndex        =   8
      Top             =   10920
      Visible         =   0   'False
      Width           =   1380
      _ExtentX        =   2434
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
      Column(1)       =   "frmAdminPagosLiqCajaListaDG.frx":322C
      Column(2)       =   "frmAdminPagosLiqCajaListaDG.frx":3350
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminPagosLiqCajaListaDG.frx":3444
      FormatStyle(2)  =   "frmAdminPagosLiqCajaListaDG.frx":357C
      FormatStyle(3)  =   "frmAdminPagosLiqCajaListaDG.frx":362C
      FormatStyle(4)  =   "frmAdminPagosLiqCajaListaDG.frx":36E0
      FormatStyle(5)  =   "frmAdminPagosLiqCajaListaDG.frx":37B8
      FormatStyle(6)  =   "frmAdminPagosLiqCajaListaDG.frx":3870
      ImageCount      =   0
      PrinterProperties=   "frmAdminPagosLiqCajaListaDG.frx":3950
   End
   Begin GridEX20.GridEX gridCajas 
      Height          =   1935
      Left            =   17880
      TabIndex        =   9
      Top             =   8280
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
      Column(1)       =   "frmAdminPagosLiqCajaListaDG.frx":3B28
      Column(2)       =   "frmAdminPagosLiqCajaListaDG.frx":3C28
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminPagosLiqCajaListaDG.frx":3D14
      FormatStyle(2)  =   "frmAdminPagosLiqCajaListaDG.frx":3E4C
      FormatStyle(3)  =   "frmAdminPagosLiqCajaListaDG.frx":3EFC
      FormatStyle(4)  =   "frmAdminPagosLiqCajaListaDG.frx":3FB0
      FormatStyle(5)  =   "frmAdminPagosLiqCajaListaDG.frx":4088
      FormatStyle(6)  =   "frmAdminPagosLiqCajaListaDG.frx":4140
      ImageCount      =   0
      PrinterProperties=   "frmAdminPagosLiqCajaListaDG.frx":4220
   End
   Begin GridEX20.GridEX gridChequesDisponibles 
      Height          =   1320
      Left            =   8880
      TabIndex        =   10
      Top             =   10800
      Visible         =   0   'False
      Width           =   6675
      _ExtentX        =   11774
      _ExtentY        =   2328
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
      Column(1)       =   "frmAdminPagosLiqCajaListaDG.frx":43F8
      Column(2)       =   "frmAdminPagosLiqCajaListaDG.frx":4578
      Column(3)       =   "frmAdminPagosLiqCajaListaDG.frx":4718
      Column(4)       =   "frmAdminPagosLiqCajaListaDG.frx":4854
      Column(5)       =   "frmAdminPagosLiqCajaListaDG.frx":4960
      Column(6)       =   "frmAdminPagosLiqCajaListaDG.frx":4A80
      Column(7)       =   "frmAdminPagosLiqCajaListaDG.frx":4B8C
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminPagosLiqCajaListaDG.frx":4C80
      FormatStyle(2)  =   "frmAdminPagosLiqCajaListaDG.frx":4DB8
      FormatStyle(3)  =   "frmAdminPagosLiqCajaListaDG.frx":4E68
      FormatStyle(4)  =   "frmAdminPagosLiqCajaListaDG.frx":4F1C
      FormatStyle(5)  =   "frmAdminPagosLiqCajaListaDG.frx":4FF4
      FormatStyle(6)  =   "frmAdminPagosLiqCajaListaDG.frx":50AC
      ImageCount      =   0
      PrinterProperties=   "frmAdminPagosLiqCajaListaDG.frx":518C
   End
   Begin GridEX20.GridEX gridChequeras 
      Height          =   1095
      Left            =   6720
      TabIndex        =   11
      Top             =   12480
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   1931
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
      Column(1)       =   "frmAdminPagosLiqCajaListaDG.frx":5364
      Column(2)       =   "frmAdminPagosLiqCajaListaDG.frx":5484
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminPagosLiqCajaListaDG.frx":5584
      FormatStyle(2)  =   "frmAdminPagosLiqCajaListaDG.frx":56BC
      FormatStyle(3)  =   "frmAdminPagosLiqCajaListaDG.frx":576C
      FormatStyle(4)  =   "frmAdminPagosLiqCajaListaDG.frx":5820
      FormatStyle(5)  =   "frmAdminPagosLiqCajaListaDG.frx":58F8
      FormatStyle(6)  =   "frmAdminPagosLiqCajaListaDG.frx":59B0
      ImageCount      =   0
      PrinterProperties=   "frmAdminPagosLiqCajaListaDG.frx":5A90
   End
   Begin GridEX20.GridEX gridChequesChequera 
      Height          =   1335
      Left            =   11880
      TabIndex        =   12
      Top             =   12480
      Width           =   6225
      _ExtentX        =   10980
      _ExtentY        =   2355
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
      Column(1)       =   "frmAdminPagosLiqCajaListaDG.frx":5C68
      Column(2)       =   "frmAdminPagosLiqCajaListaDG.frx":5D98
      SortKeysCount   =   1
      SortKey(1)      =   "frmAdminPagosLiqCajaListaDG.frx":5E98
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminPagosLiqCajaListaDG.frx":5F00
      FormatStyle(2)  =   "frmAdminPagosLiqCajaListaDG.frx":6038
      FormatStyle(3)  =   "frmAdminPagosLiqCajaListaDG.frx":60E8
      FormatStyle(4)  =   "frmAdminPagosLiqCajaListaDG.frx":619C
      FormatStyle(5)  =   "frmAdminPagosLiqCajaListaDG.frx":6274
      FormatStyle(6)  =   "frmAdminPagosLiqCajaListaDG.frx":632C
      ImageCount      =   0
      PrinterProperties=   "frmAdminPagosLiqCajaListaDG.frx":640C
   End
   Begin XtremeSuiteControls.RadioButton RadSeleccioneProveedor 
      Height          =   210
      Left            =   15720
      TabIndex        =   13
      Top             =   10320
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
         Height          =   2940
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   15180
         _Version        =   786432
         _ExtentX        =   26776
         _ExtentY        =   5186
         _StockProps     =   68
         Appearance      =   10
         Color           =   32
         PaintManager.ShowIcons=   -1  'True
         ItemCount       =   4
         SelectedItem    =   1
         Item(0).Caption =   "Cheques Propios"
         Item(0).ControlCount=   1
         Item(0).Control(0)=   "gridChequesPropios"
         Item(1).Caption =   "Cheques 3ros"
         Item(1).ControlCount=   2
         Item(1).Control(0)=   "gridNadaCheques(0)"
         Item(1).Control(1)=   "gridCheques"
         Item(2).Caption =   "Caja"
         Item(2).ControlCount=   2
         Item(2).Control(0)=   "gridCajaOperaciones"
         Item(2).Control(1)=   "gridNadaCajaOperaciones"
         Item(3).Caption =   "Banco"
         Item(3).ControlCount=   2
         Item(3).Control(0)=   "gridCompensatorios"
         Item(3).Control(1)=   "gridDepositosOperaciones"
         Begin GridEX20.GridEX gridCompensatorios 
            Height          =   4710
            Left            =   -1.39895e5
            TabIndex        =   17
            Top             =   435
            Visible         =   0   'False
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
            Column(1)       =   "frmAdminPagosLiqCajaListaDG.frx":65E4
            Column(2)       =   "frmAdminPagosLiqCajaListaDG.frx":672C
            Column(3)       =   "frmAdminPagosLiqCajaListaDG.frx":6838
            Column(4)       =   "frmAdminPagosLiqCajaListaDG.frx":6924
            Column(5)       =   "frmAdminPagosLiqCajaListaDG.frx":6A28
            FormatStylesCount=   6
            FormatStyle(1)  =   "frmAdminPagosLiqCajaListaDG.frx":6B68
            FormatStyle(2)  =   "frmAdminPagosLiqCajaListaDG.frx":6CA0
            FormatStyle(3)  =   "frmAdminPagosLiqCajaListaDG.frx":6D50
            FormatStyle(4)  =   "frmAdminPagosLiqCajaListaDG.frx":6E04
            FormatStyle(5)  =   "frmAdminPagosLiqCajaListaDG.frx":6EDC
            FormatStyle(6)  =   "frmAdminPagosLiqCajaListaDG.frx":6F94
            ImageCount      =   0
            PrinterProperties=   "frmAdminPagosLiqCajaListaDG.frx":7074
         End
         Begin GridEX20.GridEX gridDepositosOperaciones 
            Height          =   2295
            Left            =   -69880
            TabIndex        =   18
            Top             =   480
            Visible         =   0   'False
            Width           =   14850
            _ExtentX        =   26194
            _ExtentY        =   4048
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
            Column(1)       =   "frmAdminPagosLiqCajaListaDG.frx":724C
            Column(2)       =   "frmAdminPagosLiqCajaListaDG.frx":73AC
            Column(3)       =   "frmAdminPagosLiqCajaListaDG.frx":74CC
            Column(4)       =   "frmAdminPagosLiqCajaListaDG.frx":75E4
            Column(5)       =   "frmAdminPagosLiqCajaListaDG.frx":770C
            FormatStylesCount=   6
            FormatStyle(1)  =   "frmAdminPagosLiqCajaListaDG.frx":7810
            FormatStyle(2)  =   "frmAdminPagosLiqCajaListaDG.frx":7948
            FormatStyle(3)  =   "frmAdminPagosLiqCajaListaDG.frx":79F8
            FormatStyle(4)  =   "frmAdminPagosLiqCajaListaDG.frx":7AAC
            FormatStyle(5)  =   "frmAdminPagosLiqCajaListaDG.frx":7B84
            FormatStyle(6)  =   "frmAdminPagosLiqCajaListaDG.frx":7C3C
            ImageCount      =   0
            PrinterProperties=   "frmAdminPagosLiqCajaListaDG.frx":7D1C
         End
         Begin GridEX20.GridEX gridNadaCajaOperaciones 
            Height          =   2295
            Left            =   -2.09880e5
            TabIndex        =   19
            Top             =   480
            Visible         =   0   'False
            Width           =   9930
            _ExtentX        =   17515
            _ExtentY        =   4048
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
            Column(1)       =   "frmAdminPagosLiqCajaListaDG.frx":7EF4
            Column(2)       =   "frmAdminPagosLiqCajaListaDG.frx":8054
            Column(3)       =   "frmAdminPagosLiqCajaListaDG.frx":8190
            Column(4)       =   "frmAdminPagosLiqCajaListaDG.frx":82C4
            Column(5)       =   "frmAdminPagosLiqCajaListaDG.frx":83F8
            FormatStylesCount=   6
            FormatStyle(1)  =   "frmAdminPagosLiqCajaListaDG.frx":84FC
            FormatStyle(2)  =   "frmAdminPagosLiqCajaListaDG.frx":8634
            FormatStyle(3)  =   "frmAdminPagosLiqCajaListaDG.frx":86E4
            FormatStyle(4)  =   "frmAdminPagosLiqCajaListaDG.frx":8798
            FormatStyle(5)  =   "frmAdminPagosLiqCajaListaDG.frx":8870
            FormatStyle(6)  =   "frmAdminPagosLiqCajaListaDG.frx":8928
            ImageCount      =   0
            PrinterProperties=   "frmAdminPagosLiqCajaListaDG.frx":8A08
         End
         Begin GridEX20.GridEX gridNadaCheques 
            Height          =   2295
            Index           =   0
            Left            =   -1.39880e5
            TabIndex        =   35
            Top             =   480
            Visible         =   0   'False
            Width           =   9330
            _ExtentX        =   16457
            _ExtentY        =   4048
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
            Column(1)       =   "frmAdminPagosLiqCajaListaDG.frx":8BE0
            Column(2)       =   "frmAdminPagosLiqCajaListaDG.frx":8D60
            Column(3)       =   "frmAdminPagosLiqCajaListaDG.frx":8F00
            Column(4)       =   "frmAdminPagosLiqCajaListaDG.frx":8FF8
            Column(5)       =   "frmAdminPagosLiqCajaListaDG.frx":9134
            Column(6)       =   "frmAdminPagosLiqCajaListaDG.frx":9240
            Column(7)       =   "frmAdminPagosLiqCajaListaDG.frx":9310
            FormatStylesCount=   6
            FormatStyle(1)  =   "frmAdminPagosLiqCajaListaDG.frx":93FC
            FormatStyle(2)  =   "frmAdminPagosLiqCajaListaDG.frx":9534
            FormatStyle(3)  =   "frmAdminPagosLiqCajaListaDG.frx":95E4
            FormatStyle(4)  =   "frmAdminPagosLiqCajaListaDG.frx":9698
            FormatStyle(5)  =   "frmAdminPagosLiqCajaListaDG.frx":9770
            FormatStyle(6)  =   "frmAdminPagosLiqCajaListaDG.frx":9828
            ImageCount      =   0
            PrinterProperties=   "frmAdminPagosLiqCajaListaDG.frx":9908
         End
         Begin GridEX20.GridEX gridChequesPropios 
            Height          =   2295
            Left            =   -69880
            TabIndex        =   36
            Top             =   480
            Visible         =   0   'False
            Width           =   14850
            _ExtentX        =   26194
            _ExtentY        =   4048
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
            Column(1)       =   "frmAdminPagosLiqCajaListaDG.frx":9AE0
            Column(2)       =   "frmAdminPagosLiqCajaListaDG.frx":9C48
            Column(3)       =   "frmAdminPagosLiqCajaListaDG.frx":9D7C
            Column(4)       =   "frmAdminPagosLiqCajaListaDG.frx":9EB8
            Column(5)       =   "frmAdminPagosLiqCajaListaDG.frx":A020
            FormatStylesCount=   6
            FormatStyle(1)  =   "frmAdminPagosLiqCajaListaDG.frx":A118
            FormatStyle(2)  =   "frmAdminPagosLiqCajaListaDG.frx":A250
            FormatStyle(3)  =   "frmAdminPagosLiqCajaListaDG.frx":A300
            FormatStyle(4)  =   "frmAdminPagosLiqCajaListaDG.frx":A3B4
            FormatStyle(5)  =   "frmAdminPagosLiqCajaListaDG.frx":A48C
            FormatStyle(6)  =   "frmAdminPagosLiqCajaListaDG.frx":A544
            ImageCount      =   0
            PrinterProperties=   "frmAdminPagosLiqCajaListaDG.frx":A624
         End
         Begin GridEX20.GridEX gridCajaOperaciones 
            Height          =   2295
            Left            =   -69880
            TabIndex        =   38
            Top             =   480
            Visible         =   0   'False
            Width           =   14850
            _ExtentX        =   26194
            _ExtentY        =   4048
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
            Column(1)       =   "frmAdminPagosLiqCajaListaDG.frx":A7FC
            Column(2)       =   "frmAdminPagosLiqCajaListaDG.frx":A95C
            Column(3)       =   "frmAdminPagosLiqCajaListaDG.frx":AA7C
            Column(4)       =   "frmAdminPagosLiqCajaListaDG.frx":AB94
            Column(5)       =   "frmAdminPagosLiqCajaListaDG.frx":ACAC
            FormatStylesCount=   6
            FormatStyle(1)  =   "frmAdminPagosLiqCajaListaDG.frx":ADB0
            FormatStyle(2)  =   "frmAdminPagosLiqCajaListaDG.frx":AEE8
            FormatStyle(3)  =   "frmAdminPagosLiqCajaListaDG.frx":AF98
            FormatStyle(4)  =   "frmAdminPagosLiqCajaListaDG.frx":B04C
            FormatStyle(5)  =   "frmAdminPagosLiqCajaListaDG.frx":B124
            FormatStyle(6)  =   "frmAdminPagosLiqCajaListaDG.frx":B1DC
            ImageCount      =   0
            PrinterProperties=   "frmAdminPagosLiqCajaListaDG.frx":B2BC
         End
         Begin GridEX20.GridEX gridCheques 
            Height          =   2295
            Left            =   120
            TabIndex        =   39
            Top             =   480
            Width           =   14850
            _ExtentX        =   26194
            _ExtentY        =   4048
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
            Column(1)       =   "frmAdminPagosLiqCajaListaDG.frx":B494
            Column(2)       =   "frmAdminPagosLiqCajaListaDG.frx":B614
            Column(3)       =   "frmAdminPagosLiqCajaListaDG.frx":B7B4
            Column(4)       =   "frmAdminPagosLiqCajaListaDG.frx":B8AC
            Column(5)       =   "frmAdminPagosLiqCajaListaDG.frx":B9E8
            Column(6)       =   "frmAdminPagosLiqCajaListaDG.frx":BAF4
            Column(7)       =   "frmAdminPagosLiqCajaListaDG.frx":BBC4
            FormatStylesCount=   6
            FormatStyle(1)  =   "frmAdminPagosLiqCajaListaDG.frx":BCB0
            FormatStyle(2)  =   "frmAdminPagosLiqCajaListaDG.frx":BDE8
            FormatStyle(3)  =   "frmAdminPagosLiqCajaListaDG.frx":BE98
            FormatStyle(4)  =   "frmAdminPagosLiqCajaListaDG.frx":BF4C
            FormatStyle(5)  =   "frmAdminPagosLiqCajaListaDG.frx":C024
            FormatStyle(6)  =   "frmAdminPagosLiqCajaListaDG.frx":C0DC
            ImageCount      =   0
            PrinterProperties=   "frmAdminPagosLiqCajaListaDG.frx":C1BC
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
   Begin XtremeSuiteControls.ComboBox cboMonedas 
      Height          =   315
      Left            =   16320
      TabIndex        =   37
      Top             =   9120
      Width           =   1245
      _Version        =   786432
      _ExtentX        =   2196
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Style           =   2
      Text            =   "cboMonedas"
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
      Width           =   7455
      _Version        =   786432
      _ExtentX        =   13150
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
      Width           =   7455
      _Version        =   786432
      _ExtentX        =   13150
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
      Left            =   16320
      TabIndex        =   14
      Tag             =   "Total: "
      Top             =   8880
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
            If fac.id = Factura.id Then
            'If fac.numero = Factura.numero Then
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
        
        Dim C As Integer
        
        If facturaConfirmada.tipoDocumentoContable = tipoDocumentoContable.notaCredito Then C = -1 Else C = 1
        
        xlWorksheet.Cells(idx, 6).value = facturaConfirmada.total * C
       
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
        
        Dim C As Integer
        
        If Factura.tipoDocumentoContable = tipoDocumentoContable.notaCredito Then C = -1 Else C = 1
        
        xlWorksheet.Cells(idx, 6).value = Factura.total * C
        
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
    Set Factura = facturas.item(grilla.RowIndex(grilla.row))
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
            If facturasConfirmadas(i).id = facturaConfirmada.id Then
                ' Encontrar la factura a eliminar
                Set facturaAEliminar = facturasConfirmadas(i)
                
                Dim q As String
                If facturaConfirmada.estado = Saldada Then
                    q = "UPDATE AdminComprasFacturasProveedores SET estado = 2 WHERE id = " & facturaConfirmada.id
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

    formLoading = True
    
    FormHelper.Customize Me

    GridEXHelper.CustomizeGrid Me.grilla, False
    GridEXHelper.CustomizeGrid Me.grillaConfirmados, False
    
    Me.grilla.ItemCount = 0
    Me.grillaConfirmados.ItemCount = 0

    Me.grilla.Refresh
    Me.grillaConfirmados.Refresh

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
    'GridEXHelper.CustomizeGrid Me.gridRetenciones, False, True

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
    
    
''    Set colProveedores = DAOProveedor.FindAllProveedoresWithFacturasImpagas
''    For Each prov In colProveedores
''        cboProveedores.AddItem prov.RazonSocial
''        cboProveedores.ItemData(cboProveedores.NewIndex) = prov.id
''    Next

    Dim cuentasContables As Collection
    Set cuentasContables = DAOCuentaContable.GetAll()
    Dim cc As clsCuentaContable
    
''    Me.cboCuentas.Clear
''    For Each cc In cuentasContables
''        cboCuentas.AddItem cc.nombre & " - " & cc.codigo
''        cboCuentas.ItemData(cboCuentas.NewIndex) = cc.id
''    Next cc

    Me.gridCajaOperaciones.ItemCount = LiquidacionCaja.operacionesCaja.count
    Me.gridDepositosOperaciones.ItemCount = LiquidacionCaja.operacionesBanco.count
    Me.gridCheques.ItemCount = LiquidacionCaja.ChequesTerceros.count
    Me.gridChequesPropios.ItemCount = LiquidacionCaja.ChequesPropios.count

    Set Me.gridCheques.Columns("numero").DropDownControl = Me.gridChequesDisponibles

    Set Me.gridDepositosOperaciones.Columns("moneda").DropDownControl = Me.gridMonedas
    Set Me.gridDepositosOperaciones.Columns("cuenta").DropDownControl = Me.gridCuentasBancarias

    Set Me.gridCajaOperaciones.Columns("moneda").DropDownControl = Me.gridMonedas
    Set Me.gridCajaOperaciones.Columns("caja").DropDownControl = Me.gridCajas
    
    Set Me.gridChequesPropios.Columns("chequera").DropDownControl = Me.gridChequeras
    Set Me.gridChequesPropios.Columns("numero").DropDownControl = Me.gridChequesChequera
    gridChequesChequera.ItemCount = 0
    GridEXHelper.AutoSizeColumns Me.gridChequeras

   Me.dtpFecha.value = Now

   DAOMoneda.llenarComboXtremeSuite Me.cboMonedas
   
   Me.cboMonedas.ListIndex = 0

    If Me.cboMonedas.ListIndex = -1 Then
        Set LiquidacionCaja.moneda = Nothing
    Else
        Set LiquidacionCaja.moneda = DAOMoneda.GetById(Me.cboMonedas.ItemData(Me.cboMonedas.ListIndex))
    End If
    Totalizar
        
'''    llenarGrilla
    
    Me.caption = "Creación de Liquidación de Caja"

    formLoaded = True
    formLoading = False
    
End Sub


Public Sub llenarGrilla()

    grilla.ItemCount = 0

    Dim condition As String

    condition = " AdminComprasFacturasProveedores.estado = 2 OR AdminComprasFacturasProveedores.estado = 4 "

    'condition = " AdminComprasFacturasProveedores.estado = 2 "


    Set facturas = DAOFacturaProveedor.FindAll(condition, , , Permisos.AdminFaPVerSoloPropias)

    Dim F As clsFacturaProveedor
    Dim total As Double
    Dim totalneto As Double
    Dim totIva As Double
    Dim totalno As Double
    Dim totalpercep As Double
    Dim totalsaldo As Double

    Dim C As Integer

    total = 0

    For Each F In facturas

        If F.tipoDocumentoContable = tipoDocumentoContable.notaCredito Then C = -1 Else C = 1
        total = total + MonedaConverter.Convertir(F.total * C, F.moneda.id, MonedaConverter.Patron.id)
        totalneto = totalneto + MonedaConverter.Convertir(F.Monto * C - F.TotalNetoGravadoDiscriminado(0) * C, F.moneda.id, MonedaConverter.Patron.id)
        totalno = totalno + MonedaConverter.Convertir(F.TotalNetoGravadoDiscriminado(0) * C, F.moneda.id, MonedaConverter.Patron.id)
        totIva = totIva + MonedaConverter.Convertir(F.TotalIVA * C, F.moneda.id, MonedaConverter.Patron.id)
        totalpercep = totalpercep + F.totalPercepciones * C
        totalsaldo = totalsaldo + ((F.total - (F.NetoGravadoAbonadoGlobal + F.OtrosAbonadoGlobal)) * C)

    Next

    grilla.ItemCount = facturas.count

'    Me.caption = "Cbtes. filtrados [Cantidad: " & facturas.count & "]"
    
    lblComprobantesMostrados.caption = "Comprobantes: " & facturas.count
    
End Sub


Private Sub gridCajaOperaciones0_BeforeUpdate(Index As Integer, ByVal Cancel As GridEX20.JSRetBoolean)
If Index = 1 Then
        ' Acciones específicas para el control con índice 1
         MsgBox "Este es el control con índice 1 de caja operaciones."
        
            Dim cond1 As Boolean
            Dim cond2 As Boolean
            Dim cond3 As Boolean
            Dim cond4 As Boolean
        
        
            cond1 = Not IsNumeric(Me.gridCajaOperaciones.value(1))
            cond2 = Not IsNumeric(Me.gridCajaOperaciones.value(2)) And LenB(Me.gridCajaOperaciones.value(2)) = 0
            cond3 = Not IsDate(Me.gridCajaOperaciones.value(3))
            cond4 = LenB(Me.gridCajaOperaciones.value(4)) = 0 Or IsEmpty(Me.gridCajaOperaciones.value(4))    'or Not IsNumeric(Me.gridCajaOperaciones.value(4))
        
            Cancel = cond1 Or cond2 Or cond3 Or cond4
    End If
End Sub

Private Sub gridCheques_BeforeUpdate(ByVal Cancel As GridEX20.JSRetBoolean)
    ' Verificamos si el índice es 1
        ' Acciones específicas para el control con índice 1
        ' MsgBox "Este es el control con índice 1."
        
        Dim msg As New Collection
    
        ' REVISA QUE EN LA COLECCION DE CHEQUES DE TERCEROS QUE SE ESTAN CARGANDO NO EST? INGRESADO EL MISMO CHEQUE, SI LO DETECTA GENERA MSG DE ERROR
        If funciones.BuscarEnColeccion(LiquidacionCaja.ChequesTerceros, CStr(Me.gridCheques.value(1))) Then
            msg.Add "El cheque seleccionado ya fue ingresado anteriormente."
        End If
    
        Cancel = (msg.count > 0)
        If Cancel Then MsgBox funciones.JoinCollectionValues(msg, vbNewLine), vbExclamation
End Sub

Private Sub gridCheques_UnboundAddNew(ByVal NewRowBookmark As GridEX20.JSRetVariant, ByVal Values As GridEX20.JSRowData)
    ' Verificamos si el índice es 1
    Set cheque = Nothing
    If IsNumeric(Values(1)) Then Set cheque = DAOCheques.FindById(Values(1))
    If IsSomething(cheque) Then
        LiquidacionCaja.ChequesTerceros.Add cheque, CStr(cheque.id)

    End If
    Totalizar
End Sub

Private Sub gridCheques_UnboundDelete(ByVal RowIndex As Long, ByVal Bookmark As Variant)
    ' Verificamos si el índice es 1
    If RowIndex > 0 Then
        LiquidacionCaja.ChequesTerceros.remove RowIndex
        Totalizar
    End If
End Sub


Private Sub gridCheques_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    ' Verificamos si el índice es 1
        ' Acciones específicas para el control con índice 1
    If RowIndex <= LiquidacionCaja.ChequesTerceros.count Then
        Set cheque = LiquidacionCaja.ChequesTerceros.item(RowIndex)

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
    ' Verificamos si el índice es 1
    If RowIndex > 0 And LiquidacionCaja.ChequesTerceros.count >= RowIndex Then
        Set cheque = Nothing
        If IsNumeric(Values(1)) Then Set cheque = DAOCheques.FindById(Values(1))
        If IsSomething(cheque) Then
            LiquidacionCaja.ChequesTerceros.Add cheque, , , RowIndex
            LiquidacionCaja.ChequesTerceros.remove RowIndex
        End If
        Totalizar
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
    If funciones.BuscarEnColeccion(LiquidacionCaja.ChequesPropios, CStr(Me.gridChequesPropios.value(2))) Then
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

        LiquidacionCaja.ChequesPropios.Add cheque, CStr(cheque.id)


    End If
    Totalizar
End Sub

Private Sub gridChequesPropios_UnboundDelete(ByVal RowIndex As Long, ByVal Bookmark As Variant)
    If RowIndex > 0 Then
        LiquidacionCaja.ChequesPropios.remove RowIndex
        Totalizar
    End If
End Sub

Private Sub gridChequesPropios_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If LiquidacionCaja.ChequesPropios.count >= RowIndex Then
        Set cheque = LiquidacionCaja.ChequesPropios.item(RowIndex)
        Values(1) = cheque.chequera.Description
        Values(2) = vbNullString
        'FORMATCURRENCY
        Values(3) = FormatCurrency(cheque.Monto)
        Values(4) = cheque.FechaVencimiento
        Values(5) = cheque.numero


        Totalizar
    End If
End Sub

Private Sub gridChequesPropios_UnboundUpdate(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If LiquidacionCaja.ChequesPropios.count >= RowIndex Then
        Set cheque = OrdenPago.ChequesPropios.item(RowIndex)

        '        If Values(2) <> Cheque.Id Then
        '            ordenPago.ChequesPropios.remove CStr(Cheque.Id)
        '            Set Cheque = DAOCheques.FindById(Values(2))
        '            ordenPago.ChequesPropios.Add Cheque, CStr(Cheque.Id)
        '        End If

        cheque.Monto = Values(3)
        cheque.FechaVencimiento = Values(4)
    End If

    Totalizar
End Sub


Private Sub gridCompensatorios_UnboundDelete(ByVal RowIndex As Long, ByVal Bookmark As Variant)
    OrdenPago.Compensatorios.remove (RowIndex)
End Sub

Private Sub grilla_ColumnHeaderClick(ByVal Column As GridEX20.JSColumn)
    GridEXHelper.ColumnHeaderClick Me.grilla, Column

End Sub


Private Sub grilla_FetchIcon(ByVal RowIndex As Long, ByVal ColIndex As Integer, ByVal RowBookmark As Variant, ByVal IconIndex As GridEX20.JSRetInteger)
    If ColIndex = 15 And m_Archivos.item(Factura.id) > 0 Then IconIndex = 1

End Sub


Private Sub grilla_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    SeleccionarFactura
    
    ' ME PIDE KARIN QUE CUANDO SE SELECCIONA SE AGREGUE AUTOMATICAMENTE AL OTRO LISTADO
    
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
            If fac.id = Factura.id Then
            'If fac.numero = Factura.numero Then
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

' LLENADO DE GRILLA DE COMPROBANTES APROBADOS PARA PAGAR
Private Sub grilla_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)

    Set Factura = facturas.item(RowIndex)

    Dim i As Integer

    If Factura.tipoDocumentoContable = tipoDocumentoContable.notaCredito Then i = -1 Else i = 1

    With Factura

        Values(1) = enums.EnumTipoDocumentoContableShort(Factura.tipoDocumentoContable)
        Values(2) = Factura.configFactura.TipoFactura
        Values(3) = Factura.numero
        Values(4) = Factura.FEcha
        Values(5) = Factura.moneda.NombreCorto
        Values(6) = Replace(FormatCurrency(funciones.FormatearDecimales(Factura.ImporteTotalSaldo) * i), "$", "")
        Values(7) = UCase(funciones.RazonSocialFormateada(Factura.Proveedor.RazonSocial))

    End With
    
    ' Desactivar la selección inicial en el GridEx
    grilla.row = -1
    grilla.col = -1

End Sub



Private Sub grillaConfirmados_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    SeleccionarFactura

End Sub

' LLENADO DE GRILLA DE COMPROBANTES CONFIRMADOS
Private Sub grillaConfirmados_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)

    Set facturaConfirmada = facturasConfirmadas.item(RowIndex)

    Dim i As Integer

    If facturaConfirmada.tipoDocumentoContable = tipoDocumentoContable.notaCredito Then i = -1 Else i = 1

    With facturaConfirmada

        Values(1) = enums.EnumTipoDocumentoContableShort(facturaConfirmada.tipoDocumentoContable)
        Values(2) = facturaConfirmada.configFactura.TipoFactura
        Values(3) = facturaConfirmada.numero
        Values(4) = facturaConfirmada.FEcha
        Values(5) = facturaConfirmada.moneda.NombreCorto
        Values(6) = Replace(FormatCurrency(funciones.FormatearDecimales(facturaConfirmada.total) * i), "$", "")
        Values(7) = Replace(FormatCurrency(funciones.FormatearDecimales((facturaConfirmada.ImporteTotalSaldo)) * i), "$", "")
        Values(8) = UCase(funciones.RazonSocialFormateada(facturaConfirmada.Proveedor.RazonSocial))

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

    If Me.txtNumerodeLiquidacion.Text = "" Then
        MsgBox ("El número de Liquidación no puede estar vacío.")
        Exit Sub
    Else
        LiquidacionCaja.NumeroLiq = Me.txtNumerodeLiquidacion.Text

    End If

    Set LiquidacionCaja.FacturasProveedor = New Collection


    For Each Factura In facturasConfirmadas
        LiquidacionCaja.FacturasProveedor.Add Factura
    Next


    If LiquidacionCaja.IsValid Then

        Dim n As Boolean: n = (LiquidacionCaja.id = 0)

        If DAOLiquidacionCaja.Save(LiquidacionCaja, True) Then

            If n Then
                MsgBox "Liquidación de Caja Nº " & Me.txtNumerodeLiquidacion & " creada con exito.", vbInformation
            Else
<<<<<<< HEAD
                MsgBox "Liquidaci�n de Caja modificada con �xito.", vbInformation
            End If

            If n Then
                If MsgBox("Desea crear una Liquidaci�n de Caja nueva", vbQuestion + vbYesNo) = vbYes Then
=======
                MsgBox "Liquidación de Caja modificada con exito.", vbInformation
            End If

            If n Then
                If MsgBox("¿Desea crear una Liquidación de Caja nueva", vbQuestion + vbYesNo) = vbYes Then
>>>>>>> 809a13d9c3e48791cf5eeb0815c282ed35cca3bc
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
            total = total - Factura.ImporteTotalSaldo       ' Resta el total de las facturas tipo nota de crédito
        Else
            total = total + Factura.ImporteTotalSaldo        ' Suma el total de las demás facturas
        End If
    Next i

'''    LiquidacionCaja.StaticTotalFacturas = funciones.RedondearDecimales(total)

    lblLabel1.caption = "Total Comprobantes en proceso: " & FormatCurrency(funciones.FormatearDecimales(LiquidacionCaja.StaticTotal))
    
    lblCbtesConfirmados.caption = "Comprobantes: " & facturasConfirmadas.count

End Sub


'TODAS LAS FUNCIONES QUE VIENEN DE LA LIQUIDACION DE CAJA ANTERIOR
Private Sub CargarChequesDisponibles()
    Set chequesDisponibles = DAOCheques.FindAllEnCarteraDeTerceros
    Me.gridChequesDisponibles.ItemCount = chequesDisponibles.count
End Sub


Private Sub gridBancos_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex <= bancos.count Then
        Set Banco = bancos.item(RowIndex)
        Values(1) = Banco.id
        Values(2) = Banco.nombre
    End If

End Sub


Private Sub gridCajas_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex > 0 And Cajas.count > 0 Then
        Set caja = Cajas.item(RowIndex)
        Values(1) = caja.id
        Values(2) = caja.nombre
    End If
End Sub


Private Sub gridChequeras_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex <= chequeras.count Then
        Set tmpChequera = chequeras.item(RowIndex)
        Values(1) = tmpChequera.Description
        Values(2) = tmpChequera.id
    End If
End Sub


Private Sub gridChequesChequera_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex > 0 And chequesChequeraSeleccionada.count > 0 Then
        Values(1) = chequesChequeraSeleccionada(RowIndex).numero
        Values(2) = chequesChequeraSeleccionada(RowIndex).id
    End If
End Sub

'Private Sub gridChequesDisponibles_ColumnHeaderClick(ByVal Column As GridEX20.JSColumn)
'    GridEXHelper.ColumnHeaderClick Me.gridChequesDisponibles, Column
'End Sub

Private Sub gridChequesDisponibles_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex <= chequesDisponibles.count Then
        Set cheque = chequesDisponibles.item(RowIndex)
        Values(1) = cheque.numero
        'FORMATCURRENCY
        Values(2) = FormatCurrency(cheque.Monto)
        If IsSomething(cheque.moneda) Then Values(3) = cheque.moneda.NombreCorto
        If IsSomething(cheque.Banco) Then Values(4) = cheque.Banco.nombre
        Values(5) = cheque.id
        Values(6) = cheque.OrigenCheque
        Values(7) = cheque.OrigenDestino

    End If

End Sub


Private Sub gridCuentasBancarias_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If cuentasBancarias.count >= RowIndex Then
        Set CuentaBancaria = cuentasBancarias.item(RowIndex)
        Values(1) = CuentaBancaria.id
        Values(2) = CuentaBancaria.DescripcionFormateada
    End If
End Sub

Private Sub gridMonedas_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex > 0 And Monedas.count > 0 Then
        Set moneda = Monedas.item(RowIndex)
        Values(1) = moneda.id
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
    LiquidacionCaja.operacionesCaja.Add operacion

    Totalizar

End Sub

Private Sub gridCajaOperaciones_UnboundDelete(ByVal RowIndex As Long, ByVal Bookmark As Variant)
    If RowIndex > 0 And LiquidacionCaja.operacionesCaja.count >= RowIndex Then
        LiquidacionCaja.operacionesCaja.remove RowIndex

        Totalizar

    End If
End Sub

Private Sub gridCajaOperaciones_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex <= LiquidacionCaja.operacionesCaja.count Then
        Set operacion = LiquidacionCaja.operacionesCaja.item(RowIndex)
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
    ' Activar el manejo de errores
    On Error GoTo ManejarError

    If RowIndex > 0 And LiquidacionCaja.operacionesCaja.count > 0 Then
        Set operacion = LiquidacionCaja.operacionesCaja.item(RowIndex)

        ' Asignar Monto
        If IsNumeric(Values(1)) Then
            operacion.Monto = CDbl(Values(1)) ' Convierte a Double si es necesario
        Else
            operacion.Monto = 0 ' Valor por defecto
        End If

        ' Asignar Comprobante
        operacion.Comprobante = Values(5)

        ' Asignar Moneda
        If IsNumeric(Values(2)) Then
            Set operacion.moneda = DAOMoneda.GetById(Values(2))
        End If

        ' Asignar FechaOperacion
        operacion.FechaOperacion = Values(3)

        ' Asignar Caja
        If IsNumeric(Values(4)) Then
            Set operacion.caja = DAOCaja.FindById(Values(4))
        End If

        operacion.EntradaSalida = OPSalida

        ' Llamar a la función Totalizar
        Totalizar
    End If

    ' Salir del procedimiento sin ejecutar el manejo de errores
    Exit Sub

ManejarError:
    ' Mostrar un mensaje de error al usuario
    MsgBox "Se produjo un error en la actualización de la cuadrícula." & vbCrLf & _
           "Error: " & Err.Description & vbCrLf & _
           "Número de error: " & Err.Number, vbCritical, "Error"

    ' Opcional: Registrar el error en un archivo de log
    ' Call RegistrarError(Err.Number, Err.Description)

    ' Limpiar el objeto Err
    Err.Clear
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
    LiquidacionCaja.operacionesBanco.Add operacion

    Totalizar

End Sub

Private Sub gridDepositosOperaciones_UnboundDelete(ByVal RowIndex As Long, ByVal Bookmark As Variant)
    If RowIndex > 0 And LiquidacionCaja.operacionesBanco.count >= RowIndex Then
        LiquidacionCaja.operacionesBanco.remove RowIndex

        Totalizar

    End If
End Sub

Private Sub gridDepositosOperaciones_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex <= LiquidacionCaja.operacionesBanco.count Then
        Set operacion = LiquidacionCaja.operacionesBanco.item(RowIndex)
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
    If RowIndex > 0 And LiquidacionCaja.operacionesBanco.count > 0 Then
        Set operacion = LiquidacionCaja.operacionesBanco.item(RowIndex)
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
    
    Me.lblLabel1.caption = "Total Comprobantes: " & FormatCurrency(funciones.FormatearDecimales(LiquidacionCaja.StaticTotal))
    
    Me.lblLabel2.caption = "Total Valores Cargados: " & FormatCurrency(funciones.FormatearDecimales(LiquidacionCaja.StaticTotalOrigenes + LiquidacionCaja.StaticTotalRetenido))
    
    GridEXHelper.AutoSizeColumns Me.gridCajaOperaciones
    
    GridEXHelper.AutoSizeColumns Me.gridDepositosOperaciones

End Sub

Private Sub txtFiltroNumero_Change()
    Dim filterText As String
    filterText = Trim(Me.txtFiltroNumero.Text)
    
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



Public Sub Cargar(liq As clsLiquidacionCaja)

    If Not IsSomething(liq) Then
        MsgBox "La Liquidación que está intentando visualizar está en estado PENDIENTE. " & vbNewLine & "Por lo tanto no puede ser mostrada porque puede estar siendo editada." & vbNewLine & "Verifiquelo por favor.", vbCritical, "OP Pendiente"
        Unload Me
        Exit Sub
    End If

    Set LiquidacionCaja = DAOLiquidacionCaja.FindById(liq.id)
    
    Me.caption = "Liquidación Nº " & LiquidacionCaja.NumeroLiq
    
    Set facturasConfirmadas = DAOFacturaProveedor.FindAllByLiquidacionCaja(liq.id)

    Me.grillaConfirmados.ItemCount = 0
  
    Me.grillaConfirmados.ItemCount = facturasConfirmadas.count
    
    Me.gridCajaOperaciones.ItemCount = LiquidacionCaja.operacionesCaja.count
    Me.gridDepositosOperaciones.ItemCount = LiquidacionCaja.operacionesBanco.count
    Me.gridCheques.ItemCount = LiquidacionCaja.ChequesTerceros.count
    Me.gridChequesPropios.ItemCount = LiquidacionCaja.ChequesPropios.count
    
    Me.dtpFecha.value = LiquidacionCaja.FEcha
    Me.txtOtrosDescuentos.Text = LiquidacionCaja.OtrosDescuentos

    Me.txtNumerodeLiquidacion = LiquidacionCaja.NumeroLiq
    
'''    If Not ReadOnly Then
'''    MsgBox ("Esto es Editar")
'''    Else
'''    MsgBox ("Esto es Ver")
'''    End If


    Me.gridDepositosOperaciones.AllowEdit = Not ReadOnly
    Me.gridDepositosOperaciones.AllowDelete = Not ReadOnly
    Me.gridBancos.AllowEdit = Not ReadOnly
    Me.gridCajaOperaciones.AllowEdit = Not ReadOnly
    Me.gridCajaOperaciones.AllowDelete = Not ReadOnly
    Me.gridCajas.AllowEdit = Not ReadOnly
    Me.gridCheques.AllowEdit = Not ReadOnly
    Me.gridCheques.AllowDelete = Not ReadOnly
    Me.gridChequeras.AllowEdit = Not ReadOnly
    Me.gridChequesChequera.AllowEdit = Not ReadOnly
    Me.gridChequesDisponibles.AllowEdit = Not ReadOnly
    Me.gridChequesPropios.AllowEdit = Not ReadOnly
    Me.gridChequesPropios.AllowDelete = Not ReadOnly
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

    Totalizar
    
    TotalizarComprobantes

End Sub
