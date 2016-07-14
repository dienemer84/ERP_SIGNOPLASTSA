VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~3.OCX"
Begin VB.Form frmPlaneamientoPedidosPendientes 
   BackColor       =   &H8000000B&
   Caption         =   "Ordenes de trabajo"
   ClientHeight    =   8985
   ClientLeft      =   1440
   ClientTop       =   4725
   ClientWidth     =   14265
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   18
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmVerPedidosPendientes.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8985
   ScaleWidth      =   14265
   Tag             =   "Ordenes de trabajo"
   Begin XtremeSuiteControls.PushButton Command1 
      Height          =   375
      Left            =   135
      TabIndex        =   1
      Top             =   8040
      Width           =   1215
      _Version        =   786432
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Estadísticas"
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin GridEX20.GridEX grid 
      Height          =   5385
      Left            =   -15
      TabIndex        =   0
      Top             =   2295
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   9499
      Version         =   "2.0"
      PreviewRowIndent=   500
      DefaultGroupMode=   1
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      PreviewColumn   =   5
      PreviewRowLines =   2
      CalendarTodayText=   "Hoy"
      CalendarNoneText=   "Vacio"
      ColumnAutoResize=   -1  'True
      MultiSelect     =   -1  'True
      MethodHoldFields=   -1  'True
      ContScroll      =   -1  'True
      ForeColorInfoText=   16777215
      BackColorInfoText=   8421504
      GroupByBoxInfoText=   "Arrastre una columna aqui para ordenar por dicha columna."
      AllowEdit       =   0   'False
      BackColorGBBox  =   8421504
      BackColorHeader =   16761024
      ImageCount      =   1
      ImagePicture1   =   "frmVerPedidosPendientes.frx":000C
      RowHeaders      =   -1  'True
      ItemCount       =   1
      DataMode        =   99
      HeaderFontName  =   "Tahoma"
      HeaderFontBold  =   0   'False
      HeaderFontSize  =   8.25
      HeaderFontWeight=   400
      FontName        =   "Tahoma"
      FontBold        =   0   'False
      FontSize        =   8.25
      FontWeight      =   400
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   13
      Column(1)       =   "frmVerPedidosPendientes.frx":0326
      Column(2)       =   "frmVerPedidosPendientes.frx":0436
      Column(3)       =   "frmVerPedidosPendientes.frx":0506
      Column(4)       =   "frmVerPedidosPendientes.frx":0602
      Column(5)       =   "frmVerPedidosPendientes.frx":06D6
      Column(6)       =   "frmVerPedidosPendientes.frx":07CA
      Column(7)       =   "frmVerPedidosPendientes.frx":08C6
      Column(8)       =   "frmVerPedidosPendientes.frx":0996
      Column(9)       =   "frmVerPedidosPendientes.frx":0A92
      Column(10)      =   "frmVerPedidosPendientes.frx":0B62
      Column(11)      =   "frmVerPedidosPendientes.frx":0C92
      Column(12)      =   "frmVerPedidosPendientes.frx":0DEA
      Column(13)      =   "frmVerPedidosPendientes.frx":0F06
      FormatStylesCount=   10
      FormatStyle(1)  =   "frmVerPedidosPendientes.frx":0FEA
      FormatStyle(2)  =   "frmVerPedidosPendientes.frx":1112
      FormatStyle(3)  =   "frmVerPedidosPendientes.frx":11C2
      FormatStyle(4)  =   "frmVerPedidosPendientes.frx":1276
      FormatStyle(5)  =   "frmVerPedidosPendientes.frx":132A
      FormatStyle(6)  =   "frmVerPedidosPendientes.frx":1402
      FormatStyle(7)  =   "frmVerPedidosPendientes.frx":14E2
      FormatStyle(8)  =   "frmVerPedidosPendientes.frx":15AE
      FormatStyle(9)  =   "frmVerPedidosPendientes.frx":167A
      FormatStyle(10) =   "frmVerPedidosPendientes.frx":174E
      ImageCount      =   1
      ImagePicture(1) =   "frmVerPedidosPendientes.frx":1812
      PrinterProperties=   "frmVerPedidosPendientes.frx":1B2C
   End
   Begin XtremeSuiteControls.PushButton cmdImprimir 
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   8025
      Width           =   1215
      _Version        =   786432
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Imprimir"
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton cmdConsultas 
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   8040
      Width           =   1215
      _Version        =   786432
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Consultas"
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton btnTareasOrdenes 
      Height          =   375
      Left            =   4080
      TabIndex        =   4
      Top             =   8040
      Width           =   1215
      _Version        =   786432
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Excel Tareas"
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   2190
      Left            =   45
      TabIndex        =   5
      Top             =   30
      Width           =   11475
      _Version        =   786432
      _ExtentX        =   20241
      _ExtentY        =   3863
      _StockProps     =   79
      Caption         =   "Parámetros de búsqueda"
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Begin VB.ListBox lstEstados 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   915
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   19
         Top             =   1155
         Width           =   3555
      End
      Begin VB.TextBox txtDescripcion 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6240
         TabIndex        =   16
         Top             =   705
         Width           =   3540
      End
      Begin VB.TextBox txtNroOrden 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6255
         TabIndex        =   15
         Top             =   300
         Width           =   1185
      End
      Begin VB.CheckBox chkMostrarDescripcion 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mostrar Descripcion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7560
         TabIndex        =   14
         Top             =   240
         Value           =   1  'Checked
         Width           =   2235
      End
      Begin XtremeSuiteControls.GroupBox grpFecha 
         Height          =   1035
         Left            =   5265
         TabIndex        =   6
         Top             =   1035
         Width           =   4560
         _Version        =   786432
         _ExtentX        =   8043
         _ExtentY        =   1826
         _StockProps     =   79
         Caption         =   "Fecha de Entrega"
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.DateTimePicker dtpDesde 
            Height          =   315
            Left            =   795
            TabIndex        =   7
            Top             =   600
            Width           =   1470
            _Version        =   786432
            _ExtentX        =   2593
            _ExtentY        =   556
            _StockProps     =   68
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CheckBox        =   -1  'True
            Format          =   1
         End
         Begin XtremeSuiteControls.ComboBox cboRangos 
            Height          =   315
            Left            =   795
            TabIndex        =   8
            Top             =   210
            Width           =   3645
            _Version        =   786432
            _ExtentX        =   6429
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Style           =   2
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.DateTimePicker dtpHasta 
            Height          =   315
            Left            =   2985
            TabIndex        =   9
            Top             =   600
            Width           =   1470
            _Version        =   786432
            _ExtentX        =   2593
            _ExtentY        =   556
            _StockProps     =   68
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CheckBox        =   -1  'True
            Format          =   1
            CurrentDate     =   41241.3520601852
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   195
            Left            =   225
            TabIndex        =   12
            Top             =   255
            Width           =   480
            _Version        =   786432
            _ExtentX        =   847
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Rango"
            BackColor       =   12632256
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            AutoSize        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label5 
            Height          =   195
            Left            =   240
            TabIndex        =   11
            Top             =   630
            Width           =   465
            _Version        =   786432
            _ExtentX        =   820
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Desde"
            BackColor       =   12632256
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            AutoSize        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label6 
            Height          =   195
            Left            =   2415
            TabIndex        =   10
            Top             =   645
            Width           =   420
            _Version        =   786432
            _ExtentX        =   741
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Hasta"
            BackColor       =   12632256
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            AutoSize        =   -1  'True
         End
      End
      Begin XtremeSuiteControls.PushButton cmdBuscar 
         Default         =   -1  'True
         Height          =   495
         Left            =   10050
         TabIndex        =   13
         Top             =   1560
         Width           =   1290
         _Version        =   786432
         _ExtentX        =   2275
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Buscar"
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cboClientes 
         Height          =   315
         Left            =   915
         TabIndex        =   17
         Top             =   720
         Width           =   3555
         _Version        =   786432
         _ExtentX        =   6271
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Sorted          =   -1  'True
         Style           =   2
         Appearance      =   6
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.PushButton cmdSinCliente 
         Height          =   315
         Left            =   4515
         TabIndex        =   18
         Top             =   720
         Width           =   375
         _Version        =   786432
         _ExtentX        =   661
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "X"
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cboCCostos 
         Height          =   315
         Left            =   915
         TabIndex        =   20
         Top             =   360
         Width           =   3555
         _Version        =   786432
         _ExtentX        =   6271
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Sorted          =   -1  'True
         Style           =   2
         Appearance      =   6
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.PushButton cmdSinCCosto 
         Height          =   315
         Left            =   4515
         TabIndex        =   21
         Top             =   360
         Width           =   375
         _Version        =   786432
         _ExtentX        =   661
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "X"
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.Label LBLMENSAJE2 
         Height          =   360
         Left            =   9990
         TabIndex        =   28
         Top             =   870
         Width           =   945
         _Version        =   786432
         _ExtentX        =   1667
         _ExtentY        =   635
         _StockProps     =   79
         Caption         =   "Label8"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label LBLMENSAJE 
         Height          =   360
         Left            =   9975
         TabIndex        =   27
         Top             =   390
         Width           =   945
         _Version        =   786432
         _ExtentX        =   1667
         _ExtentY        =   635
         _StockProps     =   79
         Caption         =   "Label8"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoSize        =   -1  'True
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción"
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
         Left            =   5295
         TabIndex        =   26
         Top             =   720
         Width           =   840
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Orden"
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
         Left            =   5475
         TabIndex        =   25
         Top             =   345
         Width           =   660
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente"
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
         Left            =   345
         TabIndex        =   24
         Top             =   750
         Width           =   480
      End
      Begin VB.Label lblFiltro 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estados"
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
         Left            =   240
         TabIndex        =   23
         Top             =   1200
         Width           =   570
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "C.Costos"
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
         Left            =   240
         TabIndex        =   22
         Top             =   390
         Width           =   630
      End
   End
   Begin VB.Menu pedidos 
      Caption         =   "pedidos"
      Visible         =   0   'False
      Begin VB.Menu numero 
         Caption         =   "numerio"
         Enabled         =   0   'False
      End
      Begin VB.Menu editOT 
         Caption         =   "Editar..."
      End
      Begin VB.Menu mnuPrecios 
         Caption         =   "Actualizar Precios"
      End
      Begin VB.Menu mnuSeguimiento 
         Caption         =   "Seguimiento..."
      End
      Begin VB.Menu mnuPlanificacion 
         Caption         =   "Planificar..."
      End
      Begin VB.Menu mnuCerrar 
         Caption         =   "Cerrar..."
      End
      Begin VB.Menu AprobarOT 
         Caption         =   "Aprobar..."
      End
      Begin VB.Menu verDetalles 
         Caption         =   "Ver Detalles..."
      End
      Begin VB.Menu mnuExportarResumenGeneral 
         Caption         =   "Exportar Resumen General..."
      End
      Begin VB.Menu mnuAsociarFacturaAnticipo 
         Caption         =   "Asociar a Factura de Anticipo..."
      End
      Begin VB.Menu mnuImprimir 
         Caption         =   "Imprimir..."
      End
      Begin VB.Menu verHistorial 
         Caption         =   "Ver Historial..."
      End
      Begin VB.Menu verIncidencias 
         Caption         =   "Ver Incidencias..."
      End
      Begin VB.Menu mnuFacturasAplicadas 
         Caption         =   "Facturas Aplicadas..."
      End
      Begin VB.Menu mnuRemitosEntregados 
         Caption         =   "Remitos Aplicados..."
      End
      Begin VB.Menu mnuEntregas 
         Caption         =   "Ver Entregas..."
      End
      Begin VB.Menu mnuScanear 
         Caption         =   "Scanear..."
      End
      Begin VB.Menu archivos 
         Caption         =   "Archivos Asociados..."
      End
      Begin VB.Menu mnuCopiarOT 
         Caption         =   "Copiar OT..."
      End
      Begin VB.Menu nadanda 
         Caption         =   "-"
      End
      Begin VB.Menu modificarPedido 
         Caption         =   "Modificar pedido..."
      End
      Begin VB.Menu activarPedido 
         Caption         =   "Activar Pedido..."
      End
      Begin VB.Menu desactivar 
         Caption         =   "Desactivar pedido..."
      End
   End
End
Attribute VB_Name = "frmPlaneamientoPedidosPendientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Implements ISuscriber
Option Explicit
Private Enum FiltroEstadoOrdenesTrabajo
    Pendientes
    enCurso
    Finalizadas
End Enum


Private EstadoOrdenesTrabajo As FiltroEstadoOrdenesTrabajo
Dim srow As Long
Private auxLong As Long
Private m_ordenesTrabajo As New Collection
Private aux_ordenTrabajo As OrdenTrabajo
Private m_Incidencias As New Dictionary
Private m_Archivos As New Dictionary
Private m_suscriptor_id As String
Private Sub activarPedido_Click()
    Dim A As Long

    If Not aux_ordenTrabajo Is Nothing Then

        If Me.grid.SelectedItems.count = 1 Then
            A = grid.RowIndex(grid.row)
            grid_SelectionChange


            Dim frmaa As New frmActivarPedido
            Set frmaa.Pedido = aux_ordenTrabajo
            frmaa.Show

            grid.RefreshRowIndex A
        Else

            If MsgBox("¿Desea activar los pedidos seleccionados?", vbYesNo, "Confirmación") = vbNo Then Exit Sub
            Dim ped As OrdenTrabajo
            Dim algunos_errores As Boolean
            algunos_errores = False
            Dim it As JSSelectedItem
            For Each it In Me.grid.SelectedItems
                Set ped = m_ordenesTrabajo.item(Me.grid.RowIndex(it.RowIndex))
                If ped.estado = EstadoOT_EnEspera Then
                    If Not DAOOrdenTrabajo.PonerEnProduccion(ped) Then
                        algunos_errores = True
                    End If
                End If
                grid.RefreshRowIndex it.RowIndex
            Next

            If algunos_errores Then
                MsgBox "Algunas OT no pudieron activarse!", vbCritical, "Error"
            Else
                '            MsgBox "Pedidos Activados correctamente!", vbInformation, "Información"
                Me.grid.ReBind
            End If


        End If

    End If
End Sub
Private Sub AprobarOT_Click()
    Dim A As Long
    If Not aux_ordenTrabajo Is Nothing Then
        If MsgBox("¿Está seguro de aprobar la OT " & aux_ordenTrabajo.id & "?", vbYesNo, "Confirmación") = vbNo Then Exit Sub
        A = grid.RowIndex(grid.row)
        Set aux_ordenTrabajo.Detalles = DAODetalleOrdenTrabajo.FindAllByOrdenTrabajo(aux_ordenTrabajo.id)

        If Not aux_ordenTrabajo.Detalles Is Nothing Then

            Dim aaa As New frmReservaStock
            Set aaa.Ot = aux_ordenTrabajo
            aaa.Show 1
            grid.RefreshRowIndex A
            MostrarMensajePendientes
        End If
    End If
End Sub
Private Sub archivos_Click()
    If Not aux_ordenTrabajo Is Nothing Then
        Dim F As New frmArchivos2
        F.Origen = OrigenArchivos.OA_OrdenesTrabajo
        F.ObjetoId = aux_ordenTrabajo.id
        F.caption = "Archivos OT Nº " & aux_ordenTrabajo.IdFormateado
        F.Show
    End If
End Sub

Private Sub btnTareasOrdenes_Click()
    Dim path As String
    path = mBrowseFolder.BrowseDirectory(Me.hWnd, "Seleccione una carpeta para guardar los archivos Excel")

    If LenB(path) <> 0 Then
        Dim tmpOrden As OrdenTrabajo
        Dim item As JSSelectedItem
        For Each item In Me.grid.SelectedItems
            Set tmpOrden = m_ordenesTrabajo.item(item.RowIndex)
            ExcelListadoTareas path & "\Tareas de OT " & tmpOrden.IdFormateado & ".xls", tmpOrden
        Next item
        MsgBox "Proceso finalizado.", vbInformation
    Else
        MsgBox "Debe seleccionar una carpeta para guardar los archivos Excel.", vbExclamation
    End If

End Sub

Private Sub cboRangos_Click()
    funciones.CalculateDateRange Me.cboRangos, Me.dtpDesde, Me.dtpHasta
End Sub

Private Sub cmdSinCCosto_Click()
    Me.cboCCostos.ListIndex = -1
End Sub

Private Sub chkMostrarDescripcion_Click()
    If Me.chkMostrarDescripcion.value Then
        Me.grid.PreviewRowLines = 2
        Me.grid.Gridlines = jgexGLHorizontal
    Else
        Me.grid.Gridlines = jgexGLBoth
        Me.grid.PreviewRowLines = 0
    End If
End Sub
Private Sub cmdBuscar_Click()
    LlenarNuevaLista

End Sub
Private Sub cmdConsultas_Click()
    Dim si As JSSelectedItem
    Dim col As New Collection
    For Each si In Me.grid.SelectedItems
        col.Add m_ordenesTrabajo(si.RowIndex)
    Next si
    If col.count > 0 Then
        Dim F As New frmPlaneamientoConsultasMultiples
        Set F.OTElegidas = col
        F.Show
    End If
End Sub
Private Sub cmdImprimir_Click()
    With Me.grid.PrinterProperties
        .FitColumns = True
        .RepeatHeaders = True
        .Orientation = jgexPPLandscape
        .HeaderString(jgexHFCenter) = Me.caption
        .FooterString(jgexHFCenter) = Now
    End With

    Dim F As New frmPrintPreview
    grid.PrintPreview F.GEXPreview1, (grid.SelectedItems.count > 1)
    F.WindowState = vbMaximized
    F.Show
End Sub

Private Sub CMDsINCliente_Click()
    Me.cboClientes.ListIndex = -1
End Sub

Private Sub Command1_Click()
    Dim rs As Recordset
    Dim strsql As String
    Dim selectedItem As JSSelectedItem
    Dim ped As OrdenTrabajo
    Dim dto As DTOPiezaCantidad
    Dim listadtopiezacantidad As New Collection
    Dim dp As DetalleOrdenTrabajo

    Dim ots_id As New Collection

    For Each selectedItem In Me.grid.SelectedItems
        Set ped = m_ordenesTrabajo.item(selectedItem.RowIndex)

        If ped.estado = EstadoOT_Pendiente Then
            Set ped.Detalles = DAODetalleOrdenTrabajo.FindAllByOrdenTrabajo(ped.id)
            For Each dp In ped.Detalles
                Set dto = New DTOPiezaCantidad
                Set dto.Pieza = dp.Pieza
                dto.Cantidad = dp.CantidadPedida
                listadtopiezacantidad.Add dto
            Next dp
        ElseIf ped.estado = EstadoOT_Desactivado Then    ' no hace nada
        Else
            ots_id.Add ped.id
        End If

    Next selectedItem

    Dim estadisticasOT As Collection
    Set estadisticasOT = DAOOrdenTrabajo.GetDTOSectoresTiempo(ots_id)

    '    Dim indicesFinalizacion As New Collection
    '    Set indicesFinalizacion = DAOTiemposProceso.GetAvancesOTBySectorAndTareaFinalizada(ots_id)

    Dim f123 As New frmEstadistiacasEnCurso

    f123.caption = "Estadisticas de pedidos seleccionados"
    f123.conjGrabado = True
    Set f123.IDsOTAvance = ots_id
    Set f123.col = MergeEstadisticas(estadisticasOT, DAOPieza.ListaDTOTiempoPorSector(listadtopiezacantidad))
    f123.LlenarGridDesdeOT
    f123.Show

End Sub

Private Function MergeEstadisticas(colOT As Collection, colPiezas As Collection) As Collection
    Dim tmpSectorTiempo As DTOSectoresTiempo
    Dim sectorTiempo As DTOSectoresTiempo
    Dim tareaTiempo As DTOTareaTiempo
    Dim tmpTareaTiempo As DTOTareaTiempo

    For Each sectorTiempo In colPiezas
        If funciones.BuscarEnColeccion(colOT, CStr(sectorTiempo.Sector.id)) Then
            Set tmpSectorTiempo = colOT.item(CStr(sectorTiempo.Sector.id))
        Else
            Set tmpSectorTiempo = New DTOSectoresTiempo
            Set tmpSectorTiempo.Sector = sectorTiempo.Sector
            colOT.Add tmpSectorTiempo, CStr(tmpSectorTiempo.Sector.id)
        End If

        For Each tareaTiempo In sectorTiempo.ListaDtoTareaTiempo
            If BuscarEnColeccion(tmpSectorTiempo.ListaDtoTareaTiempo, CStr(tareaTiempo.Tarea.id)) Then
                Set tmpTareaTiempo = tmpSectorTiempo.ListaDtoTareaTiempo.item(CStr(tareaTiempo.Tarea.id))
            Else
                Set tmpTareaTiempo = New DTOTareaTiempo
                Set tmpTareaTiempo.Tarea = tareaTiempo.Tarea
                tmpSectorTiempo.ListaDtoTareaTiempo.Add tmpTareaTiempo, CStr(tmpTareaTiempo.Tarea.id)
            End If

            tmpTareaTiempo.Tiempo = tmpTareaTiempo.Tiempo + tareaTiempo.Tiempo


        Next tareaTiempo

    Next sectorTiempo


    Set MergeEstadisticas = colOT
End Function

Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    Dim rs As Recordset
    On Error GoTo err1
    conectar.BeginTransaction
    Set rs = conectar.RSFactory("SELECT id,idDetalleOtPadre FROM detalles_pedidos dp WHERE dp.idDetalleOtPadre>0")
    Dim c As Long
    While Not rs.EOF And Not rs.BOF
        c = c + 1
        conectar.execute "update detalles_pedidos set idDetalleOtPadre=-1 where id=" & rs!idDetalleOtPadre


        rs.MoveNext
    Wend
    conectar.CommitTransaction
    Exit Sub
err1:
    conectar.RollBackTransaction

End Sub

Private Sub desactivar_Click()
    Dim A As Long
    A = grid.RowIndex(grid.row)
    If MsgBox("¿Está seguro de desactivar el pedido?", vbYesNo + vbQuestion, "Confirmación") = vbYes Then
        If Not aux_ordenTrabajo Is Nothing Then

            If DAOOrdenTrabajo.desactivar(aux_ordenTrabajo) Then
                MsgBox "La OT " & aux_ordenTrabajo.id & " se desactivo correctamente!", vbInformation, "Información"
                grid.RefreshRowIndex A
            Else
                MsgBox "Se produjo un error al desactivar!", vbCritical, "Error"
            End If
        End If
    End If
End Sub


Private Sub editOT_Click()
    If Not aux_ordenTrabajo Is Nothing Then
        Dim F As New frmPlaneamientoOTNueva
        F.OrdenTrabajoId = aux_ordenTrabajo.id
        F.Show
    End If
End Sub
Private Sub LlenarNuevaLista()
    Dim backIndex As Long
    backIndex = srow
    Set m_Incidencias = DAOIncidencias.GetCantidadIncidenciasPorReferencia(OI_OrdenesTrabajo)
    Set m_Archivos = DAOArchivo.GetCantidadArchivosPorReferencia(OA_OrdenesTrabajo)
    Dim condition As String
    condition = "{pedido}.{activo} = 1"
    Dim i As Long
    If ListBoxHasCheckedItems(Me.lstEstados) Then
        condition = condition & " AND {pedido}.{estado} IN ("
        For i = 0 To Me.lstEstados.ListCount - 1
            If Me.lstEstados.Selected(i) Then
                condition = condition & Me.lstEstados.ItemData(i) & ", "
            End If
        Next i

        condition = condition & " -1)"

        condition = Replace$(condition, "{estado}", DAOOrdenTrabajo.CAMPO_ESTADO)
    End If

    If Me.cboClientes.ListIndex <> -1 Then
        condition = condition & " AND idClienteFacturar = " & Me.cboClientes.ItemData(Me.cboClientes.ListIndex)

    End If

    If LenB(Me.txtNroOrden.text) > 0 And IsNumeric(Me.txtNroOrden.text) Then
        condition = condition & " AND {pedido}.{pedido_id} = " & CLng(Me.txtNroOrden)
        condition = Replace$(condition, "{pedido_id}", DAOOrdenTrabajo.CAMPO_ID)
    End If

    If LenB(Me.txtDescripcion.text) > 0 Then
        condition = condition & " AND {pedido}.{descripcion} LIKE '%" & Me.txtDescripcion.text & "%'"
        condition = Replace$(condition, "{descripcion}", DAOOrdenTrabajo.CAMPO_DESCRIPCION)
    End If

    If Not IsNull(Me.dtpDesde.value) Then
        condition = condition & " AND {pedido}.{fecha} >= " & conectar.Escape(Me.dtpDesde.value)
        condition = Replace$(condition, "{fecha}", DAOOrdenTrabajo.CAMPO_FECHA_ENTREGA)
    End If

    If Not IsNull(Me.dtpHasta.value) Then
        condition = condition & " AND {pedido}.{fecha} <= " & conectar.Escape(Me.dtpHasta.value)
        condition = Replace$(condition, "{fecha}", DAOOrdenTrabajo.CAMPO_FECHA_ENTREGA)
    End If

    If Me.cboCCostos.ListIndex <> -1 Then
        condition = condition & " AND {pedido}.{cliente_id} = " & Me.cboCCostos.ItemData(Me.cboCCostos.ListIndex)
        condition = Replace$(condition, "{cliente_id}", DAOOrdenTrabajo.CAMPO_CLIENTE_ID)
    End If

    condition = Replace$(condition, "{pedido}", DAOOrdenTrabajo.TABLA_PEDIDO)
    condition = Replace$(condition, "{activo}", DAOOrdenTrabajo.CAMPO_ACTIVO)
    Set m_ordenesTrabajo = DAOOrdenTrabajo.FindAll(condition)
    Me.grid.ItemCount = 0
    Me.grid.ItemCount = m_ordenesTrabajo.count
    Me.grid.RowIndex backIndex
    Me.grid.Refresh
    Me.caption = Me.Tag & " [ Cantidad: " & m_ordenesTrabajo.count & " ]"
    '    GridEXHelper.AutoSizeColumns Me.grid, True
    SeleccionarOrden
    MostrarMensajePendientes

End Sub
Private Sub SeleccionarOrden()
    Dim RowPosition As Long
    RowPosition = grid.row
    srow = Me.grid.RowIndex(Me.grid.row)
    If grid.RowIndex(RowPosition) > 0 Then
        Set aux_ordenTrabajo = m_ordenesTrabajo(grid.RowIndex(RowPosition))
    Else
        Set aux_ordenTrabajo = Nothing
    End If
End Sub

Private Sub Form_Activate()
    txtNroOrden.SetFocus
End Sub

Private Sub Form_Load()
    FormHelper.Customize Me

    Me.grid.ItemCount = 0
    GridEXHelper.CustomizeGrid Me.grid, True
    Dim i As Long
    Me.lstEstados.Clear
    For i = LBound(funciones.estados_pedidos) To UBound(funciones.estados_pedidos)
        Me.lstEstados.AddItem estados_pedidos(i)
        Me.lstEstados.ItemData(Me.lstEstados.NewIndex) = i
    Next i

    Dim tmpCliente As clsCliente
    Me.cboClientes.Clear
    For Each tmpCliente In DAOCliente.FindAll()
        Me.cboClientes.AddItem tmpCliente.razon
        Me.cboClientes.ItemData(Me.cboClientes.NewIndex) = tmpCliente.id
    Next


    DAOCliente.llenarComboXtremeSuite Me.cboCCostos, False, True, False


    funciones.FillComboBoxDateRanges Me.cboRangos
    For i = 0 To Me.cboRangos.ListCount - 1
        If Me.cboRangos.ItemData(i) = DateRangeValue.DRV_YearCurrent Then Exit For
    Next i
    Me.cboRangos.ListIndex = i


    Me.cboCCostos.ListIndex = -1
    dtpHasta.value = Null
    chkMostrarDescripcion_Click
    m_suscriptor_id = funciones.CreateGUID
    Channel.AgregarSuscriptor Me, ordenesTrabajo

    LlenarNuevaLista
    MostrarMensajePendientes

End Sub

Private Sub MostrarMensajePendientes()
    Dim msg As String
    Dim MSG2 As String
    Dim haypend As Long
    Dim hayactiv As Long

    haypend = HayPendientes
    hayactiv = HayParaActivar
    msg = "hay " & haypend & " ordenes pendientes para aprobar"
    MSG2 = "hay " & hayactiv & " ordenes aprobadas para activar"

    If haypend > 0 And hayactiv < 1 Then
        Me.LBLMENSAJE.caption = msg
    End If

    If haypend < 1 And hayactiv > 0 Then
        Me.LBLMENSAJE.caption = MSG2
    End If

    If haypend > 0 And hayactiv > 0 Then
        Me.LBLMENSAJE.caption = msg
        Me.LBLMENSAJE2.caption = MSG2
    End If

    Me.LBLMENSAJE.Visible = (haypend > 0 Or hayactiv > 0)

    Me.LBLMENSAJE2.Visible = (haypend > 0 And hayactiv > 0)

End Sub
Private Function HayPendientes() As Integer

    If Permisos.planOTaprobaciones Then
        Dim rs As Recordset
        Dim strsql As String
        HayPendientes = 0
        strsql = "select count(id) as cantidad from pedidos where estado= " & EstadoOrdenTrabajo.EstadoOT_Pendiente
        Set rs = conectar.RSFactory(strsql)
        If Not rs.EOF And Not rs.BOF Then
            HayPendientes = rs!Cantidad
        End If
    End If

End Function

Private Function HayParaActivar() As Integer

    If Permisos.planOTaprobaciones Then
        Dim rs As Recordset
        Dim strsql As String
        HayParaActivar = 0
        strsql = "select count(id) as cantidad from pedidos where estado= " & EstadoOrdenTrabajo.EstadoOT_EnEspera

        Set rs = conectar.RSFactory(strsql)
        If Not rs.EOF And Not rs.BOF Then
            HayParaActivar = rs!Cantidad
        End If
    End If

End Function
Private Sub Form_Resize()
    On Error Resume Next
    Me.grid.Width = Me.ScaleWidth - 50
    Me.grid.Height = Me.ScaleHeight - 2900    '* 0.55
    Dim Top As Double: Top = Me.ScaleHeight - 500
    Me.GroupBox1.Width = Me.ScaleWidth - 200
    Me.Command1.Top = Top
    'Me.Command3.Top = Top
    Me.cmdImprimir.Top = Top
    Me.cmdConsultas.Top = Top
    Me.btnTareasOrdenes.Top = Top
    Me.LBLMENSAJE.Top = 300
    Me.LBLMENSAJE2.Top = Me.LBLMENSAJE.Top + Me.LBLMENSAJE.Height + 100
    Me.LBLMENSAJE.Left = (Me.GroupBox1.Width - 5500)
    Me.LBLMENSAJE2.Left = Me.LBLMENSAJE.Left

End Sub

Private Sub Form_Terminate()
    Channel.RemoverSuscripcionTotal Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Channel.RemoverSuscripcionTotal Me
End Sub

Private Sub grid_ColumnHeaderClick(ByVal Column As GridEX20.JSColumn)
    ColumnHeaderClick Me.grid, Column
End Sub
Private Sub grid_DblClick()
    If grid.RowIndex(grid.row) = 0 Then Exit Sub
    If Not aux_ordenTrabajo Is Nothing Then
        frmPlaneamientoPedidosDetalle.Pedido = aux_ordenTrabajo
        frmPlaneamientoPedidosDetalle.caption = "Pedido Nro. " & Format(aux_ordenTrabajo.id, "0000")
        frmPlaneamientoPedidosDetalle.Show
    End If
End Sub

Private Sub grid_FetchIcon(ByVal RowIndex As Long, ByVal ColIndex As Integer, ByVal RowBookmark As Variant, ByVal IconIndex As GridEX20.JSRetInteger)

    On Error Resume Next
    aux_ordenTrabajo = m_ordenesTrabajo.item(RowIndex)

    If ColIndex = 11 And m_Archivos.item(aux_ordenTrabajo.id) > 0 Then
        IconIndex = 1
    End If
End Sub

Private Sub grid_GroupByBoxHeaderClick(ByVal Group As GridEX20.JSGroup)
    GroupByBoxHeaderClick Group
End Sub

Private Sub grid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 67 And Shift = 2 Then    'CTRL + C
        GridEXHelper.Grid2Clipboard Me.grid
        DoEvents
        MsgBox "La lista ha sido copiada al portapapeles.", vbInformation + vbOKOnly
    End If
End Sub

Private Sub grid_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    SeleccionarOrden
    If Not aux_ordenTrabajo Is Nothing Then
        If Button = 2 Then
            Me.numero.caption = "[ Nro. " & aux_ordenTrabajo.IdFormateado & " ]"
            Me.activarPedido.Enabled = (aux_ordenTrabajo.estado = EstadoOrdenTrabajo.EstadoOT_EnEspera)
            Me.editOT.Enabled = (aux_ordenTrabajo.estado = EstadoOrdenTrabajo.EstadoOT_Pendiente)
            Me.AprobarOT.Enabled = (aux_ordenTrabajo.estado = EstadoOrdenTrabajo.EstadoOT_Pendiente)

            'Me.mnuPrecios.Visible = aux_ordenTrabajo.EsMarco

            Me.modificarPedido.Enabled = (aux_ordenTrabajo.estado = EstadoOT_EnEspera)

            Me.desactivar.Enabled = Not (aux_ordenTrabajo.estado = EstadoOT_Desactivado)
            Me.mnuCerrar.Enabled = (aux_ordenTrabajo.estado = EstadoOrdenTrabajo.EstadoOT_EnProceso Or aux_ordenTrabajo.estado = EstadoOrdenTrabajo.EstadoOT_ProcesoCompleto) And Not aux_ordenTrabajo.EsMarco
            Me.mnuImprimir.Enabled = True    '(aux_ordenTrabajo.estado = EstadoOrdenTrabajo.EstadoOT_EnProceso Or aux_ordenTrabajo.estado = EstadoOrdenTrabajo.EstadoOT_ProcesoCompleto)
            Me.mnuFacturasAplicadas.Enabled = Not aux_ordenTrabajo.EsMarco    '(aux_ordenTrabajo.estado = EstadoOrdenTrabajo.EstadoOT_Finalizado Or aux_ordenTrabajo.estado = EstadoOrdenTrabajo.EstadoOT_Desactivado)
            Me.mnuRemitosEntregados.Enabled = Not aux_ordenTrabajo.EsMarco    '(aux_ordenTrabajo.estado = EstadoOrdenTrabajo.EstadoOT_Finalizado Or aux_ordenTrabajo.estado = EstadoOrdenTrabajo.EstadoOT_Desactivado)
            Me.mnuEntregas.Enabled = (aux_ordenTrabajo.estado = EstadoOrdenTrabajo.EstadoOT_Finalizado Or aux_ordenTrabajo.estado = EstadoOrdenTrabajo.EstadoOT_Desactivado) And Not aux_ordenTrabajo.EsMarco
            Me.mnuSeguimiento.Enabled = Not aux_ordenTrabajo.EsMarco
            Me.mnuCopiarOT.Enabled = Not aux_ordenTrabajo.EsMarco
            Me.mnuAsociarFacturaAnticipo.Enabled = (aux_ordenTrabajo.Anticipo > 0 And Not aux_ordenTrabajo.AnticipoFacturado) Or Me.grid.SelectedItems.count > 0

            Me.PopupMenu Me.pedidos
        End If
    End If
End Sub



Private Sub grid_RowFormat(RowBuffer As GridEX20.JSRowData)
    On Error Resume Next


    Set aux_ordenTrabajo = m_ordenesTrabajo.item(RowBuffer.RowIndex)

    If aux_ordenTrabajo.estado = EstadoOT_Desactivado Then
        RowBuffer.RowStyle = "Desactivado"
    Else
        If Not IsEmpty(RowBuffer.value(3)) Then
            auxLong = DateDiff("d", RowBuffer.value(3), Now)
            If auxLong = 0 Then
                RowBuffer.CellStyle(3) = "FechaEntregaVenceHoy"
            ElseIf auxLong > 0 Then
                RowBuffer.CellStyle(3) = "FechaEntregaVencida"
            End If
        End If
        If RowBuffer.value(10) > 0 Then RowBuffer.CellStyle(10) = "TieneIncidenciasArchivos"
        'If RowBuffer.value(11) > 0 Then RowBuffer.CellStyle(11) = "TieneIncidenciasArchivos"
    End If
End Sub
Private Sub grid_SelectionChange()
    SeleccionarOrden
End Sub
Private Sub grid_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    On Error Resume Next
    If m_ordenesTrabajo.count > 0 Then
        Set aux_ordenTrabajo = m_ordenesTrabajo.item(RowIndex)
        With aux_ordenTrabajo
            Values(1) = Format(.id, "0000")
            Values(2) = .ClienteFacturar.razon
            If CDbl(.FechaEntrega) > 0 Then Values(3) = .FechaEntrega
            Values(4) = IIf(.NroPresupuesto = -1, "Manual", Format(.NroPresupuesto, "0000"))
            Values(5) = .descripcion
            Values(6) = .fechaCreado
            Values(7) = .usuario.usuario

            If aux_ordenTrabajo.EsMarco Then
                Values(8) = .FechaFinMarco
            Else
                If CDbl(.FechaCerrado) > 0 Then Values(8) = .FechaCerrado
            End If
            Values(9) = funciones.estado_pedido(.estado)

            Values(10) = IIf(IsEmpty(m_Incidencias.item(.id)), 0, m_Incidencias.item(.id))
            Values(11) = IIf(IsEmpty(m_Archivos.item(.id)), 0, "(" & m_Archivos.item(.id) & ")")
            
            If .EsMarco Then
                Values(12) = "Marco"
            Else
                Values(12) = EnumTipoOT(.TipoOrden - 1)
            End If
            
            Values(13) = .cliente.razon

            'cambio entre values2 y values13 por pedido de karin hecho el 7-4
        End With
    End If
End Sub
Private Property Get ISuscriber_id() As String
    ISuscriber_id = m_suscriptor_id
End Property
Private Function ISuscriber_Notificarse(EVENTO As clsEventoObserver) As Variant
    Dim x As Integer
    If EVENTO.EVENTO = agregar_ Then
        m_ordenesTrabajo.Add EVENTO.Elemento, CStr(EVENTO.Elemento.id)
        LlenarNuevaLista
    ElseIf EVENTO.EVENTO = modificar_ Then
        For x = 1 To m_ordenesTrabajo.count
            If m_ordenesTrabajo(x).id = EVENTO.Elemento.id Then
                m_ordenesTrabajo(x).descripcion = EVENTO.Elemento.descripcion
                m_ordenesTrabajo(x).FechaEntrega = EVENTO.Elemento.FechaEntrega
                Set m_ordenesTrabajo(x).moneda = EVENTO.Elemento.moneda
                m_ordenesTrabajo(x).CantDiasAnticipo = EVENTO.Elemento.CantDiasAnticipo
                m_ordenesTrabajo(x).CantDiasSaldo = EVENTO.Elemento.CantDiasSaldo
                m_ordenesTrabajo(x).FormaDePagoAnticipo = EVENTO.Elemento.FormaDePagoAnticipo
                m_ordenesTrabajo(x).FormaDePagoSaldo = EVENTO.Elemento.FormaDePagoSaldo
                m_ordenesTrabajo(x).Anticipo = EVENTO.Elemento.Anticipo
                Set m_ordenesTrabajo(x).cliente = EVENTO.Elemento.cliente
                Set m_ordenesTrabajo(x).ClienteFacturar = EVENTO.Elemento.ClienteFacturar
                Me.grid.RefreshRowIndex x
                Exit For
            End If
        Next x
    End If

    Dim m As OrdenTrabajo
    MostrarMensajePendientes
End Function





Private Sub mnuAsociarFacturaAnticipo_Click()
    Dim selItem As JSSelectedItem

    If Me.grid.SelectedItems.count = 0 Then Exit Sub

    Dim selecFac As New frmAdminFacturasNCElegirFC
    selecFac.EstadosDocs.Add EstadoFacturaCliente.Aprobada
    selecFac.TiposDocs.Add tipoDocumentoContable.Factura
    
    
    
    Set Selecciones.Factura = Nothing
    selecFac.Show 1

    Dim otTMP As OrdenTrabajo
    Dim fac As Factura

    If IsSomething(Selecciones.Factura) Then
        Set fac = Selecciones.Factura

        For Each selItem In Me.grid.SelectedItems
            Set otTMP = m_ordenesTrabajo.item(selItem.RowIndex)

            If otTMP.Anticipo = 0 Or otTMP.AnticipoFacturado Or otTMP.AnticipoFacturadoIdFactura <> 0 Then
                MsgBox "La OT Nº " & otTMP.IdFormateado & " no se puede asociar porque no tiene anticipo o porque ya esta facturada por otra factura de anticipo.", vbExclamation + vbOKOnly
                Exit Sub
            End If

            If Not funciones.BuscarEnColeccion(fac.OTsFacturadasAnticipo, CStr(otTMP.id)) Then
                fac.OTsFacturadasAnticipo.Add otTMP, CStr(otTMP.id)
            End If
        Next selItem

        If DAOFactura.EnlazarFacturaAnticipoConOT(fac, True) Then
            MsgBox "OT/s asociada/s con factura correctamente.", vbInformation + vbOKOnly
        Else
            MsgBox "No se pudo asociar la factura con la/s OT/s. Compruebe que los montos de las OT/s y de la factura coincidan.", vbExclamation + vbOKOnly
        End If
        fac.SetToNothingOTsFacturadasAnticipo
    End If

    'If IsSomething(aux_ordenTrabajo) Then
    '    Dim selecFac As New frmAdminFacturasNCElegirFC
    '    Set Selecciones.Factura = Nothing
    '    'selecFac.idCliente = aux_ordenTrabajo.Cliente.Id
    '    selecFac.Show 1
    '    If IsSomething(Selecciones.Factura) Then
    '        Dim oldId As Long
    '        oldId = Selecciones.Factura.IdOTAnticipo
    '        Selecciones.Factura.IdOTAnticipo = aux_ordenTrabajo.Id
    '        If DAOFactura.EnlazarFacturaAnticipoConOT(Selecciones.Factura, True) Then
    '            MsgBox "OT asociada con factura correctamente.", vbInformation + vbOKOnly
    '        Else
    '            Selecciones.Factura.IdOTAnticipo = oldId
    '            MsgBox "No se pudo asociar la factura con la OT. Compruebe los montos de ambas.", vbExclamation + vbOKOnly
    '        End If
    '    End If
    'End If
End Sub

Private Sub mnuCerrar_Click()
    If Not aux_ordenTrabajo Is Nothing Then
        '        Dim F As New frmEntregas
        '
        '        'Set f.Pedido = aux_ordenTrabajo
        '        F.lblIdOT = aux_ordenTrabajo.Id
        '        F.Show

        Dim f2 As New frmEntregas2
        f2.SetOrdenTrabajo aux_ordenTrabajo
        f2.Show

    End If
End Sub

Private Sub mnuCopiarOT_Click()
    If Not aux_ordenTrabajo Is Nothing Then
        If DAOOrdenTrabajo.CopiarOT(aux_ordenTrabajo) Then
            MsgBox "Copia Existosa!!", vbInformation
            LlenarNuevaLista
        Else
            MsgBox "Se produjo algún error y no se realizó la copia!", vbCritical
        End If

    End If

End Sub

Private Sub mnuEntregas_Click()
    If Not aux_ordenTrabajo Is Nothing Then
        Dim F As New frmEntregas2
        'Se t f.Pedido = aux_ordenTrabajo
        F.SetOrdenTrabajo aux_ordenTrabajo
        F.Show
    End If
End Sub

Private Sub mnuExportarResumenGeneral_Click()

If Not aux_ordenTrabajo Is Nothing Then
    DAOOrdenTrabajo.ExportarExcelResumenGeneral aux_ordenTrabajo.id

    End If

End Sub

Private Sub mnuFacturasAplicadas_Click()
    If Not aux_ordenTrabajo Is Nothing Then
        Dim F As New frmAdminFacturasAplicadas
        F.Origen = 1
        F.idOrigen = aux_ordenTrabajo.id
        F.Show
    End If
End Sub

Private Sub mnuImprimir_Click()
    If Not aux_ordenTrabajo Is Nothing Then
        Dim F As New frmActivarPedido
        Set F.Pedido = aux_ordenTrabajo
        F.cmdActivar.Visible = False
        F.Show
    End If
End Sub

Private Sub mnuPlanificacion_Click()
    If IsSomething(aux_ordenTrabajo) Then
        Dim frm11 As New frmPlanificacionTemporal
        frm11.Pedido = aux_ordenTrabajo
        frm11.Show
    End If
End Sub

Private Sub mnuPrecios_Click()
    If Not aux_ordenTrabajo Is Nothing Then
        Dim F As New frmPlaneamientoOTNueva
        F.ActualizacionPrecios = True
        F.OrdenTrabajoId = aux_ordenTrabajo.id
        F.Show
    End If
End Sub

Private Sub mnuRemitosEntregados_Click()
    If Not aux_ordenTrabajo Is Nothing Then
        Dim F As New frmRemitosEntregados
        F.Origen = 1
        F.idPedidoEntrega = aux_ordenTrabajo.id
        F.caption = "Nro." & aux_ordenTrabajo.id
        F.Show
    End If
End Sub
Private Sub mnuScanear_Click()
    If Not aux_ordenTrabajo Is Nothing Then
        Dim archivos As New classArchivos
        archivos.escanearDocumento OrigenArchivos.OA_OrdenesTrabajo, aux_ordenTrabajo.id
    End If
End Sub
Private Sub mnuSeguimiento_Click()
    If Not aux_ordenTrabajo Is Nothing Then
        Dim F As New frmPlaneamientoSeguimiento
        F.txtOt = aux_ordenTrabajo.id
        F.Show
    End If
End Sub

Private Sub modificarPedido_Click()
    Dim A As Long
    If Not aux_ordenTrabajo Is Nothing Then
        A = grid.RowIndex(grid.row)
        If MsgBox("¿Está seguro de hacer editable este pedido?", vbYesNo, "Confirmación") = vbNo Then Exit Sub
        If DAOOrdenTrabajo.HacerEditable(aux_ordenTrabajo) Then
            MsgBox "Pedido disponible para editar!", vbInformation, "Información"
            grid.RefreshRowIndex A
        Else
            MsgBox "Se produjo un error!", vbCritical, "Error"
        End If
    End If
End Sub

Private Sub txtDescripcion_GotFocus()
    foco Me.txtDescripcion
End Sub

Private Sub txtNroOrden_GotFocus()
    foco Me.txtNroOrden
End Sub

Private Sub VerDetalles_Click()
    If Not aux_ordenTrabajo Is Nothing Then
        If aux_ordenTrabajo.EsMarco Then
            Dim f32342 As New frmContAbiertoDetalle
            Set f32342.OTMarco = aux_ordenTrabajo
            f32342.Show
        Else
            Dim frmaaaa As New frmPlaneamientoPedidosDetalle

            frmaaaa.Pedido = aux_ordenTrabajo
            frmaaaa.Show
        End If
    End If
End Sub
Private Sub verHistorial_Click()
    If Not aux_ordenTrabajo Is Nothing Then
        DAOOrdenTrabajoHistorial.getAllByOrdenTrabajo aux_ordenTrabajo.id, True
    End If
End Sub
Private Sub verIncidencias_Click()
    If Not aux_ordenTrabajo Is Nothing Then
        frmVerIncidencias.referencia = aux_ordenTrabajo.id
        frmVerIncidencias.Origen = OrigenIncidencias.OI_OrdenesTrabajo
        frmVerIncidencias.Show
    End If
End Sub



