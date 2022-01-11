VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmAdminFacturasEmitidas 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Comprobantes Emitidos"
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12405
   Icon            =   "frmFacturasEmitidas.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6975
   ScaleWidth      =   12405
   Begin XtremeSuiteControls.GroupBox grp 
      Height          =   2055
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   19095
      _Version        =   786432
      _ExtentX        =   33681
      _ExtentY        =   3625
      _StockProps     =   79
      Caption         =   "Filtros"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.ProgressBar progreso 
         Height          =   420
         Left            =   14760
         TabIndex        =   42
         Top             =   1500
         Visible         =   0   'False
         Width           =   4215
         _Version        =   786432
         _ExtentX        =   7435
         _ExtentY        =   741
         _StockProps     =   93
         Appearance      =   6
         BarColor        =   65280
      End
      Begin XtremeSuiteControls.CheckBox chkCredito 
         Height          =   255
         Left            =   6600
         TabIndex        =   31
         Top             =   1605
         Width           =   2055
         _Version        =   786432
         _ExtentX        =   3625
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "DE CRÉDITO (MI PYME)"
         UseVisualStyle  =   -1  'True
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   315
         Left            =   17400
         TabIndex        =   23
         Top             =   480
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.TextBox txtReferencia 
         Height          =   300
         Left            =   1620
         TabIndex        =   20
         Top             =   1530
         Width           =   3490
      End
      Begin XtremeSuiteControls.ComboBox cboClientes 
         Height          =   315
         Left            =   1620
         TabIndex        =   19
         Top             =   720
         Width           =   3510
         _Version        =   786432
         _ExtentX        =   6191
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Appearance      =   6
      End
      Begin VB.TextBox txtNroFactura 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   1620
         TabIndex        =   4
         Top             =   330
         Width           =   1290
      End
      Begin VB.TextBox txtOrdenCompra 
         Height          =   300
         Left            =   1620
         TabIndex        =   3
         Top             =   1110
         Width           =   1740
      End
      Begin VB.TextBox txtRemitoAplicado 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   3960
         TabIndex        =   2
         Top             =   1110
         Width           =   1170
      End
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   285
         Left            =   5190
         TabIndex        =   5
         Top             =   735
         Width           =   375
         _Version        =   786432
         _ExtentX        =   661
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "X"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton cmdBuscar 
         Default         =   -1  'True
         Height          =   420
         Left            =   9840
         TabIndex        =   6
         Top             =   1500
         Width           =   1965
         _Version        =   786432
         _ExtentX        =   3466
         _ExtentY        =   741
         _StockProps     =   79
         Caption         =   "Buscar"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.GroupBox GroupBox2 
         Height          =   1170
         Left            =   9840
         TabIndex        =   7
         Top             =   240
         Width           =   4695
         _Version        =   786432
         _ExtentX        =   8281
         _ExtentY        =   2064
         _StockProps     =   79
         Caption         =   "Fecha Emision"
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.DateTimePicker dtpDesde 
            Height          =   315
            Left            =   960
            TabIndex        =   8
            Top             =   615
            Width           =   1470
            _Version        =   786432
            _ExtentX        =   2593
            _ExtentY        =   556
            _StockProps     =   68
            CheckBox        =   -1  'True
            Format          =   1
         End
         Begin XtremeSuiteControls.DateTimePicker dtpHasta 
            Height          =   315
            Left            =   3120
            TabIndex        =   9
            Top             =   630
            Width           =   1470
            _Version        =   786432
            _ExtentX        =   2593
            _ExtentY        =   556
            _StockProps     =   68
            CheckBox        =   -1  'True
            Format          =   1
         End
         Begin XtremeSuiteControls.ComboBox cboRangos 
            Height          =   315
            Left            =   960
            TabIndex        =   10
            Top             =   240
            Width           =   3645
            _Version        =   786432
            _ExtentX        =   6429
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Style           =   2
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.Label Label6 
            Height          =   195
            Left            =   2520
            TabIndex        =   13
            Top             =   675
            Width           =   420
            _Version        =   786432
            _ExtentX        =   741
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Hasta"
            AutoSize        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label5 
            Height          =   195
            Left            =   360
            TabIndex        =   12
            Top             =   660
            Width           =   465
            _Version        =   786432
            _ExtentX        =   820
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Desde"
            AutoSize        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label7 
            Height          =   195
            Left            =   360
            TabIndex        =   11
            Top             =   285
            Width           =   480
            _Version        =   786432
            _ExtentX        =   847
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Rango"
            AutoSize        =   -1  'True
         End
      End
      Begin XtremeSuiteControls.PushButton cmdImprimir 
         Height          =   420
         Left            =   13440
         TabIndex        =   18
         Top             =   1500
         Width           =   1050
         _Version        =   786432
         _ExtentX        =   1852
         _ExtentY        =   741
         _StockProps     =   79
         Caption         =   "Imprimir"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnExpotar 
         Height          =   420
         Left            =   12240
         TabIndex        =   22
         ToolTipText     =   "Exporta s?lo pendientes"
         Top             =   1500
         Width           =   1050
         _Version        =   786432
         _ExtentX        =   1852
         _ExtentY        =   741
         _StockProps     =   79
         Caption         =   "Exportar"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cboPuntosVenta 
         Height          =   360
         Left            =   3585
         TabIndex        =   25
         Top             =   300
         Width           =   1530
         _Version        =   786432
         _ExtentX        =   2699
         _ExtentY        =   635
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   2
         Appearance      =   6
         Text            =   "cboMoneda"
         DropDownItemCount=   3
      End
      Begin XtremeSuiteControls.PushButton PushButton3 
         Height          =   285
         Left            =   5190
         TabIndex        =   26
         Top             =   338
         Width           =   375
         _Version        =   786432
         _ExtentX        =   661
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "X"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cboEstados 
         Height          =   360
         Left            =   6600
         TabIndex        =   28
         Top             =   315
         Width           =   2355
         _Version        =   786432
         _ExtentX        =   4154
         _ExtentY        =   635
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   2
         Appearance      =   6
         Text            =   "cboMoneda"
         DropDownItemCount=   3
      End
      Begin XtremeSuiteControls.PushButton PushButton4 
         Height          =   285
         Left            =   9000
         TabIndex        =   30
         Top             =   360
         Width           =   375
         _Version        =   786432
         _ExtentX        =   661
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "X"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cboEstadosSaldada 
         Height          =   360
         Left            =   6600
         TabIndex        =   36
         Top             =   720
         Width           =   2355
         _Version        =   786432
         _ExtentX        =   4154
         _ExtentY        =   635
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   2
         Appearance      =   6
         Text            =   "cboMoneda"
         DropDownItemCount=   3
      End
      Begin XtremeSuiteControls.PushButton PushButton5 
         Height          =   285
         Left            =   9000
         TabIndex        =   37
         Top             =   780
         Width           =   375
         _Version        =   786432
         _ExtentX        =   661
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "X"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cboEstadoAfip 
         Height          =   360
         Left            =   6600
         TabIndex        =   39
         Top             =   1155
         Width           =   2355
         _Version        =   786432
         _ExtentX        =   4154
         _ExtentY        =   635
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   2
         Appearance      =   6
         Text            =   "cboMoneda"
         DropDownItemCount=   3
      End
      Begin XtremeSuiteControls.PushButton cmdLimpiarCboEstadoAfip 
         Height          =   285
         Left            =   9000
         TabIndex        =   40
         Top             =   1200
         Width           =   375
         _Version        =   786432
         _ExtentX        =   661
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "X"
         UseVisualStyle  =   -1  'True
      End
      Begin VB.Label lblExportando 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000A&
         Caption         =   "Exportando..."
         Height          =   255
         Left            =   17160
         TabIndex        =   43
         Top             =   1200
         Visible         =   0   'False
         Width           =   1815
      End
      Begin XtremeSuiteControls.Label Label14 
         Height          =   285
         Left            =   5520
         TabIndex        =   41
         Top             =   1200
         Width           =   1035
         _Version        =   786432
         _ExtentX        =   1826
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "Estado AFIP"
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label Label12 
         Height          =   285
         Left            =   5880
         TabIndex        =   38
         Top             =   758
         Width           =   675
         _Version        =   786432
         _ExtentX        =   1191
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "Saldada"
         Alignment       =   1
      End
      Begin VB.Label lblTotalNeto 
         AutoSize        =   -1  'True
         Caption         =   "Total Filtrado $:"
         Height          =   195
         Left            =   14760
         TabIndex        =   35
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblTotalIVA 
         AutoSize        =   -1  'True
         Caption         =   "Total Filtrado $:"
         Height          =   195
         Left            =   14760
         TabIndex        =   34
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblTotalPercepciones 
         AutoSize        =   -1  'True
         Caption         =   "Total Filtrado $:"
         Height          =   195
         Left            =   14760
         TabIndex        =   33
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         Caption         =   "Total Filtrado $:"
         Height          =   195
         Left            =   14760
         TabIndex        =   32
         Top             =   1080
         Width           =   1095
      End
      Begin XtremeSuiteControls.Label Label10 
         Height          =   285
         Left            =   6000
         TabIndex        =   29
         Top             =   360
         Width           =   555
         _Version        =   786432
         _ExtentX        =   979
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "Estado"
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label Label9 
         Height          =   285
         Left            =   3240
         TabIndex        =   27
         Top             =   330
         Width           =   585
         _Version        =   786432
         _ExtentX        =   1032
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "PV"
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Referencia"
         Height          =   195
         Left            =   630
         TabIndex        =   21
         Top             =   1590
         Width           =   900
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nro Comprobrante"
         Height          =   270
         Left            =   30
         TabIndex        =   17
         Top             =   375
         Width           =   1500
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Orden Compra"
         Height          =   270
         Left            =   270
         TabIndex        =   16
         Top             =   1125
         Width           =   1260
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente"
         Height          =   270
         Left            =   270
         TabIndex        =   15
         Top             =   735
         Width           =   1260
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Rto"
         Height          =   270
         Left            =   3495
         TabIndex        =   14
         Top             =   1125
         Width           =   420
      End
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   4635
      Left            =   120
      TabIndex        =   0
      Top             =   2280
      Width           =   19125
      _ExtentX        =   33734
      _ExtentY        =   8176
      Version         =   "2.0"
      PreviewRowIndent=   100
      DefaultGroupMode=   1
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      GroupFooterStyle=   2
      PreviewColumn   =   "preview"
      PreviewRowLines =   1
      RowHeight       =   26
      ColumnAutoResize=   -1  'True
      MethodHoldFields=   -1  'True
      ImageCount      =   1
      ImagePicture1   =   "frmFacturasEmitidas.frx":000C
      RowHeaders      =   -1  'True
      DataMode        =   99
      CardSpacing     =   16
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   25
      Column(1)       =   "frmFacturasEmitidas.frx":0326
      Column(2)       =   "frmFacturasEmitidas.frx":04C6
      Column(3)       =   "frmFacturasEmitidas.frx":05DA
      Column(4)       =   "frmFacturasEmitidas.frx":0716
      Column(5)       =   "frmFacturasEmitidas.frx":0856
      Column(6)       =   "frmFacturasEmitidas.frx":09AA
      Column(7)       =   "frmFacturasEmitidas.frx":0B0E
      Column(8)       =   "frmFacturasEmitidas.frx":0D3E
      Column(9)       =   "frmFacturasEmitidas.frx":0E86
      Column(10)      =   "frmFacturasEmitidas.frx":1006
      Column(11)      =   "frmFacturasEmitidas.frx":1102
      Column(12)      =   "frmFacturasEmitidas.frx":1202
      Column(13)      =   "frmFacturasEmitidas.frx":1362
      Column(14)      =   "frmFacturasEmitidas.frx":14B6
      Column(15)      =   "frmFacturasEmitidas.frx":15FE
      Column(16)      =   "frmFacturasEmitidas.frx":1756
      Column(17)      =   "frmFacturasEmitidas.frx":189E
      Column(18)      =   "frmFacturasEmitidas.frx":19E6
      Column(19)      =   "frmFacturasEmitidas.frx":1ACA
      Column(20)      =   "frmFacturasEmitidas.frx":1C1A
      Column(21)      =   "frmFacturasEmitidas.frx":1D3E
      Column(22)      =   "frmFacturasEmitidas.frx":1EB6
      Column(23)      =   "frmFacturasEmitidas.frx":2026
      Column(24)      =   "frmFacturasEmitidas.frx":2132
      Column(25)      =   "frmFacturasEmitidas.frx":221E
      FormatStylesCount=   16
      FormatStyle(1)  =   "frmFacturasEmitidas.frx":231E
      FormatStyle(2)  =   "frmFacturasEmitidas.frx":2456
      FormatStyle(3)  =   "frmFacturasEmitidas.frx":2506
      FormatStyle(4)  =   "frmFacturasEmitidas.frx":25BA
      FormatStyle(5)  =   "frmFacturasEmitidas.frx":2692
      FormatStyle(6)  =   "frmFacturasEmitidas.frx":274A
      FormatStyle(7)  =   "frmFacturasEmitidas.frx":282A
      FormatStyle(8)  =   "frmFacturasEmitidas.frx":28B6
      FormatStyle(9)  =   "frmFacturasEmitidas.frx":2996
      FormatStyle(10) =   "frmFacturasEmitidas.frx":2A46
      FormatStyle(11) =   "frmFacturasEmitidas.frx":2AFA
      FormatStyle(12) =   "frmFacturasEmitidas.frx":2BAA
      FormatStyle(13) =   "frmFacturasEmitidas.frx":2C5A
      FormatStyle(14) =   "frmFacturasEmitidas.frx":2D0E
      FormatStyle(15) =   "frmFacturasEmitidas.frx":2DE6
      FormatStyle(16) =   "frmFacturasEmitidas.frx":2ECA
      ImageCount      =   1
      ImagePicture(1) =   "frmFacturasEmitidas.frx":2FAA
      PrinterProperties=   "frmFacturasEmitidas.frx":32C4
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   15630
      Top             =   4500
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin XtremeSuiteControls.CheckBox chkVerObservaciones 
      Height          =   225
      Left            =   45
      TabIndex        =   24
      Top             =   6105
      Width           =   1695
      _Version        =   786432
      _ExtentX        =   2990
      _ExtentY        =   397
      _StockProps     =   79
      Caption         =   "Ver Observaciones"
      Appearance      =   6
      Value           =   1
   End
   Begin XtremeSuiteControls.TaskDialog taskDialog 
      Left            =   14955
      Top             =   750
      _Version        =   786432
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   0
      WindowTitle     =   "TaskDialog1"
   End
   Begin VB.Menu mnuFacturas 
      Caption         =   "armnuFacturas"
      Visible         =   0   'False
      Begin VB.Menu NRO 
         Caption         =   "nro"
         Enabled         =   0   'False
      End
      Begin VB.Menu editar 
         Caption         =   "Editar..."
      End
      Begin VB.Menu separador2 
         Caption         =   "-"
      End
      Begin VB.Menu aprobarFactura 
         Caption         =   "Aprobar localmente..."
      End
      Begin VB.Menu mnuEnviarAfip 
         Caption         =   "Enviar a AFIP..."
      End
      Begin VB.Menu separador 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAprobarEnviar 
         Caption         =   "Aprobar localmente y Enviar a AFIP..."
      End
      Begin VB.Menu mnuDesaprobarFactura 
         Caption         =   "Desaprobar..."
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu sepa3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRechazo 
         Caption         =   "Rechazo de comprobantes (FCE)"
      End
      Begin VB.Menu ImprimirFactura 
         Caption         =   "Imprimir..."
      End
      Begin VB.Menu AnularFactura 
         Caption         =   "Anular"
      End
      Begin VB.Menu desAnular 
         Caption         =   "Quitar Anulaci?n"
         Visible         =   0   'False
      End
      Begin VB.Menu aplicar 
         Caption         =   "Aplicar Recibo..."
      End
      Begin VB.Menu aplicarNCaFC 
         Caption         =   "Aplicar a Factura o ND..."
      End
      Begin VB.Menu mnuAplicarANCOLD 
         Caption         =   "Aplicar a NC..."
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu o 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCrearCopiaFactura 
         Caption         =   "Crear copia a partir de comprobante"
      End
      Begin VB.Menu mnuFechaPagoPropuesta 
         Caption         =   "Establecer Fecha Pago Propuesta"
      End
      Begin VB.Menu mnuFechaEntrega 
         Caption         =   "Establecer Fecha Entrega..."
      End
      Begin VB.Menu sdf 
         Caption         =   "-"
      End
      Begin VB.Menu verHistorialFactura 
         Caption         =   "Ver Historial..."
      End
      Begin VB.Menu mnuArchivos 
         Caption         =   "Archivos Asociados..."
      End
      Begin VB.Menu LineaUlt 
         Caption         =   "-"
      End
      Begin VB.Menu verFactura 
         Caption         =   "Ver Detalle..."
      End
      Begin VB.Menu archivos 
         Caption         =   "Archivos Asociados..."
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu scanear 
         Caption         =   "Adquirir..."
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmAdminFacturasEmitidas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements ISuscriber

Dim vId As String
Dim facturas As Collection
Dim Factura As Factura
Dim m_Archivos As Dictionary


Private Sub AnularFactura_Click()
    Dim r As Long
    r = Me.GridEX1.RowIndex(Me.GridEX1.row)
    If MsgBox("¿Desea anular el comprobante?", vbYesNo, "Confirmacion") = vbYes Then

        If DAOFactura.Anular(Factura) Then
            MsgBox "Comprobante anulado con éxito!", vbInformation, "Información"
            Me.GridEX1.RefreshRowIndex r
        Else
            MsgBox "Hubo un error. No se anulo el comprobante!", vbCritical, "Error"
        End If

    End If
End Sub



Private Sub aplicarNCaFC_Click()
On Error GoTo err1
    If MsgBox("¿Seguro de aplicar comprobante?", vbYesNo, "Confirmación") = vbYes Then
        'seleccionar factura para aplicar
        Set Selecciones.Factura = Nothing
          Dim F As New frmAdminFacturasNCElegirFC
        
        F.idCliente = Factura.cliente.id
            F.TiposDocs.Add tipoDocumentoContable.Factura
            
            If Factura.TipoDocumento = tipoDocumentoContable.notaCredito Then
                F.TiposDocs.Add tipoDocumentoContable.notaDebito
            End If
                 If Factura.TipoDocumento = tipoDocumentoContable.notaDebito Then
                F.TiposDocs.Add tipoDocumentoContable.notaCredito
            End If
            
            F.EstadosDocs.Add EstadoFacturaCliente.Aprobada
            F.Show 1

        If IsSomething(Selecciones.Factura) Then
             If DAOFactura.aplicarNCaFC(Selecciones.Factura.id, Factura.id) Then
                MsgBox "Aplicación existosa!", vbInformation, "Información"
            End If
        End If
    End If
Exit Sub
err1:
MsgBox Err.Description, vbCritical, "Error"
End Sub




Private Sub aprobarFactura_Click()
    On Error GoTo err1
    Dim g As Long
    Dim msgadicional As String
    msgadicional = ""
    If MsgBox("¿Desea aprobar localmente el comprobante?", vbYesNo + vbQuestion, "Confirmacion") = vbYes Then
        g = Me.GridEX1.RowIndex(Me.GridEX1.row)
        If DAOFactura.aprobarV2(Factura, True, False) Then
           
            If Factura.Tipo.PuntoVenta.EsElectronico And Not Factura.Tipo.PuntoVenta.CaeManual And Not Factura.AprobadaAFIP Then
              msgadicional = "Esta factura deberá enviarse a la afip"
           End If
            If Factura.Tipo.PuntoVenta.EsElectronico And Factura.Tipo.PuntoVenta.CaeManual And Not Factura.AprobadaAFIP Then
              msgadicional = "Recuerde agregar al comprobante: CAE y fecha de vencimiento del CAE "
           End If
            
            Dim msg As String
            msg = "Comprobante aprobado con éxito!"
            If IsSomething(Factura.CaeSolicitarResponse) Then
             If LenB(Factura.CaeSolicitarResponse.observaciones) > 5 Then
            
              msg = msg & Chr(10) & Factura.CaeSolicitarResponse.observaciones
            End If
            
            If LenB(msgadicional) > 0 Then
                msg = msg & Chr(10) & msgadicional
            End If
            
            End If
            MsgBox msg, vbInformation, "Información"
            
            Me.GridEX1.RefreshRowIndex g
            Me.txtNroFactura.SetFocus
            
        Else
            GoTo err1
        End If
    End If
    Exit Sub
err1:
    'MsgBox "Factura no aprobada, compruebe:" & vbNewLine & "Si la factura es de anticipo, compruebe que el valor de la misma sea el mismo que el anticipo de la OT." & vbNewLine & "Que el detalle del remito no este ya facturado." & vbNewLine & Err.Description, vbCritical

    MsgBox Err.Description, vbCritical, Err.Source
    Me.GridEX1.RefreshRowIndex g
End Sub

Private Sub archivos_Click()
    Dim F As New frmArchivos2
    F.Origen = 101
    F.ObjetoId = Factura.id
    F.caption = "Comprobante " & Factura.GetShortDescription(False, True)
    F.Show
End Sub


Private Sub btnExpotar_Click()

'FUNCIÓN PARA EXPORTAR A EXCEL
    
'    Dim id As Long
'
'    If (Me.cboClientes.ListIndex > 0) Then
'        id = Me.cboClientes.ItemData(Me.cboClientes.ListIndex)
'    Else
'        id = -1
'    End If
'
'    Dim col As New Collection
'
'    If (id > 0) Then
'        Set col = DAOFactura.FindAllByEstadoSaldoAndCliente(NoSaldada, EstadoFacturaCliente.Aprobada, id)
'    Else
'        Set col = DAOFactura.FindAllByEstadoSaldoAndCliente(NoSaldada, EstadoFacturaCliente.Aprobada)
'
'    End If
    
'INICIA EL PROGRESSBAR Y LO MUESTRA
    Me.progreso.Visible = True
    Me.lblExportando.Visible = True
    
'DEFINE EL VALOR MINIMO Y EL MAXIMO DEL PROGRESSBAR (CANTIDAD DE DATOS EN LA COLECCIÓN COL)
    progreso.min = 0
    progreso.max = facturas.count
    

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

    xlWorksheet.Cells(1, 1).value = "Reporte de comprobantes emitidos"
    
'    If (id > 0) Then
'        xlWorksheet.Cells(1, 2).value = DAOCliente.BuscarPorID(id).razon
'    Else
'        xlWorksheet.Cells(1, 2).value = "Todos"
'    End If

    xlWorksheet.Columns(4).HorizontalAlignment = xlLeft
    xlWorksheet.Columns(12).HorizontalAlignment = xlLeft
    
    xlWorksheet.Cells(2, 1).value = "Comprobante"
    xlWorksheet.Cells(2, 2).value = "Emision"
    xlWorksheet.Cells(2, 3).value = "Moneda"
    xlWorksheet.Cells(2, 4).value = "Detalle"
    xlWorksheet.Cells(2, 5).value = "Importe en " & DAOMoneda.FindFirstByPatronOrDefault.NombreCorto

    xlWorksheet.Cells(2, 6).value = "Vencimiento"
    xlWorksheet.Cells(2, 7).value = "Atraso"
    xlWorksheet.Cells(2, 8).value = "Entrega"
    xlWorksheet.Cells(2, 9).value = "Atraso"
    
'    If (id < 0) Then xlWorksheet.Cells(2, 10).value = "Cliente"

    xlWorksheet.Cells(2, 10).value = "Cliente"
    
    xlWorksheet.Cells(2, 11).value = "Cuit"
    
    xlWorksheet.Cells(2, 12).value = "Observaciones"
    
    xlWorksheet.Cells(2, 13).value = "Observaciones Cancela"
    
    
    Dim idx As Integer
    idx = 3
    
    Dim fac As Factura
  
'DEFINE EL CONTADOR DEL PROGRESSBAR Y LO INICIA EN 0
    Dim d As Long
    d = 0
    
    
'   Set facturas = DAOFactura.FindAll(filtro)
'
'   Dim F As Factura
   
'   For Each F In facturas
'
'   For Each fac In col
   
   For Each fac In facturas
        
        xlWorksheet.Cells(idx, 1).value = fac.GetShortDescription(False, True)
        'xlWorksheet.Cells(idx, 1).value = fac.NumeroFormateado
        'xlWorksheet.Cells(idx, 1).value = fac.numero
        xlWorksheet.Cells(idx, 2).value = fac.FechaEmision
        xlWorksheet.Cells(idx, 3).value = fac.moneda.NombreCorto
        xlWorksheet.Cells(idx, 4).value = fac.OrdenCompra

        If fac.TipoDocumento = tipoDocumentoContable.notaCredito Then
            xlWorksheet.Cells(idx, 5).value = funciones.RedondearDecimales(fac.TotalEstatico.Total * fac.CambioAPatron) * -1
        Else
            xlWorksheet.Cells(idx, 5).value = funciones.RedondearDecimales(fac.TotalEstatico.Total * fac.CambioAPatron)
        End If
        
        xlWorksheet.Cells(idx, 6).value = fac.Vencimiento
        xlWorksheet.Cells(idx, 7).value = fac.StringDiasAtraso
        
        If (fac.DiferenciaDiasEntrega <> -1) Then
            xlWorksheet.Cells(idx, 8).value = Format(fac.FechaEntrega, "dd/mm/yyyy")
            xlWorksheet.Cells(idx, 9).value = fac.DiferenciaDiasEntrega & " dias"
        Else
            xlWorksheet.Cells(idx, 8).value = "no definida"
            xlWorksheet.Cells(idx, 9).value = 0
        End If

'        If (id < 0) Then xlWorksheet.Cells(idx, 10).value = fac.cliente.razon

        xlWorksheet.Cells(idx, 10).value = fac.cliente.razon
        
        xlWorksheet.Cells(idx, 11).value = fac.cliente.Cuit
        
        
        xlWorksheet.Cells(idx, 12).value = fac.observaciones
        xlWorksheet.Cells(idx, 13).value = fac.observaciones_cancela
        
        xlWorksheet.Cells(idx, 14).value = fac.id
        xlWorksheet.Cells(idx, 15).value = fac.Cancelada
        

        idx = idx + 1
        
'POR CADA ITERACION SUMA UN VALOR A LA VARIABLE D DEL PROGRESSBAR
        d = d + 1
        progreso.value = d
        
        
        Next
        
        xlWorksheet.Cells(idx, 5).Formula = "=SUM(E3:E" & idx - 1 & ")"

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
        
'REINICIA EL PROGRESSBAR Y LO OCULTA
        progreso.value = 0
        Me.progreso.Visible = False
        Me.lblExportando.Visible = False
        
End Sub

Private Sub cboRangos_Click()
    funciones.CalculateDateRange Me.cboRangos, Me.dtpDesde, Me.dtpHasta
End Sub


Private Sub chkVerObservaciones_Click()
    verObservaciones
End Sub


Private Sub verObservaciones()
    If Me.chkVerObservaciones Then
        Me.GridEX1.PreviewRowLines = 1
    Else
        Me.GridEX1.PreviewRowLines = 0
    End If
End Sub


Private Sub cmdBuscar_Click()
    llenarGrilla
End Sub

Private Sub cmdImprimir_Click()


    With Me.GridEX1.PrinterProperties
        .FitColumns = True
        .RepeatHeaders = True
        .Orientation = jgexPPLandscape
        .HeaderString(jgexHFCenter) = "Emitidos"
        .FooterString(jgexHFCenter) = Now
     '202
        .FooterDistance = 1500
    .FooterString(jgexHFLeft) = lblTotalNeto & Chr(10) & lblTotalIVA & Chr(10) & lblTotalPercepciones & Chr(10) & lblTotal
    '202
    
    End With
    Load frmPrintPreview
    frmPrintPreview.Move Me.Left, Me.Top, Me.Width, Me.Height
    GridEX1.PrintPreview frmPrintPreview.GEXPreview1
    frmPrintPreview.Show 1
End Sub



Private Sub cmdLimpiarCboEstadoAfip_Click()
    Me.cboEstadoAfip.ListIndex = -1
End Sub

Private Sub Command1_Click()
    DAODetalleOrdenTrabajo.arreglarCagada
End Sub

Private Sub editar_Click()

    Dim f_c3h3 As New frmFacturaEdicion
    f_c3h3.idFactura = Factura.id
    f_c3h3.Show

End Sub

Private Sub Form_Load()
    FormHelper.Customize Me
    GridEXHelper.CustomizeGrid Me.GridEX1, True, False
    DAOCliente.llenarComboXtremeSuite Me.cboClientes, False, True, False
    Me.cboClientes.ListIndex = -1

    vId = funciones.CreateGUID
    Channel.AgregarSuscriptor Me, FacturaCliente_
    
'Modificaci?n 15/05/20 (Se muestran todos los comprobanes sin filtrar por punto de venta)
    DAOPuntoVenta.llenarComboXtremeSuite Me.cboPuntosVenta, False
    
    cboEstados.Clear
    cboEstados.AddItem "En Edición"
    cboEstados.ItemData(cboEstados.NewIndex) = 1
    cboEstados.AddItem "Aprobadas"
    cboEstados.ItemData(cboEstados.NewIndex) = 2
    cboEstados.AddItem "Anuladas"
    cboEstados.ItemData(cboEstados.NewIndex) = 3
    
    Me.cboEstadosSaldada.Clear
    cboEstadosSaldada.AddItem "No Saldado"
    cboEstadosSaldada.ItemData(cboEstadosSaldada.NewIndex) = 0
    cboEstadosSaldada.AddItem "Saldado total"
    cboEstadosSaldada.ItemData(cboEstadosSaldada.NewIndex) = 1
    cboEstadosSaldada.AddItem "Saldado parcial"
    cboEstadosSaldada.ItemData(cboEstadosSaldada.NewIndex) = 2
    cboEstadosSaldada.AddItem "Cancelado total por NC"
    cboEstadosSaldada.ItemData(cboEstadosSaldada.NewIndex) = 3
    cboEstadosSaldada.AddItem "Cancelado parcial por NC"
    cboEstadosSaldada.ItemData(cboEstadosSaldada.NewIndex) = 3
    
    
     Me.cboEstadoAfip.Clear
    cboEstadoAfip.AddItem "Sólo informadas"
    cboEstadoAfip.ItemData(cboEstadoAfip.NewIndex) = 0
    cboEstadoAfip.AddItem "Sólo no informadas"
    cboEstadoAfip.ItemData(cboEstadoAfip.NewIndex) = 1

    Dim i As Integer
    funciones.FillComboBoxDateRanges Me.cboRangos
    For i = 0 To Me.cboRangos.ListCount - 1
        If Me.cboRangos.ItemData(i) = DateRangeValue.DRV_YearCurrent Then Exit For
    Next i
    Me.cboRangos.ListIndex = i
    llenarGrilla
    verObservaciones
End Sub

Private Sub llenarGrilla()

    Dim cliente As clsCliente
    Dim filtro As String
    Set m_Archivos = DAOArchivo.GetCantidadArchivosPorReferencia(OA_factura)

    Me.GridEX1.ItemCount = 0
    filtro = "1=1"
    
    If Me.cboClientes.ListIndex >= 0 Then
        filtro = filtro & " and idCliente=" & cboClientes.ItemData(Me.cboClientes.ListIndex)
    End If

    If Me.cboPuntosVenta.ListIndex >= 0 Then
        filtro = filtro & " and pv.id=" & cboPuntosVenta.ItemData(Me.cboPuntosVenta.ListIndex)
    End If
    
    If Me.cboEstados.ListIndex >= 0 Then
        filtro = filtro & " and AdminFacturas.estado=" & cboEstados.ItemData(Me.cboEstados.ListIndex)
    End If

    If Me.cboEstadosSaldada.ListIndex >= 0 Then
        filtro = filtro & " and AdminFacturas.saldada=" & cboEstadosSaldada.ItemData(Me.cboEstadosSaldada.ListIndex)
    End If
    
    If Me.chkCredito.value > 0 Then
        filtro = filtro & " and AdminFacturas.EsCredito=" & Me.chkCredito.value
    End If
    
    If LenB(Me.txtOrdenCompra) > 0 Then
        filtro = filtro & " and OrdenCompra like '%" & Trim(Me.txtOrdenCompra) & "%'"
    End If
    
    If LenB(Me.txtNroFactura) > 0 And IsNumeric(Me.txtNroFactura) Then
        filtro = filtro & " and nroFactura=" & Me.txtNroFactura
    End If

    If Not IsNull(Me.dtpDesde.value) Then
        filtro = filtro & " AND AdminFacturas.FechaEmision >= " & conectar.Escape(Me.dtpDesde.value)
    End If

    If Not IsNull(Me.dtpHasta.value) Then
        filtro = filtro & " AND AdminFacturas.FechaEmision <= " & conectar.Escape(Me.dtpHasta.value)
    End If

    If LenB(Me.txtRemitoAplicado.text) > 0 Then
        filtro = filtro & " and AdminFacturas.id IN (SELECT fd.idFactura FROM AdminFacturasDetalleNueva fd INNER JOIN entregas e ON e.id = fd.idEntrega INNER JOIN remitos r ON r.id = e.Remito WHERE r.numero = " & Me.txtRemitoAplicado.text & ")"
    End If

    If LenB(Me.txtReferencia.text) > 0 Then
        filtro = filtro & " and AdminFacturas.OrdenCompra like '%" & Trim(Me.txtReferencia.text) & "%'"
    End If
    
   If Me.cboEstadoAfip.ListIndex = 0 Then
        filtro = filtro & " and AdminFacturas.aprobacion_afip=1"
   End If
   
   If Me.cboEstadoAfip.ListIndex = 1 Then
        filtro = filtro & " and AdminFacturas.aprobacion_afip=0"
   End If
    

   Set facturas = DAOFactura.FindAll(filtro)
   
   Dim F As Factura
   Dim c As Integer
   
   For Each F In facturas

        Dim Total As Double
        Dim totalNG As Double
        Dim TotalIVATodo As Double
        Dim totalPercepcionesIIBB As Double
        
        Dim Percepcion As Double


        If F.TipoDocumento = tipoDocumentoContable.notaCredito Then c = -1 Else c = 1

      
           Total = Total + MonedaConverter.ConvertirForzado2(F.TotalEstatico.Total * c, MonedaConverter.Patron.id, F.moneda.id, F.CambioAPatron)
    
           TotalIVATodo = TotalIVATodo + MonedaConverter.ConvertirForzado2(F.TotalEstatico.TotalIVADiscrimandoONo * c, MonedaConverter.Patron.id, F.moneda.id, F.CambioAPatron)
    
           totalNG = totalNG + MonedaConverter.ConvertirForzado2(F.TotalEstatico.TotalNetoGravado * c, MonedaConverter.Patron.id, F.moneda.id, F.CambioAPatron)
           
           Percepcion = F.TotalEstatico.TotalPercepcionesIB * c
           
           totalPercepcionesIIBB = totalPercepcionesIIBB + MonedaConverter.ConvertirForzado2(F.TotalEstatico.TotalPercepcionesIB * c, MonedaConverter.Patron.id, F.moneda.id, F.CambioAPatron)
        
    Next


    Me.lblTotal = "Total: $ " & funciones.FormatearDecimales(Total)
    Me.lblTotalPercepciones = "Total Percepciones: $ " & funciones.FormatearDecimales(totalPercepcionesIIBB)
    Me.lblTotalIVA = "Total IVA: $ " & funciones.FormatearDecimales(TotalIVATodo)
    Me.lblTotalNeto = "Total NG: $ " & funciones.FormatearDecimales(totalNG)

    Me.GridEX1.ItemCount = 0
    Me.GridEX1.ItemCount = facturas.count
    
    Me.caption = "Emitidos [Cantidad: " & facturas.count & "]"

' Desabilito la apertura directa de la Factura al encontrar exacto
    'If facturas.count = 1 Then
    '   Dim f_c3h3 As New frmFacturaEdicion
    '    f_c3h3.idFactura = facturas(1).id
    '    f_c3h3.Show
    'End If

    'GridEXHelper.AutoSizeColumns Me.GridEX1
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Me.GridEX1.Width = Me.ScaleWidth
    Me.GridEX1.Height = Me.ScaleHeight - 1900
    Me.grp.Width = Me.GridEX1.Width - 180
End Sub

Private Sub Form_Terminate()
    Channel.RemoverSuscripcionTotal Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Channel.RemoverSuscripcionTotal Me
End Sub

Private Sub GridEX1_BeforePrintPage(ByVal PageNumber As Long, ByVal nPages As Long)
    GridEX1.PrinterProperties.FooterString(jgexHFRight) = "Página" & PageNumber & " de " & nPages
End Sub

Private Sub GridEX1_ColumnHeaderClick(ByVal Column As GridEX20.JSColumn)
    GridEXHelper.ColumnHeaderClick Me.GridEX1, Column
End Sub

Private Sub GridEX1_DblClick()
    verFactura_Click
End Sub

Private Sub GridEX1_FetchIcon(ByVal RowIndex As Long, ByVal ColIndex As Integer, ByVal RowBookmark As Variant, ByVal IconIndex As GridEX20.JSRetInteger)
    If ColIndex = 20 And m_Archivos.item(Factura.id) > 0 Then IconIndex = 1
End Sub

Private Sub GridEX1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If facturas.count > 0 Then
        SeleccionarFactura
        If Button = 2 Then
            Me.NRO.caption = "[ Nro. " & Format(Factura.numero, "0000") & " ]"

            If Factura.Tipo.PuntoVenta.CaeManual Then
              Me.mnuEnviarAfip.caption = "Cargar CAE manualmente"
            Else
             Me.mnuEnviarAfip.caption = "Informar a AFIP"
            End If

            'Me.mnuFechaPagoPropuesta.Enabled = False


'actualizo leyenda para aplicación

'Aplicar a Factura o ND...
If Factura.TipoDocumento = tipoDocumentoContable.notaCredito Then
    Me.aplicarNCaFC.caption = "Aplicar a Factura o ND..."
End If

If Factura.TipoDocumento = tipoDocumentoContable.notaDebito Then
        Me.aplicarNCaFC.caption = "Aplicar a Factura o NC..."
End If



' Si el estado del comprobante es EN PROCESO
            If Factura.estado = EstadoFacturaCliente.EnProceso Then   'no se aprob? localmente
                Me.aplicarNCaFC.Enabled = False
                Me.aplicarNCaFC.Visible = False
                Me.editar.Enabled = True
                Me.editar.Visible = True
                Me.desAnular.Visible = False
                Me.AnularFactura.Visible = False
                Me.AnularFactura.Enabled = False
                Me.aprobarFactura.Enabled = Permisos.AdminFacturasAprobaciones
                Me.aprobarFactura.Visible = True
                Me.mnuEnviarAfip.Visible = False
                Me.ImprimirFactura.Enabled = False
                Me.ImprimirFactura.Visible = False
                Me.mnuDesaprobarFactura.Visible = False
                Me.aplicar.Enabled = False
                Me.aplicar.Visible = False
                Me.mnuFechaPagoPropuesta.Enabled = True
                Me.mnuFechaPagoPropuesta.Visible = True
                Me.mnuFechaEntrega.Enabled = True
                Me.mnuFechaEntrega.Visible = True
                                      
                                      
               'opci?n combinada solo v?lida para comprobantes electr?nicos no aprobados localmente
               '23-08-2020
               If Factura.esCredito Then
                      If Factura.TipoDocumento = tipoDocumentoContable.Factura Then
                            Me.mnuAprobarEnviar.Visible = Factura.Tipo.PuntoVenta.EsElectronico And Permisos.AdminFacturasAprobaciones
                            Me.mnuAprobarEnviar.Enabled = Factura.Tipo.PuntoVenta.EsElectronico And Permisos.AdminFacturasAprobaciones
                     Else
                            Me.mnuAprobarEnviar.Visible = False
                            Me.mnuAprobarEnviar.Enabled = False
                     End If
               Else
                          Me.mnuAprobarEnviar.Visible = Factura.Tipo.PuntoVenta.EsElectronico And Permisos.AdminFacturasAprobaciones And Not Factura.Tipo.PuntoVenta.CaeManual
                          Me.mnuAprobarEnviar.Enabled = Factura.Tipo.PuntoVenta.EsElectronico And Permisos.AdminFacturasAprobaciones And Not Factura.Tipo.PuntoVenta.CaeManual
                     End If
               
             End If
             
              
' Si el comprobante NO EST? EN PROCESO
            If Factura.estado <> EstadoFacturaCliente.EnProceso And Factura.estado <> EstadoFacturaCliente.Anulada Then     'se aprobo localmente y no est? anulada
                Me.editar.Enabled = False
                Me.editar.Visible = False
                Me.desAnular.Visible = False
                Me.aprobarFactura.Enabled = False
                Me.aprobarFactura.Visible = False
                
                Me.mnuFechaEntrega.Enabled = True
                Me.mnuFechaEntrega.Visible = True
                Me.mnuFechaPagoPropuesta.Enabled = True
                Me.mnuFechaPagoPropuesta.Visible = True
               
               'opci?n combinada solo v?lida para comprobantes electr?nicos no aprobados localmente
               '23-08-2020
                Me.mnuAprobarEnviar.Visible = False
                Me.mnuAprobarEnviar.Enabled = False
                
                If Factura.Tipo.PuntoVenta.EsElectronico Then
                    Me.AnularFactura.Visible = False 'si es electronico no se puede anular comprobante
                    Me.AnularFactura.Enabled = False
                    
                            If Factura.AprobadaAFIP Then
                            
                                        'Me.mnuEditarCAE.Enabled = False
                                        'Me.mnuEditarCAE.Visible = False
                                        
                                        Me.mnuEnviarAfip.Enabled = False
                                        Me.mnuEnviarAfip.Visible = False
                                        
                                        Me.aplicar.Visible = False
                                        Me.aplicar.Enabled = False '(factura.Saldado = TipoSaldadoFactura.NoSaldada Or factura.Saldado = TipoSaldadoFactura.saldadoTotal)
                                        
                                        'Desde una NC
                                        'Si es Credito no muestra la posibilidad de aplicar NC a Factura
                                        Me.aplicarNCaFC.Visible = Not Factura.esCredito And Factura.TipoDocumento <> tipoDocumentoContable.Factura And Not Factura.Tipo.PuntoVenta.CaeManual
                                        Me.aplicarNCaFC.Enabled = Not Factura.esCredito And Factura.TipoDocumento <> tipoDocumentoContable.Factura And Not Factura.Tipo.PuntoVenta.CaeManual
                                        
                                        '----------------------------------------------------------------------------------------------------------------------------------------
                                        '21/01/2021 dnemer
                                        'Agrego estas dos lineas para que se habilite el menu aplicar sobre las NC cuando sean PV 6 o sea PV Cae Manual
                                        'Me.aplicarNCaFC.Visible = Factura.Tipo.PuntoVenta.CaeManual And Factura.TipoDocumento <> tipoDocumentoContable.Factura
                                        'Me.aplicarNCaFC.Enabled = Factura.Tipo.PuntoVenta.CaeManual And Factura.TipoDocumento <> tipoDocumentoContable.Factura
                                        
                                        
                                        
                                        'Desde una FC
                                        'Si es de Credito no muestra la posibilidad de aplicar Factura a una NC
                                        'Me.mnuAplicarANC.Visible = Not Factura.esCredito And Factura.TipoDocumento = tipoDocumentoContable.Factura
                                        'Me.mnuAplicarANC.Enabled = Not Factura.esCredito And Factura.TipoDocumento = tipoDocumentoContable.Factura
                                        
                                        'Me.mnuAplicarANC.Visible = False
                                        'Me.mnuAplicarANC.Enabled = False '(factura.TipoDocumento = tipoDocumentoContable.notaDebito Or factura.TipoDocumento = tipoDocumentoContable.factura) And (factura.estado = EstadoFacturaCliente.Aprobada)
                        
                              Else
                              
                             'Me.mnuEditarCAE.Enabled = Permisos.AdminFacturasAprobaciones ' Not factura.EstaImpresa And factura.Tipo.PuntoVenta.CaeManual
                             'Me.mnuEditarCAE.Visible = True
                             'Me.mnuEnviarAfip.Enabled = Permisos.AdminFacturasAprobaciones
                                       
                             Me.mnuEnviarAfip.Visible = True
                             Me.mnuEnviarAfip.Enabled = True
                                       
                             Me.aplicar.Visible = (Factura.Saldado = TipoSaldadoFactura.NoSaldada Or Factura.Saldado <> TipoSaldadoFactura.SaldadoParcial)
                             Me.aplicar.Enabled = (Factura.Saldado = TipoSaldadoFactura.NoSaldada Or Factura.Saldado <> TipoSaldadoFactura.saldadoTotal)
                             Me.aplicarNCaFC.Visible = (Factura.TipoDocumento = tipoDocumentoContable.notaCredito Or Factura.TipoDocumento = tipoDocumentoContable.notaDebito) And (Factura.estado = EstadoFacturaCliente.Aprobada)
                             Me.aplicarNCaFC.Enabled = (Factura.TipoDocumento = tipoDocumentoContable.notaCredito Or Factura.TipoDocumento = tipoDocumentoContable.notaDebito) And (Factura.estado = EstadoFacturaCliente.Aprobada)
                             'Me.mnuAplicarANC.Visible = (Factura.TipoDocumento = tipoDocumentoContable.notaDebito Or Factura.TipoDocumento = tipoDocumentoContable.Factura) And (Factura.estado = EstadoFacturaCliente.Aprobada)
                             'Me.mnuAplicarANC.Enabled = (Factura.TipoDocumento = tipoDocumentoContable.notaDebito Or Factura.TipoDocumento = tipoDocumentoContable.Factura) And (Factura.estado = EstadoFacturaCliente.Aprobada)
                                
                            End If
                
                    Else
                        Me.mnuEnviarAfip.Enabled = False
                        Me.mnuEnviarAfip.Visible = False
                        'Me.mnuEditarCAE.Enabled = False ' Not factura.EstaImpresa And factura.Tipo.PuntoVenta.CaeManual
                        'Me.mnuEditarCAE.Visible = False '  Not factura.EstaImpresa And factura.Tipo.PuntoVenta.CaeManual
                        Me.AnularFactura.Visible = True
                        Me.AnularFactura.Enabled = True
                        Me.mnuDesaprobarFactura.Visible = False
                        Me.aplicar.Enabled = (Factura.Saldado = TipoSaldadoFactura.NoSaldada Or Factura.Saldado <> TipoSaldadoFactura.saldadoTotal)
                        Me.aplicar.Visible = (Factura.Saldado = TipoSaldadoFactura.NoSaldada Or Factura.Saldado <> TipoSaldadoFactura.saldadoTotal)
                     
                        Me.aplicarNCaFC.Enabled = (Factura.TipoDocumento = tipoDocumentoContable.notaCredito Or Factura.TipoDocumento = tipoDocumentoContable.notaDebito) And (Factura.estado = EstadoFacturaCliente.Aprobada)
                        'Me.mnuAplicarANC.Enabled = (Factura.TipoDocumento = tipoDocumentoContable.notaDebito Or Factura.TipoDocumento = tipoDocumentoContable.Factura) And (Factura.estado = EstadoFacturaCliente.Aprobada)
                        'Me.aplicarNCaFC.Visible = (Factura.TipoDocumento = tipoDocumentoContable.notaCredito) And (Factura.estado = EstadoFacturaCliente.Aprobada)
                        'Me.mnuAplicarANC.Visible = (Factura.TipoDocumento = tipoDocumentoContable.notaDebito Or Factura.TipoDocumento = tipoDocumentoContable.Factura) And (Factura.estado = EstadoFacturaCliente.Aprobada)
                   
                    End If
         
     Me.ImprimirFactura.Enabled = True
     Me.ImprimirFactura.Visible = True
         'si es FCE muestro el form para cambiar el estado de rechazo
          Me.mnuRechazo.Visible = Factura.esCredito
        
            
            End If
     
            
     
            If Factura.estado = EstadoFacturaCliente.Anulada Then
                Me.mnuFechaEntrega.Enabled = False
                Me.editar.Enabled = False
                Me.AnularFactura.Visible = False
                Me.aprobarFactura.Enabled = False
                Me.ImprimirFactura.Enabled = False
                Me.aplicar.Enabled = False
                Me.aplicarNCaFC.Enabled = False
                'Me.mnuAplicarANC = False
                
             End If
                
                
            If Factura.estado = EstadoFacturaCliente.CanceladaNC Then
                Me.editar.Enabled = False
                'Me.mnuFechaEntrega.Enabled = False
                Me.AnularFactura.Enabled = False
                Me.AnularFactura.Visible = False
                Me.aprobarFactura.Enabled = False
                Me.ImprimirFactura.Enabled = True
                Me.aplicar.Enabled = False
                Me.aplicarNCaFC.Enabled = False
                'Me.mnuAplicarANC = False
                
         End If
                        
            Me.archivos.Enabled = Permisos.SistemaArchivosVer
            
            Me.separador.Visible = Me.mnuEnviarAfip.Visible Or Me.aprobarFactura
            Me.sepa3.Visible = Me.mnuDesaprobarFactura.Visible Or Me.mnuAprobarEnviar.Visible
         
            
            If Factura.Saldado <> NoSaldada Then
                Me.mnuFechaEntrega.Enabled = False
                Me.mnuFechaEntrega.Visible = False
                Me.mnuFechaPagoPropuesta.Enabled = False
                Me.mnuFechaPagoPropuesta.Visible = False
            End If

            Me.PopupMenu Me.mnuFacturas
            
        End If
    End If
End Sub

Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)
    On Error GoTo err1
    Set Factura = facturas.item(RowBuffer.RowIndex)
    
    If Factura.estado = EstadoFacturaCliente.Anulada Then
        RowBuffer.RowStyle = "anulada"
         Else
         If Factura.estado = EstadoFacturaCliente.EnProceso Then
            RowBuffer.CellStyle(12) = "pendiente"
        ElseIf Factura.estado = EstadoFacturaCliente.Aprobada Then
            RowBuffer.CellStyle(12) = "aprobada"
        End If

'        If factura.Saldado = TipoSaldadoFactura.NoSaldada Or factura.Saldado = TipoSaldadoFactura.SaldadoParcial Or factura.Saldado = TipoSaldadoFactura.notaCredito Then
'            If factura.EstaAtrasada Then
'                RowBuffer.CellStyle(16) = "no_saldada"
'            Else
'                RowBuffer.CellStyle(16) = "no_vencida"
'            End If
'        ElseIf factura.Saldado = saldadoTotal Then
'            RowBuffer.CellStyle(16) = "saldada"
'        End If

'Nemer agrega formato especial a la Celda
         If Factura.AprobadaAFIP = True Then
            RowBuffer.CellStyle(13) = "informadaAfip"
        Else
            RowBuffer.CellStyle(13) = "No_informadaAfip"
        End If
        
    End If
    Exit Sub
err1:

End Sub

Private Sub GridEX1_SelectionChange()
    SeleccionarFactura
End Sub

Private Sub SeleccionarFactura()
    On Error Resume Next
    Set Factura = facturas.item(Me.GridEX1.RowIndex(Me.GridEX1.row))

End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    On Error GoTo err1
    Set Factura = facturas.item(RowIndex)
    
    
    Values(1) = Factura.GetShortDescription(True, False)    'enums.EnumTipoDocumentoContable(Factura.TipoDocumento)

    If IsSomething(Factura.Tipo) Then
        Values(2) = Factura.Tipo.TipoFactura.Tipo
    End If

    Values(3) = Factura.Tipo.PuntoVenta.PuntoVenta

    If Factura.esCredito Then
        Values(4) = "(FCE)"
    Else
        Values(4) = ""
    End If


If Factura.Tipo.PuntoVenta.EsElectronico And Not Factura.AprobadaAFIP And Factura.estado <> EstadoFacturaCliente.EnProceso Then
    Values(5) = "Nro. Pendiente"
Else
    Values(5) = Factura.NumeroFormateado
End If
    
    Values(6) = Factura.FechaEmision
    Values(7) = funciones.FormatearDecimales(Factura.TotalEstatico.Total)

    If Factura.moneda.id = 0 Then
        Values(8) = Factura.moneda.NombreCorto
    Else
        Values(8) = Factura.moneda.NombreCorto & " " & Factura.CambioAPatron
    End If

    Values(9) = funciones.FormatearDecimales(Factura.TotalEstatico.Total * Factura.CambioAPatron)

    Values(10) = Factura.OrdenCompra
    Values(11) = Factura.cliente.razon
    Values(12) = enums.EnumEstadoDocumentoContable(Factura.estado)
    
    If Factura.AprobadaAFIP = True Then
    Values(13) = "Informada"
    Else
    Values(13) = "No Informada"
    End If
    
    
    Values(14) = EnumTipoSaldadoFactura(Factura.Saldado)
    

    Values(15) = Factura.Vencimiento

    
    Values(16) = Factura.StringDiasAtraso
    
    
    Values(17) = Factura.usuarioCreador.usuario
    
    
    'Values(18) = Factura.observaciones
  
    If Factura.Tipo.PuntoVenta.EsElectronico Or Factura.Tipo.PuntoVenta.CaeManual Then
    
    If Factura.estado = EstadoFacturaCliente.EnProceso Then
         Values(18) = "Comprobante en proceso"
    Else
    
        If LenB(Factura.CAE) <= 2 Then
          Values(18) = "/ CAE no definido"
        Else
   
          Values(18) = "CAE: " & Factura.CAE & " / " & Factura.observaciones & " / " & Factura.observaciones_cancela
        End If
    End If
End If

    If IsSomething(Factura.UsuarioAprobacion) Then
        Values(19) = Factura.UsuarioAprobacion.usuario
    Else
        Values(19) = vbNullString
    End If
     
    
    If CDbl(Factura.FechaPropuestaPago) > 0 Then Values(20) = Factura.FechaPropuestaPago

    If Factura.DiferenciaDiasEntrega = -1 Then
        Values(20) = "Defina fecha"
    Else

        If CDbl(Factura.FechaEntrega) > 0 And Factura.estado <> EstadoFacturaCliente.Anulada Then
            If Factura.Saldado = NoSaldada Then Values(21) = Format(Factura.FechaEntrega, "dd/mm/yyyy") & " (" & Factura.DiferenciaDiasEntrega & ")"

            '       If Factura.Saldado = SaldadoTotal Then
            '      Values(18) = "Saldada"
            '    Else
            '           Values(18) = "Anulada"
            '    End If
            '        Values(18) = Factura.FechaEntrega
            '    End If

        Else
            If Factura.estado = EstadoFacturaCliente.Anulada Then
                Values(21) = "Anulada"
            Else
                Values(21) = Factura.FechaEntrega
            End If
        End If

    End If


    Values(22) = Factura.TasaAjusteMensual

    Values(23) = "(" & Val(m_Archivos.item(Factura.id)) & ")"
    
    Values(24) = Factura.id
    
    Values(25) = Factura.RecibosAplicadosId
   
    Exit Sub
err1:
End Sub

Private Sub ImprimirFactura_Click()

    On Error GoTo err451:
    Dim clasea As New classAdministracion
    Dim veces As Long


    If Factura.Tipo.PuntoVenta.EsElectronico Or Factura.Tipo.PuntoVenta.CaeManual Then
        veces = clasea.facturaImpresa(Factura.id)
        If veces > 0 Then
            If MsgBox("Este comprobante ya fue generarlo" & Chr(10) & "¿Desea volver a generarlo?", vbYesNo, "Confirmación") = vbYes Then
                'DAOFactura.GenerarPdf (Factura.id)
                DAOFactura.VerFacturaElectronicaParaImpresion (Factura.id)
            End If
        Else
                DAOFactura.VerFacturaElectronicaParaImpresion (Factura.id)
            End If
    Else

        veces = clasea.facturaImpresa(Factura.id)
        If veces = 0 Or veces = -1 Then
            If MsgBox("'¿Desea imprimir este comprobante?", vbYesNo, "Confirmación") = vbYes Then
               cd.Flags = cdlPDUseDevModeCopies
                cd.Copies = 3
                cd.ShowPrinter
                Dim i As Long
                For i = 1 To cd.Copies
                    DAOFactura.Imprimir Factura.id
                Next
            End If

        ElseIf veces > 0 Then
            If MsgBox("Este comprobante ya fue impreso." & Chr(10) & "¿Desea volver a imprimirlo?", vbYesNo, "Confirmación") = vbYes Then
                cd.Flags = cdlPDUseDevModeCopies
                cd.Copies = 3
                cd.ShowPrinter

                For i = 1 To cd.Copies
                    DAOFactura.Imprimir Factura.id
                Next i
            End If

        End If
    End If
    Exit Sub
err451:

End Sub

Private Property Get ISuscriber_id() As String
    ISuscriber_id = vId
End Property

Private Function ISuscriber_Notificarse(EVENTO As clsEventoObserver) As Variant
    Dim tmp As Factura
    If EVENTO.EVENTO = agregar_ Then
        llenarGrilla
        Me.GridEX1.Refresh
    ElseIf EVENTO.EVENTO = modificar_ Then
        Set tmp = EVENTO.Elemento

        Dim i As Long
        For i = facturas.count To 1 Step -1

            If facturas(i).id = tmp.id Then

                '                Set Factura = facturas(i)
                '                Factura.Id = tmp.Id
                '                Factura.Detalles = tmp.Detalles
                '                Factura.estado = tmp.estado
                '                Factura.OrdenCompra = tmp.OrdenCompra
                '                Factura.estado = tmp.estado
                '                Factura.Observaciones = tmp.Observaciones
                '                Factura.TasaAjusteMensual = tmp.TasaAjusteMensual
                '                Set Factura.Cliente = tmp.Cliente

                facturas.remove i
                If facturas.count > 0 Then
                    If i = 1 Then    'ver esto cuand oes un solo item
                        facturas.Add tmp, CStr(tmp.id), 1
                    ElseIf (i - 1) = facturas.count Then
                        facturas.Add tmp, CStr(tmp.id), , i - 1
                    Else
                        facturas.Add tmp, CStr(tmp.id), i
                    End If
                Else
                    facturas.Add tmp, CStr(tmp.id)
                End If

                'DAOFactura.FindById(tmp.Id, True)

                Me.GridEX1.RefreshRowIndex i
                Exit For

            End If

        Next

    End If


End Function

'Private Sub mnuAplicarANC_Click()
'  If MsgBox("?Seguro de aplicar a FC a NC?", vbYesNo, "Confirmaci?n") = vbYes Then
'        'seleccionar factura para aplicar
'        Set Selecciones.Factura = Nothing
'          Dim F As New frmAdminFacturasNCElegirFC
'
'        F.idCliente = Factura.cliente.id
'            F.TiposDocs.Add tipoDocumentoContable.notaCredito
'            F.EstadosDocs.Add EstadoFacturaCliente.Aprobada
'            F.Show 1
'
'        If IsSomething(Selecciones.Factura) Then
'            If DAOFactura.aplicarNCaFC(Factura.id, Selecciones.Factura.id) Then
'                MsgBox "Aplicaci?n existosa!", vbInformation, "Informaci?n"
'            Else
'                MsgBox "Se produjo un error, se abortan los cambios!", vbCritical, "Error"
'            End If
'        End If
'    End If
'End Sub

Private Sub mnuAprobarSinEnvio_Click()

'On Error GoTo err1
'    Dim g As Long
'
'    If MsgBox("?Desea aprobar el comprobante SIN ENV?AR A LA AFIP?", vbYesNo + vbQuestion, "Confirmacion") = vbYes Then
'        g = Me.GridEX1.RowIndex(Me.GridEX1.row)
'
'        If DAOFactura.aprobar(factura, False) Then
'
'
'              MsgBox "Recuerde agregar al comprobante: CAE y fecha de vencimiento del CAE ", vbInformation, "Informaci?n"
'
'
''            Dim msg As String
''            msg = "Comprobante aprobado con ?xito!"
''            If IsSomething(Factura.CaeSolicitarResponse) Then
''             If LenB(Factura.CaeSolicitarResponse.observaciones) > 5 Then
''
''              msg = msg & Chr(10) & Factura.CaeSolicitarResponse.observaciones
''            End If
''            End If
''            MsgBox msg, vbInformation, "Informaci?n"
'
'            Me.GridEX1.RefreshRowIndex g
'            Me.txtNroFactura.SetFocus
'        Else
'            GoTo err1
'        End If
'    End If
'    Exit Sub
'err1:
'    'MsgBox "Factura no aprobada, compruebe:" & vbNewLine & "Si la factura es de anticipo, compruebe que el valor de la misma sea el mismo que el anticipo de la OT." & vbNewLine & "Que el detalle del remito no este ya facturado." & vbNewLine & Err.Description, vbCritical
'
'    MsgBox Err.Description, vbCritical, Err.Source
'    Me.GridEX1.RefreshRowIndex g



End Sub

Private Sub mnuAprobarEnviar_Click()
    On Error GoTo err1
    Dim g As Long
    Dim msgadicional As String
    msgadicional = ""
    If MsgBox("¿Desea aprobar localmente el comprobante e informarlo a AFIP?", vbYesNo + vbQuestion, "Confirmacion") = vbYes Then
        g = Me.GridEX1.RowIndex(Me.GridEX1.row)
        If DAOFactura.aprobarV2(Factura, True, True) Then
            
            
            
            If Factura.Tipo.PuntoVenta.EsElectronico And Not Factura.Tipo.PuntoVenta.CaeManual And Not Factura.AprobadaAFIP Then
              msgadicional = "Esta factura deberá enviarse a la afip"
           End If
            If Factura.Tipo.PuntoVenta.EsElectronico And Factura.Tipo.PuntoVenta.CaeManual And Not Factura.AprobadaAFIP Then
              msgadicional = "Recuerde agregar al comprobante: CAE y fecha de vencimiento del CAE "
           End If
            
            Dim msg As String
            msg = "Comprobante aprobado con exito!"
            If IsSomething(Factura.CaeSolicitarResponse) Then
                 If LenB(Factura.CaeSolicitarResponse.observaciones) > 5 Then
                
                  msg = msg & Chr(10) & Factura.CaeSolicitarResponse.observaciones
                End If
                
                If LenB(msgadicional) > 0 Then
                    msg = msg & Chr(10) & msgadicional
                End If
                
            End If
            MsgBox msg, vbInformation, "Información"
            
            Me.GridEX1.RefreshRowIndex g
            Me.txtNroFactura.SetFocus
        Else
            GoTo err1
        End If
    End If
    Exit Sub
err1:
    'MsgBox "Factura no aprobada, compruebe:" & vbNewLine & "Si la factura es de anticipo, compruebe que el valor de la misma sea el mismo que el anticipo de la OT." & vbNewLine & "Que el detalle del remito no este ya facturado." & vbNewLine & Err.Description, vbCritical

    MsgBox Err.Description, vbCritical, Err.Source
    Me.GridEX1.RefreshRowIndex g
End Sub

Private Sub mnuArchivos_Click()
    Dim archi As New frmArchivos2

    archi.Origen = OrigenArchivos.OA_factura
    archi.ObjetoId = Factura.id
    archi.caption = Factura.GetShortDescription(False, True)
    archi.Show

End Sub

Private Sub mnuCrearCopiaFactura_Click()
    Me.taskDialog.Reset
    Me.taskDialog.MessageBoxStyle = True
    Me.taskDialog.WindowTitle = "Copia fiel de Comprobante"
    Me.taskDialog.MainInstructionText = "¿De que tipo es el nuevo comprobante?"
    Me.taskDialog.ContentText = "Elija el tipo de comprobante para el nuevo comprobante."
    taskDialog.RelativePosition = False

    Me.taskDialog.CommonButtons = 0
    taskDialog.CommonButtons = taskDialog.CommonButtons Or xtpTaskButtonOk
    taskDialog.CommonButtons = taskDialog.CommonButtons Or xtpTaskButtonCancel

    taskDialog.DefaultRadioButton = -1
    taskDialog.AddRadioButton "Factura", tipoDocumentoContable.Factura
    taskDialog.AddRadioButton "Nota de Débito", tipoDocumentoContable.notaDebito
    taskDialog.AddRadioButton "Nota de Crédito", tipoDocumentoContable.notaCredito


    taskDialog.MainIcon = xtpTaskIconInformation

    If taskDialog.ShowDialog = xtpTaskButtonOk Then
        If Me.taskDialog.DefaultRadioButton = -1 Then
            MsgBox "Debe seleccionar un tipo para el nuevo comprobante.", vbExclamation + vbOKOnly
        Else
            Dim newFact As Factura
            Set newFact = DAOFactura.CrearCopiaFiel(Factura, Me.taskDialog.DefaultRadioButton)
            If IsSomething(newFact) Then
                MsgBox "Se creó un nuevo comprobante (" & newFact.GetShortDescription(False, True) & ")", vbInformation + vbOKOnly
            Else
                MsgBox "Hubo un error al copiar la factura.", vbCritical + vbOKOnly
            End If
        End If
    End If



End Sub

Private Sub mnuDesaprobarFactura_Click()

    On Error GoTo err1
    Dim g As Long

    If MsgBox("¿Desea desaprobar localmente el comprobante?", vbYesNo + vbQuestion, "Confirmacion") = vbYes Then
        g = Me.GridEX1.RowIndex(Me.GridEX1.row)
        If DAOFactura.desaprobar(Factura) Then
            MsgBox "Comprobante desaprobado con éxito!", vbInformation, "Información"
            Me.GridEX1.RefreshRowIndex g
            Me.txtNroFactura.SetFocus
        Else
            GoTo err1
        End If
    End If
    Exit Sub
err1:
    MsgBox "Factura no aprobada, compruebe:" & vbNewLine & "Si la factura es de anticipo, compruebe que el valor de la misma sea el mismo que el anticipo de la OT." & vbNewLine & "Que el detalle del remito no este ya facturado." & vbNewLine & Err.Description, vbCritical
End Sub

Private Sub mnuEditarCAE_Click()
'    Dim g As Long
'    g = Me.GridEX1.RowIndex(Me.GridEX1.row)
'
'    Dim F As New frmAdminFacturasAprobarSinAfip
'    Set F.factura = factura
'    F.Show 1
'
' Me.GridEX1.RefreshRowIndex g


End Sub

Private Sub mnuEnviarAfip_Click()
On Error GoTo err1
    Dim g As Long

    If Not Factura.Tipo.PuntoVenta.EsElectronico Then
      Err.Raise 300, "Informar AFIP", "No puede informar un comprobante de un PV no catalogado como electrónico."
    End If
      
'    If factura.Tipo.PuntoVenta.EsElectronico And factura.Tipo.PuntoVenta.CaeManual Then
'       Err.Raise 301, "Informar AFIP", "No puede informar un comprobante del PV indicado"
'   End If

    If Factura.Tipo.PuntoVenta.EsElectronico And Factura.AprobadaAFIP Then
            Err.Raise 302, "Informar AFIP", "No puede informar un comprobante que ya fue informado."
    End If
    
    If Factura.Tipo.PuntoVenta.CaeManual Then
    
            Dim gg As Long
            gg = Me.GridEX1.RowIndex(Me.GridEX1.row)
        
            Dim F As New frmAdminFacturasAprobarSinAfip
            Set F.Factura = Factura
            F.Show 1
        
         Me.GridEX1.RefreshRowIndex gg
            
    Else
     If MsgBox("¿Desea informar  el comprobante?", vbYesNo + vbQuestion, "Confirmacion") = vbYes Then
        g = Me.GridEX1.RowIndex(Me.GridEX1.row)
        If DAOFactura.aprobarV2(Factura, False, True) Then
            
   
         
            Dim msg As String
            msg = "Comprobante informado con éxito!"
            If IsSomething(Factura.CaeSolicitarResponse) Then
             If LenB(Factura.CaeSolicitarResponse.observaciones) > 5 Then
            
                  msg = msg & Chr(10) & Factura.CaeSolicitarResponse.observaciones
                End If
            End If
            MsgBox msg, vbInformation, "Información"
            
            Me.GridEX1.RefreshRowIndex g
            Me.txtNroFactura.SetFocus
        Else
            GoTo err1
        End If
    End If
End If
    
    
    Exit Sub
err1:
   
    MsgBox Err.Description, vbCritical, Err.Source
    Me.GridEX1.RefreshRowIndex g
    
End Sub

Private Sub mnuFechaEntrega_Click()
    Dim fechaAnterior As String
    Dim fechaPosterior As String
    Dim nuevaFecha As Date
    Dim Update As Boolean

    If CDbl(Factura.FechaEntrega) > 0 Then fechaAnterior = Factura.FechaEntrega

    fechaPosterior = InputBox("Establezca fecha de entrega", "Fecha de Entrega", fechaAnterior)

    If LenB(fechaPosterior) = 0 Then
        nuevaFecha = 1 / 1 / 2005
        Update = True
    Else
        If IsDate(fechaPosterior) Then
            nuevaFecha = CDate(fechaPosterior)
            Update = True
        Else
            MsgBox "La fecha no es válida.", vbOKOnly + vbExclamation, "Fecha"
        End If
    End If

    If Update Then
        Factura.FechaEntrega = nuevaFecha
        If DAOFactura.Guardar(Factura) Then
            Me.GridEX1.RefreshRowIndex (Me.GridEX1.row)
        Else
            MsgBox "Error al guardar la factura.", vbOKOnly + vbCritical, "Error"
        End If
    End If

End Sub

Private Sub mnuFechaPagoPropuesta_Click()
    Dim fechaAnterior As String
    Dim fechaPosterior As String
    Dim nuevaFecha As Date
    Dim Update As Boolean

    If CDbl(Factura.FechaPropuestaPago) > 0 Then fechaAnterior = Factura.FechaPropuestaPago

    fechaPosterior = InputBox("Establezca fecha de pago propuesta", "Fecha de Pago", fechaAnterior)

    If LenB(fechaPosterior) = 0 Then
        Update = (MsgBox("¿Desea dejar en blanco la fecha de pago propuesta?", vbYesNo + vbQuestion) = vbYes)
    Else
        If IsDate(fechaPosterior) Then
            nuevaFecha = CDate(fechaPosterior)
            Update = True
        Else
            MsgBox "La fecha no es válida.", vbOKOnly + vbExclamation, "Fecha"
        End If
    End If

    If Update Then
        Factura.FechaPropuestaPago = nuevaFecha
        If DAOFactura.Guardar(Factura) Then
            Me.GridEX1.ReBind
        Else
            MsgBox "Error al guardar la factura.", vbOKOnly + vbCritical, "Error"
        End If
    End If

End Sub

Private Sub mnuRechazo_Click()
Dim F As New frmAdminFacturaRechazoAfip
            Set F.Factura = Factura
            F.Show
End Sub

Private Sub PushButton1_Click()
    Me.cboClientes.ListIndex = -1
End Sub



Private Sub PushButton3_Click()
    Me.cboPuntosVenta.ListIndex = -1
End Sub

Private Sub PushButton4_Click()
    Me.cboEstados.ListIndex = -1
End Sub



Private Sub PushButton5_Click()
    Me.cboEstadosSaldada.ListIndex = -1
End Sub

Private Sub scanear_Click()
    On Error Resume Next
    Dim archivos As New classArchivos
    If archivos.escanearDocumento(OrigenArchivos.OA_factura, Factura.id) Then
        Set m_Archivos = DAOArchivo.GetCantidadArchivosPorReferencia(OA_factura)
        Me.GridEX1.RefreshRowIndex (Factura.id)
    End If
End Sub

Private Sub txtOrdenCompra_GotFocus()
    foco Me.txtOrdenCompra
End Sub


Private Sub verFactura_Click()
    Dim f_c3h3 As New frmFacturaEdicion
    f_c3h3.ReadOnly = True
    f_c3h3.idFactura = Factura.id
    f_c3h3.Show

End Sub

Private Sub verHistorialFactura_Click()
    Set Factura.Historial = DAOFacturaHistorial.getAllByIdFactura(Factura.id)
    frmHistoriales.lista = Factura.Historial
    frmHistoriales.Show
End Sub
