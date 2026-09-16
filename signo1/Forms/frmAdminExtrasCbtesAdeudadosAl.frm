VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmAdminExtrasCbtesAdeudadosAl 
   Caption         =   "Comprobantes Compra adeudados al"
   ClientHeight    =   8940
   ClientLeft      =   60
   ClientTop       =   465
   ClientWidth     =   14415
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8572.338
   ScaleMode       =   0  'User
   ScaleWidth      =   14415
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.GroupBox GroupBox 
      Height          =   2160
      Index           =   2
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   26250
      _Version        =   786432
      _ExtentX        =   46302
      _ExtentY        =   3810
      _StockProps     =   79
      Caption         =   "Comprobantes de proveedores"
      BackColor       =   16744576
      UseVisualStyle  =   -1  'True
      Begin VB.TextBox txtComprobante 
         Height          =   315
         Left            =   1440
         TabIndex        =   1
         Top             =   1125
         Width           =   3885
      End
      Begin XtremeSuiteControls.ComboBox cboProveedores 
         Height          =   315
         Left            =   1440
         TabIndex        =   2
         Top             =   720
         Width           =   3885
         _Version        =   786432
         _ExtentX        =   6853
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Text            =   "cboProveedores"
      End
      Begin XtremeSuiteControls.PushButton btnRemoveProveedor 
         Height          =   255
         Left            =   5400
         TabIndex        =   3
         Top             =   765
         Width           =   420
         _Version        =   786432
         _ExtentX        =   741
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "X"
         BackColor       =   12632256
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.GroupBox GroupBox 
         Height          =   1695
         Index           =   4
         Left            =   6120
         TabIndex        =   4
         Top             =   240
         Width           =   4695
         _Version        =   786432
         _ExtentX        =   8281
         _ExtentY        =   2990
         _StockProps     =   79
         Caption         =   "Fecha Comprobante"
         BackColor       =   16744576
         Appearance      =   4
         Begin XtremeSuiteControls.DateTimePicker dtpDesde 
            Height          =   315
            Index           =   0
            Left            =   840
            TabIndex        =   5
            Top             =   840
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
            Index           =   0
            Left            =   2925
            TabIndex        =   6
            Top             =   840
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
            Left            =   840
            TabIndex        =   20
            Top             =   240
            Width           =   3555
            _Version        =   786432
            _ExtentX        =   6271
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Style           =   2
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.Label lblTotalSaldo 
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   21
            Top             =   300
            Width           =   480
            _Version        =   786432
            _ExtentX        =   847
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Rango"
            BackColor       =   12632256
            AutoSize        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label5 
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   8
            Top             =   900
            Width           =   465
            _Version        =   786432
            _ExtentX        =   820
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Desde"
            BackColor       =   12632256
            AutoSize        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label6 
            Height          =   195
            Index           =   0
            Left            =   2400
            TabIndex        =   7
            Top             =   900
            Width           =   420
            _Version        =   786432
            _ExtentX        =   741
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Hasta"
            BackColor       =   12632256
            AutoSize        =   -1  'True
         End
      End
      Begin XtremeSuiteControls.GroupBox gbBotones 
         Height          =   735
         Index           =   0
         Left            =   10920
         TabIndex        =   12
         Top             =   1200
         Width           =   3255
         _Version        =   786432
         _ExtentX        =   5741
         _ExtentY        =   1296
         _StockProps     =   79
         BackColor       =   16744576
         Appearance      =   4
         Begin XtremeSuiteControls.PushButton btnBuscar 
            Height          =   390
            Index           =   0
            Left            =   240
            TabIndex        =   13
            Top             =   240
            Width           =   1245
            _Version        =   786432
            _ExtentX        =   2196
            _ExtentY        =   688
            _StockProps     =   79
            Caption         =   "Buscar"
            BackColor       =   16744576
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton btnExportar 
            Height          =   390
            Index           =   0
            Left            =   1800
            TabIndex        =   14
            Top             =   240
            Width           =   1245
            _Version        =   786432
            _ExtentX        =   2196
            _ExtentY        =   688
            _StockProps     =   79
            Caption         =   "Exportar"
            BackColor       =   16744576
            UseVisualStyle  =   -1  'True
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox 
         Height          =   735
         Index           =   1
         Left            =   14280
         TabIndex        =   15
         Top             =   1200
         Width           =   4095
         _Version        =   786432
         _ExtentX        =   7223
         _ExtentY        =   1296
         _StockProps     =   79
         BackColor       =   16744576
         Appearance      =   4
         Begin XtremeSuiteControls.ProgressBar progreso 
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   3855
            _Version        =   786432
            _ExtentX        =   6800
            _ExtentY        =   661
            _StockProps     =   93
            Appearance      =   6
         End
      End
      Begin XtremeSuiteControls.PushButton btnCargarProveedores 
         Height          =   375
         Left            =   3240
         TabIndex        =   17
         Top             =   240
         Width           =   2100
         _Version        =   786432
         _ExtentX        =   3704
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Cargar Proveedores"
         BackColor       =   12632256
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.DateTimePicker dtpHastaFIN 
         Height          =   315
         Index           =   1
         Left            =   3840
         TabIndex        =   18
         Top             =   1560
         Width           =   1470
         _Version        =   786432
         _ExtentX        =   2593
         _ExtentY        =   556
         _StockProps     =   68
         CheckBox        =   -1  'True
         Format          =   1
         CurrentDate     =   45133.6457523148
      End
      Begin XtremeSuiteControls.GroupBox GroupBox 
         Height          =   855
         Index           =   0
         Left            =   10920
         TabIndex        =   22
         Top             =   240
         Width           =   7455
         _Version        =   786432
         _ExtentX        =   13150
         _ExtentY        =   1508
         _StockProps     =   79
         BackColor       =   16744576
         Appearance      =   4
         Begin XtremeSuiteControls.Label lblTotalSaldo 
            Height          =   195
            Index           =   4
            Left            =   3360
            TabIndex        =   28
            Top             =   240
            Width           =   825
            _Version        =   786432
            _ExtentX        =   1455
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Total Pago:"
            BackColor       =   12632256
            AutoSize        =   -1  'True
         End
         Begin XtremeSuiteControls.Label lblTotalSaldo 
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   27
            Top             =   480
            Width           =   960
            _Version        =   786432
            _ExtentX        =   1693
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Total Filtrado:"
            BackColor       =   12632256
            AutoSize        =   -1  'True
         End
         Begin XtremeSuiteControls.Label lblTotalSaldo 
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   855
            _Version        =   786432
            _ExtentX        =   1508
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Total Saldo:"
            BackColor       =   12632256
            AutoSize        =   -1  'True
         End
         Begin XtremeSuiteControls.Label lblTotalPagado 
            Height          =   195
            Index           =   3
            Left            =   4320
            TabIndex        =   25
            Top             =   240
            Width           =   1770
            _Version        =   786432
            _ExtentX        =   3122
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "$ 00,00"
            BackColor       =   12632256
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
         End
         Begin XtremeSuiteControls.Label lblTotalTotal 
            Height          =   195
            Index           =   2
            Left            =   1080
            TabIndex        =   24
            Top             =   480
            Width           =   1770
            _Version        =   786432
            _ExtentX        =   3122
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "$ 00,00"
            BackColor       =   12632256
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
         End
         Begin XtremeSuiteControls.Label lblTotalSaldo 
            Height          =   195
            Index           =   1
            Left            =   1080
            TabIndex        =   23
            Top             =   240
            Width           =   1770
            _Version        =   786432
            _ExtentX        =   3122
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "$ 00,00"
            BackColor       =   12632256
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
         End
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   195
         Index           =   1
         Left            =   2880
         TabIndex        =   19
         Top             =   1620
         Width           =   885
         _Version        =   786432
         _ExtentX        =   1561
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Pagos hasta"
         BackColor       =   12632256
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   195
         Left            =   600
         TabIndex        =   10
         Top             =   780
         Width           =   735
         _Version        =   786432
         _ExtentX        =   1296
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Proveedor"
         BackColor       =   12632256
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   195
         Index           =   0
         Left            =   165
         TabIndex        =   9
         Top             =   1200
         Width           =   1170
         _Version        =   786432
         _ExtentX        =   2064
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Nº Comprobante"
         BackColor       =   12632256
         AutoSize        =   -1  'True
      End
   End
   Begin GridEX20.GridEX grilla 
      Height          =   3975
      Left            =   120
      TabIndex        =   11
      Top             =   2400
      Width           =   26250
      _ExtentX        =   46302
      _ExtentY        =   7011
      Version         =   "2.0"
      PreviewRowIndent=   100
      AutomaticSort   =   -1  'True
      DefaultGroupMode=   1
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      PreviewRowLines =   1
      OLEDropMode     =   1
      ColumnAutoResize=   -1  'True
      HeaderStyle     =   2
      MethodHoldFields=   -1  'True
      ContScroll      =   -1  'True
      LockType        =   4
      GroupByBoxInfoText=   ""
      AllowEdit       =   0   'False
      BorderStyle     =   0
      BackColorGBBox  =   16744576
      BackColorHeader =   16761024
      ImageCount      =   1
      ImagePicture1   =   "frmAdminExtrasCbtesAdeudadosAl.frx":0000
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   13
      Column(1)       =   "frmAdminExtrasCbtesAdeudadosAl.frx":031A
      Column(2)       =   "frmAdminExtrasCbtesAdeudadosAl.frx":0476
      Column(3)       =   "frmAdminExtrasCbtesAdeudadosAl.frx":05A6
      Column(4)       =   "frmAdminExtrasCbtesAdeudadosAl.frx":06E6
      Column(5)       =   "frmAdminExtrasCbtesAdeudadosAl.frx":0836
      Column(6)       =   "frmAdminExtrasCbtesAdeudadosAl.frx":0976
      Column(7)       =   "frmAdminExtrasCbtesAdeudadosAl.frx":0ABE
      Column(8)       =   "frmAdminExtrasCbtesAdeudadosAl.frx":0BFE
      Column(9)       =   "frmAdminExtrasCbtesAdeudadosAl.frx":0D46
      Column(10)      =   "frmAdminExtrasCbtesAdeudadosAl.frx":0E86
      Column(11)      =   "frmAdminExtrasCbtesAdeudadosAl.frx":0FCE
      Column(12)      =   "frmAdminExtrasCbtesAdeudadosAl.frx":110E
      Column(13)      =   "frmAdminExtrasCbtesAdeudadosAl.frx":1272
      FormatStylesCount=   9
      FormatStyle(1)  =   "frmAdminExtrasCbtesAdeudadosAl.frx":1402
      FormatStyle(2)  =   "frmAdminExtrasCbtesAdeudadosAl.frx":153A
      FormatStyle(3)  =   "frmAdminExtrasCbtesAdeudadosAl.frx":15EA
      FormatStyle(4)  =   "frmAdminExtrasCbtesAdeudadosAl.frx":169E
      FormatStyle(5)  =   "frmAdminExtrasCbtesAdeudadosAl.frx":1776
      FormatStyle(6)  =   "frmAdminExtrasCbtesAdeudadosAl.frx":182E
      FormatStyle(7)  =   "frmAdminExtrasCbtesAdeudadosAl.frx":190E
      FormatStyle(8)  =   "frmAdminExtrasCbtesAdeudadosAl.frx":19CE
      FormatStyle(9)  =   "frmAdminExtrasCbtesAdeudadosAl.frx":1A92
      ImageCount      =   1
      ImagePicture(1) =   "frmAdminExtrasCbtesAdeudadosAl.frx":1B52
      PrinterProperties=   "frmAdminExtrasCbtesAdeudadosAl.frx":1E6C
   End
   Begin XtremeSuiteControls.Label lblDeudaNetaProveedor 
      Height          =   375
      Left            =   120
      TabIndex        =   32
      Top             =   7440
      Width           =   2655
      _Version        =   786432
      _ExtentX        =   4683
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Label3"
   End
   Begin XtremeSuiteControls.Label lblAnticiposProveedor 
      Height          =   255
      Left            =   120
      TabIndex        =   31
      Top             =   7200
      Width           =   2535
      _Version        =   786432
      _ExtentX        =   4471
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Label3"
   End
   Begin XtremeSuiteControls.Label lblDeudaProveedor 
      Height          =   255
      Left            =   120
      TabIndex        =   30
      Top             =   6840
      Width           =   2535
      _Version        =   786432
      _ExtentX        =   4471
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Label3"
   End
   Begin XtremeSuiteControls.Label lblAnticiposPendientes 
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   6480
      Width           =   2535
      _Version        =   786432
      _ExtentX        =   4471
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Anticipos AR$ del proveedor: -"
      BackColor       =   65535
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmAdminExtrasCbtesAdeudadosAl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vId As String
Private desde
Private Factura As clsFacturaProveedor
Private facturas As Collection
Dim m_Archivos As Dictionary
Private DatosAplicacionesPosteriores As Dictionary

Private mSaldoFacturasProveedor As Double
Private mAnticiposProveedor As Double
Private mSaldoNetoProveedor As Double

Private mIdProveedorResumen As Long
Private mResumenValido As Boolean



Private Sub MostrarAnticiposAlCorte()

    On Error GoTo err1

    Dim IdProveedor As Long
    Dim FechaCorte As Date
    Dim TotalAnticipos As Double

    'Valor inicial
    Me.lblAnticiposPendientes.caption = _
        "Anticipos AR$ del proveedor: -"

    '--------------------------------------------------
    ' DEBE HABER UN PROVEEDOR SELECCIONADO
    '--------------------------------------------------
    If Me.cboProveedores.ListIndex < 0 Then
        Exit Sub
    End If

    '--------------------------------------------------
    ' DEBE HABER FECHA DE CORTE
    '--------------------------------------------------
    If IsNull(Me.dtpHastaFIN(1).value) Then
        Exit Sub
    End If

    IdProveedor = CLng( _
        Me.cboProveedores.ItemData( _
            Me.cboProveedores.ListIndex))

    FechaCorte = CDate(Me.dtpHastaFIN(1).value)

    '--------------------------------------------------
    ' CONSULTAR ANTICIPOS PENDIENTES A ESA FECHA
    '--------------------------------------------------
    TotalAnticipos = _
        DAOFacturaProveedor.TotalAnticiposPendientesAl( _
            IdProveedor, _
            FechaCorte, _
            0)

    '--------------------------------------------------
    ' MOSTRAR RESULTADO
    '--------------------------------------------------
    Me.lblAnticiposPendientes.caption = _
        "Anticipos AR$ del proveedor sin aplicar al " & _
        Format$(FechaCorte, "dd/mm/yyyy") & ": " & _
        FormatCurrency(TotalAnticipos)
    Me.lblAnticiposPendientes.backColor = &HFFFF&
    Exit Sub

err1:

    Me.lblAnticiposPendientes.caption = _
        "No se pudieron calcular los anticipos."

    MsgBox _
        "Error al consultar los Pagos a Cuenta." & _
        vbCrLf & vbCrLf & _
        Err.Number & " - " & Err.Description, _
        vbCritical, _
        "Anticipos de proveedores"

End Sub


Private Sub btnBuscar_Click(Index As Integer)
    If IsNull(dtpHastaFIN(1).value) Then
                MsgBox ("Tiene que selecionar una fecha de fin de pagos!")
                    Else
        llenarGrilla
    End If
End Sub

Private Sub btnCargarProveedores_Click()

'''    Set colProveedores = DAOProveedor.FindAll
'''
'''    For Each prov In colProveedores
'''        cboProveedores.AddItem prov.RazonSocial
'''        cboProveedores.ItemData(cboProveedores.NewIndex) = prov.Id
'''    Next

    Call DAOProveedor.llenarComboProveedores(cboProveedores)
    
End Sub


Private Sub btnExportar_Click(Index As Integer)
    Me.progreso(0).Visible = True

    Dim FechaFin As String

    FechaFin = Me.dtpHastaFIN(1).value
    
    If IsSomething(facturas) Then
       If mResumenValido Then

        If Not DAOFacturaProveedor.ExportarColeccionTotalizadores( _
            facturas, _
            Me.progreso, _
            FechaFin, _
            mIdProveedorResumen, _
            mSaldoFacturasProveedor, _
            mAnticiposProveedor) Then
    
            GoTo err1
    
        End If
    
    Else
    
        If Not DAOFacturaProveedor.ExportarColeccionTotalizadores( _
            facturas, _
            Me.progreso, _
            FechaFin) Then
    
            GoTo err1
    
        End If
    
    End If

    End If

    Me.progreso(0).Visible = False

    Exit Sub
err1:
    MsgBox "Se produjo un error al exportar!", vbCritical, "Error"

End Sub

Private Sub btnRemoveProveedor_Click()
    Me.cboProveedores.ListIndex = -1
End Sub


Private Sub Form_Load()
    Set m_Archivos = DAOArchivo.GetCantidadArchivosPorReferencia(OA_FacturaProveedor)
    vId = funciones.CreateGUID
    
    FormHelper.Customize Me
    
    GridEXHelper.CustomizeGrid Me.grilla, True
    
    Me.grilla.ItemCount = 0
    
    dtpHastaFIN(1) = Now()
    
    btnRemoveProveedor_Click
    desde = DateSerial(Year(Date), Month(Date), 1)
    funciones.FillComboBoxDateRanges Me.cboRangos

    For i = 0 To Me.cboRangos.ListCount - 1
        If Me.cboRangos.ItemData(i) = DateRangeValue.DRV_YearCurrent Then Exit For
    Next i
    Me.cboRangos.ListIndex = i
   
    Me.grilla.Refresh

End Sub


Private Function CargarAplicacionesPosteriores() As Boolean

    On Error GoTo err1

    Dim q As String
    Dim rs As Recordset

    Dim fac As clsFacturaProveedor
    Dim IdsFacturas As String

    Dim FechaCorte As Date

    Dim claveFactura As String
    Dim claveOP As String
    Dim ultimaOP As String

    CargarAplicacionesPosteriores = False

    Set DatosAplicacionesPosteriores = New Dictionary

    If facturas Is Nothing Then
        CargarAplicacionesPosteriores = True
        Exit Function
    End If

    If facturas.count = 0 Then
        CargarAplicacionesPosteriores = True
        Exit Function
    End If

    If IsNull(Me.dtpHastaFIN(1).value) Then
        Exit Function
    End If

    FechaCorte = CDate(Me.dtpHastaFIN(1).value)

    '==================================================
    ' ARMAR LISTA DE FACTURAS QUE ESTAN EN LA GRILLA
    '==================================================

    IdsFacturas = ""

    For Each fac In facturas

        If LenB(IdsFacturas) > 0 Then
            IdsFacturas = IdsFacturas & ","
        End If

        IdsFacturas = IdsFacturas & CStr(fac.Id)

    Next fac

    If LenB(IdsFacturas) = 0 Then
        CargarAplicacionesPosteriores = True
        Exit Function
    End If

    '==================================================
    ' CONSULTAR OP APROBADAS POSTERIORES AL CORTE
    ' Y LOS PAGOS A CUENTA ASOCIADOS
    '==================================================

    q = "SELECT DISTINCT " _
      & "opf.id_factura_proveedor AS id_factura, " _
      & "op.id AS id_op, " _
      & "op.fecha AS fecha_op, " _
      & "p.id AS id_pcta, " _
      & "p.fecha AS fecha_pcta " _
      & "FROM ordenes_pago_facturas opf " _
      & "INNER JOIN ordenes_pago op " _
      & " ON op.id = opf.id_orden_pago " _
      & "LEFT JOIN ordenes_pago_pagos_a_cuenta vinc " _
      & " ON vinc.id_orden_pago = op.id " _
      & "LEFT JOIN pagos_a_cuenta p " _
      & " ON p.id = vinc.id_pago_a_cuenta " _
      & "WHERE opf.id_factura_proveedor IN (" _
      & IdsFacturas & ") " _
      & "AND op.estado = 1 " _
      & "AND op.fecha > " _
      & conectar.Escape(FechaCorte) & " " _
      & "ORDER BY " _
      & "opf.id_factura_proveedor, " _
      & "op.fecha, op.id, p.id"

    Set rs = conectar.RSFactory(q)

    ultimaOP = ""

    '==================================================
    ' GUARDAR EL TEXTO CORRESPONDIENTE A CADA FACTURA
    '==================================================

    While Not rs.EOF

        claveFactura = CStr(rs!id_factura)

        claveOP = claveFactura & ":" & CStr(rs!id_op)

        If Not DatosAplicacionesPosteriores.Exists( _
                    claveFactura) Then

            DatosAplicacionesPosteriores.Add _
                claveFactura, ""

        End If

        'Agregar cada OP una sola vez.
        If claveOP <> ultimaOP Then

            If LenB(DatosAplicacionesPosteriores( _
                    claveFactura)) > 0 Then

                DatosAplicacionesPosteriores( _
                    claveFactura) = _
                    DatosAplicacionesPosteriores( _
                    claveFactura) & " ; "

            End If

            DatosAplicacionesPosteriores( _
                claveFactura) = _
                DatosAplicacionesPosteriores( _
                claveFactura) & _
                "OP " & CStr(rs!id_op) & _
                " (" & Format$(rs!fecha_op, _
                              "dd/mm/yyyy") & ")"

            ultimaOP = claveOP

        End If

        'Agregar PCTA asociado a esa OP.
        If Not IsNull(rs!id_pcta) Then

            DatosAplicacionesPosteriores( _
                claveFactura) = _
                DatosAplicacionesPosteriores( _
                claveFactura) & _
                " | PCTA " & CStr(rs!id_pcta) & _
                " (" & Format$(rs!fecha_pcta, _
                              "dd/mm/yyyy") & ")"

        End If

        rs.MoveNext

    Wend

    Set rs = Nothing

    CargarAplicacionesPosteriores = True

    Exit Function

err1:

    CargarAplicacionesPosteriores = False

    MsgBox _
        "Error al consultar las aplicaciones posteriores." & _
        vbCrLf & vbCrLf & _
        Err.Number & " - " & Err.Description, _
        vbCritical, _
        "Comprobantes adeudados"

End Function


Private Function CalcularResumenProveedor() As Boolean

    On Error GoTo err1

    Dim IdProveedor As Long
    Dim FechaCorte As Date

    Dim filtro As String
    Dim fechaSQL As String

    Dim facturasProveedor As Collection
    Dim fac As clsFacturaProveedor

    Dim signo As Integer
    Dim importeFactura As Double

    CalcularResumenProveedor = False

    mResumenValido = False
    mIdProveedorResumen = 0

    mSaldoFacturasProveedor = 0
    mAnticiposProveedor = 0
    mSaldoNetoProveedor = 0

    Me.lblDeudaProveedor.caption = _
        "Deuda del proveedor AR$: -"

    Me.lblAnticiposProveedor.caption = _
        "Anticipos sin aplicar AR$: -"

    Me.lblDeudaNetaProveedor.caption = _
        "Deuda neta AR$: -"

    'Debe haber un proveedor seleccionado.
    If Me.cboProveedores.ListIndex < 0 Then
        CalcularResumenProveedor = True
        Exit Function
    End If

    If IsNull(Me.dtpHastaFIN(1).value) Then
        Exit Function
    End If

    'No presentar una deuda global calculada
    'sobre facturas que el usuario no puede consultar.
    If Permisos.AdminFaPVerSoloPropias Then

        Me.lblDeudaNetaProveedor.caption = _
            "Deuda neta: no disponible con permisos parciales"

        CalcularResumenProveedor = True
        Exit Function

    End If

    IdProveedor = CLng( _
        Me.cboProveedores.ItemData( _
            Me.cboProveedores.ListIndex))

    FechaCorte = DateValue( _
        Me.dtpHastaFIN(1).value)

    fechaSQL = conectar.Escape(FechaCorte)

    '--------------------------------------------------
    ' TODAS LAS FACTURAS DEL PROVEEDOR EN PESOS
    ' HASTA LA FECHA DE CORTE.
    '
    ' No aplicamos los filtros de la grilla.
    '--------------------------------------------------

    filtro = _
        "AdminComprasFacturasProveedores.id_proveedor = " & _
        IdProveedor

    filtro = filtro & _
        " AND AdminComprasFacturasProveedores.id_moneda = 0"

    filtro = filtro & _
        " AND AdminComprasFacturasProveedores.fecha <= " & _
        fechaSQL

    Set facturasProveedor = _
        DAOFacturaProveedor.FindAllTotalizadores( _
            filtro, fechaSQL)

    If facturasProveedor Is Nothing Then
        Err.Raise 5, , _
            "No se pudieron obtener las facturas del proveedor."
    End If

    '--------------------------------------------------
    ' CALCULAR DEUDA POR FACTURAS
    '--------------------------------------------------

    For Each fac In facturasProveedor

        signo = 1

        If fac.tipoDocumentoContable = _
                tipoDocumentoContable.notaCredito Then

            signo = -1

        End If

        importeFactura = _
            fac.Monto + _
            fac.TotalIVA + _
            fac.totalPercepciones + _
            fac.ImpuestoInterno + _
            fac.redondeo

        mSaldoFacturasProveedor = _
            mSaldoFacturasProveedor + _
            ((importeFactura - fac.TotalAbonadoGlobal) * signo)

    Next fac

    '--------------------------------------------------
    ' ANTICIPOS EXISTENTES SIN APLICAR AL CORTE
    '--------------------------------------------------

    mAnticiposProveedor = _
        DAOFacturaProveedor.TotalAnticiposPendientesAl( _
            IdProveedor, FechaCorte, 0)

    '--------------------------------------------------
    ' SALDO NETO
    '--------------------------------------------------

    mSaldoNetoProveedor = _
        funciones.FormatearDecimales( _
            mSaldoFacturasProveedor - mAnticiposProveedor)

    mIdProveedorResumen = IdProveedor
    mResumenValido = True

    '--------------------------------------------------
    ' MOSTRAR RESULTADOS
    '--------------------------------------------------

    Me.lblDeudaProveedor.caption = _
        "Deuda del proveedor AR$: " & _
        FormatCurrency(mSaldoFacturasProveedor)

    Me.lblAnticiposProveedor.caption = _
        "Anticipos sin aplicar AR$: -" & _
        FormatCurrency(mAnticiposProveedor)

    Me.lblDeudaNetaProveedor.caption = _
        "Deuda neta AR$ al " & _
        Format$(FechaCorte, "dd/mm/yyyy") & ": " & _
        FormatCurrency(mSaldoNetoProveedor)

    CalcularResumenProveedor = True

    Exit Function

err1:

    mResumenValido = False

    Me.lblDeudaNetaProveedor.caption = _
        "No se pudo calcular la deuda neta."

    MsgBox _
        "Error al calcular el resumen del proveedor." & _
        vbCrLf & vbCrLf & _
        Err.Number & " - " & Err.Description, _
        vbCritical, _
        "Comprobantes adeudados"

End Function



Public Sub llenarGrilla()
    grilla.ItemCount = 0
    Dim condition As String
    condition = " 1 = 1 "
    Dim FechaFin As String
    
    If Not IsNull(Me.dtpDesde(0).value) Then
        condition = condition & " AND AdminComprasFacturasProveedores.fecha >= " & conectar.Escape(Me.dtpDesde(0).value)
    End If

    If Not IsNull(Me.dtpHasta(0).value) Then
        condition = condition & " AND AdminComprasFacturasProveedores.fecha <= " & conectar.Escape(Me.dtpHasta(0).value)
    End If

    If cboProveedores.ListIndex > -1 Then
        condition = condition & " AND AdminComprasFacturasProveedores.id_proveedor = " & cboProveedores.ItemData(Me.cboProveedores.ListIndex)
    End If

    If LenB(Me.txtComprobante) > 0 Then
        condition = condition & " AND AdminComprasFacturasProveedores.numero_factura like '%" & Trim(Me.txtComprobante.Text) & "%'"
    End If
    
    If Not IsNull(dtpHastaFIN(1).value) Then
    
        FechaFin = conectar.Escape( _
            dtpHastaFIN(1).value)
    
        'No incluir facturas posteriores al corte.
        condition = condition & _
            " AND AdminComprasFacturasProveedores.fecha <= " & _
            FechaFin
    
    End If
    
    Set facturas = DAOFacturaProveedor.FindAllTotalizadores(condition, FechaFin, , , Permisos.AdminFaPVerSoloPropias)
    
    '==================================================
    ' CARGAR OP POSTERIORES Y ANTICIPOS ASOCIADOS
    '==================================================
    
    If Not CargarAplicacionesPosteriores() Then
    
        grilla.ItemCount = 0
        Exit Sub
    
    End If

    ''''''''''''''''
    
    Dim total As Double
    Dim pagado As Double
    Dim saldo As Double
    Dim TotalFactura As Double
    Dim TotalPagado As Double
    Dim c As Integer

    total = 0

    For Each Factura In facturas

        If Factura.tipoDocumentoContable = tipoDocumentoContable.notaCredito Then c = -1 Else c = 1
        
        TotalFactura = ((Factura.Monto - Factura.TotalNetoGravadoDiscriminado(0)) + Factura.TotalIVA + Factura.TotalNetoGravadoDiscriminado(0) + Factura.totalPercepciones + Factura.ImpuestoInterno + Factura.redondeo) * c
        total = total + TotalFactura
              
        TotalPagado = (Factura.TotalAbonadoGlobal) * c
        pagado = pagado + TotalPagado
        
        
        TotalSaldado = TotalFactura - TotalPagado
        saldo = saldo + TotalSaldado
    Next

    Me.lblTotalTotal(2).caption = FormatCurrency(funciones.FormatearDecimales(total))
    Me.lblTotalSaldo(1).caption = FormatCurrency(funciones.FormatearDecimales(saldo))
    Me.lblTotalPagado(3).caption = FormatCurrency(funciones.FormatearDecimales(pagado))
    
        If Not CalcularResumenProveedor() Then
        Exit Sub
        
    End If

    '==================================================
    ' MOSTRAR ANTICIPOS PENDIENTES A LA FECHA DE CORTE
    '==================================================
    
    MostrarAnticiposAlCorte
    
    '''''''''''''''
    

    grilla.ItemCount = facturas.count

    GridEXHelper.AutoSizeColumns Me.grilla, True

    Me.caption = "Cbtes. filtrados [Cantidad: " & facturas.count & "]"

End Sub


Private Sub cboRangos_Click()
    funciones.CalculateDateRange Me.cboRangos, Me.dtpDesde(0), Me.dtpHasta(0)
End Sub


Private Sub Form_Resize()

    On Error Resume Next

    'Ajustar ancho de la grilla
    Me.grilla.Width = Me.ScaleWidth - 200

    'Reservar espacio debajo de la grilla
    Me.grilla.Height = _
        Me.ScaleHeight - Me.grilla.Top - 650

    'Ajustar ancho del grupo superior
    Me.GroupBox(2).Width = Me.grilla.Width

    'Ubicar el Label debajo de la grilla
    Me.lblAnticiposPendientes.Left = Me.grilla.Left

    Me.lblAnticiposPendientes.Top = _
        Me.grilla.Top + Me.grilla.Height + 120

    Me.lblAnticiposPendientes.Width = _
        Me.grilla.Width

End Sub


Private Sub grilla_ColumnHeaderClick(ByVal Column As GridEX20.JSColumn)
    GridEXHelper.ColumnHeaderClick Me.grilla, Column
End Sub


Private Sub grilla_DblClick()
    verDetalle_Click
End Sub


Private Sub grilla_RowFormat(RowBuffer As GridEX20.JSRowData)
    On Error GoTo err1
    Set Factura = facturas(RowBuffer.RowIndex)

    If Factura.estado = EstadoFacturaProveedor.Aprobada Then
        RowBuffer.CellStyle(15) = "EstadoAprobado"
    ElseIf Factura.estado = EstadoFacturaProveedor.EnProceso Then
        RowBuffer.CellStyle(15) = " EstadoEnProceso"
    ElseIf Factura.estado = EstadoFacturaProveedor.Saldada Then
        RowBuffer.CellStyle(15) = "EstadoSaldado"
    End If
    Exit Sub
err1:
End Sub


Private Sub grilla_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)

    Set Factura = facturas.item(RowIndex)

    Dim i As Integer

    If Factura.tipoDocumentoContable = tipoDocumentoContable.notaCredito Then i = -1 Else i = 1

    With Factura

        Values(1) = Factura.Id
        
        If IsSomething(Factura.Proveedor) Then
            Values(2) = UCase(funciones.RazonSocialFormateada(Factura.Proveedor.RazonSocial))
            Values(3) = Factura.Proveedor.cuit
        End If

        Values(4) = enums.EnumTipoDocumentoContableShort(Factura.tipoDocumentoContable)
        Values(5) = Factura.configFactura.TipoFactura
        Values(6) = Factura.numero
        Values(7) = Factura.FEcha
        Values(8) = Factura.moneda.NombreCorto

        TotalFactura = (Factura.Monto - Factura.TotalNetoGravadoDiscriminado(0)) + Factura.TotalIVA + Factura.TotalNetoGravadoDiscriminado(0) + Factura.totalPercepciones + Factura.ImpuestoInterno
 
        Values(9) = Replace(FormatCurrency(funciones.FormatearDecimales(TotalFactura + Factura.redondeo) * i), "$", "")
        Values(10) = Replace(FormatCurrency(funciones.FormatearDecimales(Factura.TotalAbonadoGlobal) * i), "$", "")
        Values(11) = Replace(FormatCurrency(funciones.FormatearDecimales((TotalFactura + Factura.redondeo) - Factura.TotalAbonadoGlobal) * i), "$", "")

        '==================================================
        ' COLUMNA 12 - SITUACION AL CORTE
        '==================================================
        
        Dim SaldoAlCorte As Double
        
        SaldoAlCorte = funciones.FormatearDecimales( _
            Factura.total - Factura.TotalAbonadoGlobal)
        
        If Factura.tipoDocumentoContable = _
                tipoDocumentoContable.notaCredito Then
        
            Values(12) = "Nota de crédito - ver saldo"
        
        ElseIf Abs(SaldoAlCorte) < 0.01 Then
        
            Values(12) = "Saldada al corte"
        
        ElseIf Factura.TotalAbonadoGlobal > 0 Then
        
            Values(12) = "Pago parcial al corte"
        
        Else
        
            Values(12) = "Pendiente al corte"
        
        End If
        
        
        '==================================================
        ' COLUMNA 13 - APLICACION POSTERIOR
        '==================================================
        
        Values(13) = ""
        
        If Not DatosAplicacionesPosteriores Is Nothing Then
        
            If DatosAplicacionesPosteriores.Exists( _
                    CStr(Factura.Id)) Then
        
                Values(13) = _
                    DatosAplicacionesPosteriores( _
                        CStr(Factura.Id))
        
            End If
        
        End If

    End With

End Sub



Private Sub txtComprobante_GotFocus()
    foco Me.txtComprobante
End Sub


Private Sub verDetalle_Click()
    SeleccionarFactura
    Dim frm As frmAdminComprasNuevaFCProveedor
    Set frm = New frmAdminComprasNuevaFCProveedor

    frm.ver = True
    frm.Factura = Factura
    frm.Show
    
End Sub


Private Sub SeleccionarFactura()
    On Error Resume Next
    Set Factura = facturas.item(grilla.RowIndex(grilla.row))
    
End Sub


