VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmResumenBancario 
   Caption         =   "Reporte Bancario"
   ClientHeight    =   10290
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   18735
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10290
   ScaleWidth      =   18735
   WindowState     =   2  'Maximized
   Begin GridEX20.GridEX gridResumenBancario 
      Height          =   6615
      Left            =   120
      TabIndex        =   1
      Top             =   2880
      Width           =   20055
      _ExtentX        =   35375
      _ExtentY        =   11668
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      ColumnAutoResize=   -1  'True
      MethodHoldFields=   -1  'True
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   10
      Column(1)       =   "frmResumenBancario.frx":0000
      Column(2)       =   "frmResumenBancario.frx":0184
      Column(3)       =   "frmResumenBancario.frx":029C
      Column(4)       =   "frmResumenBancario.frx":03DC
      Column(5)       =   "frmResumenBancario.frx":0514
      Column(6)       =   "frmResumenBancario.frx":0634
      Column(7)       =   "frmResumenBancario.frx":077C
      Column(8)       =   "frmResumenBancario.frx":08D4
      Column(9)       =   "frmResumenBancario.frx":0A1C
      Column(10)      =   "frmResumenBancario.frx":0B9C
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmResumenBancario.frx":0D0C
      FormatStyle(2)  =   "frmResumenBancario.frx":0E44
      FormatStyle(3)  =   "frmResumenBancario.frx":0EF4
      FormatStyle(4)  =   "frmResumenBancario.frx":0FA8
      FormatStyle(5)  =   "frmResumenBancario.frx":1080
      FormatStyle(6)  =   "frmResumenBancario.frx":1138
      ImageCount      =   0
      PrinterProperties=   "frmResumenBancario.frx":1218
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   20055
      _Version        =   786432
      _ExtentX        =   35375
      _ExtentY        =   4260
      _StockProps     =   79
      Caption         =   "Fiiltro de Búsqueda"
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
      Begin XtremeSuiteControls.GroupBox GroupBox2 
         Height          =   2055
         Left            =   6000
         TabIndex        =   24
         Top             =   240
         Width           =   3615
         _Version        =   786432
         _ExtentX        =   6376
         _ExtentY        =   3625
         _StockProps     =   79
         Caption         =   "Monto inicial"
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.PushButton cmdLimpiarMontoInicial 
            Height          =   375
            Left            =   3000
            TabIndex        =   27
            Top             =   930
            Width           =   375
            _Version        =   786432
            _ExtentX        =   661
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "X"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton cmdEstablecerMontoInicial 
            Height          =   495
            Left            =   240
            TabIndex        =   26
            Top             =   1440
            Width           =   3135
            _Version        =   786432
            _ExtentX        =   5530
            _ExtentY        =   873
            _StockProps     =   79
            Caption         =   "Establecer"
            UseVisualStyle  =   -1  'True
         End
         Begin VB.TextBox txtMontoInicial 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   240
            TabIndex        =   25
            Text            =   "0"
            Top             =   960
            Width           =   2655
         End
      End
      Begin XtremeSuiteControls.PushButton cmdReestablecerTipoMov 
         Height          =   375
         Left            =   5520
         TabIndex        =   23
         Top             =   1890
         Width           =   375
         _Version        =   786432
         _ExtentX        =   661
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "X"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton cmdReestablecerMoneda 
         Height          =   375
         Left            =   3000
         TabIndex        =   22
         Top             =   1440
         Width           =   375
         _Version        =   786432
         _ExtentX        =   661
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "X"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton cmdReestablecerOrigen 
         Height          =   375
         Left            =   5520
         TabIndex        =   21
         Top             =   930
         Width           =   375
         _Version        =   786432
         _ExtentX        =   661
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "X"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton cmdReestablecerCtaBancaria 
         Height          =   375
         Left            =   5520
         TabIndex        =   20
         Top             =   450
         Width           =   375
         _Version        =   786432
         _ExtentX        =   661
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "X"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton cmdExportar 
         Height          =   495
         Left            =   17760
         TabIndex        =   19
         Top             =   1800
         Width           =   2175
         _Version        =   786432
         _ExtentX        =   3836
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Exportar"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   495
         Left            =   17760
         TabIndex        =   14
         Top             =   1200
         Width           =   2175
         _Version        =   786432
         _ExtentX        =   3836
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Limpiar"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cboOrigen 
         Height          =   315
         Left            =   1200
         TabIndex        =   13
         Top             =   960
         Width           =   4215
         _Version        =   786432
         _ExtentX        =   7435
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cboTipoMovimiento 
         Height          =   315
         Left            =   1200
         TabIndex        =   12
         Top             =   1920
         Width           =   4215
         _Version        =   786432
         _ExtentX        =   7435
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cboMonedas 
         Height          =   315
         Left            =   1200
         TabIndex        =   11
         Top             =   1470
         Width           =   1695
         _Version        =   786432
         _ExtentX        =   2990
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cboCuentasBancarias 
         Height          =   315
         Left            =   1200
         TabIndex        =   10
         Top             =   480
         Width           =   4215
         _Version        =   786432
         _ExtentX        =   7435
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.PushButton cmdProbarResumen 
         Height          =   495
         Left            =   14640
         TabIndex        =   2
         Top             =   1800
         Width           =   2175
         _Version        =   786432
         _ExtentX        =   3836
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Reportar"
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
      End
      Begin XtremeSuiteControls.GroupBox GroFechaComprobante 
         Height          =   2055
         Index           =   1
         Left            =   9720
         TabIndex        =   3
         Top             =   240
         Width           =   4695
         _Version        =   786432
         _ExtentX        =   8281
         _ExtentY        =   3625
         _StockProps     =   79
         Caption         =   "Fecha Movimiento"
         BackColor       =   16744576
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   4
         Begin XtremeSuiteControls.DateTimePicker dtpDesde 
            Height          =   315
            Index           =   1
            Left            =   720
            TabIndex        =   4
            Top             =   720
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
            Index           =   1
            Left            =   2925
            TabIndex        =   5
            Top             =   720
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
            Left            =   720
            TabIndex        =   6
            Top             =   300
            Width           =   3675
            _Version        =   786432
            _ExtentX        =   6482
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Style           =   2
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.Label lblHasta 
            Height          =   195
            Index           =   1
            Left            =   2400
            TabIndex        =   9
            Top             =   780
            Width           =   420
            _Version        =   786432
            _ExtentX        =   741
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Hasta"
            BackColor       =   12632256
            AutoSize        =   -1  'True
         End
         Begin XtremeSuiteControls.Label lblDesde 
            Height          =   195
            Index           =   1
            Left            =   165
            TabIndex        =   8
            Top             =   780
            Width           =   465
            _Version        =   786432
            _ExtentX        =   820
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Desde"
            BackColor       =   12632256
            AutoSize        =   -1  'True
         End
         Begin XtremeSuiteControls.Label lblRango 
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   7
            Top             =   360
            Width           =   480
            _Version        =   786432
            _ExtentX        =   847
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Rango"
            BackColor       =   12632256
            AutoSize        =   -1  'True
         End
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   18
         Top             =   1950
         Width           =   975
         _Version        =   786432
         _ExtentX        =   1720
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Tipo de Mov."
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   17
         Top             =   1500
         Width           =   975
         _Version        =   786432
         _ExtentX        =   1720
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Moneda"
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   990
         Width           =   975
         _Version        =   786432
         _ExtentX        =   1720
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Origen"
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   450
         Width           =   975
         _Version        =   786432
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Cta. Bcaria."
         Alignment       =   1
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   2640
      Width           =   6015
   End
   Begin VB.Menu mnuContextualResumen 
      Caption         =   "Opciones"
      Begin VB.Menu mnuAbrirMovimientoCajaBancos 
         Caption         =   "Abrir movimiento de Caja y Bancos"
      End
   End
End
Attribute VB_Name = "frmResumenBancario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Movimientos As Collection
Private MovimientosBase As Collection

Private MontoInicial As Double
Private MontoInicialEstablecido As Boolean

Private desde As Date
Private CargandoFiltros As Boolean
Private i As Integer


Private Sub cmdEstablecerMontoInicial_Click()

    On Error GoTo err1

    Dim IdCuentaBancaria As Long
    Dim IdMoneda As Long
    Dim textoMonto As String

    IdCuentaBancaria = ObtenerIdCombo(Me.cboCuentasBancarias)
    IdMoneda = ObtenerIdCombo(Me.cboMonedas)

    If IdCuentaBancaria = 0 Then

        MsgBox "Debe seleccionar una cuenta bancaria específica " & _
               "para establecer el monto inicial.", _
               vbExclamation, _
               "Monto inicial"

        Exit Sub

    End If

'''    If IdMoneda = 0 Then
'''
'''        MsgBox "Debe seleccionar una moneda específica " & _
'''               "para establecer el monto inicial.", _
'''               vbExclamation, _
'''               "Monto inicial"
'''
'''        Exit Sub
'''
'''    End If

    textoMonto = Trim$(Me.txtMontoInicial.Text)

    If LenB(textoMonto) = 0 Then

        MsgBox "Debe ingresar un monto inicial.", _
               vbExclamation, _
               "Monto inicial"

        Me.txtMontoInicial.SetFocus
        Exit Sub

    End If

    If Not IsNumeric(textoMonto) Then

        MsgBox "El monto inicial ingresado no es válido.", _
               vbExclamation, _
               "Monto inicial"

        Me.txtMontoInicial.SetFocus
        Exit Sub

    End If

    ' También se permiten saldos negativos.
    MontoInicial = CDbl(textoMonto)
    MontoInicialEstablecido = True

    If MovimientosBase Is Nothing Then
        Set MovimientosBase = New Collection
    End If

    ReconstruirMovimientosMostrados
    ActualizarGridResumen

    MsgBox "Monto inicial establecido correctamente.", _
           vbInformation, _
           "Monto inicial"

    Exit Sub

err1:
    MsgBox "No se pudo establecer el monto inicial:" & vbCrLf & _
           Err.Description, _
           vbCritical, _
           "Monto inicial"

End Sub

Private Sub cmdExportar_Click()

    If Movimientos Is Nothing Then
        MsgBox "Primero debe generar el reporte.", _
               vbExclamation, _
               "Reporte bancario"
        Exit Sub
    End If

    If Movimientos.count = 0 Then
        MsgBox "No hay movimientos para exportar.", _
               vbInformation, _
               "Reporte bancario"
        Exit Sub
    End If

    Me.MousePointer = vbHourglass

    If ExportarResumenBancario(Movimientos) Then
        MsgBox "El reporte fue exportado correctamente.", _
               vbInformation, _
               "Reporte bancario"
    Else
        MsgBox "No se pudo exportar el reporte.", _
               vbCritical, _
               "Reporte bancario"
    End If

    Me.MousePointer = vbDefault

End Sub




Private Function textoCombo(ByVal cbo As Object) As String

    If cbo.ListIndex >= 0 Then
        textoCombo = cbo.list(cbo.ListIndex)
    Else
        textoCombo = "TODOS"
    End If

End Function


Private Sub cmdLimpiarMontoInicial_Click()
    MontoInicial = 0
    MontoInicialEstablecido = False

    Me.txtMontoInicial.Text = "0"

    ReconstruirMovimientosMostrados
    ActualizarGridResumen
    
End Sub

Private Sub cmdProbarResumen_Click()

    On Error GoTo err1

    Dim IdBanco As Long
    Dim IdCuentaBancaria As Long
    Dim IdMoneda As Long

    Dim TipoMovimiento As String
    Dim Origen As String

    If Me.dtpDesde(1).value > Me.dtpHasta(1).value Then
        MsgBox "La fecha desde no puede ser mayor que la fecha hasta.", _
               vbExclamation
        Exit Sub
    End If

    IdCuentaBancaria = ObtenerIdCombo(Me.cboCuentasBancarias)
    IdMoneda = ObtenerIdCombo(Me.cboMonedas)

    TipoMovimiento = ObtenerTipoMovimiento
    Origen = ObtenerOrigen

    Me.MousePointer = vbHourglass

    Set MovimientosBase = DAOResumenBancario.FindAll( _
        Me.dtpDesde(1).value, _
        Me.dtpHasta(1).value, _
        IdCuentaBancaria, _
        IdMoneda, _
        TipoMovimiento, _
        Origen)

    If MovimientosBase Is Nothing Then
    
        Set Movimientos = New Collection
    
        Me.gridResumenBancario.ItemCount = 0
        Me.gridResumenBancario.Refresh
    
        MsgBox "Ocurrió un error al generar el resumen bancario.", _
               vbCritical
    
        GoTo salir
    
    End If
    
    ReconstruirMovimientosMostrados
    ActualizarGridResumen

salir:
    Me.MousePointer = vbDefault
    Exit Sub

err1:
    Me.MousePointer = vbDefault

    MsgBox "Error al cargar el resumen bancario:" & vbCrLf & _
           Err.Description, _
           vbCritical

End Sub

Private Sub cmdReestablecerCtaBancaria_Click()
    If Not SeleccionarItemDataCombo( _
        Me.cboCuentasBancarias, 0) Then

        MsgBox "No se encontró la opción TODAS.", _
               vbExclamation, _
               "Reporte bancario"

        Exit Sub
    End If

    Me.cboCuentasBancarias.SetFocus
End Sub

Private Sub cmdReestablecerMoneda_Click()
    If Not SeleccionarItemDataCombo( _
        Me.cboMonedas, 0) Then

        MsgBox "No se encontró la opción TODAS.", _
               vbExclamation, _
               "Reporte bancario"

        Exit Sub
    End If

    Me.cboMonedas.SetFocus
End Sub

Private Sub cmdReestablecerOrigen_Click()
    If Not SeleccionarItemDataCombo( _
        Me.cboOrigen, 0) Then

        MsgBox "No se encontró la opción TODOS.", _
               vbExclamation, _
               "Reporte bancario"

        Exit Sub
    End If

    Me.cboOrigen.SetFocus
End Sub

Private Sub cmdReestablecerTipoMov_Click()
    If Not SeleccionarItemDataCombo( _
        Me.cboTipoMovimiento, 0) Then

        MsgBox "No se encontró la opción TODOS.", _
               vbExclamation, _
               "Reporte bancario"

        Exit Sub
    End If

    Me.cboTipoMovimiento.SetFocus

End Sub

Private Sub Form_Load()

    FormHelper.Customize Me
    
    Set Movimientos = New Collection
    Set MovimientosBase = New Collection
    
    MontoInicial = 0
    MontoInicialEstablecido = False
    
    Me.txtMontoInicial.Text = "0"


    GridEXHelper.CustomizeGrid _
        Me.gridResumenBancario, _
        False, _
        False
    
    ConfigurarFormatoGrilla
    
    Me.dtpDesde(1).value = Year(Now) & "-01-01"

    desde = DateSerial(Year(Date), Month(Date), 1)   ' CDate(1 & "-" & Month(Now) & "-" & Year(Now))
    funciones.FillComboBoxDateRanges Me.cboRangos

    Me.cboRangos.ListIndex = i
    
    For i = 0 To Me.cboRangos.ListCount - 1
        If Me.cboRangos.ItemData(i) = DateRangeValue.DRV_YearCurrent Then Exit For
    Next i
    Me.cboRangos.ListIndex = i
    
    ConfigurarFormatoImportes

    Me.gridResumenBancario.ItemCount = 0
    
    Me.Label3.caption = "Registros mostrados: 0"
    

    CargarFiltros


End Sub


Private Sub gridResumenBancario_MouseDown( _
    Button As Integer, _
    Shift As Integer, _
    x As Single, _
    y As Single)

    Dim Movimiento As DTOResumenBancario
    Dim origenMovimiento As String

    If Button <> 2 Then Exit Sub

    Set Movimiento = ObtenerMovimientoResumenSeleccionado()

    If Movimiento Is Nothing Then Exit Sub

    origenMovimiento = _
        UCase$(Trim$(Movimiento.Origen))

    '------------------------------------------------------
    ' SOLAMENTE LOS REGISTROS QUE PROVIENEN
    ' DE MOVIMIENTOS DE CAJA Y BANCOS
    '------------------------------------------------------
    Select Case origenMovimiento

        Case "MOVIMIENTO CAJA/BANCOS", _
             "TRANSFERENCIA INTERBANCARIA"

            Me.PopupMenu Me.mnuContextualResumen

    End Select

End Sub

Private Sub gridResumenBancario_UnboundReadData( _
    ByVal RowIndex As Long, _
    ByVal Bookmark As Variant, _
    ByVal Values As GridEX20.JSRowData)

    Dim Movimiento As DTOResumenBancario

    If Movimientos Is Nothing Then Exit Sub

    ' En este GridEX, RowIndex comienza en 1,
    ' igual que una Collection de VB6.
    If RowIndex < 1 Then Exit Sub
    If RowIndex > Movimientos.count Then Exit Sub

    Set Movimiento = Movimientos.item(RowIndex)

    Values(1) = Movimiento.FEcha
    Values(2) = Movimiento.Banco
    Values(3) = Movimiento.CuentaBancaria
    Values(4) = Movimiento.CuentaOrigen
    Values(5) = Movimiento.Origen
    Values(6) = Movimiento.NumeroOrigen
    Values(7) = Movimiento.Comprobante

    ' Mantener como valores numéricos.
    Values(8) = Replace(FormatCurrency(funciones.FormatearDecimales(Movimiento.Ingreso)), "$", "")
    Values(9) = Replace(FormatCurrency(funciones.FormatearDecimales(Movimiento.Egreso)), "$", "")
    Values(10) = Replace(FormatCurrency(funciones.FormatearDecimales(Movimiento.SaldoAcumulado)), "$", "")


End Sub


Private Sub CargarFiltros()

    CargandoFiltros = True

    Me.dtpDesde(1).value = _
        DateSerial(Year(Date), Month(Date), 1)

    Me.dtpHasta(1).value = Date

    CargarComboCuentasBancarias 0
    CargarComboMonedas
    CargarComboTipoMovimiento
    CargarComboOrigen

    ' Seleccionar TODAS/TODOS mediante ItemData = 0.
    Call SeleccionarItemDataCombo( _
        Me.cboCuentasBancarias, 0)

    Call SeleccionarItemDataCombo( _
        Me.cboMonedas, 0)

    Call SeleccionarItemDataCombo( _
        Me.cboTipoMovimiento, 0)

    Call SeleccionarItemDataCombo( _
        Me.cboOrigen, 0)

    CargandoFiltros = False

End Sub



Private Sub CargarComboCuentasBancarias( _
    ByVal IdBanco As Long)

    Dim col As Collection
    Dim c As CuentaBancaria
    Dim filtro As String

    Me.cboCuentasBancarias.Clear

    Me.cboCuentasBancarias.AddItem "TODAS"
    Me.cboCuentasBancarias.ItemData( _
        Me.cboCuentasBancarias.NewIndex) = 0

    filtro = "1 = 1"

    If IdBanco > 0 Then
        filtro = "c.idBanco = " & IdBanco
    End If

    Set col = DAOCuentaBancaria.FindAll(filtro)

    For Each c In col
        Me.cboCuentasBancarias.AddItem _
            c.DescripcionFormateada

        Me.cboCuentasBancarias.ItemData( _
            Me.cboCuentasBancarias.NewIndex) = c.Id
    Next c

End Sub


Private Sub CargarComboMonedas()

    Dim col As Collection
    Dim m As clsMoneda

    Me.cboMonedas.Clear

    Me.cboMonedas.AddItem "TODAS"
    Me.cboMonedas.ItemData(Me.cboMonedas.NewIndex) = 0

    Set col = DAOMoneda.GetAll()

    For Each m In col
        Me.cboMonedas.AddItem m.NombreCorto
        Me.cboMonedas.ItemData(Me.cboMonedas.NewIndex) = m.Id
    Next m

End Sub


Private Sub CargarComboTipoMovimiento()

    Me.cboTipoMovimiento.Clear

    Me.cboTipoMovimiento.AddItem "TODOS"
    Me.cboTipoMovimiento.ItemData( _
        Me.cboTipoMovimiento.NewIndex) = 0

    Me.cboTipoMovimiento.AddItem "INGRESO"
    Me.cboTipoMovimiento.ItemData( _
        Me.cboTipoMovimiento.NewIndex) = 1

    Me.cboTipoMovimiento.AddItem "EGRESO"
    Me.cboTipoMovimiento.ItemData( _
        Me.cboTipoMovimiento.NewIndex) = 2
    
    Me.cboTipoMovimiento.AddItem "TRANSFERENCIA INTERBANCARIA"
    Me.cboTipoMovimiento.ItemData( _
        Me.cboTipoMovimiento.NewIndex) = 3
        
End Sub


Private Sub CargarComboOrigen()

    Me.cboOrigen.Clear

    Me.cboOrigen.AddItem "TODOS"
    Me.cboOrigen.ItemData(Me.cboOrigen.NewIndex) = 0

    Me.cboOrigen.AddItem "RECIBO"
    Me.cboOrigen.ItemData(Me.cboOrigen.NewIndex) = 1

    Me.cboOrigen.AddItem "ORDEN DE PAGO"
    Me.cboOrigen.ItemData(Me.cboOrigen.NewIndex) = 2

    Me.cboOrigen.AddItem "LIQUIDACION DE CAJA"
    Me.cboOrigen.ItemData(Me.cboOrigen.NewIndex) = 3

    Me.cboOrigen.AddItem "MOVIMIENTO CAJA/BANCOS"
    Me.cboOrigen.ItemData(Me.cboOrigen.NewIndex) = 4
    
    Me.cboOrigen.AddItem "PAGO A CUENTA"
    Me.cboOrigen.ItemData(Me.cboOrigen.NewIndex) = 5
    
    Me.cboOrigen.AddItem "DEPOSITO"
    Me.cboOrigen.ItemData(Me.cboOrigen.NewIndex) = 6
    
End Sub


Private Function ObtenerIdCombo( _
    ByVal cbo As Object) As Long

    If cbo.ListIndex = -1 Then
        ObtenerIdCombo = 0
    Else
        ObtenerIdCombo = cbo.ItemData(cbo.ListIndex)
    End If

End Function

Private Function ObtenerTipoMovimiento() As String

    If Me.cboTipoMovimiento.ListIndex = -1 Then
        ObtenerTipoMovimiento = vbNullString
        Exit Function
    End If

    Select Case Me.cboTipoMovimiento.ItemData( _
        Me.cboTipoMovimiento.ListIndex)

        Case 1
            ObtenerTipoMovimiento = "INGRESO"

        Case 2
            ObtenerTipoMovimiento = "EGRESO"
            
        Case 3
            ObtenerTipoMovimiento = "TRANSFERENCIA"

        Case Else
            ObtenerTipoMovimiento = vbNullString

    End Select

End Function

Private Function ObtenerOrigen() As String

    If Me.cboOrigen.ListIndex = -1 Then
        ObtenerOrigen = vbNullString
        Exit Function
    End If

    Select Case Me.cboOrigen.ItemData(Me.cboOrigen.ListIndex)

        Case 1
            ObtenerOrigen = "RECIBO"

        Case 2
            ObtenerOrigen = "ORDEN DE PAGO"

        Case 3
            ObtenerOrigen = "LIQUIDACION DE CAJA"

        Case 4
            ObtenerOrigen = "MOVIMIENTO CAJA/BANCOS"
            
        Case 5
            ObtenerOrigen = "PAGO A CUENTA"
            
        Case 6
            ObtenerOrigen = "DEPOSITO"

        Case Else
            ObtenerOrigen = vbNullString

    End Select

End Function


Private Sub PushButton1_Click()
    CargarFiltros

    Set Movimientos = New Collection

    Me.gridResumenBancario.ItemCount = 0
    Me.gridResumenBancario.Refresh
    
End Sub


Private Sub cboRangos_Click()
    funciones.CalculateDateRange Me.cboRangos, Me.dtpDesde(1), Me.dtpHasta(1)
End Sub


Private Sub ConfigurarFormatoImportes()
    With Me.gridResumenBancario
        .Columns("Ingreso").Format = "#,##0.00"
        .Columns("Egreso").Format = "#,##0.00"
        .Columns("saldo_del_período").Format = "#,##0.00"
    End With

End Sub


Private Function ExportarResumenBancario( _
    ByRef col As Collection _
) As Boolean

    On Error GoTo err1

    Const xlCenter As Long = -4108
    Const xlRight As Long = -4152
    Const xlMaximized As Long = -4137
    Const xlOpenXMLWorkbook As Long = 51
    Const xlContinuous As Long = 1

    Dim xlApplication As Object
    Dim xlWorkbook As Object
    Dim xlWorksheet As Object

    Dim Movimiento As DTOResumenBancario
    Dim datos() As Variant

    Dim filaEncabezado As Long
    Dim primeraFilaDatos As Long
    Dim ultimaFila As Long
    Dim indice As Long

    Dim ruta As String

    ExportarResumenBancario = False

    '=========================================================
    ' VALIDACIONES
    '=========================================================
    If col Is Nothing Then
        Err.Raise 5, _
                  "ExportarResumenBancario", _
                  "No se recibió una colección para exportar."
    End If

    If col.count = 0 Then
        Err.Raise 5, _
                  "ExportarResumenBancario", _
                  "La colección no contiene movimientos."
    End If

    '=========================================================
    ' CREAR EXCEL
    '=========================================================
    Set xlApplication = CreateObject("Excel.Application")
    Set xlWorkbook = xlApplication.Workbooks.Add
    Set xlWorksheet = xlWorkbook.Worksheets.item(1)

    xlApplication.Visible = False
    xlApplication.DisplayAlerts = False
    xlApplication.ScreenUpdating = False

    xlWorksheet.Name = "Resumen Bancario"

    '=========================================================
    ' TÍTULO
    '=========================================================
    With xlWorksheet.Range("A1:I1")
        .Merge
        .value = "REPORTE BANCARIO"
        .Font.Bold = True
        .Font.Size = 14
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    '=========================================================
    ' FILTROS APLICADOS
    '=========================================================
    xlWorksheet.Cells(2, 1).value = "Desde:"
    xlWorksheet.Cells(2, 2).value = CDate(Me.dtpDesde(1).value)

    xlWorksheet.Cells(2, 3).value = "Hasta:"
    xlWorksheet.Cells(2, 4).value = CDate(Me.dtpHasta(1).value)

    xlWorksheet.Cells(3, 1).value = "Cuenta:"
    xlWorksheet.Cells(3, 2).value = _
        textoCombo(Me.cboCuentasBancarias)

    xlWorksheet.Cells(3, 3).value = "Moneda:"
    xlWorksheet.Cells(3, 4).value = _
        textoCombo(Me.cboMonedas)

    xlWorksheet.Cells(4, 1).value = "Origen:"
    xlWorksheet.Cells(4, 2).value = _
        textoCombo(Me.cboOrigen)

    xlWorksheet.Cells(4, 3).value = "Tipo:"
    xlWorksheet.Cells(4, 4).value = _
        textoCombo(Me.cboTipoMovimiento)

    xlWorksheet.Range("B2").NumberFormat = "dd/mm/yyyy"
    xlWorksheet.Range("D2").NumberFormat = "dd/mm/yyyy"

    xlWorksheet.Range("A2:A4").Font.Bold = True
    xlWorksheet.Range("C2:C4").Font.Bold = True

    '=========================================================
    ' ENCABEZADOS
    '=========================================================
    filaEncabezado = 6
    primeraFilaDatos = filaEncabezado + 1

    xlWorksheet.Cells(filaEncabezado, 1).value = "Fecha"
    xlWorksheet.Cells(filaEncabezado, 2).value = "Banco"
    xlWorksheet.Cells(filaEncabezado, 3).value = "Cuenta bancaria"
    xlWorksheet.Cells(filaEncabezado, 4).value = "Origen"
    xlWorksheet.Cells(filaEncabezado, 5).value = "Número"
    xlWorksheet.Cells(filaEncabezado, 6).value = "Comprobante"
    xlWorksheet.Cells(filaEncabezado, 7).value = "Ingreso"
    xlWorksheet.Cells(filaEncabezado, 8).value = "Egreso"
    xlWorksheet.Cells(filaEncabezado, 9).value = _
        "Saldo del período"

    With xlWorksheet.Range( _
        xlWorksheet.Cells(filaEncabezado, 1), _
        xlWorksheet.Cells(filaEncabezado, 9))

        .Font.Bold = True
        .Interior.Color = &HC0C0C0
        .HorizontalAlignment = xlCenter
    End With

    '=========================================================
    ' PREPARAR LOS DATOS EN MEMORIA
    '=========================================================
    ReDim datos(1 To col.count, 1 To 9)

    indice = 0

    For Each Movimiento In col

        indice = indice + 1

        datos(indice, 1) = Movimiento.FEcha
        datos(indice, 2) = Movimiento.Banco
        datos(indice, 3) = Movimiento.CuentaBancaria
        datos(indice, 4) = Movimiento.Origen
        datos(indice, 5) = Movimiento.NumeroOrigen
        datos(indice, 6) = Movimiento.Comprobante

        ' Se envían como valores numéricos.
        datos(indice, 7) = CDbl(Movimiento.Ingreso)
        datos(indice, 8) = CDbl(Movimiento.Egreso)
        datos(indice, 9) = CDbl(Movimiento.SaldoAcumulado)

    Next Movimiento

    ultimaFila = primeraFilaDatos + col.count - 1

    ' Escribir todos los movimientos de una vez.
    xlWorksheet.Range( _
        xlWorksheet.Cells(primeraFilaDatos, 1), _
        xlWorksheet.Cells(ultimaFila, 9) _
    ).value = datos

    '=========================================================
    ' FORMATOS
    '=========================================================
    xlWorksheet.Range( _
        xlWorksheet.Cells(primeraFilaDatos, 1), _
        xlWorksheet.Cells(ultimaFila, 1) _
    ).NumberFormat = "dd/mm/yyyy"

    xlWorksheet.Range( _
        xlWorksheet.Cells(primeraFilaDatos, 7), _
        xlWorksheet.Cells(ultimaFila, 9) _
    ).NumberFormat = "#,##0.00"

    xlWorksheet.Range( _
        xlWorksheet.Cells(primeraFilaDatos, 7), _
        xlWorksheet.Cells(ultimaFila, 9) _
    ).HorizontalAlignment = xlRight

    With xlWorksheet.Range( _
        xlWorksheet.Cells(filaEncabezado, 1), _
        xlWorksheet.Cells(ultimaFila, 9))

        .Borders.LineStyle = xlContinuous
    End With

    ' Filtro automático.
    xlWorksheet.Range( _
        xlWorksheet.Cells(filaEncabezado, 1), _
        xlWorksheet.Cells(ultimaFila, 9) _
    ).AutoFilter

    ' Ajustar columnas.
    xlWorksheet.Columns("A:I").EntireColumn.AutoFit

    If xlWorksheet.Columns("A").ColumnWidth < 12 Then
        xlWorksheet.Columns("A").ColumnWidth = 12
    End If

    If xlWorksheet.Columns("B").ColumnWidth < 15 Then
        xlWorksheet.Columns("B").ColumnWidth = 15
    End If

    If xlWorksheet.Columns("C").ColumnWidth < 18 Then
        xlWorksheet.Columns("C").ColumnWidth = 18
    End If

    If xlWorksheet.Columns("D").ColumnWidth < 22 Then
        xlWorksheet.Columns("D").ColumnWidth = 22
    End If

    If xlWorksheet.Columns("F").ColumnWidth < 18 Then
        xlWorksheet.Columns("F").ColumnWidth = 18
    End If

    If xlWorksheet.Columns("G").ColumnWidth < 14 Then
        xlWorksheet.Columns("G").ColumnWidth = 14
    End If

    If xlWorksheet.Columns("H").ColumnWidth < 14 Then
        xlWorksheet.Columns("H").ColumnWidth = 14
    End If

    If xlWorksheet.Columns("I").ColumnWidth < 18 Then
        xlWorksheet.Columns("I").ColumnWidth = 18
    End If

    '=========================================================
    ' CONGELAR ENCABEZADOS
    '=========================================================
    xlWorkbook.Activate
    xlWorksheet.Activate
    xlWorksheet.Range("A7").Select

    xlApplication.ActiveWindow.FreezePanes = True

    '=========================================================
    ' GENERAR RUTA
    '=========================================================
    ruta = Environ$("TEMP")

    If LenB(ruta) = 0 Then
        ruta = Environ$("TMP")
    End If

    If LenB(ruta) = 0 Then
        ruta = App.path
    End If

    If Right$(ruta, 1) <> "\" Then
        ruta = ruta & "\"
    End If

    ruta = ruta & _
           "Reporte_Bancario_" & _
           Format$(Now, "yyyymmdd_hhnnss") & _
           ".xlsx"

    If LenB(Dir$(ruta)) > 0 Then
        Kill ruta
    End If


    If xlWorksheet.Cells(1, 1).value <> _
       "REPORTE BANCARIO" Then

        Err.Raise 5, _
                  "ExportarResumenBancario", _
                  "No se pudo completar el título del archivo."
    End If

    If IsEmpty( _
        xlWorksheet.Cells(primeraFilaDatos, 1).value) Then

        Err.Raise 5, _
                  "ExportarResumenBancario", _
                  "No se escribieron los movimientos en Excel."
    End If

    '=========================================================
    ' GUARDAR
    '=========================================================
    xlWorkbook.SaveAs ruta, xlOpenXMLWorkbook

    If LenB(Dir$(ruta)) = 0 Then
        Err.Raise 5, _
                  "ExportarResumenBancario", _
                  "No se creó el archivo de Excel."
    End If

    If FileLen(ruta) = 0 Then
        Err.Raise 5, _
                  "ExportarResumenBancario", _
                  "El archivo creado está vacío."
    End If

    Debug.Print "Tamaño archivo: "; FileLen(ruta)

    '=========================================================
    ' MOSTRAR EL MISMO LIBRO QUE SE COMPLETÓ
    '=========================================================
    xlApplication.ScreenUpdating = True
    xlApplication.DisplayAlerts = True

    xlWorkbook.Activate
    xlWorksheet.Activate
    xlWorksheet.Range("A1").Select

    xlApplication.Visible = True
    xlApplication.WindowState = xlMaximized

    ' Importante:
    ' No cerrar el libro.
    ' No ejecutar Excel.Quit.
    ' No abrirlo nuevamente con ShellExecute.

    ExportarResumenBancario = True
    Exit Function

err1:
    Debug.Print String$(80, "-")
    Debug.Print "Error ExportarResumenBancario"
    Debug.Print "Número: "; Err.Number
    Debug.Print "Descripción: "; Err.Description
    Debug.Print "Ruta: "; ruta
    Debug.Print String$(80, "-")

    On Error Resume Next

    If Not xlWorkbook Is Nothing Then
        xlWorkbook.Close False
    End If

    If Not xlApplication Is Nothing Then
        xlApplication.Quit
    End If

    Set xlWorksheet = Nothing
    Set xlWorkbook = Nothing
    Set xlApplication = Nothing

    ExportarResumenBancario = False

End Function


Private Sub ConfigurarFormatoGrilla()

    With Me.gridResumenBancario

        .Columns("Fecha").Format = "dd/MM/yyyy"

        .Columns("Ingreso").Format = "#,##0.00"
        .Columns("Egreso").Format = "#,##0.00"
        .Columns("saldo_del_período").Format = "#,##0.00"

    End With

End Sub


Private Sub ReconstruirMovimientosMostrados()

    Dim Movimiento As DTOResumenBancario
    Dim MovimientoInicial As DTOResumenBancario

    Set Movimientos = New Collection

    If MontoInicialEstablecido Then

        Set MovimientoInicial = CrearMovimientoSaldoInicial

        If Not MovimientoInicial Is Nothing Then
            Movimientos.Add MovimientoInicial
        End If

    End If

    If Not MovimientosBase Is Nothing Then

        For Each Movimiento In MovimientosBase
            Movimientos.Add Movimiento
        Next Movimiento

    End If

    RecalcularSaldosMostrados

End Sub


Private Function CrearMovimientoSaldoInicial() _
    As DTOResumenBancario

    On Error GoTo err1

    Dim Movimiento As DTOResumenBancario
    Dim cuenta As CuentaBancaria

    Dim IdCuentaBancaria As Long
    Dim IdMoneda As Long

    IdCuentaBancaria = ObtenerIdCombo( _
        Me.cboCuentasBancarias)

    IdMoneda = ObtenerIdCombo(Me.cboMonedas)

    If IdCuentaBancaria = 0 Then
        Set CrearMovimientoSaldoInicial = Nothing
        Exit Function
    End If

    Set cuenta = DAOCuentaBancaria.FindById( _
        IdCuentaBancaria)

    If cuenta Is Nothing Then
        Set CrearMovimientoSaldoInicial = Nothing
        Exit Function
    End If

    Set Movimiento = New DTOResumenBancario

    ' La línea inicial se ubica en el comienzo del período.
    Movimiento.FEcha = Me.dtpDesde(1).value
    Movimiento.FechaCarga = Me.dtpDesde(1).value

    Movimiento.IdCuentaBancaria = cuenta.Id
    Movimiento.CuentaBancaria = cuenta.numero

    Movimiento.CBU = cuenta.CBU
    Movimiento.IdMoneda = IdMoneda

    If Not cuenta.Banco Is Nothing Then

        Movimiento.IdBanco = cuenta.Banco.Id
        Movimiento.Banco = cuenta.Banco.nombre

    Else

        Movimiento.IdBanco = 0
        Movimiento.Banco = "SIN BANCO"

    End If

    Movimiento.TipoMovimiento = "SALDO INICIAL"
    Movimiento.Origen = "SALDO INICIAL"

    Movimiento.IdOrigen = 0
    Movimiento.NumeroOrigen = "-"

    Movimiento.IdOperacion = 0
    Movimiento.Comprobante = "Monto base"
    Movimiento.detalle = "Monto inicial establecido manualmente"

    ' No es un ingreso ni un egreso real.
    Movimiento.Ingreso = 0
    Movimiento.Egreso = 0

    Movimiento.SaldoAcumulado = MontoInicial

    Set CrearMovimientoSaldoInicial = Movimiento
    Exit Function

err1:
    Set CrearMovimientoSaldoInicial = Nothing

End Function


Private Sub RecalcularSaldosMostrados()

    Dim Movimiento As DTOResumenBancario
    Dim saldo As Double
    Dim i As Long

    saldo = 0

    If MontoInicialEstablecido Then
        saldo = MontoInicial
    End If

    For i = 1 To Movimientos.count

        Set Movimiento = Movimientos.item(i)

        If MontoInicialEstablecido _
           And i = 1 _
           And Movimiento.Origen = "SALDO INICIAL" Then

            Movimiento.SaldoAcumulado = saldo

        Else

            saldo = saldo + _
                    Movimiento.Ingreso - _
                    Movimiento.Egreso

            Movimiento.SaldoAcumulado = saldo

        End If

    Next i

End Sub


Private Sub ActualizarGridResumen()

    Me.gridResumenBancario.ItemCount = 0
    Me.gridResumenBancario.Refresh

    If Movimientos Is Nothing Then
        Me.Label3.caption = "Registros mostrados: 0"
        Exit Sub
    End If

    Me.gridResumenBancario.ItemCount = Movimientos.count
    Me.gridResumenBancario.Refresh

    Me.Label3.caption = _
        "Registros mostrados: " & Movimientos.count

    GridEXHelper.AutoSizeColumns _
        Me.gridResumenBancario

End Sub


Private Sub RestablecerComboFiltro(ByVal cbo As Object)

    On Error GoTo err1

    If cbo.ListCount > 0 Then
        cbo.ListIndex = 0
    Else
        cbo.ListIndex = -1
    End If

    Exit Sub

err1:
    MsgBox "No se pudo restablecer el filtro:" & vbCrLf & _
           Err.Description, _
           vbExclamation, _
           "Reporte bancario"

End Sub


Private Function SeleccionarItemDataCombo( _
    ByVal cbo As Object, _
    ByVal valorBuscado As Long _
) As Boolean

    Dim indice As Long

    SeleccionarItemDataCombo = False

    For indice = 0 To cbo.ListCount - 1

        If cbo.ItemData(indice) = valorBuscado Then

            cbo.ListIndex = indice
            SeleccionarItemDataCombo = True
            Exit Function

        End If

    Next indice

    cbo.ListIndex = -1

End Function


Private Function ObtenerMovimientoResumenSeleccionado() _
    As DTOResumenBancario

    On Error GoTo err1

    Dim indice As Long

    If Movimientos Is Nothing Then Exit Function
    If Movimientos.count = 0 Then Exit Function
    If Me.gridResumenBancario.ItemCount = 0 Then Exit Function

    indice = Me.gridResumenBancario.RowIndex( _
        Me.gridResumenBancario.row)

    If indice < 1 Then Exit Function
    If indice > Movimientos.count Then Exit Function

    Set ObtenerMovimientoResumenSeleccionado = _
        Movimientos.item(indice)

    Exit Function

err1:
    Set ObtenerMovimientoResumenSeleccionado = Nothing

End Function

Private Sub mnuAbrirMovimientoCajaBancos_Click()

    On Error GoTo err1

    Dim MovimientoResumen As DTOResumenBancario
    Dim MovimientoCajaBanco As clsAsientoContable

    Dim F As frmAdminCajaBancosCrearAsientoBancario

    Set MovimientoResumen = _
        ObtenerMovimientoResumenSeleccionado()

    If MovimientoResumen Is Nothing Then

        MsgBox "No se pudo determinar el movimiento seleccionado.", _
               vbExclamation, _
               "Reporte bancario"

        Exit Sub

    End If

    '------------------------------------------------------
    ' VALIDAR ORIGEN
    '------------------------------------------------------
    Select Case UCase$(Trim$(MovimientoResumen.Origen))

        Case "MOVIMIENTO CAJA/BANCOS", _
             "TRANSFERENCIA INTERBANCARIA"

            ' Correcto.

        Case Else

            MsgBox "El registro seleccionado no corresponde " & _
                   "a un movimiento de Caja y Bancos.", _
                   vbExclamation, _
                   "Reporte bancario"

            Exit Sub

    End Select

    '------------------------------------------------------
    ' VALIDAR ID
    '------------------------------------------------------
    If MovimientoResumen.IdOrigen <= 0 Then

        MsgBox "El movimiento seleccionado no tiene " & _
               "un identificador válido.", _
               vbExclamation, _
               "Reporte bancario"

        Exit Sub

    End If

    '------------------------------------------------------
    ' BUSCAR MOVIMIENTO COMPLETO
    '------------------------------------------------------
    Set MovimientoCajaBanco = _
        DAOAsientoContable.FindById( _
            MovimientoResumen.IdOrigen)

    If MovimientoCajaBanco Is Nothing Then

        MsgBox "No se encontró el movimiento de Caja y Bancos Nro " & _
               MovimientoResumen.IdOrigen & ".", _
               vbExclamation, _
               "Reporte bancario"

        Exit Sub

    End If

    '------------------------------------------------------
    ' ABRIR EN SOLO LECTURA
    '------------------------------------------------------
    Set F = New frmAdminCajaBancosCrearAsientoBancario

    Load F

    F.ReadOnly = True
    F.Cargar MovimientoCajaBanco

    F.Show

    Exit Sub

err1:

    MsgBox "No se pudo abrir el movimiento de Caja y Bancos." & _
           vbCrLf & _
           "Error: " & Err.Number & vbCrLf & _
           Err.Description, _
           vbCritical, _
           "Reporte bancario"

End Sub


