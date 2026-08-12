VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmAdminBoletasDepositoLista 
   Caption         =   "Listado de Boletas de Deposito"
   ClientHeight    =   9840
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   18585
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9840
   ScaleWidth      =   18585
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   15195
      _Version        =   786432
      _ExtentX        =   26802
      _ExtentY        =   3413
      _StockProps     =   79
      Caption         =   "Parámetros de búsqueda"
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
      Begin VB.Frame Frame1 
         Height          =   865
         Index           =   0
         Left            =   9960
         TabIndex        =   6
         Top             =   960
         Width           =   5055
         Begin XtremeSuiteControls.PushButton cmdBuscar 
            Default         =   -1  'True
            Height          =   450
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   1350
            _Version        =   786432
            _ExtentX        =   2381
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "Buscar"
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
      End
      Begin VB.Frame Frame1 
         Height          =   735
         Index           =   1
         Left            =   9960
         TabIndex        =   4
         Top             =   240
         Width           =   5055
         Begin XtremeSuiteControls.ProgressBar progreso 
            Height          =   375
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   4815
            _Version        =   786432
            _ExtentX        =   8493
            _ExtentY        =   661
            _StockProps     =   93
            Appearance      =   6
         End
      End
      Begin VB.TextBox txtNumeroBoleta 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   945
         TabIndex        =   3
         Top             =   285
         Width           =   2280
      End
      Begin XtremeSuiteControls.PushButton btnClearCtaBcaria 
         Height          =   360
         Left            =   4530
         TabIndex        =   1
         Top             =   577
         Width           =   420
         _Version        =   786432
         _ExtentX        =   741
         _ExtentY        =   635
         _StockProps     =   79
         Caption         =   "X"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cboCuenta 
         Height          =   315
         Index           =   1
         Left            =   960
         TabIndex        =   2
         Top             =   600
         Width           =   3495
         _Version        =   786432
         _ExtentX        =   6165
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Text            =   "cboCuenta"
      End
      Begin XtremeSuiteControls.DateTimePicker dtpDesde 
         Height          =   315
         Index           =   0
         Left            =   3405
         TabIndex        =   8
         Top             =   2100
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
         Left            =   3390
         TabIndex        =   9
         Top             =   2595
         Width           =   1470
         _Version        =   786432
         _ExtentX        =   2593
         _ExtentY        =   556
         _StockProps     =   68
         CheckBox        =   -1  'True
         Format          =   1
      End
      Begin XtremeSuiteControls.GroupBox GroFechaComprobante 
         Height          =   1575
         Index           =   1
         Left            =   5160
         TabIndex        =   10
         Top             =   240
         Width           =   4695
         _Version        =   786432
         _ExtentX        =   8281
         _ExtentY        =   2778
         _StockProps     =   79
         Caption         =   "Fecha Boleta"
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
            TabIndex        =   11
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
            TabIndex        =   12
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
            TabIndex        =   13
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
            TabIndex        =   16
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
            TabIndex        =   15
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
            TabIndex        =   14
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   195
         Left            =   150
         TabIndex        =   20
         Top             =   360
         Width           =   675
         _Version        =   786432
         _ExtentX        =   1191
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Nº Boleta"
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label lbl 
         Height          =   195
         Left            =   315
         TabIndex        =   19
         Top             =   660
         Width           =   510
         _Version        =   786432
         _ExtentX        =   900
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Cuenta"
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblDesde 
         Height          =   195
         Index           =   0
         Left            =   2865
         TabIndex        =   18
         Top             =   2145
         Width           =   465
         _Version        =   786432
         _ExtentX        =   820
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Desde"
         BackColor       =   12632256
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblHasta 
         Height          =   195
         Index           =   0
         Left            =   2880
         TabIndex        =   17
         Top             =   2655
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
   Begin GridEX20.GridEX gridBoletas 
      Height          =   4185
      Left            =   120
      TabIndex        =   21
      Top             =   2400
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   7382
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      GroupFooterStyle=   2
      ColumnAutoResize=   -1  'True
      MethodHoldFields=   -1  'True
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   8
      Column(1)       =   "frmAdminBoletasDepositoLista.frx":0000
      Column(2)       =   "frmAdminBoletasDepositoLista.frx":015C
      Column(3)       =   "frmAdminBoletasDepositoLista.frx":02A4
      Column(4)       =   "frmAdminBoletasDepositoLista.frx":0404
      Column(5)       =   "frmAdminBoletasDepositoLista.frx":0544
      Column(6)       =   "frmAdminBoletasDepositoLista.frx":068C
      Column(7)       =   "frmAdminBoletasDepositoLista.frx":07D4
      Column(8)       =   "frmAdminBoletasDepositoLista.frx":091C
      FormatStylesCount=   13
      FormatStyle(1)  =   "frmAdminBoletasDepositoLista.frx":0AB4
      FormatStyle(2)  =   "frmAdminBoletasDepositoLista.frx":0BDC
      FormatStyle(3)  =   "frmAdminBoletasDepositoLista.frx":0C8C
      FormatStyle(4)  =   "frmAdminBoletasDepositoLista.frx":0D40
      FormatStyle(5)  =   "frmAdminBoletasDepositoLista.frx":0E18
      FormatStyle(6)  =   "frmAdminBoletasDepositoLista.frx":0ED0
      FormatStyle(7)  =   "frmAdminBoletasDepositoLista.frx":0FB0
      FormatStyle(8)  =   "frmAdminBoletasDepositoLista.frx":1064
      FormatStyle(9)  =   "frmAdminBoletasDepositoLista.frx":111C
      FormatStyle(10) =   "frmAdminBoletasDepositoLista.frx":11D0
      FormatStyle(11) =   "frmAdminBoletasDepositoLista.frx":128C
      FormatStyle(12) =   "frmAdminBoletasDepositoLista.frx":1340
      FormatStyle(13) =   "frmAdminBoletasDepositoLista.frx":13F0
      ImageCount      =   0
      PrinterProperties=   "frmAdminBoletasDepositoLista.frx":148C
   End
   Begin GridEX20.GridEX gridDetalle 
      Height          =   2865
      Left            =   120
      TabIndex        =   23
      Top             =   7080
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   5054
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      GroupFooterStyle=   2
      ColumnAutoResize=   -1  'True
      MethodHoldFields=   -1  'True
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   8
      Column(1)       =   "frmAdminBoletasDepositoLista.frx":165C
      Column(2)       =   "frmAdminBoletasDepositoLista.frx":17B8
      Column(3)       =   "frmAdminBoletasDepositoLista.frx":1900
      Column(4)       =   "frmAdminBoletasDepositoLista.frx":1A78
      Column(5)       =   "frmAdminBoletasDepositoLista.frx":1BB8
      Column(6)       =   "frmAdminBoletasDepositoLista.frx":1CD8
      Column(7)       =   "frmAdminBoletasDepositoLista.frx":1E20
      Column(8)       =   "frmAdminBoletasDepositoLista.frx":1F60
      FormatStylesCount=   13
      FormatStyle(1)  =   "frmAdminBoletasDepositoLista.frx":20D0
      FormatStyle(2)  =   "frmAdminBoletasDepositoLista.frx":21F8
      FormatStyle(3)  =   "frmAdminBoletasDepositoLista.frx":22A8
      FormatStyle(4)  =   "frmAdminBoletasDepositoLista.frx":235C
      FormatStyle(5)  =   "frmAdminBoletasDepositoLista.frx":2434
      FormatStyle(6)  =   "frmAdminBoletasDepositoLista.frx":24EC
      FormatStyle(7)  =   "frmAdminBoletasDepositoLista.frx":25CC
      FormatStyle(8)  =   "frmAdminBoletasDepositoLista.frx":2680
      FormatStyle(9)  =   "frmAdminBoletasDepositoLista.frx":2738
      FormatStyle(10) =   "frmAdminBoletasDepositoLista.frx":27EC
      FormatStyle(11) =   "frmAdminBoletasDepositoLista.frx":28A8
      FormatStyle(12) =   "frmAdminBoletasDepositoLista.frx":295C
      FormatStyle(13) =   "frmAdminBoletasDepositoLista.frx":2A0C
      ImageCount      =   0
      PrinterProperties=   "frmAdminBoletasDepositoLista.frx":2AA8
   End
   Begin VB.Label Label2 
      Caption         =   "Detalle de Cheques"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   6840
      Width           =   5175
   End
   Begin VB.Label lblCantidad 
      Caption         =   "Boletas mostradas [ 0 ]"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   2160
      Width           =   6375
   End
End
Attribute VB_Name = "frmAdminBoletasDepositoLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private desde
Private mBoletas As Collection
Private mCheques As Collection
Private mCargando As Boolean
Dim i As Integer


Private Sub btnClearCtaBcaria_Click()
    Me.cboCuenta(1).ListIndex = -1
End Sub

Private Sub Form_Load()

    On Error GoTo err1

    mCargando = True
    
    Customize Me
    
    GridEXHelper.CustomizeGrid Me.gridBoletas, True
    GridEXHelper.CustomizeGrid Me.gridDetalle, True
    
    GridEXHelper.AutoSizeColumns Me.gridBoletas
    GridEXHelper.AutoSizeColumns Me.gridDetalle
    
    Me.dtpHasta(1).value = Now
    
    Me.dtpDesde(1).value = Year(Now) & "-01-01"

    desde = DateSerial(Year(Date), Month(Date), 1)   ' CDate(1 & "-" & Month(Now) & "-" & Year(Now))
    funciones.FillComboBoxDateRanges Me.cboRangos

    Me.cboRangos.ListIndex = i
    
    For i = 0 To Me.cboRangos.ListCount - 1
        If Me.cboRangos.ItemData(i) = DateRangeValue.DRV_YearCurrent Then Exit For
    Next i
    Me.cboRangos.ListIndex = i


    Set mBoletas = New Collection
    Set mCheques = New Collection

    CargarCuentas

    'Mes actual
    Me.dtpDesde(1).value = _
        DateSerial(Year(Date), Month(Date), 1)

    Me.dtpHasta(1).value = Date

    Me.txtNumeroBoleta.Text = vbNullString

    mCargando = False

    Exit Sub

err1:

    mCargando = False

    MsgBox "No se pudo abrir el historial de boletas." & vbCrLf & _
           Err.Description, _
           vbCritical, "Boletas de depósito"

End Sub


Private Sub cboRangos_Click()
    funciones.CalculateDateRange Me.cboRangos, Me.dtpDesde(1), Me.dtpHasta(1)
End Sub


Private Sub CargarCuentas()

    Dim cuentas As Collection
    Dim cuenta As CuentaBancaria


    Set cuentas = DAOCuentaBancaria.FindAll()


    Me.cboCuenta(1).Clear


    'Primero siempre TODAS
    Me.cboCuenta(1).AddItem "TODAS"
    Me.cboCuenta(1).ItemData(Me.cboCuenta(1).NewIndex) = 0


    For Each cuenta In cuentas

        Me.cboCuenta(1).AddItem cuenta.DescripcionFormateada

        Me.cboCuenta(1).ItemData( _
            Me.cboCuenta(1).NewIndex) = cuenta.Id

    Next cuenta


    Me.cboCuenta(1).ListIndex = -1

End Sub


Private Sub CargarHistorial()

    On Error GoTo err1


    Dim numeroBoleta As Long
    Dim idCuenta As Long


    mCargando = True


    '-------------------------------------------------------
    ' NUMERO DE BOLETA
    '-------------------------------------------------------

    numeroBoleta = 0


    If LenB(Trim$(Me.txtNumeroBoleta.Text)) > 0 Then

        If Not IsNumeric(Me.txtNumeroBoleta.Text) Then

            MsgBox "El número de boleta debe ser numérico.", _
                   vbExclamation, "Boletas de depósito"

            Me.txtNumeroBoleta.SetFocus

            GoTo salir

        End If


        numeroBoleta = CLng(Me.txtNumeroBoleta.Text)

    End If


    '-------------------------------------------------------
    ' CUENTA
    '-------------------------------------------------------

    idCuenta = 0


    If Me.cboCuenta(1).ListIndex >= 0 Then

        idCuenta = Me.cboCuenta(1).ItemData( _
                        Me.cboCuenta(1).ListIndex)

    End If


    '-------------------------------------------------------
    ' CONSULTAR
    '-------------------------------------------------------

    Set mBoletas = DAOBoletaDeposito.FindAll( _
                        Me.dtpDesde(1).value, _
                        Me.dtpHasta(1).value, _
                        numeroBoleta, _
                        idCuenta)
    
    
    If mBoletas Is Nothing Then
    
        Set mBoletas = New Collection
    
        MsgBox "No se pudo cargar el historial." & vbCrLf & _
               DAOBoletaDeposito.UltimoError, _
               vbCritical, "Boletas de depósito"
    
    End If


    Me.gridBoletas.ItemCount = 0
    Me.gridBoletas.ItemCount = mBoletas.count

    Me.gridBoletas.Update


    Me.lblCantidad.caption = _
        "Boletas: " & mBoletas.count


    'Limpiar detalle
    Set mCheques = New Collection

    Me.gridDetalle.ItemCount = 0

'''    Me.lblDetalle.caption = _
'''        "Seleccione una boleta para ver sus cheques."


salir:

    Me.gridBoletas.ItemCount = 0
    Me.gridBoletas.ItemCount = mBoletas.count
    Me.gridBoletas.Update
    
    Me.lblCantidad.caption = _
        "Boletas: " & mBoletas.count
    
    Set mCheques = New Collection
    Me.gridDetalle.ItemCount = 0
    
    mCargando = False
    
    'Si existe alguna boleta, cargar automáticamente
    'el detalle de la seleccionada.
    If mBoletas.count > 0 Then
        CargarDetalleBoletaSeleccionada
    End If
    
    Exit Sub


err1:

    mCargando = False

    MsgBox "Error al cargar las boletas." & vbCrLf & _
           Err.Description, _
           vbCritical, "Boletas de depósito"

End Sub

Private Sub gridBoletas_UnboundReadData( _
            ByVal RowIndex As Long, _
            ByVal Bookmark As Variant, _
            ByVal Values As GridEX20.JSRowData)

    Dim B As BoletaDeposito


    If RowIndex <= 0 Then Exit Sub
    If mBoletas Is Nothing Then Exit Sub
    If RowIndex > mBoletas.count Then Exit Sub


    Set B = mBoletas.item(RowIndex)

    Values(1) = B.Id
    Values(2) = B.numero
    Values(3) = B.fechaDeposito


    If Not B.CuentaDestino Is Nothing Then

        If Not B.CuentaDestino.Banco Is Nothing Then
            Values(4) = B.CuentaDestino.Banco.nombre
        End If

        Values(5) = B.CuentaDestino.numero


        If Not B.CuentaDestino.moneda Is Nothing Then
            Values(6) = B.CuentaDestino.moneda.NombreCorto
        End If

    End If


    Values(7) = B.CantidadCheques
    Values(8) = Replace(FormatCurrency(funciones.FormatearDecimales(B.Monto)), "$", "")

End Sub


Private Sub gridBoletas_SelectionChange()

    CargarDetalleBoletaSeleccionada

End Sub


Private Sub gridBoletas_Click()

    CargarDetalleBoletaSeleccionada

End Sub


Private Sub cmdBuscar_Click()

    CargarHistorial

End Sub


Private Sub cmdRestablecer_Click()

    Me.dtpDesde(1).value = _
        DateSerial(Year(Date), Month(Date), 1)

    Me.dtpHasta(1).value = Date

    Me.txtNumeroBoleta.Text = vbNullString

    If Me.cboCuenta(1).ListCount > 0 Then
        Me.cboCuenta(1).ListIndex = 0
    End If


    CargarHistorial

End Sub


Private Sub gridDetalle_UnboundReadData( _
            ByVal RowIndex As Long, _
            ByVal Bookmark As Variant, _
            ByVal Values As GridEX20.JSRowData)

    Dim ch As cheque

    If RowIndex <= 0 Then Exit Sub
    If mCheques Is Nothing Then Exit Sub
    If RowIndex > mCheques.count Then Exit Sub

    Set ch = mCheques.item(RowIndex)

    '1 - ID
    Values(1) = ch.Id

    '2 - Número de cheque
    Values(2) = ch.numero

    '3 - Fecha de vencimiento
    If ch.FechaVencimiento > 0 Then
        Values(3) = ch.FechaVencimiento
    Else
        Values(3) = Null
    End If

    '4 - Banco
    If Not ch.Banco Is Nothing Then
        Values(4) = ch.Banco.nombre
    Else
        Values(4) = vbNullString
    End If

    '5 - Tipo / origen del cheque
    Values(5) = ch.OrigenCheque

    '6 - Moneda
    If Not ch.moneda Is Nothing Then
        Values(6) = ch.moneda.NombreCorto
    Else
        Values(6) = vbNullString
    End If

    '7 - Monto
    Values(7) = Replace(FormatCurrency(funciones.FormatearDecimales(ch.Monto)), "$", "")

    '8 - Fecha recibido
    If ch.FechaRecibido > 0 Then
        Values(8) = ch.FechaVencimiento
    Else
        Values(8) = Null
    End If

End Sub


Private Sub CargarDetalleBoletaSeleccionada()

    On Error GoTo err1

    Dim idx As Long
    Dim B As BoletaDeposito
    Dim ch As cheque
    Dim total As Double

    Dim NombreBanco As String
    Dim numeroCuenta As String
    Dim nombreMoneda As String

    If mCargando Then Exit Sub
    If mBoletas Is Nothing Then Exit Sub
    If mBoletas.count = 0 Then Exit Sub

    idx = Me.gridBoletas.RowIndex(Me.gridBoletas.row)

    If idx <= 0 Then Exit Sub
    If idx > mBoletas.count Then Exit Sub

    Set B = mBoletas.item(idx)

    '-------------------------------------------------------
    ' CARGAR CHEQUES DE LA BOLETA
    '-------------------------------------------------------

    Set mCheques = _
        DAOBoletaDeposito.FindChequesByBoleta(B.Id)

    If mCheques Is Nothing Then

        Set mCheques = New Collection

        MsgBox "No se pudo obtener el detalle." & vbCrLf & _
               DAOBoletaDeposito.UltimoError, _
               vbCritical, "Boletas de depósito"

        Exit Sub

    End If

    'DEBUG TEMPORAL
    'MsgBox "Boleta ID: " & b.Id & vbCrLf & _
    '       "Cheques encontrados: " & mCheques.count

    Me.gridDetalle.ItemCount = 0
    Me.gridDetalle.ItemCount = mCheques.count
    Me.gridDetalle.Update

    GridEXHelper.AutoSizeColumns Me.gridDetalle, True

    '-------------------------------------------------------
    ' DATOS CUENTA
    '-------------------------------------------------------

    NombreBanco = vbNullString
    numeroCuenta = vbNullString
    nombreMoneda = vbNullString

    If Not B.CuentaDestino Is Nothing Then

        numeroCuenta = B.CuentaDestino.numero

        If Not B.CuentaDestino.Banco Is Nothing Then
            NombreBanco = B.CuentaDestino.Banco.nombre
        End If

        If Not B.CuentaDestino.moneda Is Nothing Then
            nombreMoneda = B.CuentaDestino.moneda.NombreCorto
        End If

    End If

    '-------------------------------------------------------
    ' TOTAL
    '-------------------------------------------------------

    total = 0

    For Each ch In mCheques
        total = total + ch.Monto
    Next ch


    Exit Sub

err1:

    MsgBox "No se pudo cargar el detalle de la boleta." & _
           vbCrLf & Err.Description, _
           vbCritical, "Boletas de depósito"

End Sub



