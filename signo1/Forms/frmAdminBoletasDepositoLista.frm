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
   ScaleHeight     =   9840
   ScaleWidth      =   18585
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   21885
      _Version        =   786432
      _ExtentX        =   38603
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
         TabIndex        =   8
         Top             =   960
         Width           =   5055
         Begin XtremeSuiteControls.PushButton cmdRestablecer 
            Height          =   450
            Left            =   1440
            TabIndex        =   30
            Top             =   240
            Width           =   615
            _Version        =   786432
            _ExtentX        =   1085
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "Reset"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton btnExportar 
            Height          =   450
            Left            =   2160
            TabIndex        =   9
            Top             =   240
            Width           =   1335
            _Version        =   786432
            _ExtentX        =   2355
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "Exportar"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton cmdBuscar 
            Default         =   -1  'True
            Height          =   450
            Left            =   0
            TabIndex        =   10
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
         Begin XtremeSuiteControls.PushButton cmdImprimir 
            Height          =   450
            Left            =   3600
            TabIndex        =   11
            Top             =   240
            Width           =   1350
            _Version        =   786432
            _ExtentX        =   2381
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "Imprimir"
            UseVisualStyle  =   -1  'True
         End
      End
      Begin VB.Frame Frame1 
         Height          =   735
         Index           =   1
         Left            =   9960
         TabIndex        =   6
         Top             =   240
         Width           =   5055
         Begin XtremeSuiteControls.ProgressBar progreso 
            Height          =   375
            Left            =   120
            TabIndex        =   7
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
         TabIndex        =   5
         Top             =   285
         Width           =   2280
      End
      Begin XtremeSuiteControls.GroupBox Totales 
         Height          =   1575
         Left            =   15120
         TabIndex        =   1
         Top             =   240
         Width           =   6615
         _Version        =   786432
         _ExtentX        =   11668
         _ExtentY        =   2778
         _StockProps     =   79
         Caption         =   "Resumen"
         UseVisualStyle  =   -1  'True
         Begin VB.Label Label4 
            Caption         =   "Label4"
            Height          =   375
            Left            =   120
            TabIndex        =   2
            Top             =   360
            Width           =   3375
         End
      End
      Begin XtremeSuiteControls.PushButton btnClearCtaBcaria 
         Height          =   255
         Left            =   4530
         TabIndex        =   3
         Top             =   610
         Width           =   420
         _Version        =   786432
         _ExtentX        =   741
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "X"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cboCuenta 
         Height          =   315
         Index           =   1
         Left            =   960
         TabIndex        =   4
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
         TabIndex        =   12
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
         TabIndex        =   13
         Top             =   2595
         Width           =   1470
         _Version        =   786432
         _ExtentX        =   2593
         _ExtentY        =   556
         _StockProps     =   68
         CheckBox        =   -1  'True
         Format          =   1
      End
      Begin XtremeSuiteControls.ComboBox cboEstado 
         Height          =   315
         Left            =   945
         TabIndex        =   14
         Top             =   1320
         Width           =   3510
         _Version        =   786432
         _ExtentX        =   6191
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.PushButton cmdLimpiaEstado 
         Height          =   255
         Left            =   4530
         TabIndex        =   15
         Top             =   1350
         Width           =   420
         _Version        =   786432
         _ExtentX        =   741
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "X"
         BackColor       =   12632256
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.GroupBox GroFechaComprobante 
         Height          =   1575
         Index           =   1
         Left            =   5160
         TabIndex        =   16
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
            TabIndex        =   17
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
            TabIndex        =   18
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
            TabIndex        =   19
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
            TabIndex        =   22
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
            TabIndex        =   21
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
            TabIndex        =   20
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
         TabIndex        =   27
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
         TabIndex        =   26
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
         TabIndex        =   25
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
         TabIndex        =   24
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
      Begin XtremeSuiteControls.Label lblRango 
         Height          =   195
         Index           =   0
         Left            =   330
         TabIndex        =   23
         Top             =   1380
         Width           =   495
         _Version        =   786432
         _ExtentX        =   873
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Estado"
         BackColor       =   12632256
         AutoSize        =   -1  'True
      End
   End
   Begin GridEX20.GridEX gridBoletas 
      Height          =   4185
      Left            =   120
      TabIndex        =   28
      Top             =   2400
      Width           =   14775
      _ExtentX        =   26061
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
      Column(3)       =   "frmAdminBoletasDepositoLista.frx":027C
      Column(4)       =   "frmAdminBoletasDepositoLista.frx":03DC
      Column(5)       =   "frmAdminBoletasDepositoLista.frx":051C
      Column(6)       =   "frmAdminBoletasDepositoLista.frx":063C
      Column(7)       =   "frmAdminBoletasDepositoLista.frx":0784
      Column(8)       =   "frmAdminBoletasDepositoLista.frx":08CC
      FormatStylesCount=   13
      FormatStyle(1)  =   "frmAdminBoletasDepositoLista.frx":0A1C
      FormatStyle(2)  =   "frmAdminBoletasDepositoLista.frx":0B44
      FormatStyle(3)  =   "frmAdminBoletasDepositoLista.frx":0BF4
      FormatStyle(4)  =   "frmAdminBoletasDepositoLista.frx":0CA8
      FormatStyle(5)  =   "frmAdminBoletasDepositoLista.frx":0D80
      FormatStyle(6)  =   "frmAdminBoletasDepositoLista.frx":0E38
      FormatStyle(7)  =   "frmAdminBoletasDepositoLista.frx":0F18
      FormatStyle(8)  =   "frmAdminBoletasDepositoLista.frx":0FCC
      FormatStyle(9)  =   "frmAdminBoletasDepositoLista.frx":1084
      FormatStyle(10) =   "frmAdminBoletasDepositoLista.frx":1138
      FormatStyle(11) =   "frmAdminBoletasDepositoLista.frx":11F4
      FormatStyle(12) =   "frmAdminBoletasDepositoLista.frx":12A8
      FormatStyle(13) =   "frmAdminBoletasDepositoLista.frx":1358
      ImageCount      =   0
      PrinterProperties=   "frmAdminBoletasDepositoLista.frx":13F4
   End
   Begin GridEX20.GridEX gridDetalle 
      Height          =   2985
      Left            =   120
      TabIndex        =   31
      Top             =   6720
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   5265
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
      ColumnsCount    =   7
      Column(1)       =   "frmAdminBoletasDepositoLista.frx":15C4
      Column(2)       =   "frmAdminBoletasDepositoLista.frx":16DC
      Column(3)       =   "frmAdminBoletasDepositoLista.frx":17E0
      Column(4)       =   "frmAdminBoletasDepositoLista.frx":18CC
      Column(5)       =   "frmAdminBoletasDepositoLista.frx":19C0
      Column(6)       =   "frmAdminBoletasDepositoLista.frx":1AB4
      Column(7)       =   "frmAdminBoletasDepositoLista.frx":1BA0
      FormatStylesCount=   13
      FormatStyle(1)  =   "frmAdminBoletasDepositoLista.frx":1C9C
      FormatStyle(2)  =   "frmAdminBoletasDepositoLista.frx":1DC4
      FormatStyle(3)  =   "frmAdminBoletasDepositoLista.frx":1E74
      FormatStyle(4)  =   "frmAdminBoletasDepositoLista.frx":1F28
      FormatStyle(5)  =   "frmAdminBoletasDepositoLista.frx":2000
      FormatStyle(6)  =   "frmAdminBoletasDepositoLista.frx":20B8
      FormatStyle(7)  =   "frmAdminBoletasDepositoLista.frx":2198
      FormatStyle(8)  =   "frmAdminBoletasDepositoLista.frx":224C
      FormatStyle(9)  =   "frmAdminBoletasDepositoLista.frx":2304
      FormatStyle(10) =   "frmAdminBoletasDepositoLista.frx":23B8
      FormatStyle(11) =   "frmAdminBoletasDepositoLista.frx":2474
      FormatStyle(12) =   "frmAdminBoletasDepositoLista.frx":2528
      FormatStyle(13) =   "frmAdminBoletasDepositoLista.frx":25D8
      ImageCount      =   0
      PrinterProperties=   "frmAdminBoletasDepositoLista.frx":2674
   End
   Begin VB.Label lblCantidad 
      Caption         =   "Boletas mostradas [ 0 ]"
      Height          =   255
      Left            =   120
      TabIndex        =   29
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

Private mBoletas As Collection
Private mCheques As Collection
Private mCargando As Boolean


Private Sub Form_Load()

    On Error GoTo err1

    Customize Me

    GridEXHelper.CustomizeGrid Me.gridBoletas, False, False
    GridEXHelper.CustomizeGrid Me.gridDetalle, False, False


    Set mBoletas = New Collection
    Set mCheques = New Collection


    CargarCuentas


    'Mes actual
    Me.dtpDesde(1).value = _
        DateSerial(Year(Date), Month(Date), 1)

    Me.dtpHasta(1).value = Date


    Me.txtNumeroBoleta.Text = vbNullString


    CargarHistorial

    Exit Sub


err1:

    MsgBox "No se pudo abrir el historial de depósitos." & _
           vbCrLf & Err.Description, _
           vbCritical, "Boletas de depósito"

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


    Me.cboCuenta(1).ListIndex = 0

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

'''    Set mBoletas = _
'''        DAOBoletaDeposito.FindAllHistorialCheques( _
'''            Me.dtpDesde(1).value, _
'''            Me.dtpHasta(1).value, _
'''            numeroBoleta, _
'''            idCuenta)


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

    mCargando = False

    Exit Sub


err1:

    mCargando = False

    MsgBox "Error al cargar las boletas." & vbCrLf & _
           Err.Description, _
           vbCritical, "Boletas de depósito"

End Sub

Private Sub gridBoletas_UnboundReadData( _
            ByVal rowIndex As Long, _
            ByVal Bookmark As Variant, _
            ByVal Values As GridEX20.JSRowData)


    Dim b As DTOBoletaDepositoHistorial


    If rowIndex <= 0 Then Exit Sub

    If mBoletas Is Nothing Then Exit Sub

    If rowIndex > mBoletas.count Then Exit Sub


    Set b = mBoletas.item(rowIndex)


    Values(1) = b.numeroBoleta
    Values(2) = b.fechaDeposito
    Values(3) = b.BancoNombre
    Values(4) = b.CuentaNumero
    Values(5) = b.MonedaNombre
    Values(6) = b.CantidadCheques
    Values(7) = b.Monto

End Sub

Private Sub gridBoletas_SelectionChange()

    On Error GoTo err1


    Dim idx As Long
    Dim b As DTOBoletaDepositoHistorial

    Dim ch As cheque
    Dim total As Double


    If mCargando Then Exit Sub

    If mBoletas Is Nothing Then Exit Sub


    idx = Me.gridBoletas.rowIndex( _
            Me.gridBoletas.row)


    If idx <= 0 Then Exit Sub

    If idx > mBoletas.count Then Exit Sub


    Set b = mBoletas.item(idx)


    Set mCheques = _
        DAOBoletaDeposito.FindChequesByBoleta(b.Id)


    If mCheques Is Nothing Then

        Set mCheques = New Collection

        MsgBox "No se pudo obtener el detalle." & vbCrLf & _
               DAOBoletaDeposito.UltimoError, _
               vbCritical, "Boletas de depósito"

        Exit Sub

    End If


    Me.gridDetalle.ItemCount = 0
    Me.gridDetalle.ItemCount = mCheques.count

    Me.gridDetalle.Update


    '-------------------------------------------------------
    ' TOTAL DETALLE
    '-------------------------------------------------------

    total = 0

    For Each ch In mCheques
        total = total + ch.Monto
    Next ch


    Me.lblDetalle.caption = _
        "Boleta Nº " & b.numeroBoleta & _
        "  |  " & _
        Format$(b.fechaDeposito, "dd/mm/yyyy") & _
        "  |  " & _
        b.BancoNombre & _
        "  |  " & _
        b.CuentaNumero & _
        "  |  Total: " & _
        b.MonedaNombre & " " & _
        Format$(total, "#,##0.00")


    Exit Sub


err1:

    MsgBox "No se pudo cargar el detalle de la boleta." & _
           vbCrLf & Err.Description, _
           vbCritical, "Boletas de depósito"

End Sub

Private Sub gridDetalle_UnboundReadData( _
            ByVal rowIndex As Long, _
            ByVal Bookmark As Variant, _
            ByVal Values As GridEX20.JSRowData)


    Dim ch As cheque


    If rowIndex <= 0 Then Exit Sub

    If mCheques Is Nothing Then Exit Sub

    If rowIndex > mCheques.count Then Exit Sub


    Set ch = mCheques.item(rowIndex)


    Values(1) = ch.numero
    Values(2) = ch.FechaVencimiento


    If Not ch.Banco Is Nothing Then
        Values(3) = ch.Banco.nombre
    Else
        Values(3) = vbNullString
    End If


    Values(4) = ch.OrigenCheque


    If Not ch.moneda Is Nothing Then
        Values(5) = ch.moneda.NombreCorto
    Else
        Values(5) = vbNullString
    End If


    Values(6) = ch.Monto
    Values(7) = ch.FechaRecibido

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
