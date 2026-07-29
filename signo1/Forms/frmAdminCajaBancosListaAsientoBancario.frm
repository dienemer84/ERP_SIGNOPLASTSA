VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmAdminCajaBancosListaAsientoBancario 
   Caption         =   "Movimientos de caja y bancos"
   ClientHeight    =   8880
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   18015
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   24657.09
   ScaleMode       =   0  'User
   ScaleWidth      =   18015
   WindowState     =   2  'Maximized
   Begin VB.PictureBox pic 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   0
      Top             =   8160
      Visible         =   0   'False
      Width           =   615
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1935
      Left            =   120
      TabIndex        =   1
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
      Begin XtremeSuiteControls.GroupBox Totales 
         Height          =   1575
         Left            =   15120
         TabIndex        =   32
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
            TabIndex        =   33
            Top             =   360
            Width           =   3375
         End
      End
      Begin XtremeSuiteControls.PushButton btnClearCtaBcaria 
         Height          =   255
         Left            =   4530
         TabIndex        =   30
         Top             =   610
         Width           =   420
         _Version        =   786432
         _ExtentX        =   741
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "X"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cboCtaBcaria 
         Height          =   315
         Left            =   960
         TabIndex        =   28
         Top             =   600
         Width           =   3495
         _Version        =   786432
         _ExtentX        =   6165
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Text            =   "cboCtaBcaria"
      End
      Begin VB.TextBox txtNro 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   945
         TabIndex        =   8
         Top             =   285
         Width           =   2280
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
      Begin VB.Frame Frame1 
         Height          =   865
         Index           =   0
         Left            =   9960
         TabIndex        =   2
         Top             =   960
         Width           =   5055
         Begin XtremeSuiteControls.PushButton btnExportar 
            Height          =   450
            Left            =   1920
            TabIndex        =   3
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
            Left            =   120
            TabIndex        =   4
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
            TabIndex        =   5
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
      Begin XtremeSuiteControls.ComboBox cboCuenta 
         Height          =   315
         Left            =   945
         TabIndex        =   9
         Top             =   975
         Width           =   3510
         _Version        =   786432
         _ExtentX        =   6191
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Text            =   "cboCtaContable"
      End
      Begin XtremeSuiteControls.PushButton btnClearCtaCble 
         Height          =   255
         Left            =   4530
         TabIndex        =   10
         Top             =   980
         Width           =   420
         _Version        =   786432
         _ExtentX        =   741
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "X"
         BackColor       =   12632256
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.DateTimePicker dtpDesde 
         Height          =   315
         Index           =   0
         Left            =   3405
         TabIndex        =   11
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
         TabIndex        =   12
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
         TabIndex        =   13
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
         TabIndex        =   14
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
         TabIndex        =   15
         Top             =   240
         Width           =   4695
         _Version        =   786432
         _ExtentX        =   8281
         _ExtentY        =   2778
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
            TabIndex        =   16
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
         Begin XtremeSuiteControls.ComboBox cboRangos 
            Height          =   315
            Left            =   720
            TabIndex        =   18
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
         Begin XtremeSuiteControls.Label lblRango 
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   21
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
         Begin XtremeSuiteControls.Label lblDesde 
            Height          =   195
            Index           =   1
            Left            =   165
            TabIndex        =   20
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
         Begin XtremeSuiteControls.Label lblHasta 
            Height          =   195
            Index           =   1
            Left            =   2400
            TabIndex        =   19
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
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Cta. Cble"
         Height          =   255
         Left            =   -150
         TabIndex        =   29
         Top             =   1005
         Width           =   975
      End
      Begin XtremeSuiteControls.Label lblRango 
         Height          =   195
         Index           =   0
         Left            =   330
         TabIndex        =   26
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
      Begin XtremeSuiteControls.Label lblHasta 
         Height          =   195
         Index           =   0
         Left            =   2880
         TabIndex        =   25
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
      Begin XtremeSuiteControls.Label lblDesde 
         Height          =   195
         Index           =   0
         Left            =   2865
         TabIndex        =   24
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
      Begin XtremeSuiteControls.Label lbl 
         Height          =   195
         Left            =   0
         TabIndex        =   23
         Top             =   660
         Width           =   825
         _Version        =   786432
         _ExtentX        =   1455
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Cta. Bcaria."
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   195
         Left            =   240
         TabIndex        =   22
         Top             =   330
         Width           =   585
         _Version        =   786432
         _ExtentX        =   1032
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Nº Mov."
         AutoSize        =   -1  'True
      End
   End
   Begin GridEX20.GridEX gridOrdenes 
      Height          =   5505
      Left            =   120
      TabIndex        =   27
      Top             =   2400
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   9710
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
      ColumnsCount    =   11
      Column(1)       =   "frmAdminCajaBancosListaAsientoBancario.frx":0000
      Column(2)       =   "frmAdminCajaBancosListaAsientoBancario.frx":0178
      Column(3)       =   "frmAdminCajaBancosListaAsientoBancario.frx":02B8
      Column(4)       =   "frmAdminCajaBancosListaAsientoBancario.frx":03F8
      Column(5)       =   "frmAdminCajaBancosListaAsientoBancario.frx":0538
      Column(6)       =   "frmAdminCajaBancosListaAsientoBancario.frx":06A0
      Column(7)       =   "frmAdminCajaBancosListaAsientoBancario.frx":07E8
      Column(8)       =   "frmAdminCajaBancosListaAsientoBancario.frx":0930
      Column(9)       =   "frmAdminCajaBancosListaAsientoBancario.frx":0A70
      Column(10)      =   "frmAdminCajaBancosListaAsientoBancario.frx":0BA8
      Column(11)      =   "frmAdminCajaBancosListaAsientoBancario.frx":0CF0
      FormatStylesCount=   13
      FormatStyle(1)  =   "frmAdminCajaBancosListaAsientoBancario.frx":0E10
      FormatStyle(2)  =   "frmAdminCajaBancosListaAsientoBancario.frx":0F38
      FormatStyle(3)  =   "frmAdminCajaBancosListaAsientoBancario.frx":0FE8
      FormatStyle(4)  =   "frmAdminCajaBancosListaAsientoBancario.frx":109C
      FormatStyle(5)  =   "frmAdminCajaBancosListaAsientoBancario.frx":1174
      FormatStyle(6)  =   "frmAdminCajaBancosListaAsientoBancario.frx":122C
      FormatStyle(7)  =   "frmAdminCajaBancosListaAsientoBancario.frx":130C
      FormatStyle(8)  =   "frmAdminCajaBancosListaAsientoBancario.frx":13C0
      FormatStyle(9)  =   "frmAdminCajaBancosListaAsientoBancario.frx":1478
      FormatStyle(10) =   "frmAdminCajaBancosListaAsientoBancario.frx":152C
      FormatStyle(11) =   "frmAdminCajaBancosListaAsientoBancario.frx":15E8
      FormatStyle(12) =   "frmAdminCajaBancosListaAsientoBancario.frx":169C
      FormatStyle(13) =   "frmAdminCajaBancosListaAsientoBancario.frx":174C
      ImageCount      =   0
      PrinterProperties=   "frmAdminCajaBancosListaAsientoBancario.frx":17E8
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   840
      Top             =   8160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Flags           =   1
   End
   Begin VB.Label Label3 
      Caption         =   "Movimientos mostrados [ 0 ]"
      Height          =   255
      Left            =   120
      TabIndex        =   31
      Top             =   2160
      Width           =   6375
   End
   Begin VB.Menu menu 
      Caption         =   "menu"
      Begin VB.Menu mnuEditar 
         Caption         =   "Editar"
      End
      Begin VB.Menu mnuAprobar 
         Caption         =   "Aprobar"
      End
      Begin VB.Menu mnuAnular 
         Caption         =   "Anular"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuVer 
         Caption         =   "Ver"
      End
      Begin VB.Menu mnuImprimir 
         Caption         =   "Imprimir"
      End
   End
End
Attribute VB_Name = "frmAdminCajaBancosListaAsientoBancario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private desde
Dim ids As String
Private Movimientos As New Collection
Private AsientoContable As clsAsientoContable
Dim i As Integer


Private Sub btnClearProveedor_Click()
    Me.cboCuenta.ListIndex = -1
End Sub


Private Sub btnClearCtaBcaria_Click()
    Me.cboCtaBcaria.ListIndex = -1
End Sub

Private Sub btnClearCtaCble_Click()
    Me.cboCuenta.ListIndex = -1
End Sub

Private Sub btnExportar_Click()

    Me.progreso.Visible = True

    If IsSomething(Movimientos) Then
        If Not DAOAsientoContable.ExportarColeccion(Movimientos, Me.progreso) Then GoTo err1
    End If

    Me.progreso.Visible = False

    Exit Sub
err1:
    MsgBox "Se produjo un error al exportar!", vbCritical, "Error"

End Sub


Private Sub cmdBuscar_Click()
    If 1 = 1 Then llenarLista Else Me.gridOrdenes.ItemCount = 0

End Sub


Private Sub cmdImprimir_Click()

    Dim pro As String
    If Me.cboCuenta.ListIndex > -1 Then
        pro = " Cuenta Contable: " & Me.cboCuenta.Text
    End If

    With Me.gridOrdenes.PrinterProperties

        .FitColumns = True
        .RepeatHeaders = True
        .Orientation = jgexPPLandscape
        .HeaderString(jgexHFCenter) = "Listado de Movimientos de Caja y Bancos"
        If LenB(pro) > 1 Then
            .HeaderString(jgexHFLeft) = pro
        End If
        .FooterString(jgexHFCenter) = Now

    End With
    Load frmPrintPreview
    frmPrintPreview.Move Me.Left, Me.Top, Me.Width, Me.Height
    Me.gridOrdenes.PrintPreview frmPrintPreview.GEXPreview1
    frmPrintPreview.Show 1

End Sub


Private Sub cmdLimpiaEstado_Click()
    Me.cboEstado.ListIndex = -1
End Sub


Private Sub Form_Load()

    Customize Me
    GridEXHelper.CustomizeGrid Me.gridOrdenes, True
    
    DAOCuentaBancaria.llenarComboXtremeSuite Me.cboCtaBcaria
    Me.cboCtaBcaria.ListIndex = -1
    
    DAOCuentaContable.llenarComboXtremeSuite Me.cboCuenta, True, True, True
    Me.cboCuenta.ListIndex = -1

    Me.dtpHasta(1).value = Now
    
    Me.gridOrdenes.ItemCount = 0
    
    Me.Label4.caption = "Total: " & FormatCurrency(0)
    
    GridEXHelper.AutoSizeColumns Me.gridOrdenes
    ids = funciones.CreateGUID
      
    Me.cboEstado.Clear
    Me.cboEstado.AddItem enums.enumEstadoMovimientosCajaYBancos(EstadoMovimientoCajaYBancos.EnEdicion)
    Me.cboEstado.ItemData(Me.cboEstado.NewIndex) = EstadoMovimientoCajaYBancos.EnEdicion
    Me.cboEstado.AddItem enums.enumEstadoMovimientosCajaYBancos(EstadoMovimientoCajaYBancos.Aprobado)
    Me.cboEstado.ItemData(Me.cboEstado.NewIndex) = EstadoMovimientoCajaYBancos.Aprobado
    
    Me.dtpDesde(1).value = Year(Now) & "-01-01"

    desde = DateSerial(Year(Date), Month(Date), 1)   ' CDate(1 & "-" & Month(Now) & "-" & Year(Now))
    funciones.FillComboBoxDateRanges Me.cboRangos

    Me.cboRangos.ListIndex = i
    
    For i = 0 To Me.cboRangos.ListCount - 1
        If Me.cboRangos.ItemData(i) = DateRangeValue.DRV_YearCurrent Then Exit For
    Next i
    Me.cboRangos.ListIndex = i
    
End Sub


Private Sub llenarLista()
    
    Dim filter As String
    filter = "1 = 1"
    
    If Me.cboCtaBcaria.ListIndex > -1 Then
        filter = filter & " AND movimientos_caja_bancos.id_cuenta_bancaria_principal = " & Me.cboCtaBcaria.ItemData(Me.cboCtaBcaria.ListIndex)
    End If

    If Me.cboCuenta.ListIndex > -1 Then
        filter = filter & " AND movimientos_caja_bancos.id_cuentacontable = " & Me.cboCuenta.ItemData(Me.cboCuenta.ListIndex)
    End If

    If LenB(Me.txtNro.Text) > 0 Then
        filter = filter & " AND  movimientos_caja_bancos.id LIKE '%" & val(Me.txtNro.Text) & "%'"
    End If
    
    If Me.cboEstado.ListIndex > -1 Then
        filter = filter & " AND movimientos_caja_bancos.estado = " & Me.cboEstado.ItemData(Me.cboEstado.ListIndex)
    End If

    Dim filtroor As String

    If Not IsNull(Me.dtpDesde(1).value) Then
        filter = filter & " AND movimientos_caja_bancos.fecha >= " & conectar.Escape(Me.dtpDesde(1).value)
    End If

    If Not IsNull(Me.dtpHasta(1).value) Then
        filter = filter & " AND movimientos_caja_bancos.fecha <= " & conectar.Escape(Me.dtpHasta(1).value)
    End If

    If LenB(filtroor) > 0 Then
        filtroor = " AND (" & Right(filtroor, Len(filtroor) - 3) & " )"
        filter = filter & filtroor
    End If

    Me.gridOrdenes.ItemCount = 0
    
    Set Movimientos = DAOAsientoContable.FindAll(filter, "movimientos_caja_bancos.id DESC")
    
    Me.gridOrdenes.ItemCount = Movimientos.count
    
    Me.caption = _
        "Listado de Movimientos" & _
        " [Cant: " & Movimientos.count & "]"
    
    Me.Label3.caption = _
        "Movimientos mostrados" & _
        " [Cant: " & Movimientos.count & "]"
    
    TotalizarMovimientos
    

End Sub


Private Sub Form_Resize()
    On Error Resume Next
    Me.gridOrdenes.Width = Me.ScaleWidth - 300
    Me.gridOrdenes.Height = (Me.ScaleHeight * 75) / 100

    Me.GroupBox1.Width = Me.gridOrdenes.Width
    GridEXHelper.AutoSizeColumns Me.gridOrdenes
    
End Sub


Private Sub gridOrdenes_ColumnHeaderClick(ByVal Column As GridEX20.JSColumn)
    GridEXHelper.ColumnHeaderClick Me.gridOrdenes, Column
End Sub


Private Sub gridOrdenes_DblClick()

    Dim movSeleccionado As clsAsientoContable

    Set movSeleccionado = ObtenerMovimientoSeleccionado()

    If Not IsSomething(movSeleccionado) Then Exit Sub

    If movSeleccionado.estado = _
            EstadoMovimientoCajaYBancos.EnEdicion Then

        AbrirMovimiento movSeleccionado, False

    Else

        AbrirMovimiento movSeleccionado, True

    End If

End Sub


Private Sub gridOrdenes_SelectionChange()
    SeleccionarOP
End Sub


Private Sub gridOrdenes_MouseUp( _
    Button As Integer, _
    Shift As Integer, _
    x As Single, _
    y As Single)

    If Button <> 2 Then Exit Sub
    If Movimientos.count = 0 Then Exit Sub

    Set AsientoContable = ObtenerMovimientoSeleccionado()

    If Not IsSomething(AsientoContable) Then Exit Sub

    Me.mnuEditar.Enabled = _
        (AsientoContable.estado = _
         EstadoMovimientoCajaYBancos.EnEdicion)

    Me.mnuAprobar.Enabled = _
        (AsientoContable.estado = _
         EstadoMovimientoCajaYBancos.EnEdicion)

    Me.mnuAnular.Enabled = _
        (AsientoContable.estado <> _
         EstadoMovimientoCajaYBancos.EnEdicion)

    Me.mnuVer.Enabled = True

    Me.mnuImprimir.Enabled = _
        (AsientoContable.estado = _
         EstadoMovimientoCajaYBancos.Aprobado)

    Me.PopupMenu menu

End Sub


Private Sub gridOrdenes_RowFormat( _
    RowBuffer As GridEX20.JSRowData)

    On Error GoTo salir

    If RowBuffer.RowIndex <= 0 Then Exit Sub
    If RowBuffer.RowIndex > Movimientos.count Then Exit Sub

    Dim mov As clsAsientoContable

    Set mov = Movimientos.item(RowBuffer.RowIndex)

    Select Case mov.estado

        Case EstadoMovimientoCajaYBancos.EnEdicion
            RowBuffer.CellStyle(6) = "pendiente"

        Case EstadoMovimientoCajaYBancos.Aprobado
            RowBuffer.CellStyle(6) = "aprobada"

    End Select

    Select Case UCase$(mov.TipoMovimiento)

        Case "INGRESO"
            RowBuffer.CellStyle(3) = "INGRESO"

        Case "EGRESO", "SALIDA"
            RowBuffer.CellStyle(3) = "EGRESO"

        Case "TRANSFERENCIA"
            RowBuffer.CellStyle(3) = "TRANSFERENCIA"

    End Select

salir:

End Sub


Private Sub gridOrdenes_UnboundReadData( _
    ByVal RowIndex As Long, _
    ByVal Bookmark As Variant, _
    ByVal Values As GridEX20.JSRowData)

    If RowIndex <= 0 Then Exit Sub
    If RowIndex > Movimientos.count Then Exit Sub

    Dim mov As clsAsientoContable

    Set mov = Movimientos.item(RowIndex)

    Values(1) = mov.Id
    Values(2) = mov.FEcha
    Values(3) = mov.TipoMovimiento

        If mov.TipoMovimiento = "TRANSFERENCIA" Then
        
            If IsSomething(mov.CuentaBancaria) And IsSomething(mov.CuentaBancariaDestino) Then
                Values(4) = mov.CuentaBancaria.DescripcionFormateada _
                            & "  ->  " _
                            & mov.CuentaBancariaDestino.DescripcionFormateada
        
            ElseIf IsSomething(mov.CuentaBancaria) Then
                Values(4) = mov.CuentaBancaria.DescripcionFormateada
        
            Else
                Values(4) = vbNullString
            End If
        
        Else
        
            If IsSomething(mov.CuentaBancaria) Then
                Values(4) = mov.CuentaBancaria.DescripcionFormateadaCompleta
            Else
                Values(4) = vbNullString
            End If
        
        End If

        If IsSomething(mov.CuentaContable) Then
            Values(5) = mov.CuentaContable.nombre
        Else
            Values(5) = vbNullString
        End If

        Values(6) = enums.enumEstadoMovimientosCajaYBancos(mov.estado)

        If IsSomething(mov.moneda) Then
            Values(7) = mov.moneda.NombreCorto
        Else
            Values(7) = vbNullString
        End If

        Values(8) = Replace(FormatCurrency(funciones.FormatearDecimales(mov.StaticTotalOrigenes)), "$", "")
        Values(9) = mov.Observaciones
        
        If IsSomething(mov.Usuario) Then
            Values(10) = mov.Usuario.Usuario
        Else
            Values(10) = vbNullString
        End If
        
        Values(11) = mov.Creada



End Sub


Private Property Get ISuscriber_id() As String
    ISuscriber_id = ids
End Property


Private Sub mnuAprobar_Click()
    SeleccionarOP
    
    If Not IsSomething(AsientoContable) Then
        MsgBox "Debe seleccionar un movimiento para aprobar.", vbExclamation, "Aprobación"
        Exit Sub
    End If
    
    If AsientoContable.estado <> EstadoMovimientoCajaYBancos.EnEdicion Then
        MsgBox "Solo se pueden aprobar movimientos que estén en edición.", vbExclamation, "Aprobación"
        Exit Sub
    End If
    
    Dim cuentaTxt As String
    
    cuentaTxt = vbNullString
    
    If AsientoContable.TipoMovimiento = "TRANSFERENCIA" Then
        
        If IsSomething(AsientoContable.CuentaBancaria) Then
            cuentaTxt = AsientoContable.CuentaBancaria.DescripcionFormateada
        End If
        
        If IsSomething(AsientoContable.CuentaBancariaDestino) Then
            cuentaTxt = cuentaTxt & "  ->  " & AsientoContable.CuentaBancariaDestino.DescripcionFormateada
        End If
        
    Else
        
        If IsSomething(AsientoContable.CuentaBancaria) Then
            cuentaTxt = AsientoContable.CuentaBancaria.DescripcionFormateada
        End If
        
    End If
    
    If MsgBox("Está por aprobar el movimiento Nº " & AsientoContable.Id & "." & vbCrLf & vbCrLf & _
              "Tipo: " & AsientoContable.TipoMovimiento & vbCrLf & _
              "Cuenta: " & cuentaTxt & vbCrLf & _
              "Importe: " & Replace(FormatCurrency(funciones.FormatearDecimales(AsientoContable.StaticTotalOrigenes)), "$", "") & vbCrLf & vbCrLf & _
              "Una vez aprobado, el movimiento no podrá editarse." & vbCrLf & _
              "¿Desea continuar?", _
              vbQuestion + vbYesNo + vbDefaultButton2, _
              "Confirmar aprobación") = vbNo Then
        Exit Sub
    End If
    
    If DAOAsientoContable.aprobar(AsientoContable, False) Then
        MsgBox "Aprobación exitosa!", vbInformation + vbOKOnly
        Me.gridOrdenes.RefreshRowIndex Me.gridOrdenes.RowIndex(Me.gridOrdenes.row)
        cmdBuscar_Click
    Else
        MsgBox "Error, no se aprobó el movimiento!", vbCritical + vbOKOnly
    End If

End Sub




Private Sub mnuImprimir_Click()

    On Error GoTo err4
    Me.CommonDialog.ShowPrinter

   If Not DAOAsientoContable.PrintMovimiento(AsientoContable) Then GoTo err4
   Exit Sub

err4:
 
End Sub


Private Sub mnuVer_Click()

    Dim movSeleccionado As clsAsientoContable

    Set movSeleccionado = ObtenerMovimientoSeleccionado()

    If Not IsSomething(movSeleccionado) Then

        MsgBox "Debe seleccionar un movimiento.", _
               vbExclamation

        Exit Sub

    End If

    AbrirMovimiento movSeleccionado, True

End Sub


Private Sub cboRangos_Click()
    funciones.CalculateDateRange Me.cboRangos, Me.dtpDesde(1), Me.dtpHasta(1)
End Sub


Private Sub SeleccionarOP()

    Set AsientoContable = ObtenerMovimientoSeleccionado()

End Sub


Private Sub AbrirMovimiento( _
    ByVal mov As clsAsientoContable, _
    ByVal modoSoloLectura As Boolean)

    If Not IsSomething(mov) Then

        MsgBox "Debe seleccionar un movimiento.", _
               vbExclamation

        Exit Sub

    End If

    Dim f22 As New frmAdminCajaBancosCrearAsientoBancario

    Load f22

    f22.ReadOnly = modoSoloLectura
    f22.Cargar mov
    f22.Show

End Sub


Private Sub mnuEditar_Click()

    Dim movSeleccionado As clsAsientoContable

    Set movSeleccionado = ObtenerMovimientoSeleccionado()

    If Not IsSomething(movSeleccionado) Then

        MsgBox "Debe seleccionar un movimiento.", _
               vbExclamation

        Exit Sub

    End If

    If movSeleccionado.estado <> _
            EstadoMovimientoCajaYBancos.EnEdicion Then

        MsgBox "Solo se pueden editar movimientos en edición.", _
               vbExclamation

        Exit Sub

    End If

    AbrirMovimiento movSeleccionado, False

End Sub

Private Function ObtenerMovimientoSeleccionado() _
    As clsAsientoContable

    On Error GoTo err1

    Dim idMovimiento As Long
    Dim mov As clsAsientoContable
    Dim valorId As Variant

    Set ObtenerMovimientoSeleccionado = Nothing

    If Movimientos.count = 0 Then Exit Function
    If Me.gridOrdenes.row <= 0 Then Exit Function

    valorId = Me.gridOrdenes.value(1)

    If IsNull(valorId) Or IsEmpty(valorId) Then Exit Function
    If Not IsNumeric(valorId) Then Exit Function

    idMovimiento = CLng(valorId)

    For Each mov In Movimientos

        If mov.Id = idMovimiento Then
            Set ObtenerMovimientoSeleccionado = mov
            Exit Function
        End If

    Next mov

    Exit Function

err1:
    Set ObtenerMovimientoSeleccionado = Nothing

End Function


Private Sub TotalizarMovimientos()

    Dim mov As clsAsientoContable
    Dim totalValores As Double

    totalValores = 0

    If Movimientos Is Nothing Then
        Me.Label4.caption = "Total: " & FormatCurrency(0)
        Exit Sub
    End If

    For Each mov In Movimientos

        If IsSomething(mov) Then
            totalValores = totalValores + mov.StaticTotalOrigenes
        End If

    Next mov

    Me.Label4.caption = _
        "Total: " & _
        FormatCurrency( _
            funciones.FormatearDecimales(totalValores))

End Sub


