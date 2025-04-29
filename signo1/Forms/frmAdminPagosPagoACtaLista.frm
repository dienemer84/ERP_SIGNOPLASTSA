VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmAdminPagosPagoACtaLista 
   Caption         =   "Pagos a cuenta"
   ClientHeight    =   8580
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   13725
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8580
   ScaleWidth      =   13725
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
      TabIndex        =   31
      Top             =   7920
      Visible         =   0   'False
      Width           =   615
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   15165
      _Version        =   786432
      _ExtentX        =   26749
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
         TabIndex        =   27
         Top             =   960
         Width           =   5055
         Begin XtremeSuiteControls.PushButton btnExportar 
            Height          =   450
            Left            =   1920
            TabIndex        =   28
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
            TabIndex        =   29
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
            TabIndex        =   30
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
         TabIndex        =   25
         Top             =   240
         Width           =   5055
         Begin XtremeSuiteControls.ProgressBar progreso 
            Height          =   375
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   4815
            _Version        =   786432
            _ExtentX        =   8493
            _ExtentY        =   661
            _StockProps     =   93
            Appearance      =   6
         End
      End
      Begin VB.TextBox txtNro 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   945
         TabIndex        =   5
         Top             =   285
         Width           =   2280
      End
      Begin XtremeSuiteControls.GroupBox GroFechaComprobante 
         Height          =   1215
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   1920
         Visible         =   0   'False
         Width           =   2355
         _Version        =   786432
         _ExtentX        =   4154
         _ExtentY        =   2143
         _StockProps     =   79
         Caption         =   "Estado Proveedor"
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.CheckBox chkContado 
            Height          =   195
            Left            =   405
            TabIndex        =   2
            Top             =   225
            Width           =   1635
            _Version        =   786432
            _ExtentX        =   2884
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Contado"
            UseVisualStyle  =   -1  'True
            Value           =   1
         End
         Begin XtremeSuiteControls.CheckBox chkCtaCte 
            Height          =   315
            Left            =   405
            TabIndex        =   3
            Top             =   465
            Width           =   1800
            _Version        =   786432
            _ExtentX        =   3175
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "Cuenta Corriente"
            UseVisualStyle  =   -1  'True
            Value           =   1
         End
         Begin XtremeSuiteControls.CheckBox chkEliminado 
            Height          =   315
            Left            =   405
            TabIndex        =   4
            Top             =   765
            Width           =   1800
            _Version        =   786432
            _ExtentX        =   3175
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "Inactivos"
            UseVisualStyle  =   -1  'True
            Value           =   1
         End
      End
      Begin XtremeSuiteControls.ComboBox cboProveedores 
         Height          =   315
         Left            =   945
         TabIndex        =   6
         Top             =   615
         Width           =   3525
         _Version        =   786432
         _ExtentX        =   6218
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Text            =   "cboProveedores"
      End
      Begin XtremeSuiteControls.PushButton btnClearProveedor 
         Height          =   255
         Left            =   4530
         TabIndex        =   7
         Top             =   630
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
      Begin XtremeSuiteControls.ComboBox cboEstado 
         Height          =   315
         Left            =   945
         TabIndex        =   10
         Top             =   960
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
         TabIndex        =   11
         Top             =   995
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
         TabIndex        =   12
         Top             =   240
         Width           =   4695
         _Version        =   786432
         _ExtentX        =   8281
         _ExtentY        =   2778
         _StockProps     =   79
         Caption         =   "Fecha Pago"
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
            TabIndex        =   13
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
            TabIndex        =   14
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
            TabIndex        =   15
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
            TabIndex        =   18
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
            TabIndex        =   17
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
            TabIndex        =   16
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
         Left            =   240
         TabIndex        =   23
         Top             =   330
         Width           =   600
         _Version        =   786432
         _ExtentX        =   1058
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Nº Pago"
         Alignment       =   1
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label lbl 
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   660
         Width           =   750
         _Version        =   786432
         _ExtentX        =   1323
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Proveedor"
         Alignment       =   1
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblDesde 
         Height          =   195
         Index           =   0
         Left            =   2865
         TabIndex        =   21
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
         TabIndex        =   20
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
         Left            =   360
         TabIndex        =   19
         Top             =   1020
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
   Begin GridEX20.GridEX gridOrdenes 
      Height          =   5505
      Left            =   120
      TabIndex        =   24
      Top             =   2160
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
      ColumnsCount    =   7
      Column(1)       =   "frmAdminPagosPagoACtaLista.frx":0000
      Column(2)       =   "frmAdminPagosPagoACtaLista.frx":0188
      Column(3)       =   "frmAdminPagosPagoACtaLista.frx":02B0
      Column(4)       =   "frmAdminPagosPagoACtaLista.frx":03F0
      Column(5)       =   "frmAdminPagosPagoACtaLista.frx":0538
      Column(6)       =   "frmAdminPagosPagoACtaLista.frx":0678
      Column(7)       =   "frmAdminPagosPagoACtaLista.frx":07C0
      FormatStylesCount=   11
      FormatStyle(1)  =   "frmAdminPagosPagoACtaLista.frx":0908
      FormatStyle(2)  =   "frmAdminPagosPagoACtaLista.frx":0A30
      FormatStyle(3)  =   "frmAdminPagosPagoACtaLista.frx":0AE0
      FormatStyle(4)  =   "frmAdminPagosPagoACtaLista.frx":0B94
      FormatStyle(5)  =   "frmAdminPagosPagoACtaLista.frx":0C6C
      FormatStyle(6)  =   "frmAdminPagosPagoACtaLista.frx":0D24
      FormatStyle(7)  =   "frmAdminPagosPagoACtaLista.frx":0E04
      FormatStyle(8)  =   "frmAdminPagosPagoACtaLista.frx":0E24
      FormatStyle(9)  =   "frmAdminPagosPagoACtaLista.frx":0ED8
      FormatStyle(10) =   "frmAdminPagosPagoACtaLista.frx":0F90
      FormatStyle(11) =   "frmAdminPagosPagoACtaLista.frx":1044
      ImageCount      =   0
      PrinterProperties=   "frmAdminPagosPagoACtaLista.frx":1100
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   840
      Top             =   8160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Flags           =   1
   End
   Begin VB.Menu menu 
      Caption         =   "menu"
      Begin VB.Menu mnuEditar 
         Caption         =   "Editar"
      End
      Begin VB.Menu mnuAprobar 
         Caption         =   "Aprobar"
         Visible         =   0   'False
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
Attribute VB_Name = "frmAdminPagosPagoACtaLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private desde
Dim ids As String
Private pagosacuenta As New Collection
Private PagoACuenta As clsPagoACta
Dim i As Integer


Private Sub btnClearProveedor_Click()
    Me.cboProveedores.ListIndex = -1
End Sub


Private Sub btnExportar_Click()

    Me.progreso.Visible = True

    If IsSomething(pagosacuenta) Then
        If Not DAOPagoACta.ExportarColeccion(pagosacuenta, Me.progreso) Then GoTo err1
    End If

    Me.progreso.Visible = False

    Exit Sub
err1:
    MsgBox "Se produjo un error al exportar!", vbCritical, "Error"

End Sub


Private Sub cmdBuscar_Click()
    If (Me.chkContado.Value = xtpChecked Or Me.chkCtaCte.Value = xtpChecked Or Me.chkEliminado.Value = xtpGrayed) Then llenarLista Else Me.gridOrdenes.ItemCount = 0

End Sub


Private Sub cmdImprimir_Click()

    Dim pro As String
    If Me.cboProveedores.ListIndex > -1 Then
        pro = " Proveedor: " & Me.cboProveedores.Text
    End If

    With Me.gridOrdenes.PrinterProperties

        .FitColumns = True
        .RepeatHeaders = True
        .Orientation = jgexPPLandscape
        .HeaderString(jgexHFCenter) = "Listado de Pagos a cuenta"
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
    DAOProveedor.llenarComboXtremeSuite Me.cboProveedores, True, True, True
    
    Me.cboProveedores.ListIndex = -1
    '    llenarLista

    Me.dtpHasta(1).Value = Now
    
    Me.gridOrdenes.ItemCount = 0
    GridEXHelper.AutoSizeColumns Me.gridOrdenes
    ids = funciones.CreateGUID

    Me.cboEstado.Clear
    Me.cboEstado.AddItem enums.enumEstadoPagoACuenta(EstadoPagoACuenta.Disponible)
    Me.cboEstado.ItemData(Me.cboEstado.NewIndex) = EstadoPagoACuenta.Disponible
    Me.cboEstado.AddItem enums.enumEstadoPagoACuenta(EstadoPagoACuenta.Procesada)
    Me.cboEstado.ItemData(Me.cboEstado.NewIndex) = EstadoPagoACuenta.Procesada

    Me.dtpDesde(1).Value = Year(Now) & "-01-01"

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

    If Me.cboProveedores.ListIndex > -1 Then
        filter = filter & " AND proveedores.id = " & Me.cboProveedores.ItemData(Me.cboProveedores.ListIndex)
    End If

    If LenB(Me.txtNro.Text) > 0 Then
        filter = filter & " AND  pagos_a_cuenta.id LIKE '%" & Val(Me.txtNro.Text) & "%'"
    End If

    Dim filtroor As String

    If Not IsNull(Me.dtpDesde(1).Value) Then
        filter = filter & " AND pagos_a_cuenta.fecha >= " & conectar.Escape(Me.dtpDesde(1).Value)
    End If

    If Not IsNull(Me.dtpHasta(1).Value) Then
        filter = filter & " AND pagos_a_cuenta.fecha <= " & conectar.Escape(Me.dtpHasta(1).Value)
    End If


    If Me.chkContado.Value = xtpChecked Then
        filtroor = filtroor & " OR proveedores.estado = " & EstadoProveedor.EstadoProveedorContado
    End If

    If Me.chkCtaCte.Value = xtpChecked Then
        filtroor = filtroor & " OR proveedores.estado = " & EstadoProveedor.EstadoProveedorCuentaCorriente
    End If

    If Me.chkEliminado.Value = xtpChecked Then
        filtroor = filtroor & " OR proveedores.estado = " & EstadoProveedor.EstadoProveedorEliminado
    End If


    If Me.cboEstado.ListIndex > -1 Then
        filter = filter & " AND pagos_a_cuenta.estado = " & Me.cboEstado.ItemData(Me.cboEstado.ListIndex)
    End If


    If LenB(filtroor) > 0 Then
        filtroor = " AND (" & Right(filtroor, Len(filtroor) - 3) & " )"
        filter = filter & filtroor
    End If

    Me.gridOrdenes.ItemCount = 0
    
    Set pagosacuenta = DAOPagoACta.FindAll(filter, "pagos_a_cuenta.id DESC")
    
    Me.gridOrdenes.ItemCount = pagosacuenta.count

    Me.caption = "Listado de Pagos a cta" & " [Cant: " & pagosacuenta.count & "]"


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
    mnuVer_Click
End Sub


Private Sub gridOrdenes_SelectionChange()
    SeleccionarOP
End Sub


Private Sub SeleccionarOP()
    On Error Resume Next
    Set PagoACuenta = pagosacuenta.item(gridOrdenes.rowIndex(gridOrdenes.row))

End Sub


Private Sub gridOrdenes_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If pagosacuenta.count > 0 Then
        gridOrdenes_SelectionChange
        If Button = 2 Then
            Me.mnuEditar.Enabled = (PagoACuenta.estado = EstadoPagoACuenta.Disponible)
            Me.mnuAprobar.Enabled = (PagoACuenta.estado = EstadoPagoACuenta.Procesada)
            Me.mnuAnular.Enabled = Not (PagoACuenta.estado = Disponible)
            Me.mnuVer.Enabled = PagoACuenta.estado = Disponible Or PagoACuenta.estado = Procesada

            Me.PopupMenu menu
        End If
    End If
End Sub


Private Sub gridOrdenes_RowFormat(RowBuffer As GridEX20.JSRowData)
    If RowBuffer.rowIndex > 0 And pagosacuenta.count > 0 Then
        Set PagoACuenta = pagosacuenta.item(RowBuffer.rowIndex)
        If PagoACuenta.estado = EstadoPagoACuenta.Disponible Then
            RowBuffer.CellStyle(6) = "pendiente"
        ElseIf PagoACuenta.estado = EstadoPagoACuenta.Disponible Then
            RowBuffer.RowStyle = "anulada2"

            RowBuffer.CellStyle(6) = "anulada"
        ElseIf PagoACuenta.estado = EstadoPagoACuenta.Procesada Then
            RowBuffer.CellStyle(6) = "aprobada"
        End If
    End If
End Sub


Private Sub gridOrdenes_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If rowIndex > 0 And pagosacuenta.count > 0 Then
        Set PagoACuenta = pagosacuenta.item(rowIndex)
        Values(1) = PagoACuenta.Id
        Values(2) = PagoACuenta.Proveedor.RazonSocial
        Values(3) = PagoACuenta.FEcha
        Values(4) = PagoACuenta.moneda.NombreCorto
        Values(5) = Replace(FormatCurrency(funciones.FormatearDecimales(PagoACuenta.StaticTotalOrigenes)), "$", "")
        Values(6) = enums.enumEstadoPagoACuenta(PagoACuenta.estado)
        Values(7) = PagoACuenta.Creada
        
    End If
End Sub


Private Property Get ISuscriber_id() As String
    ISuscriber_id = ids
End Property



Private Sub mnuAnular_Click()
    SeleccionarOP
    
    If MsgBox("Â¿Desea anular la OP?", vbQuestion + vbYesNo) = vbYes Then
        If DAOPagoACta.Delete(PagoACuenta.Id, True) Then
            MsgBox "Anulación éxitosa.", vbInformation + vbOKOnly
            Me.gridOrdenes.ItemCount = 0
            ordenes.remove CStr(Orden.Id)
            Me.gridOrdenes.ItemCount = ordenes.count
            cmdBuscar_Click
        Else
            MsgBox "No se pudo borrar el Pago a cuenta.", vbCritical + vbOKOnly
        End If
    End If
End Sub


Private Sub mnuAprobar_Click()
    SeleccionarOP
    
    If DAOPagoACta.aprobar(PagoACuenta, True) Then
        MsgBox "Aprobación éxitosa!", vbInformation + vbOKOnly
        Me.gridOrdenes.RefreshRowIndex Me.gridOrdenes.rowIndex(Me.gridOrdenes.row)
        cmdBuscar_Click
    Else
        MsgBox "Error, no se aprobó el Pago a cuenta!", vbCritical + vbOKOnly
    End If

End Sub


Private Sub mnuEditar_Click()
    SeleccionarOP
    
    Dim f22 As New frmAdminPagosCrearPagoACta
    f22.Show
    f22.Cargar PagoACuenta
End Sub


Private Sub mnuHistorial_Click()
    SeleccionarOP
    
    Dim F As New frmHistorico
    F.Configurar "orden_pago_historial", PagoACuenta.Id, "pago a cuenta Nro " & PagoACuenta.Id
    F.Show
    
End Sub

Private Sub mnuImprimir_Click()

    On Error GoTo err4
    Me.CommonDialog.ShowPrinter

   If Not DAOPagoACta.PrintPCTA(PagoACuenta) Then GoTo err4
   Exit Sub

err4:
 
End Sub


Private Sub mnuVer_Click()
    Dim f22 As New frmAdminPagosCrearPagoACta
    f22.Show
    SeleccionarOP

    f22.ReadOnly = True

    f22.Cargar PagoACuenta

    
End Sub



Private Sub cboRangos_Click()
    funciones.CalculateDateRange Me.cboRangos, Me.dtpDesde(1), Me.dtpHasta(1)
End Sub



