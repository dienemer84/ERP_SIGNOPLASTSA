VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmAdminPagosOrdenesPagoLista 
   Caption         =   "Ordenes de Pago"
   ClientHeight    =   9105
   ClientLeft      =   8445
   ClientTop       =   3465
   ClientWidth     =   12885
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOrdenesPago.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9105
   ScaleWidth      =   12885
   WindowState     =   2  'Maximized
   Begin VB.PictureBox pic 
      Height          =   495
      Left            =   240
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   31
      Top             =   8400
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Index           =   1
      Left            =   9960
      TabIndex        =   29
      Top             =   360
      Width           =   5055
      Begin XtremeSuiteControls.ProgressBar progreso 
         Height          =   375
         Left            =   120
         TabIndex        =   30
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
      TabIndex        =   18
      Top             =   1080
      Width           =   5055
      Begin XtremeSuiteControls.PushButton btnExportar 
         Height          =   450
         Left            =   1920
         TabIndex        =   19
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
         Height          =   450
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1350
         _Version        =   786432
         _ExtentX        =   2381
         _ExtentY        =   794
         _StockProps     =   79
         Caption         =   "Buscar"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton cmdImprimir 
         Height          =   450
         Left            =   3600
         TabIndex        =   21
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
   Begin GridEX20.GridEX gridOrdenes 
      Height          =   5505
      Left            =   120
      TabIndex        =   0
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
      ColumnsCount    =   9
      Column(1)       =   "frmOrdenesPago.frx":000C
      Column(2)       =   "frmOrdenesPago.frx":01A4
      Column(3)       =   "frmOrdenesPago.frx":0304
      Column(4)       =   "frmOrdenesPago.frx":044C
      Column(5)       =   "frmOrdenesPago.frx":05B0
      Column(6)       =   "frmOrdenesPago.frx":07E4
      Column(7)       =   "frmOrdenesPago.frx":0944
      Column(8)       =   "frmOrdenesPago.frx":0A84
      Column(9)       =   "frmOrdenesPago.frx":0BA4
      FormatStylesCount=   10
      FormatStyle(1)  =   "frmOrdenesPago.frx":0CEC
      FormatStyle(2)  =   "frmOrdenesPago.frx":0E14
      FormatStyle(3)  =   "frmOrdenesPago.frx":0EC4
      FormatStyle(4)  =   "frmOrdenesPago.frx":0F78
      FormatStyle(5)  =   "frmOrdenesPago.frx":1050
      FormatStyle(6)  =   "frmOrdenesPago.frx":1108
      FormatStyle(7)  =   "frmOrdenesPago.frx":11E8
      FormatStyle(8)  =   "frmOrdenesPago.frx":129C
      FormatStyle(9)  =   "frmOrdenesPago.frx":1350
      FormatStyle(10) =   "frmOrdenesPago.frx":1430
      ImageCount      =   0
      PrinterProperties=   "frmOrdenesPago.frx":14E8
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1935
      Left            =   45
      TabIndex        =   1
      Top             =   120
      Width           =   14805
      _Version        =   786432
      _ExtentX        =   26114
      _ExtentY        =   3413
      _StockProps     =   79
      Caption         =   "Parámetros de búsqueda"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.GroupBox GroFechaComprobante 
         Height          =   1215
         Index           =   0
         Left            =   240
         TabIndex        =   7
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
            TabIndex        =   8
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
            TabIndex        =   9
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
            TabIndex        =   10
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
      Begin VB.TextBox txtNro 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   945
         TabIndex        =   2
         Top             =   285
         Width           =   840
      End
      Begin XtremeSuiteControls.ComboBox cboProveedores 
         Height          =   315
         Left            =   945
         TabIndex        =   3
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
         TabIndex        =   4
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
         TabIndex        =   15
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
         TabIndex        =   17
         Top             =   990
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
         TabIndex        =   22
         Top             =   240
         Width           =   4695
         _Version        =   786432
         _ExtentX        =   8281
         _ExtentY        =   2778
         _StockProps     =   79
         Caption         =   "Fecha OP"
         BackColor       =   16744576
         Appearance      =   4
         Begin XtremeSuiteControls.DateTimePicker dtpDesde 
            Height          =   315
            Index           =   1
            Left            =   720
            TabIndex        =   23
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
            TabIndex        =   24
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
            TabIndex        =   25
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
            TabIndex        =   28
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
            TabIndex        =   27
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
            TabIndex        =   26
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
      Begin XtremeSuiteControls.Label lblRango 
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   16
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
      Begin XtremeSuiteControls.Label lblHasta 
         Height          =   195
         Index           =   0
         Left            =   2880
         TabIndex        =   14
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
         TabIndex        =   13
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
         Left            =   165
         TabIndex        =   6
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   195
         Left            =   195
         TabIndex        =   5
         Top             =   330
         Width           =   675
         _Version        =   786432
         _ExtentX        =   1191
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Nº Orden"
         Alignment       =   1
         AutoSize        =   -1  'True
      End
   End
   Begin VB.Menu menu 
      Caption         =   "menu"
      Visible         =   0   'False
      Begin VB.Menu mnuEditar 
         Caption         =   "Editar"
      End
      Begin VB.Menu mnuAprobar 
         Caption         =   "Aprobar"
      End
      Begin VB.Menu mnuAnular 
         Caption         =   "Anular"
      End
      Begin VB.Menu mnuVer 
         Caption         =   "Ver"
      End
      Begin VB.Menu mnuImprimir 
         Caption         =   "Imprimir"
      End
      Begin VB.Menu mnuVerCertificado 
         Caption         =   "Ver Certificado IIBB"
      End
      Begin VB.Menu nada 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHistorial 
         Caption         =   "Ver Historial"
      End
   End
End
Attribute VB_Name = "frmAdminPagosOrdenesPagoLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Implements ISuscriber
Private desde
Dim ids As String
Private ordenes As New Collection
Private Orden As OrdenPago
Private fac As clsFacturaProveedor
    Dim i As Integer

Private Sub btnClearProveedor_Click()
    Me.cboProveedores.ListIndex = -1
End Sub

Private Sub btnExportar_Click()

    Me.progreso.Visible = True

    If IsSomething(ordenes) Then
        If Not DAOOrdenPago.ExportarColeccion(ordenes, Me.progreso) Then GoTo err1
    End If

    Me.progreso.Visible = False

    Exit Sub
err1:
    MsgBox "Se produjo un error al exportar!", vbCritical, "Error"

End Sub

Private Sub cmdBuscar_Click()
    If (Me.chkContado.value = xtpChecked Or Me.chkCtaCte.value = xtpChecked Or Me.chkEliminado.value = xtpGrayed) Then llenarLista Else Me.gridOrdenes.ItemCount = 0

End Sub

Private Sub cmdImprimir_Click()

    Dim pro As String
    If Me.cboProveedores.ListIndex > -1 Then
        pro = " Proveedor: " & Me.cboProveedores.text
    End If

    With Me.gridOrdenes.PrinterProperties

        .FitColumns = True
        .RepeatHeaders = True
        .Orientation = jgexPPLandscape
        .HeaderString(jgexHFCenter) = "Listado de Ordenes de Pago "
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

    Me.dtpHasta(1).value = Now
    
    Me.gridOrdenes.ItemCount = 0
    GridEXHelper.AutoSizeColumns Me.gridOrdenes
    ids = funciones.CreateGUID
    Channel.AgregarSuscriptor Me, OrdenesPago_

    Me.cboEstado.Clear
    Me.cboEstado.AddItem enums.EnumEstadoOrdenPago(EstadoOrdenPago.EstadoOrdenPago_pendiente)
    Me.cboEstado.ItemData(Me.cboEstado.NewIndex) = EstadoOrdenPago.EstadoOrdenPago_pendiente
    Me.cboEstado.AddItem enums.EnumEstadoOrdenPago(EstadoOrdenPago.EstadoOrdenPago_Aprobada)
    Me.cboEstado.ItemData(Me.cboEstado.NewIndex) = EstadoOrdenPago.EstadoOrdenPago_Aprobada
    Me.cboEstado.AddItem enums.EnumEstadoOrdenPago(EstadoOrdenPago.EstadoOrdenPago_Anulada)
    Me.cboEstado.ItemData(Me.cboEstado.NewIndex) = EstadoOrdenPago.EstadoOrdenPago_Anulada

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

    If Me.cboProveedores.ListIndex > -1 Then
        filter = filter & " AND AdminComprasFacturasProveedores.id_proveedor = " & Me.cboProveedores.ItemData(Me.cboProveedores.ListIndex)
    End If

    If LenB(Me.txtNro.text) > 0 Then
        filter = filter & " AND  ordenes_pago.id  = " & Val(Me.txtNro.text)
    End If

    Dim filtroor As String

    If Not IsNull(Me.dtpDesde(1).value) Then
        filter = filter & " AND ordenes_pago.fecha >= " & conectar.Escape(Me.dtpDesde(1).value)
    End If

    If Not IsNull(Me.dtpHasta(1).value) Then
        filter = filter & " AND ordenes_pago.fecha <= " & conectar.Escape(Me.dtpHasta(1).value)
    End If


    If Me.chkContado.value = xtpChecked Then
        filtroor = filtroor & " OR proveedores.estado = " & EstadoProveedor.EstadoProveedorContado
    End If

    If Me.chkCtaCte.value = xtpChecked Then
        filtroor = filtroor & " OR proveedores.estado = " & EstadoProveedor.EstadoProveedorCuentaCorriente
    End If

    If Me.chkEliminado.value = xtpChecked Then
        filtroor = filtroor & " OR proveedores.estado = " & EstadoProveedor.EstadoProveedorEliminado
    End If


    If Me.cboEstado.ListIndex > -1 Then
        filter = filter & " AND ordenes_pago.estado = " & Me.cboEstado.ItemData(Me.cboEstado.ListIndex)
    End If


    If LenB(filtroor) > 0 Then
        filtroor = " AND (" & Right(filtroor, Len(filtroor) - 3) & " )"
        filter = filter & filtroor
    End If

    Me.gridOrdenes.ItemCount = 0
    
    Set ordenes = DAOOrdenPago.FindAll(filter, "ordenes_pago.id DESC")
    
    Me.gridOrdenes.ItemCount = ordenes.count

    Me.caption = "Listado de OP" & " [Cant: " & ordenes.count & "]"


End Sub
Private Sub Form_Resize()
    On Error Resume Next
    Me.gridOrdenes.Width = Me.ScaleWidth - 300
    Me.gridOrdenes.Height = (Me.ScaleHeight * 75) / 100

    Me.GroupBox1.Width = Me.gridOrdenes.Width
    GridEXHelper.AutoSizeColumns Me.gridOrdenes
    
End Sub

Private Sub Form_Terminate()
    Channel.RemoverSuscripcionTotal Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Channel.RemoverSuscripcionTotal Me
    
    
End Sub


Private Sub gridOrdenes_ColumnHeaderClick(ByVal Column As GridEX20.JSColumn)
    GridEXHelper.ColumnHeaderClick Me.gridOrdenes, Column
End Sub


'Private Sub gridOrdenes_Click()
'    gridOrdenes_SelectionChange
'End Sub


Private Sub gridOrdenes_DblClick()
'    gridOrdenes_SelectionChange
    mnuVer_Click
End Sub


Private Sub gridOrdenes_SelectionChange()
    SeleccionarOP
End Sub


Private Sub SeleccionarOP()
    On Error Resume Next
    Set Orden = ordenes.item(gridOrdenes.rowIndex(gridOrdenes.row))

End Sub


Private Sub gridOrdenes_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If ordenes.count > 0 Then
        gridOrdenes_SelectionChange
        If Button = 2 Then
            Me.mnuVerCertificado.Enabled = Orden.EsParaFacturaProveedor And (Orden.estado = EstadoOrdenPago_Aprobada)
            Me.mnuEditar.Enabled = (Orden.estado = EstadoOrdenPago_pendiente)
            Me.mnuAprobar.Enabled = (Orden.estado = EstadoOrdenPago_pendiente)
            Me.mnuAnular.Enabled = Not (Orden.estado = EstadoOrdenPago_Anulada)
            Me.mnuVer.Enabled = Not (Orden.estado = EstadoOrdenPago_Anulada)

            Me.PopupMenu menu
        End If
    End If
End Sub

Private Sub gridOrdenes_RowFormat(RowBuffer As GridEX20.JSRowData)
    If RowBuffer.rowIndex > 0 And ordenes.count > 0 Then
        Set Orden = ordenes.item(RowBuffer.rowIndex)
        If Orden.estado = EstadoOrdenPago.EstadoOrdenPago_Aprobada Then
            RowBuffer.CellStyle(9) = "aprobada"
        ElseIf Orden.estado = EstadoOrdenPago_Anulada Then
            RowBuffer.RowStyle = "anulada2"

            RowBuffer.CellStyle(9) = "anulada"
        ElseIf Orden.estado = EstadoOrdenPago_pendiente Then
            RowBuffer.CellStyle(9) = "pendiente"
        End If
    End If
End Sub


Private Sub gridOrdenes_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If rowIndex > 0 And ordenes.count > 0 Then
        Set Orden = ordenes.item(rowIndex)
        Values(1) = Orden.Id
        Values(2) = Orden.FEcha

        Values(3) = Orden.moneda.NombreCorto

        Values(4) = Replace(FormatCurrency(funciones.FormatearDecimales(Orden.StaticTotalOrigenes)), "$", "")
        Values(5) = Replace(FormatCurrency(funciones.FormatearDecimales(Orden.StaticTotalRetenido)), "$", "")
        Values(6) = Replace(FormatCurrency(funciones.FormatearDecimales(Orden.StaticTotalOrigenes + Orden.StaticTotalRetenido)), "$", "")

        If Orden.EsParaFacturaProveedor Then
            Set fac = Orden.FacturasProveedor.item(1)
            Values(7) = "Factura Proveedor"
            Values(8) = fac.Proveedor.RazonSocial
        Else
            Values(7) = "Cuenta Contable"
            If IsSomething(Orden.CuentaContable) Then
                Values(8) = Orden.CuentaContable.nombre & " (" & Orden.CuentaContable.codigo & ")"
            End If
        End If

        Values(9) = enums.EnumEstadoOrdenPago(Orden.estado)
    End If
End Sub

Private Property Get ISuscriber_id() As String
    ISuscriber_id = ids
End Property

Private Function ISuscriber_Notificarse(EVENTO As clsEventoObserver) As Variant
    Dim tmp As OrdenPago
    Dim i As Long

    If EVENTO.EVENTO = agregar_ Then
        ordenes.Add EVENTO.Elemento
        llenarLista
    ElseIf EVENTO.EVENTO = modificar_ Then
        For i = ordenes.count To 1 Step -1
            Set tmp = EVENTO.Elemento
            If ordenes(i).Id = tmp.Id Then
                Set Orden = ordenes(i)
                Orden.Id = tmp.Id
                Orden.estado = tmp.estado
                Me.gridOrdenes.RefreshRowIndex i
                Exit For
            End If
        Next
    End If
End Function

Private Sub mnuAnular_Click()
    SeleccionarOP
    
    If MsgBox("¿Desea anular la OP?", vbQuestion + vbYesNo) = vbYes Then
        If DAOOrdenPago.Delete(Orden.Id, True) Then
            MsgBox "Anulación Exitosa.", vbInformation + vbOKOnly
            Me.gridOrdenes.ItemCount = 0
            ordenes.remove CStr(Orden.Id)
            Me.gridOrdenes.ItemCount = ordenes.count
            cmdBuscar_Click
        Else
            MsgBox "No se pudo borrar.", vbCritical + vbOKOnly
        End If
    End If
End Sub

Private Sub mnuAprobar_Click()

    SeleccionarOP
    
    If DAOOrdenPago.aprobar(Orden, True) Then
        MsgBox "Aprobación Exitosa!", vbInformation + vbOKOnly
        Me.gridOrdenes.RefreshRowIndex Me.gridOrdenes.rowIndex(Me.gridOrdenes.row)
        cmdBuscar_Click
    Else
        MsgBox "Error, no se aprobó la OP!", vbCritical + vbOKOnly
    End If

End Sub

Private Sub mnuEditar_Click()
    SeleccionarOP
    
    Dim f22 As New frmAdminPagosCrearOrdenPago
    f22.Show
    f22.Cargar Orden
End Sub

Private Sub mnuHistorial_Click()
    SeleccionarOP
    
    Dim F As New frmHistorico
    F.Configurar "orden_pago_historial", Orden.Id, "orden de pago Nro " & Orden.Id
    F.Show
    
End Sub

Private Sub mnuImprimir_Click()
    Dim dlg As Object
    Set dlg = CreateObject("MSComDlg.CommonDialog")
    
    dlg.ShowPrinter
    
    If Not DAOOrdenPago.PrintOP(Orden, Me.pic) Then GoTo err1

    ' Limpia el objeto dlg
    Set dlg = Nothing
    
        Exit Sub
err1:
End Sub


Private Sub mnuVer_Click()

    

    Dim f22 As New frmAdminPagosCrearOrdenPago
    f22.Show
    SeleccionarOP

    f22.ReadOnly = True

    f22.Cargar Orden
 
    
End Sub

Private Sub mnuVerCertificado_Click()
    Dim cr As Collection
    Set cr = DAOCertificadoRetencion.FindAllByOrdenPago(Orden.Id)

    If IsSomething(cr) Then

        Dim c As CertificadoRetencion
        For Each c In cr
            DAOCertificadoRetencion.VerCertificado c
        Next
    Else
        MsgBox "La orden de pago no tiene certificado.", vbInformation
    End If
End Sub


Private Sub cboRangos_Click()
    funciones.CalculateDateRange Me.cboRangos, Me.dtpDesde(1), Me.dtpHasta(1)
End Sub


