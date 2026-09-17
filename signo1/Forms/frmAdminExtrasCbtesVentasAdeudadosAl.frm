VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmAdminExtrasCbtesVentasAdeudadosAl 
   Caption         =   "Comprobantes Ventas adeudados al"
   ClientHeight    =   6645
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14415
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6645
   ScaleWidth      =   14415
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.GroupBox GroupBox 
      Height          =   2160
      Index           =   2
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   18570
      _Version        =   786432
      _ExtentX        =   32755
      _ExtentY        =   3810
      _StockProps     =   79
      Caption         =   "Comprobantes de clientes"
      BackColor       =   16744576
      UseVisualStyle  =   -1  'True
      Begin VB.TextBox txtNroFactura 
         Height          =   315
         Left            =   1440
         TabIndex        =   1
         Top             =   1125
         Width           =   3885
      End
      Begin XtremeSuiteControls.ComboBox cboClientes 
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
         Text            =   "cboClientes"
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
            TabIndex        =   7
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
         Begin XtremeSuiteControls.Label Label6 
            Height          =   195
            Index           =   0
            Left            =   2400
            TabIndex        =   10
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
         Begin XtremeSuiteControls.Label Label5 
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   9
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
         Begin XtremeSuiteControls.Label Label4 
            Height          =   195
            Left            =   240
            TabIndex        =   8
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
      End
      Begin XtremeSuiteControls.GroupBox gbBotones 
         Height          =   735
         Left            =   10920
         TabIndex        =   11
         Top             =   1200
         Width           =   3255
         _Version        =   786432
         _ExtentX        =   5741
         _ExtentY        =   1296
         _StockProps     =   79
         BackColor       =   16744576
         Appearance      =   4
         Begin XtremeSuiteControls.PushButton btnBuscar 
            Default         =   -1  'True
            Height          =   390
            Left            =   240
            TabIndex        =   12
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
            Left            =   1800
            TabIndex        =   13
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
         TabIndex        =   14
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
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   3855
            _Version        =   786432
            _ExtentX        =   6800
            _ExtentY        =   661
            _StockProps     =   93
            Appearance      =   6
         End
      End
      Begin XtremeSuiteControls.DateTimePicker dtpHastaFIN 
         Height          =   315
         Index           =   1
         Left            =   3840
         TabIndex        =   16
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
         TabIndex        =   21
         Top             =   240
         Width           =   7455
         _Version        =   786432
         _ExtentX        =   13150
         _ExtentY        =   1508
         _StockProps     =   79
         BackColor       =   16744576
         Appearance      =   4
         Begin VB.Label lblTotalNuevoSaldo 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Caption         =   "$ 00,00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   4320
            TabIndex        =   25
            Top             =   240
            Width           =   1740
         End
         Begin VB.Label lblTotalNuevoSaldo 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Nuevo Saldo:"
            Height          =   195
            Index           =   0
            Left            =   3360
            TabIndex        =   24
            Top             =   240
            Width           =   975
         End
         Begin XtremeSuiteControls.Label lbl 
            Height          =   195
            Index           =   2
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
         Begin XtremeSuiteControls.Label lblTotalFiltrado 
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Width           =   960
            _Version        =   786432
            _ExtentX        =   1693
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Total Filtrado:"
            BackColor       =   12632256
            AutoSize        =   -1  'True
         End
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   195
         Index           =   0
         Left            =   165
         TabIndex        =   19
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   195
         Left            =   840
         TabIndex        =   18
         Top             =   780
         Width           =   480
         _Version        =   786432
         _ExtentX        =   847
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Cliente"
         BackColor       =   12632256
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   195
         Index           =   1
         Left            =   1800
         TabIndex        =   17
         Top             =   1620
         Width           =   2010
         _Version        =   786432
         _ExtentX        =   3545
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Pagos/Cancelaciones hasta"
         BackColor       =   12632256
         AutoSize        =   -1  'True
      End
   End
   Begin GridEX20.GridEX grilla 
      Height          =   3975
      Left            =   120
      TabIndex        =   20
      Top             =   2400
      Width           =   18570
      _ExtentX        =   32755
      _ExtentY        =   7011
      Version         =   "2.0"
      PreviewRowIndent=   100
      AutomaticSort   =   -1  'True
      HoldSortSettings=   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      GroupFooterStyle=   2
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
      ImagePicture1   =   "frmAdminExtrasCbtesVentasAdeudadosAl.frx":0000
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   18
      Column(1)       =   "frmAdminExtrasCbtesVentasAdeudadosAl.frx":031A
      Column(2)       =   "frmAdminExtrasCbtesVentasAdeudadosAl.frx":0476
      Column(3)       =   "frmAdminExtrasCbtesVentasAdeudadosAl.frx":05A6
      Column(4)       =   "frmAdminExtrasCbtesVentasAdeudadosAl.frx":06E6
      Column(5)       =   "frmAdminExtrasCbtesVentasAdeudadosAl.frx":0836
      Column(6)       =   "frmAdminExtrasCbtesVentasAdeudadosAl.frx":09B6
      Column(7)       =   "frmAdminExtrasCbtesVentasAdeudadosAl.frx":0AFE
      Column(8)       =   "frmAdminExtrasCbtesVentasAdeudadosAl.frx":0C3E
      Column(9)       =   "frmAdminExtrasCbtesVentasAdeudadosAl.frx":0D86
      Column(10)      =   "frmAdminExtrasCbtesVentasAdeudadosAl.frx":0F6E
      Column(11)      =   "frmAdminExtrasCbtesVentasAdeudadosAl.frx":115E
      Column(12)      =   "frmAdminExtrasCbtesVentasAdeudadosAl.frx":129E
      Column(13)      =   "frmAdminExtrasCbtesVentasAdeudadosAl.frx":13E6
      Column(14)      =   "frmAdminExtrasCbtesVentasAdeudadosAl.frx":1566
      Column(15)      =   "frmAdminExtrasCbtesVentasAdeudadosAl.frx":16E6
      Column(16)      =   "frmAdminExtrasCbtesVentasAdeudadosAl.frx":1852
      Column(17)      =   "frmAdminExtrasCbtesVentasAdeudadosAl.frx":19BE
      Column(18)      =   "frmAdminExtrasCbtesVentasAdeudadosAl.frx":1BD6
      FormatStylesCount=   10
      FormatStyle(1)  =   "frmAdminExtrasCbtesVentasAdeudadosAl.frx":1CEA
      FormatStyle(2)  =   "frmAdminExtrasCbtesVentasAdeudadosAl.frx":1E22
      FormatStyle(3)  =   "frmAdminExtrasCbtesVentasAdeudadosAl.frx":1ED2
      FormatStyle(4)  =   "frmAdminExtrasCbtesVentasAdeudadosAl.frx":1F86
      FormatStyle(5)  =   "frmAdminExtrasCbtesVentasAdeudadosAl.frx":205E
      FormatStyle(6)  =   "frmAdminExtrasCbtesVentasAdeudadosAl.frx":2116
      FormatStyle(7)  =   "frmAdminExtrasCbtesVentasAdeudadosAl.frx":21F6
      FormatStyle(8)  =   "frmAdminExtrasCbtesVentasAdeudadosAl.frx":22B6
      FormatStyle(9)  =   "frmAdminExtrasCbtesVentasAdeudadosAl.frx":237A
      FormatStyle(10) =   "frmAdminExtrasCbtesVentasAdeudadosAl.frx":243A
      ImageCount      =   1
      ImagePicture(1) =   "frmAdminExtrasCbtesVentasAdeudadosAl.frx":24BE
      PrinterProperties=   "frmAdminExtrasCbtesVentasAdeudadosAl.frx":27D8
   End
End
Attribute VB_Name = "frmAdminExtrasCbtesVentasAdeudadosAl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vId As String
Private desde
Private Factura As Factura
Private facturas As Collection
Dim m_Archivos As Dictionary

Private Sub btnExportar_Click()
    Me.progreso.Visible = True
    Dim FechaFin As String
    FechaFin = Me.dtpHastaFIN(1).value
    
    If IsSomething(facturas) Then
        If Not DAOFactura.ExportarColeccionTotalizadores(facturas, Me.progreso, FechaFin) Then GoTo err1
    End If

    Me.progreso.Visible = False

    Exit Sub
err1:
    MsgBox "Se produjo un error al exportar!", vbCritical, "Error"
    
End Sub


Private Sub btnRemoveProveedor_Click()
    Me.cboClientes.ListIndex = -1
End Sub


Private Sub btnBuscar_Click()

    If IsNull(Me.dtpHastaFIN(1).value) Then
        MsgBox "Tiene que seleccionar una fecha de fin de cobro.", _
               vbExclamation, _
               "Fecha requerida"
        Exit Sub
    End If

    If Not IsNull(Me.dtpDesde(0).value) _
       And Not IsNull(Me.dtpHasta(0).value) Then

        If DateValue(Me.dtpDesde(0).value) > _
           DateValue(Me.dtpHasta(0).value) Then

            MsgBox "La fecha Desde no puede ser posterior a la fecha Hasta.", _
                   vbExclamation, _
                   "Rango incorrecto"
            Exit Sub
        End If

    End If

    If Not IsNull(Me.dtpHasta(0).value) Then

        If DateValue(Me.dtpHasta(0).value) > _
           DateValue(Me.dtpHastaFIN(1).value) Then

            MsgBox "La fecha final de los comprobantes no puede ser posterior " & _
                   "a la fecha de corte de pagos y cancelaciones.", _
                   vbExclamation, _
                   "Fecha de corte incorrecta"
            Exit Sub
        End If

    End If

    llenarGrilla

End Sub


Private Sub Form_Load()
    Set m_Archivos = DAOArchivo.GetCantidadArchivosPorReferencia(OA_FacturaProveedor)
    vId = funciones.CreateGUID
    
      
    FormHelper.Customize Me
    
    GridEXHelper.CustomizeGrid Me.grilla, True
    
    Me.grilla.ItemCount = 0
        
    dtpHastaFIN(1) = Now()
    
    DAOCliente.llenarComboXtremeSuite Me.cboClientes, False, True, False
    Me.cboClientes.ListIndex = -1
    
    
    btnRemoveProveedor_Click
    desde = DateSerial(Year(Date), Month(Date), 1)
    funciones.FillComboBoxDateRanges Me.cboRangos

    For i = 0 To Me.cboRangos.ListCount - 1
        If Me.cboRangos.ItemData(i) = DateRangeValue.DRV_YearCurrent Then Exit For
    Next i
    
    Me.cboRangos.ListIndex = i
   
    Me.grilla.Refresh
    
'Se agrega estas lineas para pre-setear las fechas ejemplo...
'    Me.dtpDesde(0).value = "01/04/2022"
'    Me.dtpHasta(0).value = "31/03/2023"
'    Me.dtpHastaFIN(1).value = "31/03/2023"

End Sub


Public Sub llenarGrilla()
    
    Dim filtro As String
    Dim FechaFin As String

    Me.grilla.ItemCount = 0
    
    filtro = "1=1"
    
    'Incluir solamente comprobantes válidos para la cuenta corriente.
    filtro = filtro & _
        " AND AdminFacturas.estado IN (" & _
        EstadoFacturaCliente.Aprobada & ", " & _
        EstadoFacturaCliente.CanceladaNC & ", " & _
        EstadoFacturaCliente.CanceladaNCParcial & ", " & _
        EstadoFacturaCliente.AplicadaND & ", " & _
        EstadoFacturaCliente.AplicadaACbte & ")"

    If Me.cboClientes.ListIndex >= 0 Then
        filtro = filtro & " and AdminFacturas.idCliente=" & cboClientes.ItemData(Me.cboClientes.ListIndex)
    End If

    If LenB(Me.txtNroFactura) > 0 And IsNumeric(Me.txtNroFactura) Then
        filtro = filtro & " and AdminFacturas.nroFactura=" & Me.txtNroFactura
    End If

    If Not IsNull(Me.dtpDesde(0).value) Then
        filtro = filtro & " AND AdminFacturas.FechaEmision >= " & conectar.Escape(Me.dtpDesde(0).value)
    End If

    If Not IsNull(Me.dtpHasta(0).value) Then
        filtro = filtro & " AND AdminFacturas.FechaEmision <= " & conectar.Escape(Me.dtpHasta(0).value)
    End If

    If Not IsNull(Me.dtpHastaFIN(1).value) Then
    
        FechaFin = conectar.Escape(Me.dtpHastaFIN(1).value)
    
        'Un reporte al corte no puede contener comprobantes
        'emitidos después de esa fecha.
        filtro = filtro & _
            " AND AdminFacturas.FechaEmision <= " & FechaFin
    
    End If
    
    Set facturas = DAOFactura.FindAllTotalizadores(filtro, , , FechaFin)
    
    ''''''''''''''''    ''''''''''''''''    ''''''''''''''''
    
    Dim F As Factura
    Dim Signo As Integer
    
    Dim TotalFiltrado As Double
    Dim NuevoSaldo As Double
    
    Dim TotalComprobante As Double
    Dim TotalCobrado As Double
    Dim TotalNotasCredito As Double
    Dim SaldoFactura As Double
    
    TotalFiltrado = 0
    NuevoSaldo = 0
    
    For Each F In facturas
    
        If F.TipoDocumento = tipoDocumentoContable.notaCredito Then
            Signo = -1
        Else
            Signo = 1
        End If
    
        'Total del comprobante convertido a moneda patrón.
        TotalComprobante = funciones.RedondearDecimales( _
            F.TotalEstatico.total * F.CambioAPatron, 2)
    
        'Total neto de los comprobantes filtrados.
        TotalFiltrado = TotalFiltrado + _
                        (TotalComprobante * Signo)
    
        If F.TipoDocumento <> tipoDocumentoContable.notaCredito Then
    
            'Recibos aprobados aplicados hasta la fecha de corte.
            TotalCobrado = funciones.RedondearDecimales( _
                F.MontoCobrado * F.CambioAPatron, 2)
    
            'Suma de todas las notas de crédito aplicadas.
            TotalNotasCredito = funciones.RedondearDecimales( _
                F.CbteAsociadoMonto, 2)
    
            'Saldo definitivo de la factura.
            SaldoFactura = funciones.RedondearDecimales( _
                TotalComprobante - _
                TotalCobrado - _
                TotalNotasCredito, 2)
    
            NuevoSaldo = NuevoSaldo + SaldoFactura
    
        End If
    
    Next F
    
    TotalFiltrado = funciones.RedondearDecimales( _
        TotalFiltrado, 2)
    
    NuevoSaldo = funciones.RedondearDecimales( _
        NuevoSaldo, 2)
    
    Me.grilla.ItemCount = facturas.count

    GridEXHelper.AutoSizeColumns Me.grilla, True
    
    Me.caption = "Cbtes. filtrados [Cantidad: " & facturas.count & "]"
    
    Me.lbl(2).caption = FormatCurrency(TotalFiltrado)
    Me.lblTotalNuevoSaldo(1).caption = FormatCurrency(NuevoSaldo)
    
End Sub


Private Sub cboRangos_Click()
    funciones.CalculateDateRange Me.cboRangos, Me.dtpDesde(0), Me.dtpHasta(0)
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    Me.grilla.Width = Me.ScaleWidth - 200
    Me.grilla.Height = (Me.ScaleHeight * 75) / 100
    Me.GroupBox(2).Width = Me.grilla.Width

End Sub


Private Sub grilla_ColumnHeaderClick(ByVal Column As GridEX20.JSColumn)
    GridEXHelper.ColumnHeaderClick Me.grilla, Column
End Sub


Private Sub grilla_DblClick()
    verDetalle_Click
End Sub


Private Sub grilla_UnboundReadData(ByVal rowIndex As Long, _
                                   ByVal Bookmark As Variant, _
                                   ByVal Values As GridEX20.JSRowData)

    Set Factura = facturas.item(rowIndex)

    Dim Signo As Integer
    Dim TotalComprobante As Double
    Dim TotalCobrado As Double
    Dim SaldoComprobante As Double
    Dim TotalNotasCredito As Double
    Dim NuevoSaldo As Double

    With Factura

        Values(1) = .Id
        Values(2) = .Cliente.razon
        Values(3) = .Cliente.cuit

        If .esCredito Then
            Values(4) = .GetShortDescription(True, False) & " (FCE)"
        Else
            Values(4) = .GetShortDescription(True, False)
        End If

        If IsSomething(.Tipo) Then
            Values(5) = .Tipo.TipoFactura.Tipo

            If .Tipo.PuntoVenta.EsElectronico _
               And Not .AprobadaAFIP _
               And .estado <> EstadoFacturaCliente.EnProceso Then

                Values(6) = "Nro. Pendiente"
            Else
                Values(6) = .NumeroFormateado
            End If
        Else
            Values(5) = ""
            Values(6) = .NumeroFormateado
        End If

        Values(7) = .FechaEmision
        Values(8) = .moneda.NombreCorto

        If .TipoDocumento = tipoDocumentoContable.notaCredito Then
            Signo = -1
        Else
            Signo = 1
        End If

        'Importe original convertido a moneda patrón.
        TotalComprobante = funciones.RedondearDecimales( _
            .TotalEstatico.total * .CambioAPatron, 2)

        'Recibos aprobados y aplicados hasta la fecha seleccionada.
        TotalCobrado = funciones.RedondearDecimales( _
            .MontoCobrado * .CambioAPatron, 2)

        'Saldo sin considerar todavía las notas de crédito.
        SaldoComprobante = funciones.RedondearDecimales( _
            TotalComprobante - TotalCobrado, 2)

        'Suma de todas las notas de crédito asociadas.
        TotalNotasCredito = funciones.RedondearDecimales( _
            .CbteAsociadoMonto, 2)

        Values(9) = Replace(FormatCurrency( _
            TotalComprobante * Signo), "$", "")

        Values(10) = Replace(FormatCurrency( _
            TotalCobrado * Signo), "$", "")

        Values(11) = Replace(FormatCurrency( _
            SaldoComprobante * Signo), "$", "")

        'Inicializar datos del comprobante asociado.
        Values(12) = ""
        Values(13) = ""
        Values(14) = ""
        Values(15) = ""
        Values(16) = ""

        If .TipoDocumento = tipoDocumentoContable.notaCredito Then

            'La NC ya se descuenta en la factura relacionada.
            'No debe descontarse nuevamente en el Nuevo Saldo.
            Values(17) = Replace(FormatCurrency(0), "$", "")

        Else

            If LenB(Trim$(.CbteAsociadoID)) > 0 Then

                Values(12) = .observaciones_cancela
                Values(13) = .CbteAsociadoID
                Values(14) = .CbteAsociado

                If .CbteAsociadoFecha > 0 Then
                    Values(15) = .CbteAsociadoFecha
                End If

                If TotalNotasCredito <> 0 Then
                    Values(16) = Replace(FormatCurrency( _
                        TotalNotasCredito), "$", "")
                End If

            End If

            'Fórmula definitiva:
            'Total factura - cobros - notas de crédito.
            NuevoSaldo = funciones.RedondearDecimales( _
                SaldoComprobante - TotalNotasCredito, 2)

            Values(17) = Replace(FormatCurrency( _
                NuevoSaldo), "$", "")

        End If

    End With

End Sub


Private Sub txtNroFactura_GotFocus()
    foco Me.txtNroFactura
    
End Sub


Private Sub verDetalle_Click()
    SeleccionarFactura
   
    Dim f_c3h3 As New frmAdminFacturasEdicion
    f_c3h3.ReadOnly = True
    f_c3h3.idFactura = Factura.Id
    f_c3h3.Show
    
End Sub


Private Sub SeleccionarFactura()
    On Error Resume Next
    Set Factura = facturas.item(grilla.rowIndex(grilla.row))
    
   
End Sub



