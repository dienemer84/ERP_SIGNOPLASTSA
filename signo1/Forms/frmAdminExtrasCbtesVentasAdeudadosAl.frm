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
         Text            =   "cboProveedores"
      End
      Begin XtremeSuiteControls.PushButton btnRemoveProveedor 
         Height          =   255
         Left            =   5520
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
         Height          =   855
         Left            =   11040
         TabIndex        =   11
         Top             =   1080
         Width           =   3255
         _Version        =   786432
         _ExtentX        =   5741
         _ExtentY        =   1508
         _StockProps     =   79
         BackColor       =   16744576
         Appearance      =   4
         Begin XtremeSuiteControls.PushButton btnBuscar 
            Default         =   -1  'True
            Height          =   390
            Left            =   240
            TabIndex        =   12
            Top             =   300
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
            Top             =   300
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
         Height          =   855
         Index           =   1
         Left            =   14400
         TabIndex        =   14
         Top             =   1080
         Width           =   4095
         _Version        =   786432
         _ExtentX        =   7223
         _ExtentY        =   1508
         _StockProps     =   79
         BackColor       =   16744576
         Appearance      =   4
         Begin XtremeSuiteControls.ProgressBar progreso 
            Height          =   375
            Left            =   120
            TabIndex        =   15
            Top             =   300
            Visible         =   0   'False
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
         TabIndex        =   16
         Top             =   240
         Width           =   2100
         _Version        =   786432
         _ExtentX        =   3704
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Cargar Clientes"
         BackColor       =   12632256
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.DateTimePicker dtpHastaFIN 
         Height          =   315
         Index           =   1
         Left            =   3840
         TabIndex        =   17
         Top             =   1620
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
         Left            =   11040
         TabIndex        =   22
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
            Left            =   4950
            TabIndex        =   30
            Top             =   480
            Width           =   1740
         End
         Begin VB.Label lblTotalNuevoSaldo 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Total Nuevo Saldo:"
            Height          =   195
            Index           =   0
            Left            =   3480
            TabIndex        =   29
            Top             =   480
            Width           =   1380
         End
         Begin XtremeSuiteControls.Label lbl 
            Height          =   195
            Index           =   1
            Left            =   4920
            TabIndex        =   28
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
         Begin XtremeSuiteControls.Label lbl 
            Height          =   195
            Index           =   2
            Left            =   1080
            TabIndex        =   27
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
         Begin XtremeSuiteControls.Label lbl 
            Height          =   195
            Index           =   3
            Left            =   1080
            TabIndex        =   26
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
         Begin XtremeSuiteControls.Label lblTotalSaldo 
            Height          =   195
            Index           =   0
            Left            =   3480
            TabIndex        =   25
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
         Begin XtremeSuiteControls.Label lblTotalFiltrado 
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   24
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
         Begin XtremeSuiteControls.Label lblTotalCobro 
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Width           =   870
            _Version        =   786432
            _ExtentX        =   1535
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Total Cobro:"
            BackColor       =   12632256
            AutoSize        =   -1  'True
         End
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   195
         Index           =   0
         Left            =   165
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
         Top             =   1680
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
      TabIndex        =   21
      Top             =   2400
      Width           =   18570
      _ExtentX        =   32755
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
      ImagePicture1   =   "frmAdminExtrasCbtesVentasAdeudadosAl.frx":0000
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   17
      Column(1)       =   "frmAdminExtrasCbtesVentasAdeudadosAl.frx":031A
      Column(2)       =   "frmAdminExtrasCbtesVentasAdeudadosAl.frx":0476
      Column(3)       =   "frmAdminExtrasCbtesVentasAdeudadosAl.frx":05A6
      Column(4)       =   "frmAdminExtrasCbtesVentasAdeudadosAl.frx":06E6
      Column(5)       =   "frmAdminExtrasCbtesVentasAdeudadosAl.frx":0836
      Column(6)       =   "frmAdminExtrasCbtesVentasAdeudadosAl.frx":0976
      Column(7)       =   "frmAdminExtrasCbtesVentasAdeudadosAl.frx":0ABE
      Column(8)       =   "frmAdminExtrasCbtesVentasAdeudadosAl.frx":0BFE
      Column(9)       =   "frmAdminExtrasCbtesVentasAdeudadosAl.frx":0D46
      Column(10)      =   "frmAdminExtrasCbtesVentasAdeudadosAl.frx":0E86
      Column(11)      =   "frmAdminExtrasCbtesVentasAdeudadosAl.frx":0FCE
      Column(12)      =   "frmAdminExtrasCbtesVentasAdeudadosAl.frx":110E
      Column(13)      =   "frmAdminExtrasCbtesVentasAdeudadosAl.frx":1246
      Column(14)      =   "frmAdminExtrasCbtesVentasAdeudadosAl.frx":13AA
      Column(15)      =   "frmAdminExtrasCbtesVentasAdeudadosAl.frx":150E
      Column(16)      =   "frmAdminExtrasCbtesVentasAdeudadosAl.frx":167A
      Column(17)      =   "frmAdminExtrasCbtesVentasAdeudadosAl.frx":17E6
      FormatStylesCount=   10
      FormatStyle(1)  =   "frmAdminExtrasCbtesVentasAdeudadosAl.frx":1946
      FormatStyle(2)  =   "frmAdminExtrasCbtesVentasAdeudadosAl.frx":1A7E
      FormatStyle(3)  =   "frmAdminExtrasCbtesVentasAdeudadosAl.frx":1B2E
      FormatStyle(4)  =   "frmAdminExtrasCbtesVentasAdeudadosAl.frx":1BE2
      FormatStyle(5)  =   "frmAdminExtrasCbtesVentasAdeudadosAl.frx":1CBA
      FormatStyle(6)  =   "frmAdminExtrasCbtesVentasAdeudadosAl.frx":1D72
      FormatStyle(7)  =   "frmAdminExtrasCbtesVentasAdeudadosAl.frx":1E52
      FormatStyle(8)  =   "frmAdminExtrasCbtesVentasAdeudadosAl.frx":1F12
      FormatStyle(9)  =   "frmAdminExtrasCbtesVentasAdeudadosAl.frx":1FD6
      FormatStyle(10) =   "frmAdminExtrasCbtesVentasAdeudadosAl.frx":2096
      ImageCount      =   1
      ImagePicture(1) =   "frmAdminExtrasCbtesVentasAdeudadosAl.frx":211A
      PrinterProperties=   "frmAdminExtrasCbtesVentasAdeudadosAl.frx":2434
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

Private Sub btnCargarProveedores_Click()

    DAOCliente.llenarComboXtremeSuite Me.cboClientes, False, True, False
    Me.cboClientes.ListIndex = -1
    
End Sub


Private Sub btnExportar_Click()
    Me.progreso.Visible = True
    Dim FechaFIn As String
    FechaFIn = Me.dtpHastaFIN(1).value
    
    If IsSomething(facturas) Then
        If Not DAOFactura.ExportarColeccionTotalizadores(facturas, Me.progreso, FechaFIn) Then GoTo err1
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
    If IsNull(dtpHastaFIN(1).value) Then
                MsgBox ("Tiene que selecionar una fecha de fin de cobro!")
                    Else
        llenarGrilla
    End If

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
    
    
    Me.dtpDesde(0).value = "01/04/2022"
    Me.dtpHasta(0).value = "31/03/2023"
    Me.dtpHastaFIN(1).value = "31/03/2023"

End Sub


Public Sub llenarGrilla()
    
    Dim filtro As String
    Dim FechaFIn As String
    
    Me.grilla.ItemCount = 0
    filtro = "1=1"

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


    If Not IsNull(dtpHastaFIN(1).value) Then
        FechaFIn = conectar.Escape(dtpHastaFIN(1).value)
    End If

    
    Set facturas = DAOFactura.FindAllTotalizadores(filtro, , , FechaFIn)
    
    ''''''''''''''''    ''''''''''''''''    ''''''''''''''''
    
    Dim F As Factura
    Dim c As Integer

    For Each F In facturas

        Dim total As Double
        Dim TotalCobrado As Double
        Dim cobrado As Double
        Dim nuevoSaldo As Double

        If F.TipoDocumento = tipoDocumentoContable.NotaCredito Then c = -1 Else c = 1

        total = total + MonedaConverter.ConvertirForzado2(F.TotalEstatico.total * c, MonedaConverter.Patron.Id, F.moneda.Id, F.CambioAPatron)

        TotalCobrado = F.MontoCobrado
        cobrado = cobrado + TotalCobrado
        
       
        
    Next
'PARA AJUSTAR! 11/08/2023
'    Me.lbl(2).caption = FormatCurrency(funciones.FormatearDecimales(total))
'    Me.lbl(1).caption = FormatCurrency(funciones.FormatearDecimales(total - cobrado))
'    Me.lbl(3).caption = FormatCurrency(funciones.FormatearDecimales(cobrado))
'    lblTotalNuevoSaldo(1) = FormatCurrency(funciones.FormatearDecimales(nuevoSaldo))
    
    '''''''''''''''    ''''''''''''''''    ''''''''''''''''

    Me.grilla.ItemCount = 0
    
    Me.grilla.ItemCount = facturas.count
    
    GridEXHelper.AutoSizeColumns Me.grilla, True
    
    Me.caption = "Cbtes. filtrados [Cantidad: " & facturas.count & "]"


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


Private Sub grilla_FetchIcon(ByVal rowIndex As Long, ByVal ColIndex As Integer, ByVal RowBookmark As Variant, ByVal IconIndex As GridEX20.JSRetInteger)
    If ColIndex = 15 And m_Archivos.item(Factura.Id) > 0 Then IconIndex = 1

End Sub


Private Sub grilla_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)

    Set Factura = facturas.item(rowIndex)

    Dim i As Integer

    With Factura

        Values(1) = Factura.Id
        Values(2) = Factura.cliente.razon
        Values(3) = Factura.cliente.Cuit

        If Factura.esCredito Then
            Values(4) = Factura.GetShortDescription(True, False) & " " & "(FCE)"
        Else
            Values(4) = Factura.GetShortDescription(True, False)
        End If

        If IsSomething(Factura.Tipo) Then
            Values(5) = Factura.Tipo.TipoFactura.Tipo
        End If

        If Factura.Tipo.PuntoVenta.EsElectronico And Not Factura.AprobadaAFIP And Factura.estado <> EstadoFacturaCliente.EnProceso Then
            Values(6) = "Nro. Pendiente"
        Else
            Values(6) = Factura.NumeroFormateado
        End If

        Values(7) = Factura.FechaEmision

        Values(8) = Factura.moneda.NombreCorto

        If Factura.TipoDocumento = tipoDocumentoContable.NotaCredito Then
            i = -1
        Else
            i = 1
        End If

        'MONTO TOTAL
        Dim TotalComprobante As Double
        TotalComprobante = (Factura.TotalEstatico.total * Factura.CambioAPatron)
        Values(9) = Replace(FormatCurrency(funciones.FormatearDecimales(TotalComprobante) * i), "$", "")

        'MONTO COBRADO
        Dim TotalCobrado As Double
        TotalCobrado = Factura.MontoCobrado * Factura.CambioAPatron
        Values(10) = Replace(FormatCurrency(funciones.FormatearDecimales(TotalCobrado) * i), "$", "")

        'SALDO
        Dim saldoComprobante As Double
        saldoComprobante = TotalComprobante - TotalCobrado
        Values(11) = Replace(FormatCurrency(funciones.FormatearDecimales(saldoComprobante) * i), "$", "")

    End With

    If Factura.TipoDocumento = tipoDocumentoContable.NotaCredito Then
        Values(17) = "0,00"
    Else
        If Factura.CbteAsociadoTipo <> "2" And Factura.CbteAsociadoTipo <> "5" And Factura.CbteAsociadoTipo <> "8" And Factura.CbteAsociadoTipo <> "16" And Factura.CbteAsociadoTipo <> "11" And Factura.CbteAsociadoTipo <> "22" Then

            Values(12) = ""
            Values(13) = ""
            Values(14) = ""
            Values(15) = ""
            Values(16) = ""

            Values(17) = Values(11)

        Else

            Values(12) = Factura.observaciones_cancela
            Values(13) = Factura.CbteAsociadoID
            Values(14) = Factura.CbteAsociado
            '        Values(17) = Replace(FormatCurrency(funciones.FormatearDecimales(saldoComprobante) * i), "$", "")

            If Factura.CbteAsociadoFecha = "12:00:00 a.m." Then
                Values(15) = ""
            Else
                Values(15) = Format(Factura.CbteAsociadoFecha, ddmmyyy)
            End If

            If Factura.CbteAsociadoMonto = 0 Then
                Values(16) = ""
            Else
                Values(16) = Replace(FormatCurrency(funciones.FormatearDecimales(Factura.CbteAsociadoMonto) * i), "$", "")
            End If

            Values(17) = Replace(FormatCurrency(funciones.FormatearDecimales(TotalComprobante - Factura.CbteAsociadoMonto) * i), "$", "")



        End If
    End If


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



