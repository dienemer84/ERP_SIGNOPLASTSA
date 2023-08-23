VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmAdminExtrasCbtesAdeudadosAl 
   Caption         =   "Comprobantes Compra adeudados al"
   ClientHeight    =   6930
   ClientLeft      =   60
   ClientTop       =   465
   ClientWidth     =   14415
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6645
   ScaleMode       =   0  'User
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
      ImagePicture1   =   "frmAdminExtrasCbtesAdeudadosAl.frx":0000
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   11
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
      FormatStylesCount=   9
      FormatStyle(1)  =   "frmAdminExtrasCbtesAdeudadosAl.frx":110E
      FormatStyle(2)  =   "frmAdminExtrasCbtesAdeudadosAl.frx":1246
      FormatStyle(3)  =   "frmAdminExtrasCbtesAdeudadosAl.frx":12F6
      FormatStyle(4)  =   "frmAdminExtrasCbtesAdeudadosAl.frx":13AA
      FormatStyle(5)  =   "frmAdminExtrasCbtesAdeudadosAl.frx":1482
      FormatStyle(6)  =   "frmAdminExtrasCbtesAdeudadosAl.frx":153A
      FormatStyle(7)  =   "frmAdminExtrasCbtesAdeudadosAl.frx":161A
      FormatStyle(8)  =   "frmAdminExtrasCbtesAdeudadosAl.frx":16DA
      FormatStyle(9)  =   "frmAdminExtrasCbtesAdeudadosAl.frx":179E
      ImageCount      =   1
      ImagePicture(1) =   "frmAdminExtrasCbtesAdeudadosAl.frx":185E
      PrinterProperties=   "frmAdminExtrasCbtesAdeudadosAl.frx":1B78
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

Private Sub btnBuscar_Click(Index As Integer)
    If IsNull(dtpHastaFIN(1).value) Then
                MsgBox ("Tiene que selecionar una fecha de fin de pagos!")
                    Else
        llenarGrilla
    End If
End Sub

Private Sub btnCargarProveedores_Click()

    Set colProveedores = DAOProveedor.FindAll
    For Each prov In colProveedores
        cboProveedores.AddItem prov.RazonSocial
        cboProveedores.ItemData(cboProveedores.NewIndex) = prov.Id
    Next

End Sub

Private Sub btnExportar_Click(Index As Integer)
    Me.progreso(0).Visible = True

    Dim FechaFIn As String

    FechaFIn = Me.dtpHastaFIN(1).value
    
    If IsSomething(facturas) Then
        If Not DAOFacturaProveedor.ExportarColeccionTotalizadores(facturas, Me.progreso, FechaFIn) Then GoTo err1

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


Public Sub llenarGrilla()
    grilla.ItemCount = 0
    Dim condition As String
    condition = " 1 = 1 "
    Dim FechaFIn As String
    
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
        condition = condition & " AND AdminComprasFacturasProveedores.numero_factura like '%" & Trim(Me.txtComprobante.text) & "%'"
    End If

    If Not IsNull(dtpHastaFIN(1).value) Then
        FechaFIn = conectar.Escape(dtpHastaFIN(1).value)
    End If
    
    Set facturas = DAOFacturaProveedor.FindAllTotalizadores(condition, FechaFIn, , , Permisos.AdminFaPVerSoloPropias)
    
    ''''''''''''''''
    
    Dim total As Double
    Dim pagado As Double
    Dim saldo As Double
    Dim TotalFactura As Double
    Dim TotalPagado As Double
    Dim c As Integer

    total = 0

    For Each Factura In facturas

        If Factura.tipoDocumentoContable = tipoDocumentoContable.NotaCredito Then c = -1 Else c = 1
        
        TotalFactura = ((Factura.Monto - Factura.TotalNetoGravadoDiscriminado(0)) + Factura.TotalIVA + Factura.TotalNetoGravadoDiscriminado(0) + Factura.totalPercepciones + Factura.ImpuestoInterno + Factura.Redondeo) * c
        total = total + TotalFactura
              
        TotalPagado = (Factura.TotalAbonadoGlobal) * c
        pagado = pagado + TotalPagado
        
        
        TotalSaldado = TotalFactura - TotalPagado
        saldo = saldo + TotalSaldado
    Next

    Me.lblTotalTotal(2).caption = FormatCurrency(funciones.FormatearDecimales(total))
    Me.lblTotalSaldo(1).caption = FormatCurrency(funciones.FormatearDecimales(saldo))
    Me.lblTotalPagado(3).caption = FormatCurrency(funciones.FormatearDecimales(pagado))
    
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


Private Sub grilla_RowFormat(RowBuffer As GridEX20.JSRowData)
    On Error GoTo err1
    Set Factura = facturas(RowBuffer.rowIndex)

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


Private Sub grilla_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)

    Set Factura = facturas.item(rowIndex)

    Dim i As Integer

    If Factura.tipoDocumentoContable = tipoDocumentoContable.NotaCredito Then i = -1 Else i = 1

    With Factura

        Values(1) = Factura.Id
        
        If IsSomething(Factura.Proveedor) Then
            Values(2) = UCase(funciones.RazonSocialFormateada(Factura.Proveedor.RazonSocial))
            Values(3) = Factura.Proveedor.Cuit
        End If

        Values(4) = enums.EnumTipoDocumentoContableShort(Factura.tipoDocumentoContable)
        Values(5) = Factura.configFactura.TipoFactura
        Values(6) = Factura.numero
        Values(7) = Factura.FEcha
        Values(8) = Factura.moneda.NombreCorto

        TotalFactura = (Factura.Monto - Factura.TotalNetoGravadoDiscriminado(0)) + Factura.TotalIVA + Factura.TotalNetoGravadoDiscriminado(0) + Factura.totalPercepciones + Factura.ImpuestoInterno
 
        Values(9) = Replace(FormatCurrency(funciones.FormatearDecimales(TotalFactura + Factura.Redondeo) * i), "$", "")
        Values(10) = Replace(FormatCurrency(funciones.FormatearDecimales(Factura.TotalAbonadoGlobal) * i), "$", "")
        Values(11) = Replace(FormatCurrency(funciones.FormatearDecimales((TotalFactura + Factura.Redondeo) - Factura.TotalAbonadoGlobal) * i), "$", "")

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
    Set Factura = facturas.item(grilla.rowIndex(grilla.row))
    
End Sub


