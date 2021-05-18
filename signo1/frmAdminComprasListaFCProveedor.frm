VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GRIDEX20.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmAdminComprasListaFCProveedor 
   Appearance      =   0  'Flat
   BackColor       =   &H00FF8080&
   Caption         =   "Facturas Proveedores"
   ClientHeight    =   7815
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   25005
   ClipControls    =   0   'False
   Icon            =   "frmAdminComprasListaFCProveedor.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7815
   ScaleWidth      =   25005
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1800
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   23715
      _Version        =   786432
      _ExtentX        =   41831
      _ExtentY        =   3175
      _StockProps     =   79
      Caption         =   "Parámetros de búsqueda"
      UseVisualStyle  =   -1  'True
      Begin VB.TextBox txtCtaContable 
         Height          =   285
         Left            =   7155
         TabIndex        =   28
         Top             =   630
         Width           =   1605
      End
      Begin VB.TextBox txtComprobante 
         Height          =   285
         Left            =   1440
         TabIndex        =   13
         Top             =   615
         Width           =   1965
      End
      Begin XtremeSuiteControls.ComboBox cboProveedores 
         Height          =   315
         Left            =   1440
         TabIndex        =   10
         Top             =   240
         Width           =   3885
         _Version        =   786432
         _ExtentX        =   6853
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Text            =   "cboProveedores"
      End
      Begin XtremeSuiteControls.PushButton Command2 
         Default         =   -1  'True
         Height          =   390
         Left            =   22200
         TabIndex        =   2
         Top             =   240
         Width           =   1245
         _Version        =   786432
         _ExtentX        =   2196
         _ExtentY        =   688
         _StockProps     =   79
         Caption         =   "Buscar"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.DateTimePicker dtpDesde 
         Height          =   315
         Left            =   7995
         TabIndex        =   3
         Top             =   990
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
         Left            =   10215
         TabIndex        =   4
         Top             =   990
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
         Left            =   2490
         TabIndex        =   5
         Top             =   975
         Width           =   2835
         _Version        =   786432
         _ExtentX        =   5001
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.PushButton CMDsINCliente 
         Height          =   255
         Left            =   5400
         TabIndex        =   11
         Top             =   255
         Width           =   420
         _Version        =   786432
         _ExtentX        =   741
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "X"
         BackColor       =   12632256
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cboFantasia 
         Height          =   315
         Left            =   7155
         TabIndex        =   14
         Top             =   240
         Width           =   4530
         _Version        =   786432
         _ExtentX        =   7990
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Sorted          =   -1  'True
         Text            =   "cboProveedores"
      End
      Begin XtremeSuiteControls.PushButton btnClearFantasia 
         Height          =   255
         Left            =   11760
         TabIndex        =   15
         Top             =   255
         Width           =   420
         _Version        =   786432
         _ExtentX        =   741
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "X"
         BackColor       =   12632256
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.DateTimePicker dtpDesdeCarga 
         Height          =   315
         Left            =   7995
         TabIndex        =   17
         Top             =   1350
         Width           =   1470
         _Version        =   786432
         _ExtentX        =   2593
         _ExtentY        =   556
         _StockProps     =   68
         CheckBox        =   -1  'True
         Format          =   1
      End
      Begin XtremeSuiteControls.DateTimePicker dtpHastaCarga 
         Height          =   315
         Left            =   10215
         TabIndex        =   18
         Top             =   1350
         Width           =   1470
         _Version        =   786432
         _ExtentX        =   2593
         _ExtentY        =   556
         _StockProps     =   68
         CheckBox        =   -1  'True
         Format          =   1
      End
      Begin XtremeSuiteControls.ComboBox cboRangosCarga 
         Height          =   315
         Left            =   2490
         TabIndex        =   19
         Top             =   1335
         Width           =   2835
         _Version        =   786432
         _ExtentX        =   5001
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.PushButton btnRemoveEstado 
         Height          =   255
         Left            =   5400
         TabIndex        =   24
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
      Begin XtremeSuiteControls.ComboBox cboEstado 
         Height          =   315
         Left            =   4080
         TabIndex        =   23
         Top             =   600
         Width           =   1245
         _Version        =   786432
         _ExtentX        =   2196
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Sorted          =   -1  'True
         Style           =   2
         Text            =   "cboProveedores"
      End
      Begin XtremeSuiteControls.PushButton cmdImprimir 
         Height          =   390
         Left            =   22200
         TabIndex        =   33
         Top             =   1200
         Width           =   1245
         _Version        =   786432
         _ExtentX        =   2196
         _ExtentY        =   688
         _StockProps     =   79
         Caption         =   "Imprimir"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton cmdExportar 
         Height          =   390
         Left            =   22200
         TabIndex        =   34
         Top             =   720
         Width           =   1245
         _Version        =   786432
         _ExtentX        =   2196
         _ExtentY        =   688
         _StockProps     =   79
         Caption         =   "Exportar"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnClearCtaCble 
         Height          =   255
         Left            =   11760
         TabIndex        =   36
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
      Begin XtremeSuiteControls.ComboBox cboCuentasContables 
         Height          =   315
         Left            =   9840
         TabIndex        =   35
         Top             =   615
         Width           =   1845
         _Version        =   786432
         _ExtentX        =   3254
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Sorted          =   -1  'True
      End
      Begin XtremeSuiteControls.ProgressBar progreso 
         Height          =   390
         Left            =   16440
         TabIndex        =   39
         Top             =   1200
         Visible         =   0   'False
         Width           =   5535
         _Version        =   786432
         _ExtentX        =   9763
         _ExtentY        =   688
         _StockProps     =   93
         Appearance      =   6
      End
      Begin XtremeSuiteControls.Label lblExportando 
         Height          =   375
         Left            =   15360
         TabIndex        =   40
         Top             =   1200
         Visible         =   0   'False
         Width           =   1215
         _Version        =   786432
         _ExtentX        =   2143
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Exportando..."
      End
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         Caption         =   "Total Filtrado $:"
         Height          =   195
         Left            =   12720
         TabIndex        =   38
         Top             =   1500
         Width           =   1095
      End
      Begin XtremeSuiteControls.Label Label12 
         Height          =   195
         Left            =   8880
         TabIndex        =   37
         Top             =   675
         Width           =   960
         _Version        =   786432
         _ExtentX        =   1693
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Cta. Contable"
         BackColor       =   12632256
         Alignment       =   1
         AutoSize        =   -1  'True
      End
      Begin VB.Label lblTotalNeto 
         AutoSize        =   -1  'True
         Caption         =   "Total Filtrado $:"
         Height          =   195
         Left            =   12720
         TabIndex        =   32
         Top             =   690
         Width           =   1095
      End
      Begin VB.Label lblTotalIVA 
         AutoSize        =   -1  'True
         Caption         =   "Total Filtrado $:"
         Height          =   195
         Left            =   12720
         TabIndex        =   31
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label lblTotalNoGravadoFiltrado 
         AutoSize        =   -1  'True
         Caption         =   "Total Filtrado $:"
         Height          =   195
         Left            =   12720
         TabIndex        =   30
         Top             =   420
         Width           =   1095
      End
      Begin VB.Label lblNetoGravadoFiltrado 
         AutoSize        =   -1  'True
         Caption         =   "Total Filtrado $:"
         Height          =   195
         Left            =   12720
         TabIndex        =   29
         Top             =   150
         Width           =   1095
      End
      Begin XtremeSuiteControls.Label Label11 
         Height          =   195
         Left            =   6120
         TabIndex        =   27
         Top             =   675
         Width           =   960
         _Version        =   786432
         _ExtentX        =   1693
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Cta. Contable"
         BackColor       =   12632256
         AutoSize        =   -1  'True
      End
      Begin VB.Label lblTotalPercepciones 
         AutoSize        =   -1  'True
         Caption         =   "Total Filtrado $:"
         Height          =   195
         Left            =   12720
         TabIndex        =   26
         Top             =   1230
         Width           =   1095
      End
      Begin XtremeSuiteControls.Label Label10 
         Height          =   195
         Left            =   3480
         TabIndex        =   25
         Top             =   660
         Width           =   495
         _Version        =   786432
         _ExtentX        =   873
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Estado"
         BackColor       =   12632256
         Alignment       =   1
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label9 
         Height          =   195
         Left            =   9720
         TabIndex        =   22
         Top             =   1395
         Width           =   420
         _Version        =   786432
         _ExtentX        =   741
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Hasta"
         BackColor       =   12632256
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label8 
         Height          =   195
         Left            =   7485
         TabIndex        =   21
         Top             =   1410
         Width           =   465
         _Version        =   786432
         _ExtentX        =   820
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Desde"
         BackColor       =   12632256
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label7 
         Height          =   195
         Left            =   960
         TabIndex        =   20
         Top             =   1380
         Width           =   1440
         _Version        =   786432
         _ExtentX        =   2540
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Rango Fecha Carga"
         BackColor       =   12632256
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   195
         Left            =   6135
         TabIndex        =   16
         Top             =   315
         Width           =   975
         _Version        =   786432
         _ExtentX        =   1720
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Nom Fantasia"
         BackColor       =   12632256
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   195
         Left            =   195
         TabIndex        =   12
         Top             =   660
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
         Left            =   600
         TabIndex        =   9
         Top             =   285
         Width           =   735
         _Version        =   786432
         _ExtentX        =   1296
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Proveedor"
         BackColor       =   12632256
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label4 
         Height          =   195
         Left            =   450
         TabIndex        =   8
         Top             =   1020
         Width           =   1965
         _Version        =   786432
         _ExtentX        =   3466
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Rango Fecha Comprobante"
         BackColor       =   12632256
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label5 
         Height          =   195
         Left            =   7455
         TabIndex        =   7
         Top             =   1050
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
         Left            =   9720
         TabIndex        =   6
         Top             =   1050
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
   Begin GridEX20.GridEX grilla 
      Height          =   5280
      Left            =   120
      TabIndex        =   0
      Top             =   2400
      Width           =   23700
      _ExtentX        =   41804
      _ExtentY        =   9313
      Version         =   "2.0"
      PreviewRowIndent=   100
      DefaultGroupMode=   1
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      GroupFooterStyle=   2
      PreviewRowLines =   1
      ColumnAutoResize=   -1  'True
      MultiSelect     =   -1  'True
      MethodHoldFields=   -1  'True
      GroupByBoxInfoText=   "Arrastrar una columna para agrupar"
      AllowEdit       =   0   'False
      BackColorGBBox  =   16744576
      BackColorHeader =   16761024
      ImageCount      =   1
      ImagePicture1   =   "frmAdminComprasListaFCProveedor.frx":000C
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   19
      Column(1)       =   "frmAdminComprasListaFCProveedor.frx":0326
      Column(2)       =   "frmAdminComprasListaFCProveedor.frx":045E
      Column(3)       =   "frmAdminComprasListaFCProveedor.frx":0532
      Column(4)       =   "frmAdminComprasListaFCProveedor.frx":05FE
      Column(5)       =   "frmAdminComprasListaFCProveedor.frx":0742
      Column(6)       =   "frmAdminComprasListaFCProveedor.frx":0862
      Column(7)       =   "frmAdminComprasListaFCProveedor.frx":09A6
      Column(8)       =   "frmAdminComprasListaFCProveedor.frx":0C12
      Column(9)       =   "frmAdminComprasListaFCProveedor.frx":0E0E
      Column(10)      =   "frmAdminComprasListaFCProveedor.frx":103E
      Column(11)      =   "frmAdminComprasListaFCProveedor.frx":124E
      Column(12)      =   "frmAdminComprasListaFCProveedor.frx":145A
      Column(13)      =   "frmAdminComprasListaFCProveedor.frx":165A
      Column(14)      =   "frmAdminComprasListaFCProveedor.frx":17CA
      Column(15)      =   "frmAdminComprasListaFCProveedor.frx":190E
      Column(16)      =   "frmAdminComprasListaFCProveedor.frx":1A66
      Column(17)      =   "frmAdminComprasListaFCProveedor.frx":1BD6
      Column(18)      =   "frmAdminComprasListaFCProveedor.frx":1D2E
      Column(19)      =   "frmAdminComprasListaFCProveedor.frx":1ED2
      FormatStylesCount=   9
      FormatStyle(1)  =   "frmAdminComprasListaFCProveedor.frx":1FD6
      FormatStyle(2)  =   "frmAdminComprasListaFCProveedor.frx":210E
      FormatStyle(3)  =   "frmAdminComprasListaFCProveedor.frx":21BE
      FormatStyle(4)  =   "frmAdminComprasListaFCProveedor.frx":2272
      FormatStyle(5)  =   "frmAdminComprasListaFCProveedor.frx":234A
      FormatStyle(6)  =   "frmAdminComprasListaFCProveedor.frx":2402
      FormatStyle(7)  =   "frmAdminComprasListaFCProveedor.frx":24E2
      FormatStyle(8)  =   "frmAdminComprasListaFCProveedor.frx":25A2
      FormatStyle(9)  =   "frmAdminComprasListaFCProveedor.frx":2666
      ImageCount      =   1
      ImagePicture(1) =   "frmAdminComprasListaFCProveedor.frx":2726
      PrinterProperties=   "frmAdminComprasListaFCProveedor.frx":2A40
   End
   Begin VB.Menu menu 
      Caption         =   "menu"
      Visible         =   0   'False
      Begin VB.Menu VerDetalle 
         Caption         =   "Ver Factura..."
      End
      Begin VB.Menu editar 
         Caption         =   "Editar..."
      End
      Begin VB.Menu finalizar 
         Caption         =   "Aprobar..."
      End
      Begin VB.Menu mnuPagarEnEfectivo 
         Caption         =   "Pagar en Efectivo..."
      End
      Begin VB.Menu mnuArchivos 
         Caption         =   "Archivos Asociados..."
      End
      Begin VB.Menu mnuScan 
         Caption         =   "Adquirir..."
      End
      Begin VB.Menu n1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEliminar 
         Caption         =   "Eliminar"
      End
      Begin VB.Menu n2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuVerOP 
         Caption         =   "Ver Orden de Pago..."
      End
      Begin VB.Menu verHistorial 
         Caption         =   "Historial..."
      End
      Begin VB.Menu Imprimir 
         Caption         =   "Imprimir"
      End
   End
End
Attribute VB_Name = "frmAdminComprasListaFCProveedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Implements ISuscriber
Dim vId As String
Private desde
Private Factura As clsFacturaProveedor
Private facturas As Collection
Dim m_Archivos As Dictionary

Private Sub btnClearCtaCble_Click()
Me.cboCuentasContables.ListIndex = -1
End Sub

Private Sub btnClearFantasia_Click()
    Me.cboFantasia.ListIndex = -1
End Sub

Private Sub btnRemoveEstado_Click()
    Me.cboEstado.ListIndex = -1
End Sub

Private Sub cboRangos_Click()
    funciones.CalculateDateRange Me.cboRangos, Me.dtpDesde, Me.dtpHasta
End Sub

Private Sub cboRangosCarga_Click()
    funciones.CalculateDateRange Me.cboRangosCarga, Me.dtpDesdeCarga, Me.dtpHastaCarga
End Sub

Private Sub cmdExportar_Click()

    Me.progreso.Visible = True
    Me.lblExportando.Visible = True
    
    If IsSomething(facturas) Then
        If Not DAOFacturaProveedor.ExportarColeccion(facturas, Me.progreso) Then GoTo err1
    End If
    
    Me.progreso.Visible = False
    Me.lblExportando.Visible = False
    
    Exit Sub
err1:
    MsgBox "Se produjo un error al exportar!", vbCritical, "Error"
End Sub

Private Sub cmdImprimir_Click()
    Dim elegidos As Boolean
    Dim q As String

    Dim rf As String

    rf = Me.lblNetoGravadoFiltrado & Chr(10)
    rf = rf & Me.lblTotalNoGravadoFiltrado & Chr(10)
    rf = rf & Me.lblTotalNeto & Chr(10)
    rf = rf & Me.lblTotalIVA & Chr(10)
    rf = rf & Me.lblTotal & Chr(10)


    If Not IsNull(Me.dtpDesde) Then
        q = "Desde " & Format(Me.dtpDesde, "dd-mm-yyyy") & Chr(10)
    End If
    If Not IsNull(Me.dtpHasta) Then
        q = q & "Hasta " & Format(Me.dtpHasta, "dd-mm-yyyy") & Chr(10)

    End If


    If IsNull(Me.dtpHasta) And IsNull(Me.dtpDesde) Then
        q = "PERIODO SIN ESPECIFICAR" & Chr(10)
    End If

    Dim pro As String
    If Me.cboProveedores.ListIndex > -1 Then
        pro = " Proveedor: " & Me.cboProveedores.text
    End If

    With Me.grilla.PrinterProperties

        .FitColumns = True
        .RepeatHeaders = True
        .Orientation = jgexPPLandscape
        .HeaderString(jgexHFCenter) = "Listado de Comprobantes de Proveedores" & Chr(10) & pro
        .FooterString(jgexHFCenter) = Now
        .FooterString(jgexHFLeft) = rf
        .BottomMargin = 1500
        .FooterDistance = 1400
    End With
    Load frmPrintPreview
    frmPrintPreview.Move Me.Left, Me.Top, Me.Width, Me.Height
    Me.grilla.PrintPreview frmPrintPreview.GEXPreview1, elegidos
    frmPrintPreview.Show 1
End Sub

Private Sub CMDsINCliente_Click()
    Me.cboProveedores.ListIndex = -1
End Sub

Private Sub Command1_Click()
    Dim elegidos As Boolean
    If grilla.SelectedItems.count > 1 Then
        elegidos = True
    Else
        elegidos = False
    End If

    With Me.grilla.PrinterProperties
        .FitColumns = True
        .RepeatHeaders = True
        .Orientation = jgexPPLandscape
        .HeaderString(jgexHFCenter) = "Lista de Facturas de proveedores"
        .FooterString(jgexHFCenter) = Now
    End With
    Load frmPrintPreview
    frmPrintPreview.Move Me.Left, Me.Top, Me.Width, Me.Height
    grilla.PrintPreview frmPrintPreview.GEXPreview1, elegidos
    frmPrintPreview.Show 1
End Sub

Private Sub Command2_Click()
    llenarGrilla
End Sub

Private Sub Command4_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
    Dim q As String

    conectar.BeginTransaction
    For Each Factura In facturas

        If Factura.estado = Saldada Then
            q = "UPDATE AdminComprasFacturasProveedores SET total_abonado = " & Factura.Total & " WHERE id=" & Factura.id
            If Not conectar.execute(q) Then GoTo E
        End If

    Next
    conectar.CommitTransaction
    Exit Sub
E:
    conectar.RollBackTransaction
End Sub

Private Sub editar_Click()
    Set Factura = facturas.item(grilla.RowIndex(grilla.row))
    Dim frm As frmAdminComprasNuevaFCProveedor
    Set frm = New frmAdminComprasNuevaFCProveedor

    frm.ver = False
    frm.Factura = Factura
    frm.Show
End Sub
Private Sub finalizar_Click()
    If Me.grilla.ItemCount > 0 Then
        SeleccionarFactura
        Dim l As Long
        l = grilla.RowIndex(grilla.row)
        If MsgBox("¿Desea aprobar la factura?", vbQuestion + vbYesNo) = vbYes Then
            If DAOFacturaProveedor.aprobar(Factura) Then
                MsgBox "Factura aprobada con éxito!", vbInformation, "Información"
                '--------------- added 28-1-11
                txtComprobante.SetFocus
                funciones.foco Me.txtComprobante
                '---------------------------------------
                If Not Factura.FormaPagoCuentaCorriente Then MsgBox "El pago de la factura ha sido registrado con la orden de pago Nº " & DAOOrdenPago.FindLast().id & ".", vbInformation
                
                Dim tmp As clsFacturaProveedor
                facturas.item(grilla.RowIndex(grilla.row)).estado = Factura.estado
                
                
                grilla.RefreshRowIndex l
            Else
                'MsgBox "Se produjo algún error, no se aprobó la factura!", vbCritical, "Error"
            End If
        End If
    End If
End Sub
Private Sub Form_Load()
    Set m_Archivos = DAOArchivo.GetCantidadArchivosPorReferencia(OA_FacturaProveedor)
    vId = funciones.CreateGUID
    Channel.AgregarSuscriptor Me, TipoSuscripcion.FacturaProveedor_
    FormHelper.Customize Me
    GridEXHelper.CustomizeGrid Me.grilla, True
    
   ' DAOProveedor.llenarComboXtremeSuite Me.cboProveedores, True, True
   
    Set colProveedores = DAOProveedor.FindAll
    For Each prov In colProveedores
        cboProveedores.AddItem prov.RazonSocial
        cboProveedores.ItemData(cboProveedores.NewIndex) = prov.id
    Next


    Me.cboEstado.Clear
    Me.cboEstado.AddItem enums.enumEstadoFacturaProveedor(1)
    Me.cboEstado.ItemData(Me.cboEstado.NewIndex) = EstadoFacturaProveedor.EnProceso
    Me.cboEstado.AddItem enums.enumEstadoFacturaProveedor(2)
    Me.cboEstado.ItemData(Me.cboEstado.NewIndex) = EstadoFacturaProveedor.Aprobada
    Me.cboEstado.AddItem enums.enumEstadoFacturaProveedor(3)
    Me.cboEstado.ItemData(Me.cboEstado.NewIndex) = EstadoFacturaProveedor.Saldada


'    Dim P As clsProveedor
'    For Each P In DAOProveedor.FindAll()
'        If LenB(Trim$(P.razonFantasia)) > 0 Then
'            Me.cboFantasia.AddItem P.razonFantasia
'            Me.cboFantasia.ItemData(Me.cboFantasia.NewIndex) = P.id
'        End If
'    Next P
'    Me.cboFantasia.ListIndex = -1


Dim cc As clsCuentaContable
    For Each cc In DAOCuentaContable.GetAll
        If LenB(Trim$(cc.nombre)) > 0 Then
            Me.cboCuentasContables.AddItem cc.codigo & "- " & cc.nombre
            Me.cboCuentasContables.ItemData(Me.cboCuentasContables.NewIndex) = cc.id
        End If
    Next cc
    Me.cboCuentasContables.ListIndex = -1


    Me.grilla.ItemCount = 0
    CMDsINCliente_Click
    desde = DateSerial(Year(Date), Month(Date), 1)   ' CDate(1 & "-" & Month(Now) & "-" & Year(Now))
    funciones.FillComboBoxDateRanges Me.cboRangos
    For i = 0 To Me.cboRangos.ListCount - 1
        If Me.cboRangos.ItemData(i) = DateRangeValue.DRV_YearCurrent Then Exit For
    Next i
    Me.cboRangos.ListIndex = i

    funciones.FillComboBoxDateRanges Me.cboRangosCarga
    Me.dtpDesdeCarga.value = Null
    Me.dtpHastaCarga.value = Null

    Me.grilla.Refresh
End Sub
Public Sub llenarGrilla()
    Dim tot As Double
    grilla.ItemCount = 0
    Dim condition As String
    condition = " 1 = 1 "

    If Not IsNull(Me.dtpDesde.value) Then
        condition = condition & " AND AdminComprasFacturasProveedores.fecha >= " & conectar.Escape(Me.dtpDesde.value)
    End If

    If Not IsNull(Me.dtpHasta.value) Then
        condition = condition & " AND AdminComprasFacturasProveedores.fecha <= " & conectar.Escape(Me.dtpHasta.value)
    End If

    If Not IsNull(Me.dtpDesdeCarga.value) Then
        condition = condition & " AND AdminComprasFacturasProveedores.fecha_carga >= " & conectar.Escape(CDate(Int(CDbl(Me.dtpDesdeCarga.value))))
    End If

    If Not IsNull(Me.dtpHastaCarga.value) Then
        condition = condition & " AND AdminComprasFacturasProveedores.fecha_carga <= " & conectar.Escape(CDate(Int(CDbl(Me.dtpHastaCarga.value))) + TimeSerial(23, 59, 59))
    End If

    If cboProveedores.ListIndex > -1 Then
        condition = condition & " AND AdminComprasFacturasProveedores.id_proveedor = " & cboProveedores.ItemData(Me.cboProveedores.ListIndex)
    End If

    If Me.cboFantasia.ListIndex > -1 Then
        condition = condition & " AND AdminComprasFacturasProveedores.id_proveedor = " & cboFantasia.ItemData(Me.cboFantasia.ListIndex)
    End If

    If LenB(Me.txtComprobante) > 0 Then
        condition = condition & " AND AdminComprasFacturasProveedores.numero_factura like '%" & Trim(Me.txtComprobante.text) & "%'"
    End If

    If Me.cboEstado.ListIndex > -1 Then
        condition = condition & " AND AdminComprasFacturasProveedores.estado = " & Me.cboEstado.ItemData(Me.cboEstado.ListIndex)
    End If

    If LenB(Me.txtCtaContable) > 0 Then
        condition = condition & " AND AdminComprasCuentasContables.codigo LIKE '%" & Me.txtCtaContable & "%'"

    End If
    
    '#181
   If Me.cboCuentasContables.ListIndex > -1 Then
            condition = condition & " AND AdminComprasCuentasContables.id = '" & Me.cboCuentasContables.ItemData(Me.cboCuentasContables.ListIndex) & " '"

    End If



    Set facturas = DAOFacturaProveedor.FindAll(condition, , "AdminComprasFacturasProveedores.id DESC", Permisos.AdminFaPVerSoloPropias)

    Dim F As clsFacturaProveedor
    Dim Total As Double
    Dim totalneto As Double
    Dim totIva As Double
    Dim totalno As Double
    'Agrega DNEMER 03/02/2021
    Dim totalpercep As Double
    Dim c As Integer
    Total = 0

    For Each F In facturas



        If F.tipoDocumentoContable = tipoDocumentoContable.notaCredito Then c = -1 Else c = 1
        Total = Total + MonedaConverter.Convertir(F.Total * c, F.moneda.id, MonedaConverter.Patron.id)
        totalneto = totalneto + MonedaConverter.Convertir(F.Monto * c - F.TotalNetoGravadoDiscriminado(0) * c, F.moneda.id, MonedaConverter.Patron.id)
        totalno = totalno + MonedaConverter.Convertir(F.TotalNetoGravadoDiscriminado(0) * c, F.moneda.id, MonedaConverter.Patron.id)
        totIva = totIva + MonedaConverter.Convertir(F.TotalIVA * c, F.moneda.id, MonedaConverter.Patron.id)
        'Agrega DNEMER 03/02/2021
        totalpercep = totalpercep + F.totalPercepciones * c
        
    Next



    Me.lblTotal = "Total Filtrado: $ " & funciones.FormatearDecimales(Total)
    Me.lblTotalNoGravadoFiltrado = "Total No Gravado: $ " & funciones.FormatearDecimales(totalno)
    Me.lblNetoGravadoFiltrado = "Total Neto Gravado: $ " & funciones.FormatearDecimales(totalneto)
    Me.lblTotalIVA = "Total IVA: $ " & funciones.FormatearDecimales(totIva)
    Me.lblTotalNeto = "Total Neto: $ " & funciones.FormatearDecimales(funciones.RedondearDecimales(totalneto) + funciones.RedondearDecimales(totalno))
    'Agregar totalizador de Percepciones
    Me.lblTotalPercepciones = "Total Percepciones: $ " & funciones.FormatearDecimales(totalpercep)
    
    grilla.ItemCount = facturas.count
    GridEXHelper.AutoSizeColumns Me.grilla, True

    Me.caption = "Facturas Proveedores (" & facturas.count & " comprobantes encontrados)"
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Me.grilla.Width = Me.ScaleWidth - 220
    Me.grilla.Height = Me.ScaleHeight - 2000
    Me.GroupBox1.Width = Me.grilla.Width
    Me.cmdImprimir.Left = Me.GroupBox1.Width - (Me.cmdImprimir.Width + 220)
    Me.cmdExportar.Left = Me.cmdImprimir.Left
    Me.Command2.Left = Me.cmdImprimir.Left
End Sub

Private Sub Form_Terminate()
    Set vFactura = Nothing
    Channel.RemoverSuscripcionTotal Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Channel.RemoverSuscripcionTotal Me
End Sub

Private Sub grilla_ColumnHeaderClick(ByVal Column As GridEX20.JSColumn)
    GridEXHelper.ColumnHeaderClick Me.grilla, Column
End Sub

Private Sub grilla_DblClick()
    verDetalle_Click
End Sub

Private Sub grilla_FetchIcon(ByVal RowIndex As Long, ByVal ColIndex As Integer, ByVal RowBookmark As Variant, ByVal IconIndex As GridEX20.JSRetInteger)
    If ColIndex = 15 And m_Archivos.item(Factura.id) > 0 Then IconIndex = 1

End Sub

Private Sub grilla_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Me.grilla.ItemCount > 0 Then
        If Button = 2 Then
            SeleccionarFactura
            Me.finalizar.Enabled = (Factura.estado = EstadoFacturaProveedor.EnProceso)
            Me.editar.Enabled = (Factura.estado = EstadoFacturaProveedor.EnProceso)
            Me.mnuPagarEnEfectivo.Enabled = (Factura.estado = EstadoFacturaProveedor.Aprobada)

            Me.MnuVerOP.Enabled = (Factura.estado = Saldada And Factura.OrdenPagoId > 0)
            If (Factura.estado = Saldada And Factura.OrdenPagoId > 0) Then
                Me.MnuVerOP.Visible = True
                Me.MnuVerOP.caption = "Ver OP Nº " & Factura.OrdenPagoId
            Else
                Me.MnuVerOP.Visible = False
                Me.MnuVerOP.caption = "No hay OP asociada"
            End If

            Me.PopupMenu menu
        End If
    End If
End Sub

Private Sub grilla_RowFormat(RowBuffer As GridEX20.JSRowData)
    '    If RowBuffer.RowIndex > 0 Then
    '        Set tmpRto = remitos(RowBuffer.RowIndex)
    On Error GoTo err1
    Set Factura = facturas(RowBuffer.RowIndex)

    If Factura.estado = EstadoFacturaProveedor.Aprobada Then
        RowBuffer.CellStyle(14) = "EstadoAprobado"
    ElseIf Factura.estado = EstadoFacturaProveedor.EnProceso Then
        RowBuffer.CellStyle(14) = " EstadoEnProceso"
    ElseIf Factura.estado = EstadoFacturaProveedor.Saldada Then
        RowBuffer.CellStyle(14) = "EstadoSaldado"
    End If
    Exit Sub
err1:
End Sub

Private Sub grilla_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Set Factura = facturas.item(RowIndex)


    Dim i As Integer

    If Factura.tipoDocumentoContable = tipoDocumentoContable.notaCredito Then i = -1 Else i = 1
    With Factura
    
    If IsSomething(Factura.Proveedor) Then
        Values(1) = funciones.RazonSocialFormateada(Factura.Proveedor.RazonSocial)
        End If
        
        'Values(2) = Factura.NumeroFormateado
        
        Values(2) = enums.EnumTipoDocumentoContableShort(Factura.tipoDocumentoContable)
      '  vConfigFactura.TipoFactura
        Values(3) = Factura.configFactura.TipoFactura
        
        Values(4) = Factura.numero
        
        Values(5) = Factura.FEcha
        Values(6) = Factura.moneda.NombreCorto
        Values(7) = funciones.FormatearDecimales(Factura.Monto - Factura.TotalNetoGravadoDiscriminado(0)) * i
        Values(8) = funciones.FormatearDecimales(Factura.TotalIVA) * i
        Values(9) = funciones.FormatearDecimales(Factura.TotalNetoGravadoDiscriminado(0)) * i
        Values(10) = funciones.FormatearDecimales(Factura.totalPercepciones) * i
        Values(11) = funciones.FormatearDecimales(Factura.ImpuestoInterno) * i
        Values(12) = funciones.FormatearDecimales(Factura.Total) * i
        If Factura.cuentasContables.count > 0 Then
            Values(13) = Factura.cuentasContables.item(1).cuentas.codigo
        End If
        Values(14) = enums.enumEstadoFacturaProveedor(Factura.estado)

        If Factura.FormaPagoCuentaCorriente Then
            Values(15) = "Cta. Cte."
        Else
            Values(15) = "Contado"
        End If

        If LenB(Factura.OrdenesPagoId) > 0 Then
            Values(16) = Factura.OrdenesPagoId
        Else
            If Factura.OrdenPagoId > 0 Then
                Values(16) = Factura.OrdenPagoId
            End If
        End If
        

        
              
        Values(17) = Factura.UsuarioCarga.usuario
        
        Values(18) = Factura.TipoCambio
                    
        Values(19) = "(" & Val(m_Archivos.item(Factura.id)) & ")"


  End With
End Sub



Private Property Get ISuscriber_id() As String
    ISuscriber_id = vId
End Property
Private Function ISuscriber_Notificarse(EVENTO As clsEventoObserver) As Variant

    If EVENTO.EVENTO = agregar_ Then
        facturas.Add EVENTO.Elemento
        Me.grilla.ItemCount = facturas.count
    ElseIf EVENTO.EVENTO = modificar_ Then
        Dim rectmp As clsFacturaProveedor
        Dim tmp As clsFacturaProveedor
        Set tmp = EVENTO.Elemento

        For i = facturas.count To 1 Step -1
            If facturas(i).id = tmp.id Then
                Set rectmp = facturas(i)
                rectmp.id = tmp.id
                rectmp.estado = tmp.estado
                rectmp.Proveedor = tmp.Proveedor
                rectmp.FEcha = tmp.FEcha
                rectmp.ImpuestoInterno = tmp.ImpuestoInterno
                rectmp.cuentasContables = tmp.cuentasContables
                rectmp.IvaAplicado = tmp.IvaAplicado
                rectmp.percepciones = tmp.percepciones
                rectmp.Redondeo = tmp.Redondeo
                rectmp.Monto = tmp.Monto
                rectmp.numero = tmp.numero
                rectmp.FormaPagoCuentaCorriente = tmp.FormaPagoCuentaCorriente
                Set rectmp.moneda = tmp.moneda
                Me.grilla.RefreshRowIndex i
                Exit For
            End If
        Next
    End If
End Function




Private Sub mnuArchivos_Click()
    Dim archi As New frmArchivos2

    archi.Origen = OrigenArchivos.OA_FacturaProveedor
    archi.ObjetoId = Factura.id
    archi.caption = Factura.NumeroFormateado
    archi.Show

End Sub

Private Sub mnuEliminar_Click()
    If MsgBox("¿Está seguro de eliminar la " & Factura.NumeroFormateado & " de " & Factura.Proveedor.RazonSocial & "?", vbInformation + vbYesNo) = vbYes Then
        If DAOFacturaProveedor.Delete(Factura.id) Then
            MsgBox "Factura eliminada.", vbInformation
            llenarGrilla
        Else
            MsgBox "No se pudo eliminar la factura.", vbCritical
        End If
    End If
End Sub

Private Sub mnuPagarEnEfectivo_Click()
    If MsgBox("¿Está seguro de abonar en efectivo el comprobante " & Factura.NumeroFormateado & " de " & Factura.moneda.NombreCorto & " " & Factura.Total & "?", vbInformation + vbYesNo) = vbYes Then
        Dim fechaPago As String
        'fechaPago = InputBox("Ingrese la fecha de pago de factura", , Factura.FEcha)
        MsgBox "Se creará una OP con fecha " + CStr(Factura.FEcha)
        If IsDate(Factura.FEcha) Then
            If DAOFacturaProveedor.PagarEnEfectivo(Factura, Factura.FEcha, True) Then
                MsgBox "El pago de la factura ha sido registrado con la orden de pago Nº " & DAOOrdenPago.FindLast().id & ".", vbInformation
                llenarGrilla
                Me.txtComprobante.SetFocus
            Else
                MsgBox "No se pudo registrar el pago de la factura.", vbCritical
            End If
        Else
            MsgBox "Debe ingresar una fecha de pago válida.", vbExclamation
        End If
    End If
End Sub



Private Sub mnuScan_Click()
    On Error Resume Next
    Dim archivos As New classArchivos
    If archivos.escanearDocumento(OrigenArchivos.OA_FacturaProveedor, Factura.id) Then
        Set m_Archivos = DAOArchivo.GetCantidadArchivosPorReferencia(OA_FacturaProveedor)
        Me.grilla.RefreshRowIndex (Factura.id)

    End If

End Sub

Private Sub MnuVerOP_Click()

    Dim Orden As OrdenPago
    Set Orden = DAOOrdenPago.FindByFacturaId(Factura.id)
    Dim f22 As New frmCrearOrdenPago
    f22.Show
    f22.ReadOnly = True
    f22.Cargar Orden
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
Private Sub verHistorial_Click()
    If grilla.ItemCount > 0 Then
        SeleccionarFactura
        Factura.Historial = DaoFacturaProveedorHistorial.getAllByIdFactura(Factura.id)
        frmHistoriales.lista = Factura.Historial
        frmHistoriales.Show
    End If
End Sub
Private Sub SeleccionarFactura()
    On Error Resume Next
    Set Factura = facturas.item(grilla.RowIndex(grilla.row))
End Sub
