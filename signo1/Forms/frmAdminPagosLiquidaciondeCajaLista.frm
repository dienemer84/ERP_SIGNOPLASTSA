VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmAdminPagosLiquidaciondeCajaLista 
   Caption         =   "Listado de Liquidaciones de Caja"
   ClientHeight    =   10380
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   14415
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10380
   ScaleWidth      =   14415
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   16245
      _Version        =   786432
      _ExtentX        =   28654
      _ExtentY        =   3201
      _StockProps     =   79
      Caption         =   "Parámetros de búsqueda"
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
      Begin XtremeSuiteControls.GroupBox GroupBox 
         Height          =   735
         Index           =   2
         Left            =   11040
         TabIndex        =   25
         Top             =   0
         Width           =   5055
         _Version        =   786432
         _ExtentX        =   8916
         _ExtentY        =   1296
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.ProgressBar ProgressBar 
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
      Begin VB.PictureBox pic 
         Height          =   300
         Left            =   4920
         ScaleHeight     =   240
         ScaleWidth      =   285
         TabIndex        =   24
         Top             =   240
         Visible         =   0   'False
         Width           =   345
      End
      Begin XtremeSuiteControls.GroupBox GroupBox 
         Height          =   870
         Index           =   1
         Left            =   11040
         TabIndex        =   20
         Top             =   720
         Width           =   5055
         _Version        =   786432
         _ExtentX        =   8916
         _ExtentY        =   1526
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.PushButton btnExportar 
            Height          =   450
            Left            =   2040
            TabIndex        =   23
            Top             =   240
            Width           =   1335
            _Version        =   786432
            _ExtentX        =   2355
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "Exportar"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton btnBuscar 
            Height          =   450
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   1350
            _Version        =   786432
            _ExtentX        =   2381
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "Buscar"
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
         Begin XtremeSuiteControls.PushButton btnImprimir 
            Height          =   450
            Left            =   3600
            TabIndex        =   22
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
      Begin VB.TextBox txtNro 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         TabIndex        =   6
         Top             =   285
         Width           =   840
      End
      Begin XtremeSuiteControls.PushButton PushButton2 
         Height          =   330
         Left            =   11340
         TabIndex        =   1
         Top             =   2490
         Visible         =   0   'False
         Width           =   1245
         _Version        =   786432
         _ExtentX        =   2196
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "PushButton2"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.GroupBox GroupBox2 
         Height          =   1335
         Left            =   5760
         TabIndex        =   2
         Top             =   0
         Width           =   2355
         _Version        =   786432
         _ExtentX        =   4154
         _ExtentY        =   2355
         _StockProps     =   79
         Caption         =   "Estado Proveedor"
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
         Begin XtremeSuiteControls.CheckBox chkContado 
            Height          =   195
            Left            =   405
            TabIndex        =   3
            Top             =   360
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
            TabIndex        =   4
            Top             =   600
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
            TabIndex        =   5
            Top             =   900
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
         Left            =   1305
         TabIndex        =   7
         Top             =   735
         Width           =   3525
         _Version        =   786432
         _ExtentX        =   6218
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Enabled         =   0   'False
         Text            =   "cboProveedores"
      End
      Begin XtremeSuiteControls.PushButton btnClearProveedor 
         Height          =   375
         Left            =   4920
         TabIndex        =   8
         Top             =   720
         Width           =   405
         _Version        =   786432
         _ExtentX        =   714
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "X"
         BackColor       =   12632256
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cboEstado 
         Height          =   315
         Left            =   1305
         TabIndex        =   9
         Top             =   1200
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
         Height          =   375
         Left            =   4920
         TabIndex        =   10
         Top             =   1170
         Width           =   405
         _Version        =   786432
         _ExtentX        =   714
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "X"
         BackColor       =   12632256
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.GroupBox GroupBox 
         Height          =   1335
         Index           =   0
         Left            =   8280
         TabIndex        =   14
         Top             =   0
         Width           =   2655
         _Version        =   786432
         _ExtentX        =   4683
         _ExtentY        =   2355
         _StockProps     =   79
         Caption         =   "Fecha de Creación"
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
         Begin XtremeSuiteControls.DateTimePicker dtpDesde 
            Height          =   315
            Left            =   900
            TabIndex        =   15
            Top             =   360
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
            Left            =   885
            TabIndex        =   16
            Top             =   855
            Width           =   1470
            _Version        =   786432
            _ExtentX        =   2593
            _ExtentY        =   556
            _StockProps     =   68
            CheckBox        =   -1  'True
            Format          =   1
         End
         Begin XtremeSuiteControls.Label Label5 
            Height          =   195
            Left            =   360
            TabIndex        =   18
            Top             =   405
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
            Left            =   375
            TabIndex        =   17
            Top             =   915
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   195
         Left            =   180
         TabIndex        =   13
         Top             =   330
         Width           =   1080
         _Version        =   786432
         _ExtentX        =   1905
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Nº Liquidación:"
         Alignment       =   1
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label lbl 
         Height          =   195
         Left            =   480
         TabIndex        =   12
         Top             =   780
         Width           =   780
         _Version        =   786432
         _ExtentX        =   1376
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Proveedor:"
         Enabled         =   0   'False
         Alignment       =   1
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label4 
         Height          =   195
         Left            =   720
         TabIndex        =   11
         Top             =   1260
         Width           =   540
         _Version        =   786432
         _ExtentX        =   953
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Estado:"
         BackColor       =   12632256
         AutoSize        =   -1  'True
      End
   End
   Begin GridEX20.GridEX gridOrdenes 
      Height          =   5505
      Left            =   120
      TabIndex        =   19
      Top             =   2040
      Width           =   16215
      _ExtentX        =   28601
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
      Column(1)       =   "frmAdminPagosLiquidaciondeCajaLista.frx":0000
      Column(2)       =   "frmAdminPagosLiquidaciondeCajaLista.frx":00C8
      Column(3)       =   "frmAdminPagosLiquidaciondeCajaLista.frx":016C
      Column(4)       =   "frmAdminPagosLiquidaciondeCajaLista.frx":0210
      Column(5)       =   "frmAdminPagosLiquidaciondeCajaLista.frx":02B4
      Column(6)       =   "frmAdminPagosLiquidaciondeCajaLista.frx":0358
      Column(7)       =   "frmAdminPagosLiquidaciondeCajaLista.frx":03FC
      Column(8)       =   "frmAdminPagosLiquidaciondeCajaLista.frx":04A0
      Column(9)       =   "frmAdminPagosLiquidaciondeCajaLista.frx":0544
      FormatStylesCount=   7
      FormatStyle(1)  =   "frmAdminPagosLiquidaciondeCajaLista.frx":05E8
      FormatStyle(2)  =   "frmAdminPagosLiquidaciondeCajaLista.frx":0720
      FormatStyle(3)  =   "frmAdminPagosLiquidaciondeCajaLista.frx":07D0
      FormatStyle(4)  =   "frmAdminPagosLiquidaciondeCajaLista.frx":0884
      FormatStyle(5)  =   "frmAdminPagosLiquidaciondeCajaLista.frx":095C
      FormatStyle(6)  =   "frmAdminPagosLiquidaciondeCajaLista.frx":0A14
      FormatStyle(7)  =   "frmAdminPagosLiquidaciondeCajaLista.frx":0AF4
      ImageCount      =   0
      PrinterProperties=   "frmAdminPagosLiquidaciondeCajaLista.frx":0B14
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
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuVer 
         Caption         =   "Ver"
      End
      Begin VB.Menu mnuImprimir 
         Caption         =   "Imprimir"
      End
      Begin VB.Menu mnuHistorial 
         Caption         =   "Ver Historial"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmAdminPagosLiquidaciondeCajaLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Implements ISuscriber

Dim ids As String
Private liquidaciones As New Collection
Private LiquidacionCaja As clsLiquidacionCaja
Private fac As clsFacturaProveedor


Private Sub btnExportar_Click()
    Me.ProgressBar.Visible = True

    If IsSomething(liquidaciones) Then
        If Not DAOLiquidacionCaja.ExportarColeccion(liquidaciones, Me.ProgressBar) Then GoTo err1
    End If

    Me.ProgressBar.Visible = False

    Exit Sub
err1:
    MsgBox "Se produjo un error al exportar!", vbCritical, "Error"
End Sub


Private Sub Form_Load()
    Customize Me
    GridEXHelper.CustomizeGrid Me.gridOrdenes, True
    DAOProveedor.llenarComboXtremeSuite Me.cboProveedores, True, True, True
    Me.cboProveedores.ListIndex = -1
    Me.dtpDesde.value = Year(Now) & "-01-01"
    Me.dtpHasta.value = Now
    Me.gridOrdenes.ItemCount = 0
    GridEXHelper.AutoSizeColumns Me.gridOrdenes
    ids = funciones.CreateGUID
    '    Channel.AgregarSuscriptor Me, OrdenesPago_

    Me.cboEstado.Clear
    Me.cboEstado.AddItem enums.EnumEstadoOrdenPago(EstadoOrdenPago.EstadoOrdenPago_pendiente)
    Me.cboEstado.ItemData(Me.cboEstado.NewIndex) = EstadoOrdenPago.EstadoOrdenPago_pendiente
    Me.cboEstado.AddItem enums.EnumEstadoOrdenPago(EstadoOrdenPago.EstadoOrdenPago_Aprobada)
    Me.cboEstado.ItemData(Me.cboEstado.NewIndex) = EstadoOrdenPago.EstadoOrdenPago_Aprobada
    Me.cboEstado.AddItem enums.EnumEstadoOrdenPago(EstadoOrdenPago.EstadoOrdenPago_Anulada)
    Me.cboEstado.ItemData(Me.cboEstado.NewIndex) = EstadoOrdenPago.EstadoOrdenPago_Anulada

    'llenarLista
    
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    Me.gridOrdenes.Width = Me.ScaleWidth - 150
    Me.gridOrdenes.Height = Me.ScaleHeight - 2000
    Me.GroupBox1.Width = Me.gridOrdenes.Width - 100

    GridEXHelper.AutoSizeColumns Me.gridOrdenes
End Sub


'Private Sub Form_Terminate()
'    Channel.RemoverSuscripcionTotal Me
'End Sub

'Private Sub Form_Unload(Cancel As Integer)
'    Channel.RemoverSuscripcionTotal Me
'End Sub

Private Sub btnClearProveedor_Click()
    Me.cboProveedores.ListIndex = -1
End Sub

Private Sub btnBuscar_Click()
    If (Me.chkContado.value = xtpChecked Or Me.chkCtaCte.value = xtpChecked Or Me.chkEliminado.value = xtpGrayed) Then llenarLista Else Me.gridOrdenes.ItemCount = 0

End Sub

Private Sub btnImprimir_Click()

    Dim pro As String
    If Me.cboProveedores.ListIndex > -1 Then
        pro = " Proveedor: " & Me.cboProveedores.Text
    End If

    With Me.gridOrdenes.PrinterProperties

        .FitColumns = True
        .RepeatHeaders = True
        .Orientation = jgexPPLandscape
        .HeaderString(jgexHFCenter) = "Listado de Liquidaciones de Caja "
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

Private Sub gridOrdenes_DblClick()
    gridOrdenes_SelectionChange
    mnuVer_Click
End Sub

Private Sub PushButton2_Click()
    conectar.BeginTransaction
    Dim newcol As New Collection

    Dim nop As OrdenPago

    For Each nop In liquidaciones
        If (nop.StaticTotalOrigenes + nop.StaticTotalRetenido) = 0 And nop.estado = EstadoOrdenPago_Aprobada Then
            newcol.Add nop

        End If
    Next
    Dim q As String
    Dim opeCaja As operacion

    For Each nop In newcol


        If nop.FacturasProveedor.count = 1 Then    'se pago una sola fc
            Set opeCaja = New operacion
            opeCaja.Pertenencia = OrigenOperacion.caja
            opeCaja.Monto = nop.StaticTotal
            Set opeCaja.moneda = fac.moneda
            opeCaja.FechaOperacion = nop.FEcha


            opeCaja.FechaCarga = Now
            Set opeCaja.caja = DAOCaja.FindById(1)
            opeCaja.EntradaSalida = OPSalida

            If Not DAOOperacion.Save(opeCaja) Then GoTo E
            opeCaja.Id = conectar.UltimoId2
            q = "INSERT INTO ordenes_pago_operaciones VALUES (" & nop.Id & ", " & opeCaja.Id & ")"
            If Not conectar.execute(q) Then GoTo E
            q = "update ordenes_pago set static_total_origen=" & opeCaja.Monto & " where id=" & nop.Id
            If Not conectar.execute(q) Then GoTo E


        End If

    Next
    conectar.CommitTransaction
    Exit Sub
E:
    conectar.RollBackTransaction

End Sub


Private Sub llenarLista()
    Dim filter As String
    filter = "1 = 1"

    If Me.cboProveedores.ListIndex > -1 Then
        filter = filter & " AND AdminComprasFacturasProveedores.id_proveedor = " & Me.cboProveedores.ItemData(Me.cboProveedores.ListIndex)
    End If

    If LenB(Me.txtNro.Text) > 0 Then
        filter = filter & " AND  liquidaciones_caja.numero_liq  = " & Val(Me.txtNro.Text)
    End If

    Dim filtroor As String

    If Not IsNull(Me.dtpDesde.value) Then
        filter = filter & " AND liquidaciones_caja.fecha >= " & conectar.Escape(Me.dtpDesde.value)
    End If

    If Not IsNull(Me.dtpHasta.value) Then
        filter = filter & " AND liquidaciones_caja.fecha <= " & conectar.Escape(Me.dtpHasta.value)
    End If

    If Me.cboEstado.ListIndex > -1 Then
        filter = filter & " AND liquidaciones_caja.estado = " & Me.cboEstado.ItemData(Me.cboEstado.ListIndex)
    End If

    If LenB(filtroor) > 0 Then
        filtroor = " AND (" & Right(filtroor, Len(filtroor) - 3) & " )"
        filter = filter & filtroor
    End If

    Me.gridOrdenes.ItemCount = 0
    Set liquidaciones = DAOLiquidacionCaja.FindAll(filter, "liquidaciones_caja.id DESC")
    Me.gridOrdenes.ItemCount = liquidaciones.count

    Me.caption = "Listado de Liquidaciones de Caja" & " [Cantidad: " & liquidaciones.count & "]"

End Sub


Private Sub gridOrdenes_ColumnHeaderClick(ByVal Column As GridEX20.JSColumn)
    GridEXHelper.ColumnHeaderClick Me.gridOrdenes, Column
End Sub


Private Sub gridOrdenes_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If liquidaciones.count > 0 Then
        gridOrdenes_SelectionChange
        If Button = 2 Then
            Me.mnuAprobar.Enabled = (LiquidacionCaja.estado = EstadoLiquidacionCaja_pendiente)
            Me.mnuEditar.Enabled = (LiquidacionCaja.estado = EstadoLiquidacionCaja_pendiente)
            
            'OCULTO LA OPCION DE ANULAR QUE NO ESTÁ DESARROLLADA (DNEMER 30.05.2023)
            'Me.mnuAnular.Enabled = Not (LiquidacionCaja.estado = EstadoLiquidacionCaja_Anulada)

            Me.PopupMenu menu

        End If
    End If
End Sub


Private Sub gridOrdenes_RowFormat(RowBuffer As GridEX20.JSRowData)
    If RowBuffer.RowIndex > 0 And liquidaciones.count > 0 Then
        Set LiquidacionCaja = liquidaciones.item(RowBuffer.RowIndex)
        If LiquidacionCaja.estado = EstadoLiquidacionCaja.EstadoLiquidacionCaja_Aprobada Then
            RowBuffer.CellStyle(9) = "aprobada"
        ElseIf LiquidacionCaja.estado = EstadoLiquidacionCaja_Anulada Then
            RowBuffer.RowStyle = "anulada2"

            RowBuffer.CellStyle(9) = "anulada"
        ElseIf LiquidacionCaja.estado = EstadoLiquidacionCaja_pendiente Then
            RowBuffer.CellStyle(9) = "pendiente"
        End If
    End If
End Sub


Private Sub gridOrdenes_SelectionChange()
    On Error Resume Next
    Set LiquidacionCaja = liquidaciones.item(gridOrdenes.RowIndex(gridOrdenes.row))
End Sub


Private Sub gridOrdenes_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex > 0 And liquidaciones.count > 0 Then

        Set LiquidacionCaja = liquidaciones.item(RowIndex)

        Values(1) = LiquidacionCaja.NumeroLiq
        Values(2) = LiquidacionCaja.FEcha

        Values(3) = LiquidacionCaja.moneda.NombreCorto

        Values(4) = Replace(FormatCurrency(funciones.FormatearDecimales(LiquidacionCaja.StaticTotalOrigenes)), "$", "")
        Values(5) = Replace(FormatCurrency(funciones.FormatearDecimales(LiquidacionCaja.StaticTotalRetenido)), "$", "")
        Values(6) = Replace(FormatCurrency(funciones.FormatearDecimales(LiquidacionCaja.StaticTotalOrigenes + LiquidacionCaja.StaticTotalRetenido)), "$", "")

        If LiquidacionCaja.EsParaFacturaProveedor Then
            Set fac = LiquidacionCaja.FacturasProveedor.item(1)
            Values(7) = "Factura Proveedor"
            Values(8) = "VARIOS"
        Else
            Values(7) = "Cuenta Contable"
            If IsSomething(LiquidacionCaja.CuentaContable) Then
                Values(8) = LiquidacionCaja.CuentaContable.nombre & " (" & LiquidacionCaja.CuentaContable.codigo & ")"
            End If
        End If

        Values(9) = enums.EnumEstadoOrdenPago(LiquidacionCaja.estado)
    End If
End Sub

Private Sub mnuAnular_Click()
    If MsgBox("¿Desea anular la Liquidación?", vbQuestion + vbYesNo) = vbYes Then
        If DAOLiquidacionCaja.Delete(LiquidacionCaja.Id, True) Then
            MsgBox "Anulación Exitosa.", vbInformation + vbOKOnly
            Me.gridOrdenes.ItemCount = 0
            liquidaciones.remove CStr(LiquidacionCaja.Id)
            Me.gridOrdenes.ItemCount = liquidaciones.count
            btnBuscar_Click
        Else
            MsgBox "No se pudo anular la Liquidación.", vbCritical + vbOKOnly
        End If
    End If
End Sub


Private Sub mnuAprobar_Click()
    If DAOLiquidacionCaja.aprobar(LiquidacionCaja, True) Then
        MsgBox "Aprobación Exitosa!", vbInformation + vbOKOnly
        Me.gridOrdenes.RefreshRowIndex Me.gridOrdenes.RowIndex(Me.gridOrdenes.row)
        btnBuscar_Click
    Else
        MsgBox "Error, no se aprobó la OP!", vbCritical + vbOKOnly
    End If

End Sub

Private Sub mnuEditar_Click()
'    MsgBox ("Funcion en desarrollo")
        Dim f22 As New frmAdminPagosLiqCajaListaDG
        f22.Show
'        Dim LiquidacionCaja As clsLiquidacionCaja
        f22.Cargar LiquidacionCaja
End Sub


'Private Sub mnuHistorial_Click()
'    Dim F As New frmHistorico
'    F.Configurar "orden_pago_historial", Orden.Id, "orden de pago Nro " & Orden.Id
'    F.Show
'End Sub

Private Sub mnuImprimir_Click()

    gridOrdenes_SelectionChange

    If Not DAOLiquidacionCaja.PrintLiq(LiquidacionCaja, Me.pic) Then GoTo err1

    Exit Sub
err1:
End Sub

Private Sub mnuVer_Click()
'    Dim f22 As New frmAdminPagosLiquidaciondeCajaCrear
'    f22.Show
'    f22.ReadOnly = True
'    f22.Cargar LiquidacionCaja

    Dim f22 As New frmAdminPagosLiqCajaListaDG
    f22.Show
    f22.ReadOnly = True
    f22.Cargar LiquidacionCaja

End Sub


Private Property Get ISuscriber_id() As String
    ISuscriber_id = ids
End Property

