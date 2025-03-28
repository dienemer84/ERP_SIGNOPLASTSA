VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.OCX"
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
         TabIndex        =   13
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
            TabIndex        =   14
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
         Index           =   0
         Left            =   5040
         ScaleHeight     =   240
         ScaleWidth      =   285
         TabIndex        =   12
         Top             =   240
         Visible         =   0   'False
         Width           =   345
      End
      Begin XtremeSuiteControls.GroupBox GroupBox 
         Height          =   870
         Index           =   1
         Left            =   11040
         TabIndex        =   8
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
            TabIndex        =   11
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
            Default         =   -1  'True
            Height          =   450
            Left            =   120
            TabIndex        =   9
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
            TabIndex        =   10
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
         TabIndex        =   2
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
      Begin XtremeSuiteControls.ComboBox cboEstado 
         Height          =   315
         Left            =   1305
         TabIndex        =   3
         Top             =   720
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
         TabIndex        =   4
         Top             =   690
         Width           =   405
         _Version        =   786432
         _ExtentX        =   714
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "X"
         BackColor       =   12632256
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.GroupBox GroFechaComprobante 
         Height          =   1335
         Index           =   1
         Left            =   5880
         TabIndex        =   15
         Top             =   240
         Width           =   4695
         _Version        =   786432
         _ExtentX        =   8281
         _ExtentY        =   2355
         _StockProps     =   79
         Caption         =   "Fecha Liquidación"
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
         Begin XtremeSuiteControls.Label lblHasta 
            Height          =   195
            Index           =   1
            Left            =   2400
            TabIndex        =   21
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
         Begin XtremeSuiteControls.Label lblRango 
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   19
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
         Left            =   180
         TabIndex        =   6
         Top             =   330
         Width           =   1035
         _Version        =   786432
         _ExtentX        =   1826
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "N° Liquidación"
         Alignment       =   1
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label4 
         Height          =   195
         Left            =   720
         TabIndex        =   5
         Top             =   780
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
      TabIndex        =   7
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
      Column(2)       =   "frmAdminPagosLiquidaciondeCajaLista.frx":01A4
      Column(3)       =   "frmAdminPagosLiquidaciondeCajaLista.frx":02E4
      Column(4)       =   "frmAdminPagosLiquidaciondeCajaLista.frx":042C
      Column(5)       =   "frmAdminPagosLiquidaciondeCajaLista.frx":0574
      Column(6)       =   "frmAdminPagosLiquidaciondeCajaLista.frx":0688
      Column(7)       =   "frmAdminPagosLiquidaciondeCajaLista.frx":07C8
      Column(8)       =   "frmAdminPagosLiquidaciondeCajaLista.frx":0908
      Column(9)       =   "frmAdminPagosLiquidaciondeCajaLista.frx":0A50
      FormatStylesCount=   10
      FormatStyle(1)  =   "frmAdminPagosLiquidaciondeCajaLista.frx":0B98
      FormatStyle(2)  =   "frmAdminPagosLiquidaciondeCajaLista.frx":0CD0
      FormatStyle(3)  =   "frmAdminPagosLiquidaciondeCajaLista.frx":0D80
      FormatStyle(4)  =   "frmAdminPagosLiquidaciondeCajaLista.frx":0E34
      FormatStyle(5)  =   "frmAdminPagosLiquidaciondeCajaLista.frx":0F0C
      FormatStyle(6)  =   "frmAdminPagosLiquidaciondeCajaLista.frx":0FC4
      FormatStyle(7)  =   "frmAdminPagosLiquidaciondeCajaLista.frx":10A4
      FormatStyle(8)  =   "frmAdminPagosLiquidaciondeCajaLista.frx":10C4
      FormatStyle(9)  =   "frmAdminPagosLiquidaciondeCajaLista.frx":1178
      FormatStyle(10) =   "frmAdminPagosLiquidaciondeCajaLista.frx":1230
      ImageCount      =   0
      PrinterProperties=   "frmAdminPagosLiquidaciondeCajaLista.frx":12E4
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   840
      Top             =   7920
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


Private Sub cboRangos_Click()
    funciones.CalculateDateRange Me.cboRangos, Me.dtpDesde(1), Me.dtpHasta(1)
End Sub


Private Sub Form_Load()
    Customize Me
    GridEXHelper.CustomizeGrid Me.gridOrdenes, True
    

    Me.dtpDesde(1).value = Year(Now) & "-01-01"
    Me.dtpHasta(1).value = Now
    Me.gridOrdenes.ItemCount = 0
    GridEXHelper.AutoSizeColumns Me.gridOrdenes
    ids = funciones.CreateGUID
    
   funciones.FillComboBoxDateRanges Me.cboRangos
    
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


Private Sub btnBuscar_Click()
    llenarLista

End Sub

Private Sub btnImprimir_Click()

    With Me.gridOrdenes.PrinterProperties

        .FitColumns = True
        .RepeatHeaders = True
        .Orientation = jgexPPLandscape
        .HeaderString(jgexHFCenter) = "Listado de Liquidaciones de Caja "
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

    If LenB(Me.txtNro.Text) > 0 Then
        filter = filter & " AND  liquidaciones_caja.numero_liq  = " & Val(Me.txtNro.Text)
    End If

    Dim filtroor As String

    If Not IsNull(Me.dtpDesde(1).value) Then
        filter = filter & " AND liquidaciones_caja.fecha >= " & conectar.Escape(Me.dtpDesde(1).value)
    End If

    If Not IsNull(Me.dtpHasta(1).value) Then
        filter = filter & " AND liquidaciones_caja.fecha <= " & conectar.Escape(Me.dtpHasta(1).value)
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
            Me.mnuImprimir.Enabled = (LiquidacionCaja.estado = EstadoLiquidacionCaja_Aprobada)
            Me.mnuEditar.Enabled = (LiquidacionCaja.estado = EstadoLiquidacionCaja_pendiente)
            
            'OCULTO LA OPCION DE ANULAR QUE NO ESTÃ DESARROLLADA (DNEMER 30.05.2023)
            'Me.mnuAnular.Enabled = Not (LiquidacionCaja.estado = EstadoLiquidacionCaja_Anulada)

            Me.PopupMenu menu

        End If
    End If
End Sub


Private Sub gridOrdenes_RowFormat(RowBuffer As GridEX20.JSRowData)
    If RowBuffer.rowIndex > 0 And liquidaciones.count > 0 Then
        Set LiquidacionCaja = liquidaciones.item(RowBuffer.rowIndex)
        If LiquidacionCaja.estado = EstadoLiquidacionCaja.EstadoLiquidacionCaja_Aprobada Then
            RowBuffer.CellStyle(9) = "Aprobada"
        ElseIf LiquidacionCaja.estado = EstadoLiquidacionCaja_Anulada Then
            RowBuffer.RowStyle = "Anulada"

            RowBuffer.CellStyle(9) = "Anulada"
        ElseIf LiquidacionCaja.estado = EstadoLiquidacionCaja_pendiente Then
            RowBuffer.CellStyle(9) = "Pendiente"
        End If
    End If
End Sub


Private Sub gridOrdenes_SelectionChange()
    On Error Resume Next
    Set LiquidacionCaja = liquidaciones.item(gridOrdenes.rowIndex(gridOrdenes.row))
End Sub


Private Sub gridOrdenes_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If rowIndex > 0 And liquidaciones.count > 0 Then

        Set LiquidacionCaja = liquidaciones.item(rowIndex)

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
        MsgBox "Aprobación éxitosa!", vbInformation + vbOKOnly
        Me.gridOrdenes.RefreshRowIndex Me.gridOrdenes.rowIndex(Me.gridOrdenes.row)
        btnBuscar_Click
    Else
        MsgBox "Error, no se aprobó la OP!", vbCritical + vbOKOnly
    End If

End Sub

Private Sub mnuEditar_Click()
        Dim f22 As New frmAdminPagosLiqCajaListaDG
        f22.Show
        f22.Cargar LiquidacionCaja
End Sub


Private Sub mnuImprimir_Click()
    On Error GoTo err4
    '''gridOrdenes_SelectionChange
    
    Me.CommonDialog.ShowPrinter
    
    If Not DAOLiquidacionCaja.PrintLiq(LiquidacionCaja) Then GoTo err4
        Exit Sub
        
err4:

End Sub

Private Sub mnuVer_Click()
    Dim f22 As New frmAdminPagosLiqCajaListaDG
    f22.Show
    f22.ReadOnly = True
    f22.Cargar LiquidacionCaja

End Sub


Private Property Get ISuscriber_id() As String
    ISuscriber_id = ids
End Property

