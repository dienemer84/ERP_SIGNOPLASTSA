VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GRIDEX20.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmOrdenesPago 
   Caption         =   "Ordenes de Pago"
   ClientHeight    =   7050
   ClientLeft      =   8445
   ClientTop       =   3465
   ClientWidth     =   12660
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
   ScaleHeight     =   7050
   ScaleWidth      =   12660
   Begin GridEX20.GridEX gridOrdenes 
      Height          =   5505
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   12510
      _ExtentX        =   22066
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
      Column(2)       =   "frmOrdenesPago.frx":0178
      Column(3)       =   "frmOrdenesPago.frx":02AC
      Column(4)       =   "frmOrdenesPago.frx":03A0
      Column(5)       =   "frmOrdenesPago.frx":04D8
      Column(6)       =   "frmOrdenesPago.frx":06E0
      Column(7)       =   "frmOrdenesPago.frx":0814
      Column(8)       =   "frmOrdenesPago.frx":0900
      Column(9)       =   "frmOrdenesPago.frx":09F4
      FormatStylesCount=   10
      FormatStyle(1)  =   "frmOrdenesPago.frx":0AE8
      FormatStyle(2)  =   "frmOrdenesPago.frx":0C10
      FormatStyle(3)  =   "frmOrdenesPago.frx":0CC0
      FormatStyle(4)  =   "frmOrdenesPago.frx":0D74
      FormatStyle(5)  =   "frmOrdenesPago.frx":0E4C
      FormatStyle(6)  =   "frmOrdenesPago.frx":0F04
      FormatStyle(7)  =   "frmOrdenesPago.frx":0FE4
      FormatStyle(8)  =   "frmOrdenesPago.frx":1098
      FormatStyle(9)  =   "frmOrdenesPago.frx":114C
      FormatStyle(10) =   "frmOrdenesPago.frx":122C
      ImageCount      =   0
      PrinterProperties=   "frmOrdenesPago.frx":12E4
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1350
      Left            =   45
      TabIndex        =   1
      Top             =   90
      Width           =   12525
      _Version        =   786432
      _ExtentX        =   22093
      _ExtentY        =   2381
      _StockProps     =   79
      Caption         =   "Parámetros de búsqueda"
      UseVisualStyle  =   -1  'True
      Begin VB.PictureBox pic 
         Height          =   540
         Left            =   11745
         ScaleHeight     =   480
         ScaleWidth      =   300
         TabIndex        =   18
         Top             =   195
         Visible         =   0   'False
         Width           =   360
      End
      Begin XtremeSuiteControls.PushButton PushButton2 
         Height          =   330
         Left            =   11340
         TabIndex        =   12
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
         Height          =   1095
         Left            =   4890
         TabIndex        =   8
         Top             =   105
         Width           =   2355
         _Version        =   786432
         _ExtentX        =   4154
         _ExtentY        =   1931
         _StockProps     =   79
         Caption         =   "Estado Proveedor"
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.CheckBox chkContado 
            Height          =   195
            Left            =   405
            TabIndex        =   9
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
            TabIndex        =   10
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
            TabIndex        =   11
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
      Begin XtremeSuiteControls.PushButton cmdBuscar 
         Default         =   -1  'True
         Height          =   450
         Left            =   9615
         TabIndex        =   3
         Top             =   270
         Width           =   1350
         _Version        =   786432
         _ExtentX        =   2381
         _ExtentY        =   794
         _StockProps     =   79
         Caption         =   "Buscar"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cboProveedores 
         Height          =   315
         Left            =   945
         TabIndex        =   4
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
         TabIndex        =   5
         Top             =   630
         Width           =   300
         _Version        =   786432
         _ExtentX        =   529
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "X"
         BackColor       =   12632256
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton cmdImprimir 
         Height          =   450
         Left            =   9615
         TabIndex        =   13
         Top             =   765
         Width           =   1350
         _Version        =   786432
         _ExtentX        =   2381
         _ExtentY        =   794
         _StockProps     =   79
         Caption         =   "Imprimir"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.DateTimePicker dtpDesde 
         Height          =   315
         Left            =   7845
         TabIndex        =   14
         Top             =   180
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
         Left            =   7830
         TabIndex        =   15
         Top             =   555
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
         TabIndex        =   19
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
         TabIndex        =   21
         Top             =   960
         Width           =   300
         _Version        =   786432
         _ExtentX        =   529
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "X"
         BackColor       =   12632256
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label4 
         Height          =   195
         Left            =   360
         TabIndex        =   20
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
      Begin XtremeSuiteControls.Label Label6 
         Height          =   195
         Left            =   7320
         TabIndex        =   17
         Top             =   615
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
         Left            =   7305
         TabIndex        =   16
         Top             =   225
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
         TabIndex        =   7
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
         TabIndex        =   6
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
Attribute VB_Name = "frmOrdenesPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements ISuscriber

Dim ids As String
Private ordenes As New Collection
Private Orden As OrdenPago
Private fac As clsFacturaProveedor

Private Sub btnClearProveedor_Click()
    Me.cboProveedores.ListIndex = -1
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
        .HeaderString(jgexHFCenter) = "Listado de Ordens de Pago "
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

    Me.dtpDesde.value = Year(Now) & "-01-01"

    Me.dtpHasta.value = Now
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

    If Not IsNull(Me.dtpDesde.value) Then
        filter = filter & " AND ordenes_pago.fecha >= " & conectar.Escape(Me.dtpDesde.value)
    End If

    If Not IsNull(Me.dtpHasta.value) Then
        filter = filter & " AND ordenes_pago.fecha <= " & conectar.Escape(Me.dtpHasta.value)
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

    If ordenes.count = 1 And LenB(Me.txtNro.text) > 0 Then
        Set Orden = ordenes(1)
        If Orden.estado <> EstadoOrdenPago_Anulada Then gridOrdenes_DblClick
    End If
End Sub
Private Sub Form_Resize()
    On Error Resume Next
    Me.gridOrdenes.Width = Me.ScaleWidth - 50
    Me.gridOrdenes.Height = Me.ScaleHeight - Me.gridOrdenes.Top

    Me.GroupBox1.Width = Me.gridOrdenes.Width - 100
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

Private Sub gridOrdenes_DblClick()
    gridOrdenes_SelectionChange
    mnuVer_Click
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
    If RowBuffer.RowIndex > 0 And ordenes.count > 0 Then
        Set Orden = ordenes.item(RowBuffer.RowIndex)
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

Private Sub gridOrdenes_SelectionChange()
    On Error Resume Next
    Set Orden = ordenes.item(gridOrdenes.RowIndex(gridOrdenes.row))
End Sub

Private Sub gridOrdenes_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex > 0 And ordenes.count > 0 Then
        Set Orden = ordenes.item(RowIndex)
        Values(1) = Orden.id
        Values(2) = Orden.FEcha

        Values(3) = Orden.moneda.NombreCorto

        Values(4) = funciones.FormatearDecimales(Orden.StaticTotalOrigenes)
        Values(5) = funciones.FormatearDecimales(Orden.StaticTotalRetenido)
        Values(6) = funciones.FormatearDecimales(Orden.StaticTotalOrigenes + Orden.StaticTotalRetenido)

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
            If ordenes(i).id = tmp.id Then
                Set Orden = ordenes(i)
                Orden.id = tmp.id
                Orden.estado = tmp.estado
                Me.gridOrdenes.RefreshRowIndex i
                Exit For
            End If
        Next
    End If
End Function

Private Sub mnuAnular_Click()
    If MsgBox("¿Desea borrar la OP?", vbQuestion + vbYesNo) = vbYes Then
        If DAOOrdenPago.Delete(Orden.id, True) Then
            Me.gridOrdenes.ItemCount = 0
            ordenes.remove CStr(Orden.id)
            Me.gridOrdenes.ItemCount = ordenes.count
        Else
            MsgBox "No se pudo borrar.", vbCritical + vbOKOnly
        End If
    End If
End Sub

Private Sub mnuAprobar_Click()
    If DAOOrdenPago.aprobar(Orden, True) Then
        MsgBox "Aprobación Exitosa!", vbInformation + vbOKOnly
        Me.gridOrdenes.RefreshRowIndex Me.gridOrdenes.RowIndex(Me.gridOrdenes.row)
    Else
        MsgBox "Error, no se aprobó la OP!", vbCritical + vbOKOnly
    End If

End Sub

Private Sub mnuEditar_Click()
    Dim f22 As New frmCrearOrdenPago
    f22.Show
    f22.Cargar Orden
End Sub

Private Sub mnuHistorial_Click()
Dim c As New Collection

Dim F As New frmHistorico
F.Configurar "orden_pago_historial", Orden.id, "orden de pago Nro " & Orden.id
F.Show
End Sub

Private Sub mnuImprimir_Click()

    If Not DAOOrdenPago.PrintOP(Orden, Me.pic) Then GoTo err1

    Exit Sub
err1:
End Sub

Private Sub Imprimir()
    With drpOrdenPago.Sections("seccion").Controls

        .item("lblTitulo").caption = "SIGNOPLAST S.A. - Orden de Pago Nº " & Orden.id
        .item("lblFecha").caption = Orden.FEcha

        If Orden.FacturasProveedor.count > 0 Then
            .item("lblProveedor").caption = Orden.FacturasProveedor(1).Proveedor.RazonSocial
        End If

        .item("lblAlicuota").caption = Orden.alicuota & "%"

        Dim cert As CertificadoRetencion
        Set cert = DAOCertificadoRetencion.FindByOrdenPago(Orden.id)
        If IsSomething(cert) Then
            .item("lblCertificadoIIBB").caption = cert.id
        Else
            .item("lblCertificadoIIBB").caption = "NO POSEE"
        End If

        .item("lblMoneda").caption = Orden.moneda.NombreCorto & " " & Orden.moneda.NombreLargo



        Set Orden.FacturasProveedor = DAOFacturaProveedor.FindAllByOrdenPago(Orden.id)
        Dim F As clsFacturaProveedor
        Dim facs As New Collection
        For Each F In Orden.FacturasProveedor
            'facs.Add F.NumeroFormateado & String$(8, " del ") & F.FEcha & String$(8, " por ") & F.Moneda.NombreCorto & " " & F.Total

            facs.Add F.NumeroFormateado & " del " & F.FEcha & " por " & F.moneda.NombreCorto & " " & F.Total

        Next F
        If facs.count = 0 Then
            .item("lblFacturas").caption = "NO POSEE FACTURAS"
        Else
            .item("lblFacturas").caption = funciones.JoinCollectionValues(facs, "  ||  ")
        End If


        Dim cheq As cheque
        Dim tmpCol As New Collection
        For Each cheq In Orden.ChequesPropios
            tmpCol.Add cheq.numero & String$(8, " ") & cheq.Banco.nombre & String$(24, " ") & cheq.FechaVencimiento & String$(8, " ") & cheq.moneda.NombreCorto & " " & cheq.Monto
        Next cheq
        If tmpCol.count = 0 Then
            .item("lblChequesPropios").caption = "NO POSEE CHEQUES PROPIOS"
        Else
            .item("lblChequesPropios").caption = funciones.JoinCollectionValues(tmpCol, " - ")
        End If


        Set tmpCol = New Collection
        For Each cheq In Orden.ChequesTerceros
            tmpCol.Add cheq.numero & String$(8, " ") & cheq.Banco.nombre & String$(16, " ") & cheq.FechaVencimiento & String$(8, " ") & cheq.moneda.NombreCorto & " " & cheq.Monto
        Next cheq
        If tmpCol.count = 0 Then
            .item("lblChequesTerceros").caption = "NO POSEE CHEQUES DE 3ros"
        Else
            .item("lblChequesTerceros").caption = funciones.JoinCollectionValues(tmpCol, vbNewLine)
        End If


        Dim op As operacion
        Set tmpCol = New Collection
        For Each op In Orden.OperacionesBanco
            tmpCol.Add op.FechaOperacion & String$(8, " ") & op.moneda.NombreCorto & " " & op.Monto
        Next op
        If tmpCol.count = 0 Then
            .item("lblTransferencias").caption = "NO POSEE OPERACIONES DE BANCO"
        Else
            .item("lblTransferencias").caption = funciones.JoinCollectionValues(tmpCol, vbNewLine)
        End If


        Set tmpCol = New Collection
        For Each op In Orden.OperacionesCaja
            tmpCol.Add op.FechaOperacion & String$(8, " ") & op.moneda.NombreCorto & " " & op.Monto
        Next op
        If tmpCol.count = 0 Then
            .item("lblEfectivo").caption = "NO POSEE OPERACIONES DE CAJA"
        Else
            .item("lblEfectivo").caption = funciones.JoinCollectionValues(tmpCol, vbNewLine)
        End If


        .item("lblDifTipoCambio").caption = Orden.moneda.NombreCorto & " " & Orden.DiferenciaCambio
        .item("lblOtrosDescuentos").caption = Orden.moneda.NombreCorto & " " & Orden.OtrosDescuentos

        .item("lblTotalFacturas").caption = Orden.moneda.NombreCorto & " " & Orden.StaticTotalFacturas
        .item("lblTotalRetenido").caption = Orden.moneda.NombreCorto & " " & Orden.StaticTotalRetenido
        .item("lblTotalAbonado").caption = Orden.moneda.NombreCorto & " " & Orden.StaticTotalOrigenes    '+ Orden.StaticTotalRetenido


        Dim r As Recordset
        Set r = conectar.RSFactory("SELECT 1")
        Set drpOrdenPago.DataSource = r

    End With
End Sub

Private Sub mnuVer_Click()
    Dim f22 As New frmCrearOrdenPago
    f22.Show
    f22.ReadOnly = True
    f22.Cargar Orden
End Sub

Private Sub mnuVerCertificado_Click()
    Dim cr As Collection
    Set cr = DAOCertificadoRetencion.FindAllByOrdenPago(Orden.id)

    If IsSomething(cr) Then
    
    Dim c As CertificadoRetencion
        For Each c In cr
            DAOCertificadoRetencion.VerCertificado c
        Next
    Else
        MsgBox "La orden de pago no tiene certificado.", vbInformation
    End If
End Sub

Private Sub PushButton1_Click()
    Dim ordenes As Collection
    Set ordenes = DAOOrdenPago.FindAll()
    Dim Orden As OrdenPago

    Dim d As New Dictionary
    Dim ret As Retencion
    Dim colret As Collection

    conectar.BeginTransaction


    Dim facturasPosta As Collection

    'conectar.execute "TRUNCATE certificados_retencion"
    'conectar.execute "TRUNCATE certificados_retencion_detalles"

    For Each Orden In ordenes
        If Orden.FacturasProveedor.count > 0 Then    'no traia las facturas bien, faltaban datos y no me daban los totales
            Set facturasPosta = New Collection
            Set facturasPosta = DAOFacturaProveedor.FindAll("AdminComprasFacturasProveedores.id IN (" & funciones.JoinCollectionValues(Orden.FacturasProveedor, ", ", "Id") & ")")
            Set Orden.FacturasProveedor = facturasPosta
        End If

        If Orden.StaticTotalFacturas = 0 Then

            Orden.StaticTotalFacturas = Orden.TotalFacturas
            Orden.StaticTotalFacturasNG = Orden.TotalFacturasNG
            Orden.StaticTotalOrigenes = Orden.TotalOrigenes

            Set colret = DAORetenciones.FindAllEsAgente
            Set d = DAOCertificadoRetencion.VerPosibleRetenciones(Orden.FacturasProveedor, colret, Orden.alicuota, Orden.DiferenciaCambio)
            Dim totRet As Double
            totRet = 0
            For Each ret In colret
                totRet = totRet + d.item(CStr(ret.id))
            Next ret
            Orden.StaticTotalRetenido = funciones.RedondearDecimales(totRet)

            If Not DAOOrdenPago.Guardar(Orden) Then Stop

            If Orden.estado = EstadoOrdenPago_Aprobada And Orden.StaticTotalRetenido > 0 Then
                ' If Not IsSomething(DAOCertificadoRetencion.Create(Orden,) Then Stop
                Err.Raise "ver error en frmOrdenesPago"
            End If

        End If
    Next Orden

    conectar.CommitTransaction



End Sub

Private Sub PushButton2_Click()
    conectar.BeginTransaction
    Dim newcol As New Collection

    Dim nop As OrdenPago



    For Each nop In ordenes
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


            'Q = "DELETE FROM operaciones WHERE id IN (SELECT id_operacion FROM ordenes_pago_operaciones WHERE id_orden_pago = " & op.Id & ")"
            '        If Not conectar.execute(Q) Then GoTo e
            '        Q = "DELETE FROM ordenes_pago_operaciones WHERE id_orden_pago = " & op.Id
            '        If Not conectar.execute(Q) Then GoTo e
            If Not DAOOperacion.Save(opeCaja) Then GoTo E
            opeCaja.id = conectar.UltimoId2
            q = "INSERT INTO ordenes_pago_operaciones VALUES (" & nop.id & ", " & opeCaja.id & ")"
            If Not conectar.execute(q) Then GoTo E
            q = "update ordenes_pago set static_total_origen=" & opeCaja.Monto & " where id=" & nop.id
            If Not conectar.execute(q) Then GoTo E


        End If
        Debug.Print nop.id
    Next



    conectar.CommitTransaction
    Exit Sub
E:
    conectar.RollBackTransaction



End Sub
