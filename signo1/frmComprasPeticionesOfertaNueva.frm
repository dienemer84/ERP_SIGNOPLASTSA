VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmComprasPeticionesOfertaNueva 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle Peticion de Oferta"
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   10860
   Icon            =   "frmComprasPeticionesOfertaNueva.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   10860
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1035
      Left            =   165
      TabIndex        =   13
      Top             =   105
      Width           =   8055
      _Version        =   786432
      _ExtentX        =   14208
      _ExtentY        =   1826
      _StockProps     =   79
      Caption         =   "Datos de la PO"
      UseVisualStyle  =   -1  'True
      Begin VB.ComboBox cboMoneda 
         Height          =   315
         Left            =   6180
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   570
         Width           =   1710
      End
      Begin VB.Label lblUsuario 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario: "
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   5490
         TabIndex        =   19
         Tag             =   "Usuario: "
         Top             =   300
         Width           =   630
      End
      Begin VB.Label lblFechaEmision 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Emision: "
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2625
         TabIndex        =   18
         Tag             =   "Fecha de Emision: "
         Top             =   300
         Width           =   1350
      End
      Begin VB.Label lblProveedor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Proveedor: "
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   135
         TabIndex        =   17
         Tag             =   "Proveedor: "
         Top             =   630
         Width           =   825
      End
      Begin VB.Label lblNroReq 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Req: "
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   135
         TabIndex        =   16
         Tag             =   "Nº Req: "
         Top             =   300
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Moneda:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   5475
         TabIndex        =   15
         Top             =   630
         Width           =   630
      End
   End
   Begin XtremeSuiteControls.GroupBox grpEntregas 
      Height          =   1905
      Left            =   6870
      TabIndex        =   11
      Top             =   5535
      Width           =   3930
      _Version        =   786432
      _ExtentX        =   6932
      _ExtentY        =   3360
      _StockProps     =   79
      Caption         =   "Entregas del item"
      UseVisualStyle  =   -1  'True
      Begin GridEX20.GridEX GrillaEntregas 
         Height          =   1590
         Left            =   90
         TabIndex        =   12
         Top             =   225
         Width           =   3750
         _ExtentX        =   6615
         _ExtentY        =   2805
         Version         =   "2.0"
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         CalendarTodayText=   "Hoy"
         CalendarNoneText=   "Nada"
         ColumnAutoResize=   -1  'True
         MethodHoldFields=   -1  'True
         AllowDelete     =   -1  'True
         BorderStyle     =   3
         GroupByBoxVisible=   0   'False
         BackColorHeader =   16761024
         RowHeaders      =   -1  'True
         DataMode        =   99
         AllowAddNew     =   -1  'True
         ColumnHeaderHeight=   285
         IntProp1        =   0
         ColumnsCount    =   2
         Column(1)       =   "frmComprasPeticionesOfertaNueva.frx":000C
         Column(2)       =   "frmComprasPeticionesOfertaNueva.frx":012C
         FormatStylesCount=   7
         FormatStyle(1)  =   "frmComprasPeticionesOfertaNueva.frx":0240
         FormatStyle(2)  =   "frmComprasPeticionesOfertaNueva.frx":0378
         FormatStyle(3)  =   "frmComprasPeticionesOfertaNueva.frx":0428
         FormatStyle(4)  =   "frmComprasPeticionesOfertaNueva.frx":04DC
         FormatStyle(5)  =   "frmComprasPeticionesOfertaNueva.frx":05B4
         FormatStyle(6)  =   "frmComprasPeticionesOfertaNueva.frx":066C
         FormatStyle(7)  =   "frmComprasPeticionesOfertaNueva.frx":074C
         ImageCount      =   0
         PrinterProperties=   "frmComprasPeticionesOfertaNueva.frx":07D8
      End
   End
   Begin XtremeSuiteControls.PushButton cmdGuardar 
      Height          =   390
      Left            =   9285
      TabIndex        =   7
      Top             =   780
      Width           =   1470
      _Version        =   786432
      _ExtentX        =   2593
      _ExtentY        =   688
      _StockProps     =   79
      Caption         =   "Guardar"
      UseVisualStyle  =   -1  'True
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   4185
      Left            =   0
      TabIndex        =   0
      Top             =   1320
      Width           =   10875
      _ExtentX        =   19182
      _ExtentY        =   7382
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      PreviewColumn   =   10
      PreviewRowLines =   2
      ColumnAutoResize=   -1  'True
      MultiSelect     =   -1  'True
      MethodHoldFields=   -1  'True
      ContScroll      =   -1  'True
      GroupByBoxInfoText=   "Arrastre una columna para agrupar"
      GroupByBoxVisible=   0   'False
      BackColorGBBox  =   16744576
      BackColorHeader =   16761024
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   10
      Column(1)       =   "frmComprasPeticionesOfertaNueva.frx":09B0
      Column(2)       =   "frmComprasPeticionesOfertaNueva.frx":0AD0
      Column(3)       =   "frmComprasPeticionesOfertaNueva.frx":0BAC
      Column(4)       =   "frmComprasPeticionesOfertaNueva.frx":0CBC
      Column(5)       =   "frmComprasPeticionesOfertaNueva.frx":0DCC
      Column(6)       =   "frmComprasPeticionesOfertaNueva.frx":0EDC
      Column(7)       =   "frmComprasPeticionesOfertaNueva.frx":0FD0
      Column(8)       =   "frmComprasPeticionesOfertaNueva.frx":10E4
      Column(9)       =   "frmComprasPeticionesOfertaNueva.frx":11E8
      Column(10)      =   "frmComprasPeticionesOfertaNueva.frx":1300
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmComprasPeticionesOfertaNueva.frx":13C0
      FormatStyle(2)  =   "frmComprasPeticionesOfertaNueva.frx":14F8
      FormatStyle(3)  =   "frmComprasPeticionesOfertaNueva.frx":15A8
      FormatStyle(4)  =   "frmComprasPeticionesOfertaNueva.frx":165C
      FormatStyle(5)  =   "frmComprasPeticionesOfertaNueva.frx":1734
      FormatStyle(6)  =   "frmComprasPeticionesOfertaNueva.frx":17EC
      ImageCount      =   0
      PrinterProperties=   "frmComprasPeticionesOfertaNueva.frx":18CC
   End
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   1650
      Left            =   180
      TabIndex        =   1
      Top             =   5715
      Width           =   3135
      _Version        =   786432
      _ExtentX        =   5530
      _ExtentY        =   2910
      _StockProps     =   68
      Appearance      =   10
      Color           =   32
      ItemCount       =   2
      Item(0).Caption =   "Comerciales"
      Item(0).ControlCount=   6
      Item(0).Control(0)=   "Label2"
      Item(0).Control(1)=   "Label3"
      Item(0).Control(2)=   "Label5"
      Item(0).Control(3)=   "txtPorcentajeDescuento"
      Item(0).Control(4)=   "txtCantDiasPago"
      Item(0).Control(5)=   "txtFormaPago"
      Item(1).Caption =   "Entrega"
      Item(1).ControlCount=   2
      Item(1).Control(0)=   "radRetiramos"
      Item(1).Control(1)=   "radEntregan"
      Begin VB.TextBox txtPorcentajeDescuento 
         Height          =   285
         Left            =   1500
         TabIndex        =   10
         Top             =   1155
         Width           =   1440
      End
      Begin VB.TextBox txtCantDiasPago 
         Height          =   285
         Left            =   1500
         TabIndex        =   9
         Top             =   795
         Width           =   1440
      End
      Begin VB.TextBox txtFormaPago 
         Height          =   285
         Left            =   1500
         TabIndex        =   8
         Top             =   435
         Width           =   1440
      End
      Begin XtremeSuiteControls.RadioButton radRetiramos 
         Height          =   345
         Left            =   -69040
         TabIndex        =   5
         Top             =   600
         Visible         =   0   'False
         Width           =   1050
         _Version        =   786432
         _ExtentX        =   1852
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Retiramos"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton radEntregan 
         Height          =   360
         Left            =   -69040
         TabIndex        =   6
         Top             =   930
         Visible         =   0   'False
         Width           =   1050
         _Version        =   786432
         _ExtentX        =   1852
         _ExtentY        =   635
         _StockProps     =   79
         Caption         =   "Entregan"
         UseVisualStyle  =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Forma de Pago"
         Height          =   195
         Left            =   315
         TabIndex        =   4
         Top             =   465
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad de días"
         Height          =   195
         Left            =   180
         TabIndex        =   3
         Top             =   825
         Width           =   1215
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "% Descuento"
         Height          =   195
         Left            =   420
         TabIndex        =   2
         Top             =   1170
         Width           =   945
      End
   End
   Begin VB.Menu mnuEstados 
      Caption         =   "Estados"
      Visible         =   0   'False
      Begin VB.Menu mnuCambiarEstado 
         Caption         =   "Cambiar estado a"
         Begin VB.Menu mnuActivo 
            Caption         =   "Activo"
         End
         Begin VB.Menu mnuAnulado 
            Caption         =   "Anulado"
         End
         Begin VB.Menu mnuEnEspera 
            Caption         =   "En espera"
         End
      End
   End
End
Attribute VB_Name = "frmComprasPeticionesOfertaNueva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SoloLectura As Boolean
Dim vPeticion As clsPeticionOferta
Dim detalle As Collection
Dim tmp As clsPeticionOfertaDetalle
Private detalleActual As clsPeticionOfertaDetalle
Private tmpEntrega As EntregaPetOfDetalle

Public Property Let peticion(nvalue As clsPeticionOferta)
    Set vPeticion = nvalue
    Me.lblFechaEmision.caption = Me.lblFechaEmision.Tag & nvalue.FechaEmision
    Me.lblProveedor.caption = Me.lblProveedor.Tag & nvalue.Proveedor.RazonSocial
    Me.lblUsuario.caption = Me.lblUsuario.Tag & nvalue.usuarioCreador.usuario
    Me.lblNroReq.caption = Me.lblNroReq.Tag & nvalue.idReque

    Dim mon As clsMoneda
    Me.cboMoneda.Clear
    For Each mon In DAOMoneda.GetAll()
        Me.cboMoneda.AddItem mon.NombreLargo
        Me.cboMoneda.ItemData(Me.cboMoneda.NewIndex) = mon.Id
    Next mon

    If vPeticion.moneda Is Nothing Then
        Me.cboMoneda.ListIndex = -1
    Else
        Me.cboMoneda.ListIndex = funciones.PosIndexCbo(vPeticion.moneda.Id, Me.cboMoneda)
    End If

    Me.txtCantDiasPago.Text = vPeticion.CantidadDiasPago
    Me.txtFormaPago.Text = vPeticion.FormaDePago
    Me.txtPorcentajeDescuento.Text = vPeticion.PorcentajeDescuento
    Me.radRetiramos.value = vPeticion.EntregaRetiramos
    Me.radEntregan.value = Not vPeticion.EntregaRetiramos

End Property

Private Sub cmdGuardar_Click()
    Dim tmp As clsPeticionOfertaDetalle
    For Each tmp In detalle
        If Not tmp.IsValidCantidadEnEntregas() Then
            MsgBox "No coinciden las cantidades de las entregas con los detalles. Revise.", vbCritical + vbOKOnly
            Exit Sub
        End If
    Next tmp

    conectar.BeginTransaction

    Set vPeticion.moneda = DAOMoneda.GetById(Me.cboMoneda.ItemData(Me.cboMoneda.ListIndex))
    vPeticion.CantidadDiasPago = Val(Me.txtCantDiasPago.Text)
    vPeticion.PorcentajeDescuento = Val(Me.txtPorcentajeDescuento.Text)
    vPeticion.FormaDePago = Me.txtFormaPago.Text
    vPeticion.EntregaRetiramos = Me.radRetiramos.value
    If Not DAOPeticionOferta.Update(vPeticion) Then
        conectar.RollBackTransaction
        MsgBox "Se cancelo la transaccion por algun error.", vbCritical
        Exit Sub
    End If

    For Each tmp In detalle

        If tmp.Valor > 0 Then
            If Not DAOPeticionOfertaDetalle.Update(tmp, vPeticion) Then
                conectar.RollBackTransaction
                MsgBox "Se cancelo la transaccion por algun error.", vbCritical
                Exit Sub
            End If
        End If
    Next tmp
    conectar.CommitTransaction
    Channel.Notificar Nothing, EdicionDetallePeticionOferta

    ''' traigo lo nuevo (por las entregas que las borro)
    Set detalle = DAOPeticionOfertaDetalle.FindAll(CLng(vPeticion.numero))
    'Me.GridEX1.ItemCount = detalle.count
    'GridEX1_SelectionChange
    '''

    MsgBox "Los cambios han sido guardados.", vbInformation
    'Unload Me
End Sub

Private Sub Form_Load()
    FormHelper.Customize Me
    Set vPeticion = DAOPeticionOferta.GetById(vPeticion.numero)

    GridEXHelper.CustomizeGrid Me.GridEX1, False, True
    GridEXHelper.CustomizeGrid Me.GrillaEntregas, False, True

    Set detalle = DAOPeticionOfertaDetalle.FindAll(CLng(vPeticion.numero))
    Me.GridEX1.ItemCount = 0
    Me.GridEX1.ItemCount = detalle.count
    GridEX1_SelectionChange

    Me.caption = "Detalles de peticion de oferta Nº " & vPeticion.numero


End Sub

Private Sub GridEX1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu Me.mnuEstados
    End If
End Sub

Private Sub GridEX1_SelectionChange()
    GrillaEntregas.ItemCount = 0

    If Me.GridEX1.rowIndex(Me.GridEX1.row) > 0 Then
        Set detalleActual = detalle.item(Me.GridEX1.row)
        Me.GrillaEntregas.ItemCount = detalleActual.Entregas.count
    Else
        Set detalleActual = Nothing
    End If

End Sub

Private Sub GridEX1_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Set tmp = detalle.item(rowIndex)
    With tmp
        Values(1) = .Cantidad
        Values(2) = .DetalleReque.Material.codigo & " | " & .DetalleReque.Material.descripcion
        Values(3) = .DetalleReque.Kg
        Values(4) = .DetalleReque.m2
        Values(5) = .DetalleReque.ML
        Values(6) = .Valor
        Values(7) = .total
        Values(8) = .DetalleReque.observaciones
        Values(9) = enums.EstadosPeticionOfertaDetalle.item(CStr(.estado))
        Values(10) = funciones.JoinCollectionValues(.DetalleReque.Material.Atributos, "|")
    End With
End Sub

Private Sub GridEX1_UnboundUpdate(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    On Error Resume Next
    Set tmp = detalle.item(rowIndex)
    tmp.Cantidad = Values(1)
    tmp.Valor = Values(6)
End Sub


Private Sub GrillaEntregas_BeforeDelete(ByVal Cancel As GridEX20.JSRetBoolean)
    If MsgBox("¿Está seguro de eliminar los registros?", vbYesNo + vbQuestion) = vbNo Then Cancel = True
End Sub

Private Sub GrillaEntregas_BeforeUpdate(ByVal Cancel As GridEX20.JSRetBoolean)
    Cancel = (Me.GridEX1.SelectedItems.count = 0)
End Sub

Private Sub GrillaEntregas_UnboundAddNew(ByVal NewRowBookmark As GridEX20.JSRetVariant, ByVal Values As GridEX20.JSRowData)
    If Me.GridEX1.SelectedItems.count > 0 Then
        Set tmpEntrega = New EntregaPetOfDetalle
        tmpEntrega.Cantidad = Values(1)
        tmpEntrega.FEcha = Values(2)
        detalleActual.Entregas.Add tmpEntrega
    End If
End Sub

Private Sub GrillaEntregas_UnboundDelete(ByVal rowIndex As Long, ByVal Bookmark As Variant)
    On Error Resume Next
    detalleActual.Entregas.remove (rowIndex)
End Sub

Private Sub GrillaEntregas_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If Not detalleActual Is Nothing Then
        Set tmpEntrega = detalleActual.Entregas(rowIndex)
        Values(1) = tmpEntrega.Cantidad
        Values(2) = tmpEntrega.FEcha
    End If
End Sub

Private Sub GrillaEntregas_UnboundUpdate(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    On Error Resume Next
    Set tmpEntrega = detalleActual.Entregas.item(rowIndex)
    tmpEntrega.Cantidad = Values(1)
    tmpEntrega.FEcha = Values(2)
End Sub


