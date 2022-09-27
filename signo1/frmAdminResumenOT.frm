VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmAdminResumenOT 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Resúmen..."
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   13185
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   13185
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ Acciones ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   8640
      TabIndex        =   12
      Top             =   5520
      Width           =   1335
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Exportar"
         CausesValidation=   0   'False
         Height          =   255
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Cancel          =   -1  'True
         Caption         =   "Salir"
         CausesValidation=   0   'False
         Height          =   255
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Imprimir"
         Height          =   255
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Generar"
         Default         =   -1  'True
         Height          =   255
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.ComboBox cboMondeasProceso 
      Height          =   315
      Left            =   11400
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   6000
      Width           =   1455
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5295
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   9340
      _Version        =   393216
      TabOrientation  =   3
      Style           =   1
      Tab             =   1
      TabHeight       =   520
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "En Proceso"
      TabPicture(0)   =   "frmAdminResumenOT.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lstenProceso"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Finalizados"
      TabPicture(1)   =   "frmAdminResumenOT.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lstProcesados"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Entregas"
      TabPicture(2)   =   "frmAdminResumenOT.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lstEntregas"
      Tab(2).ControlCount=   1
      Begin MSComctlLib.ListView lstenProceso 
         Height          =   5055
         Left            =   -74880
         TabIndex        =   1
         Top             =   120
         Width           =   12495
         _ExtentX        =   22040
         _ExtentY        =   8916
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nro"
            Object.Width           =   1004
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Cliente"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Creación"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Detalle"
            Object.Width           =   4762
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Entrega"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Moneda"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Monto"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Facturado"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Text            =   "No facturado"
            Object.Width           =   2117
         EndProperty
      End
      Begin MSComctlLib.ListView lstProcesados 
         Height          =   5055
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   12495
         _ExtentX        =   22040
         _ExtentY        =   8916
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nro"
            Object.Width           =   1004
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Cliente"
            Object.Width           =   3616
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Creación"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Detalle"
            Object.Width           =   4762
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Entrega"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Moneda"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Monto"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Facturado"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Text            =   "No facturado"
            Object.Width           =   2117
         EndProperty
      End
      Begin MSComctlLib.ListView lstEntregas 
         Height          =   5055
         Left            =   -74880
         TabIndex        =   3
         Top             =   120
         Width           =   12495
         _ExtentX        =   22040
         _ExtentY        =   8916
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nro"
            Object.Width           =   1004
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Cliente"
            Object.Width           =   3616
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Creación"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Detalle"
            Object.Width           =   4762
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Entrega"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Moneda"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Monto"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Facturado"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Text            =   "No facturado"
            Object.Width           =   2117
         EndProperty
      End
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10080
      TabIndex        =   11
      Top             =   6960
      Width           =   1215
   End
   Begin VB.Label lblTotalProceso 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11400
      TabIndex        =   10
      Top             =   6960
      Width           =   1455
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "Moneda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9720
      TabIndex        =   9
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Facturado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10080
      TabIndex        =   8
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "No Facturado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10080
      TabIndex        =   7
      Top             =   6600
      Width           =   1215
   End
   Begin VB.Label lblFacturadoProceso 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11400
      TabIndex        =   6
      Top             =   6360
      Width           =   1455
   End
   Begin VB.Label lblNoFacturadoProceso 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11400
      TabIndex        =   5
      Top             =   6600
      Width           =   1455
   End
   Begin VB.Menu mnu 
      Caption         =   "mnu"
      Visible         =   0   'False
      Begin VB.Menu nroOrden 
         Caption         =   "[ nro ]"
         Enabled         =   0   'False
      End
      Begin VB.Menu detalle 
         Caption         =   "Ver detalle..."
      End
      Begin VB.Menu remitos 
         Caption         =   "Remitos Aplicados"
      End
      Begin VB.Menu Facturas 
         Caption         =   "Facturas Aplicadas..."
      End
   End
End
Attribute VB_Name = "frmAdminResumenOT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New Recordset
Dim baseC As New classConfigurar
Dim baseSP As New classSignoplast
Dim baseP As New classPlaneamiento

Dim vIdCliente As Long
Private Property Let idCliente(nIdCliente)
    vIdCliente = nIdCliente
End Property

Private Sub cboMondeasProceso_Click()
    verTotales
End Sub
Private Sub Command1_Click()
    Unload Me
End Sub
Public Sub llenarListas()
    Dim strsql As String
    Me.lstenProceso.ListItems.Clear
    Me.lstEntregas.ListItems.Clear
    Me.lstProcesados.ListItems.Clear
    'o/t proceso

    idm = Me.cboMondeasProceso.ItemData(Me.cboMondeasProceso.ListIndex)

    'Set rs = baseP.CrearRS("select p.id,'O/T' as ot,p.fechaCreado,p.fechaEntrega,p.descripcion, p.idmoneda,sum(dp.precio*dp.cantidad) as total, sum(dp.precio*dp.cantidad_facturada) as facturado, sum(dp.precio*(dp.cantidad-dp.cantidad_facturada)) as noFacturado,c.razon from pedidos p, detalles_pedidos dp, clientes c where (p.estado=2 or p.estado=3) and dp.idPedido=p.id and c.id=p.idCliente and p.idMoneda=" & idm & " and dp.precio>0 group by p.id")
    strsql = "select p.id,'O/T' as ot,p.fechaCreado,p.fechaEntrega,p.descripcion, p.idmoneda,sum(dp.precio*dp.cantidad*(1-(p.dto/100))) as total, sum(dp.precio*dp.cantidad_facturada*(1-(p.dto/100))) as facturado, sum(dp.precio*(dp.cantidad-dp.cantidad_facturada)*(1-(p.dto/100))) as noFacturado,c.razon from pedidos p, detalles_pedidos dp, clientes c where (p.estado=2 or p.estado=3) and dp.idPedido=p.id and c.id=p.idCliente and p.idMoneda=" & idm & " and dp.precio>0 group by p.id"
    Set rs = conectar.RSFactory(strsql)
    Dim x As ListItem
    While Not rs.EOF
        Set x = Me.lstenProceso.ListItems.Add(, , Format(rs!Id, "0000"))
        x.SubItems(1) = rs!Ot
        x.SubItems(2) = rs!razon
        x.SubItems(3) = Format(rs!fechaCreado, "dd-mm-yyyy")
        x.SubItems(4) = rs!descripcion
        x.SubItems(5) = Format(rs!FechaEntrega, "dd-mm-yyyy")
        x.SubItems(6) = funciones.queMoneda(rs!IdMoneda)
        x.SubItems(7) = funciones.FormatearDecimales(rs!Total, 2)
        x.SubItems(8) = funciones.FormatearDecimales(rs!Facturado, 2)
        x.SubItems(9) = funciones.FormatearDecimales(rs!nofacturado, 2)
        rs.MoveNext
    Wend
    idm = Me.cboMondeasProceso.ItemData(Me.cboMondeasProceso.ListIndex)
    'o/t finalizado

    strsql = "select p.id,'O/T' as ot,p.fechaCreado,p.fechaEntrega,p.descripcion, p.idmoneda,sum(dp.precio*dp.cantidad*(1-(p.dto/100))) as total, sum(dp.precio*dp.cantidad_facturada*(1-(p.dto/100))) as facturado, sum(dp.precio*(dp.cantidad-dp.cantidad_facturada)*(1-(p.dto/100))) as noFacturado,c.razon from pedidos p, detalles_pedidos dp, clientes c where p.estado=4 and dp.idPedido=p.id and c.id=p.idCliente and p.idMoneda=" & idm & " and dp.precio>0 group by p.id"
    Set rs = conectar.RSFactory(strsql)
    While Not rs.EOF
        Set x = Me.lstProcesados.ListItems.Add(, , Format(rs!Id, "0000"))
        x.SubItems(1) = rs!Ot
        x.SubItems(2) = rs!razon
        x.SubItems(3) = Format(rs!fechaCreado, "dd-mm-yyyy")
        x.SubItems(4) = rs!descripcion
        x.SubItems(5) = Format(rs!FechaEntrega, "dd-mm-yyyy")
        x.SubItems(6) = funciones.queMoneda(rs!IdMoneda)
        x.SubItems(7) = funciones.FormatearDecimales(rs!Total, 2)
        x.SubItems(8) = funciones.FormatearDecimales(rs!Facturado, 2)
        x.SubItems(9) = funciones.FormatearDecimales(rs!nofacturado, 2)
        rs.MoveNext
    Wend

    idm = Me.cboMondeasProceso.ItemData(Me.cboMondeasProceso.ListIndex)
    'select p.id,'O/E' as ot,p.fechaCreado,p.fecha as fechaEntrega,p.referencia as descripcion, p.idmoneda,sum(dp.vale*dp.cantidad) as total, sum(dp.vale*dp.cantidad_facturada) as facturado, sum(dp.vale*(dp.cantidad-dp.cantidad_facturada)) as noFacturado,c.razon from PedidosEntregas p, detallesPedidosEntregas dp, clientes c where (p.estado=2 or p.estado=3) and dp.idPedidoEntrega=p.id and c.id=p.idCliente  group by p.id
    Set rs = conectar.RSFactory("select p.id,'O/E' as ot,p.fechaCreado,p.fecha as fechaEntrega,p.referencia as descripcion, p.idmoneda,sum(dp.vale*dp.cantidad) as total, sum(dp.vale*dp.cantidad_facturada) as facturado, sum(dp.vale*(dp.cantidad-dp.cantidad_facturada)) as noFacturado,c.razon from PedidosEntregas p, detallesPedidosEntregas dp, clientes c where (p.estado=2 or p.estado=3) and dp.idPedidoEntrega=p.id and c.id=p.idCliente and p.idmoneda=" & idm & " and dp.vale>0 group by p.id")
    While Not rs.EOF
        Set x = Me.lstEntregas.ListItems.Add(, , Format(rs!Id, "0000"))
        x.SubItems(1) = rs!Ot
        x.SubItems(2) = rs!razon
        x.SubItems(3) = Format(rs!fechaCreado, "dd-mm-yyyy")
        x.SubItems(4) = rs!descripcion
        x.SubItems(5) = Format(rs!FechaEntrega, "dd-mm-yyyy")
        x.SubItems(6) = funciones.queMoneda(rs!IdMoneda)
        x.SubItems(7) = funciones.FormatearDecimales(rs!Total, 2)
        x.SubItems(8) = funciones.FormatearDecimales(rs!Facturado, 2)
        x.SubItems(9) = funciones.FormatearDecimales(rs!nofacturado, 2)
        rs.MoveNext
    Wend





    verTotales
End Sub

Public Sub verTotales()
    On Error Resume Next
    Dim FACTU As Double
    Dim NOFACTU As Double
    Dim tot As Double
    Dim Facturado As Double
    Dim nofacturado As Double
    Dim Total As Double
    On Error Resume Next
    idm = Me.cboMondeasProceso.ItemData(Me.cboMondeasProceso.ListIndex)
    Me.lblFacturadoProceso = 0
    Me.lblNoFacturadoProceso = 0
    Me.lblTotalProceso = 0

    Dim lst As ListView
    If SSTab1.Tab = 0 Then
        Set lst = Me.lstenProceso
    ElseIf SSTab1.Tab = 1 Then
        Set lst = Me.lstProcesados
    ElseIf SSTab1.Tab = 2 Then
        Set lst = Me.lstEntregas
    End If

    'Set rs = baseP.CrearRS("select sum(dp.precio*dp.cantidad) as total, sum(dp.precio*dp.cantidad_facturada) as facturado, sum(dp.precio*(dp.cantidad-dp.cantidad_facturada)) as noFacturado from pedidos p, detalles_pedidos dp where p.idCliente=" & vIdCliente & "  and (p.estado=2 or p.estado=3) and dp.idPedido=p.id and p.idMoneda=" & idm & " group by p.idCliente")


    For x = 1 To lst.ListItems.count
        FACTU = FACTU + lst.ListItems(x).ListSubItems(7)
        NOFACTU = NOFACTU + lst.ListItems(x).ListSubItems(8)
        tot = tot + lst.ListItems(x).ListSubItems(9)


    Next x




    Me.lblFacturadoProceso = funciones.FormatearDecimales(FACTU, 2)
    Me.lblNoFacturadoProceso = funciones.FormatearDecimales(NOFACTU, 2)
    Me.lblTotalProceso = funciones.FormatearDecimales(Total, 2)
End Sub



Private Sub facturas_Click()
    If Me.SSTab1.Tab = 2 Then  'O/E
        Id = Me.lstEntregas.selectedItem
        Origen = 2
    ElseIf Me.SSTab1.Tab = 1 Then  'O/T fin
        Id = Me.lstProcesados.selectedItem
        Origen = 1
    ElseIf Me.SSTab1.Tab = 0 Then  'o/t pend
        Id = Me.lstenProceso.selectedItem
        Origen = 1
    End If
    frmAdminFacturasAplicadas.Origen = Origen
    frmAdminFacturasAplicadas.idOrigen = CLng(Id)
    'frmAdminFacturasAplicadas.Frame1.Caption = "[ Nro." & id & " ]"
    frmAdminFacturasAplicadas.Show

End Sub

Private Sub Form_Activate()
    On Error Resume Next

    DAOMoneda.LlenarCombo Me.cboMondeasProceso

    llenarListas
    Me.Refresh
End Sub

Private Sub Form_Load()
    FormHelper.Customize Me
End Sub

Private Sub lstenProceso_DblClick()
    If Me.lstenProceso.ListItems.count > 0 Then
        frmDetallePedido.lblIdPedido = Me.lstenProceso.selectedItem
        frmDetallePedido.Frame1.caption = "[ Pedido Nro. " & Format(frmDetallePedido.lblIdPedido, "0000") & " ]"
        frmDetallePedido.Show 1
    End If
End Sub

Private Sub lstEnProceso_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Me.lstenProceso.ListItems.count > 0 Then
        ida = CLng(Me.lstenProceso.selectedItem)

        If Button = 2 Then
            Me.nroOrden.caption = "[ " & Format(ida, "0000") & " ]"
            Me.PopupMenu Me.mnu
        End If
    End If

End Sub

Private Sub lstEntregas_DblClick()
    If Me.lstEntregas.ListItems.count > 0 Then
        frmVerOe.IDOE = CLng(Me.lstEntregas.selectedItem)
        frmVerOe.Show
    End If
End Sub

Private Sub lstEntregas_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Me.lstEntregas.ListItems.count > 0 Then
        ida = CLng(Me.lstEntregas.selectedItem)
        If Button = 2 Then
            Me.nroOrden.caption = "[ " & Format(ida, "0000") & " ]"
            Me.PopupMenu Me.mnu
        End If
    End If

End Sub

Private Sub lstProcesados_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Me.lstProcesados.ListItems.count > 0 Then
        ida = CLng(Me.lstProcesados.selectedItem)
        If Button = 2 Then
            Me.nroOrden.caption = "[ " & Format(ida, "0000") & " ]"
            Me.PopupMenu Me.mnu
        End If
    End If

End Sub

Private Sub remitos_Click()
    If Me.SSTab1.Tab = 2 Then  'O/E
        Id = Me.lstEntregas.selectedItem
        Origen = 1
    ElseIf Me.SSTab1.Tab = 1 Then  'O/T fin
        Id = Me.lstProcesados.selectedItem
        Origen = 0
    ElseIf Me.SSTab1.Tab = 0 Then  'o/t pend
        Id = Me.lstenProceso.selectedItem
        Origen = 0
    End If


    frmRemitosEntregados.Origen = 1
    frmRemitosEntregados.idPedidoEntrega = Id
    frmRemitosEntregados.caption = "Nro." & Id

    frmRemitosEntregados.Show

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)

    verTotales

End Sub

