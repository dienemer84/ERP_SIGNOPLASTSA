VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form frmAdminMasInfoCliente 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Más información..."
   ClientHeight    =   10395
   ClientLeft      =   690
   ClientTop       =   4320
   ClientWidth     =   18540
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10395
   ScaleWidth      =   18540
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   4215
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   7435
      _Version        =   393216
      Style           =   1
      Tab             =   2
      TabHeight       =   520
      BackColor       =   12632256
      TabCaption(0)   =   "Trabajos en proceso"
      TabPicture(0)   =   "frmMasInfoCliente.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lstEnProceso"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Trabajos Finalizados"
      TabPicture(1)   =   "frmMasInfoCliente.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lstProcesados"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Entregas"
      TabPicture(2)   =   "frmMasInfoCliente.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "lstEntregas"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin MSComctlLib.ListView lstEntregas 
         Height          =   3735
         Left            =   120
         TabIndex        =   24
         Top             =   360
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   6588
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstProcesados 
         Height          =   3735
         Left            =   -74880
         TabIndex        =   23
         Top             =   360
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   6588
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstEnProceso 
         Height          =   3735
         Left            =   -74880
         TabIndex        =   22
         Top             =   360
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   6588
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4200
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ Totales ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   8880
      TabIndex        =   13
      Top             =   4920
      Width           =   2775
      Begin VB.ComboBox cboMondeasProceso 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   240
         Width           =   1095
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
         Left            =   1440
         TabIndex        =   21
         Top             =   840
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
         Left            =   1440
         TabIndex        =   20
         Top             =   600
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
         Left            =   120
         TabIndex        =   19
         Top             =   840
         Width           =   1215
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
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Moneda "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   17
         Top             =   240
         Width           =   735
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
         Left            =   1440
         TabIndex        =   16
         Top             =   1200
         Width           =   1215
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
         Left            =   120
         TabIndex        =   15
         Top             =   1200
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ Filtro ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   6600
      TabIndex        =   7
      Top             =   4920
      Width           =   2175
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   255
         Left            =   720
         TabIndex        =   12
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   21299201
         CurrentDate     =   39204
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   255
         Left            =   720
         TabIndex        =   11
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   21299201
         CurrentDate     =   39204
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Aplicar filtro"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Hasta  "
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Desde  "
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   615
      End
   End
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
      Height          =   1575
      Left            =   5160
      TabIndex        =   3
      Top             =   4920
      Width           =   1335
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Imprimir"
         Height          =   255
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Generar"
         Default         =   -1  'True
         Height          =   255
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00E0E0E0&
         Cancel          =   -1  'True
         Caption         =   "Salir"
         Height          =   255
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1080
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ Cliente ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11655
      Begin VB.ComboBox cboClientes 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   240
         Width           =   10215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Razón Social"
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
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Menu mnu 
      Caption         =   "mnu"
      Visible         =   0   'False
      Begin VB.Menu nro 
         Caption         =   "nro"
      End
      Begin VB.Menu detalles 
         Caption         =   "Ver Detalles..."
      End
      Begin VB.Menu remitos 
         Caption         =   "Ver Remitos..."
      End
      Begin VB.Menu facturas 
         Caption         =   "Ver Facturas..."
      End
   End
End
Attribute VB_Name = "frmAdminMasInfoCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New Recordset
Dim baseC As New classConfigurar
Dim baseS As New classStock
Dim baseSP As New classSignoplast
Dim baseP As New classPlaneamiento
Dim baseV As New classVentas
Dim vIdCliente As Long
Public Property Let idCliente(nIdCliente As Long)
vIdCliente = nIdCliente
End Property


Private Sub Command1_Click()
vIdCliente = CLng(Me.cboClientes.ItemData(Me.cboClientes.ListIndex))
generarInforme

End Sub

Private Sub Command2_Click()
For X = 1 To Me.CommonDialog1.Copies
'ImprimirInforme
Next
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub detalles_Click()
If Me.SSTab1.Tab = 2 Then  'O/E
id = Me.lstEntregas.SelectedItem
ElseIf Me.SSTab1.Tab = 1 Then  'O/T fin
id = Me.lstProcesados.SelectedItem
ElseIf Me.SSTab1.Tab = 0 Then  'o/t pend
id = Me.lstEnProceso.SelectedItem
End If


    
    'frmDetallePedido.lblIdPedido = id
    'frmDetallePedido.Frame1.Caption = "[ Pedido Nro. " & Format(frmDetallePedido.lblIdPedido, "0000") & " ]"
    'frmDetallePedido.Show



End Sub

Private Sub lstenProceso_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lstEnProceso.ListItems.count > 0 Then
ida = CLng(Me.lstEnProceso.SelectedItem)
If Button = 2 Then
Me.nro.Caption = "[ " & Format(ida, "0000") & " ]"
Me.PopupMenu Me.mnu
End If
End If
End Sub

Private Sub lstEntregas_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Me.lstEntregas.ListItems.count > 0 Then
ida = CLng(Me.lstEntregas.SelectedItem)
If Button = 2 Then
Me.nro.Caption = "[ " & Format(ida, "0000") & " ]"
Me.PopupMenu Me.mnu
End If
End If
End Sub

Private Sub lstProcesados_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Me.lstProcesados.ListItems.count > 0 Then
ida = CLng(Me.lstProcesados.SelectedItem)

If Button = 2 Then
Me.nro.Caption = "[ " & Format(ida, "0000") & " ]"
Me.PopupMenu Me.mnu
End If
End If
End Sub

Private Sub remitos_Click()
If Me.SSTab1.Tab = 2 Then  'O/E
id = Me.lstEntregas.SelectedItem
origen = 1
ElseIf Me.SSTab1.Tab = 1 Then  'O/T fin
id = Me.lstProcesados.SelectedItem
origen = 0
ElseIf Me.SSTab1.Tab = 0 Then  'o/t pend
id = Me.lstEnProceso.SelectedItem
origen = 0
End If


frmRemitosEntregados.origen = 1
frmRemitosEntregados.idPedidoEntrega = id
frmRemitosEntregados.Frame1.Caption = "[ Nro." & id & " ]"
frmRemitosEntregados.Show

End Sub

Private Sub Facturas_Click()
If Me.SSTab1.Tab = 2 Then  'O/E
id = Me.lstEntregas.SelectedItem
origen = 2
ElseIf Me.SSTab1.Tab = 1 Then  'O/T fin
id = Me.lstProcesados.SelectedItem
origen = 1
ElseIf Me.SSTab1.Tab = 0 Then  'o/t pend
id = Me.lstEnProceso.SelectedItem
origen = 1
End If


frmAdminFacturasAplicadas.origen = origen
frmAdminFacturasAplicadas.idOrigen = CLng(id)
frmAdminFacturasAplicadas.Frame1.Caption = "[ Nro." & id & " ]"
frmAdminFacturasAplicadas.Show

End Sub

Private Sub Form_Activate()
On Error Resume Next
baseC.llenarCboMonedas Me.cboMondeasProceso
baseS.llenar_combo_clientes Me.cboClientes, 9999
Me.cboClientes.ListIndex = 0
Me.cboMondeasProceso.ListIndex = 0
Me.DTPicker1 = Now
Me.DTPicker2 = Now
Me.Refresh
End Sub

Public Sub llenarListas()
Me.lstEnProceso.ListItems.Clear
Me.lstEntregas.ListItems.Clear


Me.lstProcesados.ListItems.Clear
Dim strsql As String
Dim filtro As Boolean
If Me.Check1 Then
 filtro = True
 desde = Format(Me.DTPicker1, "yyyy/mm/dd")
 hasta = Format(Me.DTPicker2, "yyyy/mm/dd")
Else
 filtro = False
End If


idm = Me.cboMondeasProceso.ItemData(Me.cboMondeasProceso.ListIndex)
'o/t en proceso
strsql = "select p.id,'O/T' as ot,p.fechaCreado,p.fechaEntrega,p.descripcion, p.idmoneda,sum(dp.precio*dp.cantidad*(1-(p.dto/100))) as total, sum(dp.precio*dp.cantidad_facturada*(1-(p.dto/100))) as facturado, sum(dp.precio*(dp.cantidad-dp.cantidad_facturada)*(1-(p.dto/100))) as noFacturado from pedidos p, detalles_pedidos dp where p.idCliente=" & vIdCliente & " and (p.estado=2 or p.estado=3) and dp.idPedido=p.id and p.idMoneda=" & idm & " and dp.precio>0 "
filtros = " and p.fechaEntrega>'" & desde & "' and p.fechaEntrega<'" & hasta & "'"
grupo = " group by p.id"
If filtro Then strsql = strsql & filtros
strsql = strsql & grupo

Set rs = baseP.CrearRS(strsql)
Dim X As ListItem
While Not rs.EOF
Set X = Me.lstEnProceso.ListItems.Add(, , Format(rs!id, "0000"))
    X.SubItems(1) = rs!ot
    X.SubItems(2) = rs!FechaCreado
    X.SubItems(3) = rs!descripcion
    X.SubItems(4) = rs!fechaEntrega
    X.SubItems(5) = funciones.queMoneda(rs!idMoneda)
    X.SubItems(6) = funciones.formatearDecimales(rs!total, 2)
    X.SubItems(7) = funciones.formatearDecimales(rs!facturado, 2)
    X.SubItems(8) = funciones.formatearDecimales(rs!nofacturado, 2)
rs.MoveNext
Wend

'o/t finalizados
strsql = "select p.id,'O/T' as ot,p.fechaCreado,p.fechaEntrega,p.descripcion, p.idmoneda,sum(dp.precio*dp.cantidad*(1-(p.dto/100))) as total, sum(dp.precio*dp.cantidad_facturada*(1-(p.dto/100))) as facturado, sum(dp.precio*(dp.cantidad-dp.cantidad_facturada)*(1-(p.dto/100))) as noFacturado from pedidos p, detalles_pedidos dp where p.idCliente=" & vIdCliente & " and p.estado=4 and dp.idPedido=p.id and p.idMoneda=" & idm & " and dp.precio>=0 "
grupo = " group by p.id"
filtros = " and p.fechaEntrega>'" & desde & "' and p.fechaEntrega<'" & hasta & "'"
If filtro Then strsql = strsql & filtros
strsql = strsql & grupo

Set rs = baseP.CrearRS(strsql)
While Not rs.EOF
Set X = Me.lstProcesados.ListItems.Add(, , Format(rs!id, "0000"))
    X.SubItems(1) = rs!ot
    X.SubItems(2) = rs!FechaCreado
    X.SubItems(3) = rs!descripcion
    X.SubItems(4) = rs!fechaEntrega
    X.SubItems(5) = funciones.queMoneda(rs!idMoneda)
    X.SubItems(6) = funciones.formatearDecimales(rs!total, 2)
    X.SubItems(7) = funciones.formatearDecimales(rs!facturado, 2)
    X.SubItems(8) = funciones.formatearDecimales(rs!nofacturado, 2)
rs.MoveNext
Wend

'o/e
strsql = "select p.id,'O/E' as ot,p.fechaCreado,p.fecha as fechaEntrega,p.referencia as descripcion, p.idmoneda,sum(dp.vale*dp.cantidad) as total, sum(dp.vale*dp.cantidad_facturada) as facturado, sum(dp.vale*(dp.cantidad-dp.cantidad_facturada)) as noFacturado,c.razon from PedidosEntregas p, detallesPedidosEntregas dp, clientes c where (p.estado=2 or p.estado=3) and dp.idPedidoEntrega=p.id and c.id=p.idCliente and p.idmoneda=" & idm & " and dp.vale>0 and p.idCliente=" & vIdCliente
grupo = " group by p.id"
filtros = " and p.fecha>'" & desde & "' and p.fecha<'" & hasta & "'"
If filtro Then strsql = strsql & filtros
strsql = strsql & grupo



Set rs = baseP.CrearRS(strsql)
While Not rs.EOF
Set X = Me.lstEntregas.ListItems.Add(, , Format(rs!id, "0000"))
    X.SubItems(1) = rs!ot
    X.SubItems(2) = Format(rs!FechaCreado, "dd-mm-yyyy")
    X.SubItems(3) = rs!descripcion
    X.SubItems(4) = Format(rs!fechaEntrega, "dd-mm-yyyy")
    X.SubItems(5) = funciones.queMoneda(rs!idMoneda)
    X.SubItems(6) = funciones.formatearDecimales(rs!total, 2)
    X.SubItems(7) = funciones.formatearDecimales(rs!facturado, 2)
    X.SubItems(8) = funciones.formatearDecimales(rs!nofacturado, 2)
rs.MoveNext
Wend

verTotales
End Sub

Public Sub verTotales()
Dim lst As ListView
Dim facturado As Double, total As Double, nofacturado As Double
If Me.SSTab1.Tab = 0 Then 'ot proceso
Set lst = Me.lstEnProceso
ElseIf SSTab1.Tab = 1 Then 'ot fin
Set lst = Me.lstProcesados
ElseIf SSTab1.Tab = 2 Then 'oe
Set lst = Me.lstEntregas
End If


total = 0
facturado = 0
nofacturado = 0
For X = 1 To lst.ListItems.count
facturado = CDbl(lst.ListItems(X).ListSubItems(7)) + facturado
nofacturado = CDbl(lst.ListItems(X).ListSubItems(8)) + nofacturado
total = CDbl(lst.ListItems(X).ListSubItems(6)) + total

Next X
Me.lblFacturadoProceso = funciones.formatearDecimales(facturado, 2)
Me.lblTotalProceso = funciones.formatearDecimales(total, 2)
Me.lblNoFacturadoProceso = funciones.formatearDecimales(nofacturado, 2)

End Sub
Private Sub generarInforme()

    llenarListas
    verTotales
End Sub


Private Sub lstEntregas_DblClick()
If Me.lstEntregas.ListItems.count > 0 Then
  frmVerOe.IDOE = CLng(Me.lstEntregas.SelectedItem)
  frmVerOe.Show
End If

End Sub

Private Sub lstProcesados_DblClick()
If Me.lstProcesados.ListItems.count > 0 Then
    frmDetallePedido.lblIdPedido = Me.lstProcesados.SelectedItem
    frmDetallePedido.Frame1.Caption = "[ Pedido Nro. " & Format(frmDetallePedido.lblIdPedido, "0000") & " ]"
    frmDetallePedido.Show
End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
verTotales
End Sub

