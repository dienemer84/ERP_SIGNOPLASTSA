VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAdminMasInfoCliente 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Resúmen de facturación"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   12135
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   375
      Left            =   1200
      TabIndex        =   26
      Top             =   5760
      Width           =   1095
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
      TabIndex        =   21
      Top             =   0
      Width           =   12015
      Begin VB.ComboBox cboClientes 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   240
         Width           =   10575
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
         TabIndex        =   23
         Top             =   240
         Width           =   1215
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
      Left            =   5475
      TabIndex        =   17
      Top             =   5445
      Width           =   1335
      Begin VB.CommandButton Command3 
         BackColor       =   &H00E0E0E0&
         Cancel          =   -1  'True
         Caption         =   "Salir"
         Height          =   255
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Generar"
         Default         =   -1  'True
         Height          =   255
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   345
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Imprimir"
         Height          =   255
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   720
         Width           =   1095
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
      Left            =   6960
      TabIndex        =   11
      Top             =   5400
      Width           =   2175
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Aplicar filtro"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   255
         Left            =   720
         TabIndex        =   12
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   62521345
         CurrentDate     =   39204
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   255
         Left            =   720
         TabIndex        =   13
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   62521345
         CurrentDate     =   39204
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Desde  "
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Hasta  "
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1080
         Width           =   615
      End
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
      Left            =   9240
      TabIndex        =   2
      Top             =   5400
      Width           =   2775
      Begin VB.ComboBox cboMondeasProceso 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   1095
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
         TabIndex        =   10
         Top             =   1200
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
         Left            =   1440
         TabIndex        =   9
         Top             =   1200
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
         TabIndex        =   8
         Top             =   240
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
         Left            =   120
         TabIndex        =   7
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
         TabIndex        =   6
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
         TabIndex        =   5
         Top             =   600
         Width           =   1215
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
         TabIndex        =   4
         Top             =   840
         Width           =   1215
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4575
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   8070
      _Version        =   393216
      TabOrientation  =   3
      Style           =   1
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
      TabCaption(0)   =   "En proceso"
      TabPicture(0)   =   "frmAdminMasInfoCliente.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lstEnProceso"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Finalizados"
      TabPicture(1)   =   "frmAdminMasInfoCliente.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lstProcesados"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Entregas"
      TabPicture(2)   =   "frmAdminMasInfoCliente.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lstEntregas"
      Tab(2).ControlCount=   1
      Begin MSComctlLib.ListView lstEnProceso 
         Height          =   4335
         Left            =   120
         TabIndex        =   1
         Top             =   30
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   7646
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
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nro"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "OC"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Entrega"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Moneda"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Total"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Facturado"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "No Facturado"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "% Pend"
            Object.Width           =   1764
         EndProperty
      End
      Begin MSComctlLib.ListView lstProcesados 
         Height          =   4335
         Left            =   -74880
         TabIndex        =   24
         Top             =   30
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   7646
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
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nro"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "OC"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Entrega"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Moneda"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Total"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Facturado"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "No Facturado"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "% Pend"
            Object.Width           =   1764
         EndProperty
      End
      Begin MSComctlLib.ListView lstEntregas 
         Height          =   4335
         Left            =   -74880
         TabIndex        =   25
         Top             =   30
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   7646
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
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nro"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "OC"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Entrega"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Moneda"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Total"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Facturado"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "No Facturado"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "% Pend"
            Object.Width           =   1764
         EndProperty
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3600
      Top             =   6360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Menu deta 
      Caption         =   "Detalle"
      Visible         =   0   'False
      Begin VB.Menu nro 
         Caption         =   "numero"
         Enabled         =   0   'False
      End
      Begin VB.Menu detalle 
         Caption         =   "Detalle..."
      End
   End
End
Attribute VB_Name = "frmAdminMasInfoCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New Recordset
Dim baseA As New classAdministracion
Dim baseC As New classConfigurar
Dim baseS As New classStock
Dim baseSP As New classSignoplast
Dim baseP As New classPlaneamiento


Dim vIdCliente As Long
Public Property Let idCliente(nIdCliente As Long)
    vIdCliente = nIdCliente
End Property


Private Sub Command1_Click()
    vIdCliente = CLng(Me.cboClientes.ItemData(Me.cboClientes.ListIndex))
    generarInforme

End Sub

Private Sub Command2_Click()
    Dim FONDO1 As String, FONDO2 As String
    tit = "Detalle por orden de trabajo"
    ti = "PENDIENTES DE FACTURACION"
    FONDO1 = "No Facturado: " & Me.lblNoFacturadoProceso
    FONDO2 = "Facturado: " & Me.lblFacturadoProceso & Chr(10) & "Total: " & Me.lblTotalProceso

    Dim l As ListView
    If Me.SSTab1.Tab = 0 Then
        Set l = Me.lstEnProceso
        ti = ti & " - ot en proceso"
    ElseIf Me.SSTab1.Tab = 1 Then
        Set l = Me.lstProcesados
        ti = ti & " - ot procesados"
    ElseIf Me.SSTab1.Tab = 2 Then
        Set l = Me.lstEntregas
        ti = ti & " - entregas"
    End If

    If Me.Check1 Then
        tit = tit & " entre " & Format(Me.DTPicker1, "dd-mm-yyyy") & " y " & Format(Me.DTPicker2, "dd-mm-yyyy")
    End If


    funciones.ImprimirLista ti, l, frmPrincipal.cd, tit, FONDO1, FONDO2
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub detalles_Click()
    If Me.SSTab1.Tab = 2 Then  'O/E
        id = Me.lstEntregas.selectedItem
    ElseIf Me.SSTab1.Tab = 1 Then  'O/T fin
        id = Me.lstProcesados.selectedItem
    ElseIf Me.SSTab1.Tab = 0 Then  'o/t pend
        id = Me.lstEnProceso.selectedItem
    End If



    'frmDetallePedido.lblIdPedido = id
    'frmDetallePedido.Frame1.Caption = "[ Pedido Nro. " & Format(frmDetallePedido.lblIdPedido, "0000") & " ]"
    'frmDetallePedido.Show



End Sub

Private Sub Command4_Click()
Dim frm As New frmAdminMasInfoCliente2
frm.Show
End Sub

Private Sub detalle_Click()
    If Me.SSTab1.Tab = 2 Then  'O/E
        id = Me.lstEntregas.selectedItem.Tag
        origen = 1
    ElseIf Me.SSTab1.Tab = 1 Then  'O/T fin
        id = Me.lstProcesados.selectedItem.Tag
        origen = 0
    ElseIf Me.SSTab1.Tab = 0 Then  'o/t pend
        id = Me.lstEnProceso.selectedItem.Tag
        origen = 0
    End If
    frmAdminMasInfoClienteDetalle.origen = origen
    frmAdminMasInfoClienteDetalle.otoe = id
    frmAdminMasInfoClienteDetalle.Show
End Sub

Private Sub Form_Load()
FormHelper.Customize Me
End Sub

Private Sub lstEnProceso_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If lstEnProceso.ListItems.count > 0 Then
        ida = Me.lstEnProceso.selectedItem
        If Button = 2 Then
            Me.nro.caption = "[ " & Format(ida, "0000") & " ]"
            Me.PopupMenu Me.deta
        End If
    End If
End Sub
Private Sub remitos_Click()
    If Me.SSTab1.Tab = 2 Then  'O/E
        id = Me.lstEntregas.selectedItem
        origen = 1
    ElseIf Me.SSTab1.Tab = 1 Then  'O/T fin
        id = Me.lstProcesados.selectedItem
        origen = 0
    ElseIf Me.SSTab1.Tab = 0 Then  'o/t pend
        id = Me.lstEnProceso.selectedItem
        origen = 0
    End If
    frmRemitosEntregados.origen = 1
    frmRemitosEntregados.idPedidoEntrega = id
    frmRemitosEntregados.caption = "Nro." & id
    frmRemitosEntregados.Show
End Sub
Private Sub facturas_Click()
    If Me.SSTab1.Tab = 2 Then  'O/E
        id = Me.lstEntregas.selectedItem
        origen = 2
    ElseIf Me.SSTab1.Tab = 1 Then  'O/T fin
        id = Me.lstProcesados.selectedItem
        origen = 1
    ElseIf Me.SSTab1.Tab = 0 Then  'o/t pend
        id = Me.lstEnProceso.selectedItem
        origen = 1
    End If
    frmAdminFacturasAplicadas.origen = origen
    frmAdminFacturasAplicadas.idOrigen = CLng(id)
    frmAdminFacturasAplicadas.Show
End Sub
Private Sub Form_Activate()
    On Error Resume Next
    
    DAOMoneda.LlenarCombo Me.cboMondeasProceso
    
        
    DAOCliente.LlenarCombo Me.cboClientes
    
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
    filtros = " and p.fechaEntrega>='" & desde & "' and p.fechaEntrega<='" & hasta & "'"
    grupo = " group by p.id"
    If filtro Then strsql = strsql & filtros
    strsql = strsql & grupo
    Dim nofc As Double
    Dim Fc As Double
    Dim porcno As Double
    Dim tot As Double
    Set rs = conectar.RSFactory(strsql)
    Dim x As ListItem
Dim ots As OrdenTrabajo
    
    While Not rs.EOF
            Set ots = DAOOrdenTrabajo.FindById(rs!id)
            
            Set ots.Detalles = DAODetalleOrdenTrabajo.FindAllByOrdenTrabajo(ots.id, False, False, True)
            
        Set x = Me.lstEnProceso.ListItems.Add(, , rs!ot & " " & Format(rs!id, "0000"))
        x.SubItems(1) = rs!Descripcion
        x.SubItems(2) = rs!FechaEntrega
        x.SubItems(3) = funciones.queMoneda(rs!IdMoneda)
        x.SubItems(4) = funciones.FormatearDecimales(rs!Total, 2)
        'X.SubItems(5) = funciones.formatearDecimales(rs!Facturado, 2)
        x.SubItems(5) = funciones.FormatearDecimales(ots.TotalFacturado * (1 - (ots.Descuento / 100)), 2)
        x.SubItems(6) = funciones.FormatearDecimales(ots.Total - ots.TotalFacturado) 'rs!nofacturado, 2)
        
        
        Fc = rs!Facturado
        nofc = rs!nofacturado
        tot = rs!Total
        porcno = (nofc * 100) / tot
        x.SubItems(7) = funciones.FormatearDecimales(porcno, 1) & "%"
        x.Tag = rs!id
        rs.MoveNext
    Wend

    'o/t finalizados
    strsql = "select p.id,'O/T' as ot,p.fechaCreado,p.fechaEntrega,p.descripcion, p.idmoneda,sum(dp.precio*dp.cantidad*(1-(p.dto/100))) as total, sum(dp.precio*dp.cantidad_facturada*(1-(p.dto/100))) as facturado, sum(dp.precio*(dp.cantidad-dp.cantidad_facturada)*(1-(p.dto/100))) as noFacturado from pedidos p, detalles_pedidos dp where p.idCliente=" & vIdCliente & " and p.estado=4 and dp.idPedido=p.id and p.idMoneda=" & idm & " and dp.precio>=0 "
    grupo = " group by p.id"
    filtros = " and p.fechaEntrega>='" & desde & "' and p.fechaEntrega<='" & hasta & "'"
    If filtro Then strsql = strsql & filtros
    strsql = strsql & grupo

    Set rs = conectar.RSFactory(strsql)
    While Not rs.EOF
        Set x = Me.lstProcesados.ListItems.Add(, , rs!ot & " " & Format(rs!id, "0000"))

        x.SubItems(1) = rs!Descripcion
        x.SubItems(2) = rs!FechaEntrega
        x.SubItems(3) = funciones.queMoneda(rs!IdMoneda)
        x.SubItems(4) = funciones.FormatearDecimales(rs!Total, 2)
        x.SubItems(5) = funciones.FormatearDecimales(rs!Facturado, 2)
        x.SubItems(6) = funciones.FormatearDecimales(rs!nofacturado, 2)
        Fc = rs!Facturado
        nofc = rs!nofacturado
        tot = rs!Total
        If tot > 0 Then
            porcno = (nofc * 100) / tot
        End If
        x.SubItems(7) = funciones.FormatearDecimales(porcno, 1) & "%"
        x.Tag = rs!id
        rs.MoveNext
    Wend
    'o/e
    strsql = "select p.id,'O/E' as ot,p.fechaCreado,p.fecha as fechaEntrega,p.referencia as descripcion, p.idmoneda,sum(dp.vale*dp.cantidad) as total, sum(dp.vale*dp.cantidad_facturada) as facturado, sum(dp.vale*(dp.cantidad-dp.cantidad_facturada)) as noFacturado,c.razon from PedidosEntregas p, detallesPedidosEntregas dp, clientes c where (p.estado=2 or p.estado=3) and dp.idPedidoEntrega=p.id and c.id=p.idCliente and p.idmoneda=" & idm & " and dp.vale>0 and p.idCliente=" & vIdCliente
    grupo = " group by p.id"
    filtros = " and p.fecha>'" & desde & "' and p.fecha<'" & hasta & "'"
    If filtro Then strsql = strsql & filtros
    strsql = strsql & grupo
    Set rs = conectar.RSFactory(strsql)
    While Not rs.EOF
        Set x = Me.lstEntregas.ListItems.Add(, , rs!ot & " " & Format(rs!id, "0000"))
        x.SubItems(1) = rs!Descripcion
        x.SubItems(2) = Format(rs!FechaEntrega, "dd-mm-yyyy")
        x.SubItems(3) = funciones.queMoneda(rs!IdMoneda)
        x.SubItems(4) = funciones.FormatearDecimales(rs!Total, 2)
        x.SubItems(5) = funciones.FormatearDecimales(rs!Facturado, 2)
        x.SubItems(6) = funciones.FormatearDecimales(rs!nofacturado, 2)
        Fc = rs!Facturado
        nofc = rs!nofacturado
        tot = rs!Total
        porcno = (nofc * 100) / tot
        x.SubItems(7) = funciones.FormatearDecimales(porcno, 1) & "%"
        rs.MoveNext
    Wend

    verTotales
End Sub

Public Sub verTotales()
    Dim lst As ListView
    Dim Facturado As Double, Total As Double, nofacturado As Double
    If Me.SSTab1.Tab = 0 Then    'ot proceso
        Set lst = Me.lstEnProceso
    ElseIf SSTab1.Tab = 1 Then    'ot fin
        Set lst = Me.lstProcesados
    ElseIf SSTab1.Tab = 2 Then    'oe
        Set lst = Me.lstEntregas
    End If


    Total = 0
    Facturado = 0
    nofacturado = 0
    For x = 1 To lst.ListItems.count
        Facturado = CDbl(lst.ListItems(x).ListSubItems(5)) + Facturado
        nofacturado = CDbl(lst.ListItems(x).ListSubItems(6)) + nofacturado
        Total = CDbl(lst.ListItems(x).ListSubItems(4)) + Total

    Next x
    Me.lblFacturadoProceso = funciones.FormatearDecimales(Facturado, 2)
    Me.lblTotalProceso = funciones.FormatearDecimales(Total, 2)
    Me.lblNoFacturadoProceso = funciones.FormatearDecimales(nofacturado, 2)
    Set rs = Nothing
End Sub
Private Sub generarInforme()

    llenarListas
    verTotales
End Sub


Private Sub lstEntregas_DblClick()
    If Me.lstEntregas.ListItems.count > 0 Then
        frmVerOe.IDOE = CLng(Me.lstEntregas.selectedItem)
        frmVerOe.Show
    End If

End Sub

Private Sub lstEntregas_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If lstEntregas.ListItems.count > 0 Then
        ida = Me.lstEntregas.selectedItem
        If Button = 2 Then
            Me.nro.caption = "[ " & Format(ida, "0000") & " ]"
            Me.PopupMenu Me.deta
        End If
    End If

End Sub

Private Sub lstProcesados_DblClick()
    If Me.lstProcesados.ListItems.count > 0 Then
        frmDetallePedido.lblIdPedido = Me.lstProcesados.selectedItem
        frmDetallePedido.Frame1.caption = "[ Pedido Nro. " & Format(frmDetallePedido.lblIdPedido, "0000") & " ]"
        frmDetallePedido.Show
    End If
End Sub

Private Sub lstProcesados_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If lstProcesados.ListItems.count > 0 Then
        ida = Me.lstProcesados.selectedItem
        If Button = 2 Then
            Me.nro.caption = "[ " & Format(ida, "0000") & " ]"
            Me.PopupMenu Me.deta
        End If
    End If

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    verTotales
End Sub




