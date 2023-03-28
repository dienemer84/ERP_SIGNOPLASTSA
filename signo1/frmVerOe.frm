VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPlaneamientoOEVer 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ver Orden de entrega"
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6240
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtReferencia 
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   120
      Width           =   5175
   End
   Begin VB.TextBox txtCreado 
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   480
      Width           =   5175
   End
   Begin VB.TextBox txtEntrega 
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   840
      Width           =   5175
   End
   Begin VB.TextBox txtCliente 
      Height          =   285
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1320
      Width           =   5295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Imprimir"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   5160
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   255
      Left            =   1080
      TabIndex        =   0
      Top             =   5160
      Width           =   855
   End
   Begin MSComctlLib.ListView lstOe 
      Height          =   3255
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   5741
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Cantidad"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Detalle"
         Object.Width           =   6526
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Valor"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Referencia"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cliente"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Creado"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Entrega"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   855
   End
End
Attribute VB_Name = "frmPlaneamientoOEVer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strsql As String
Dim rs As Recordset
Dim vidOe As Long
Dim classP As New classPlaneamiento
Public Property Let IDOE(nidoe As Long)
    vidOe = nidoe
End Property

Private Sub Command1_Click()
    If Me.lstOE.ListItems.count > 0 Then
        classP.imprimirOrdenEntrega (vidOe)
    End If

End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    Me.lstOE.ListItems.Clear
    strsql = "select c.razon,p.referencia,p.fecha,p.fechaCreado from PedidosEntregas p inner join clientes c on p.idCliente=c.id where p.id=" & vidOe
    Set rs = conectar.RSFactory(strsql)
    ref = rs!referencia
    fecEnt = rs!FEcha
    fecCre = rs!fechaCreado
    razon = rs!razon
    Me.txtCliente = razon
    Me.txtCreado = fecCre
    Me.txtEntrega = fecEnt
    Me.txtReferencia = ref
    strsql = "select pe.id ,s.detalle, pe.cantidad,pe.vale from detallesPedidosEntregas pe inner join stock s on pe.idPieza = s.id where pe.IdPedidoEntrega=" & vidOe
    Set rs = conectar.RSFactory(strsql)
    While Not rs.EOF
        Set x = Me.lstOE.ListItems.Add(, , rs!Cantidad)
        x.SubItems(1) = rs!detalle
        x.SubItems(2) = funciones.FormatearDecimales(rs!vale, 2)
        rs.MoveNext
    Wend
End Sub

Private Sub Form_Load()
    FormHelper.Customize Me
    'Me.Frame1.Caption = "[ O/E " & Format(vidOe, "0000") & " ]"
End Sub

