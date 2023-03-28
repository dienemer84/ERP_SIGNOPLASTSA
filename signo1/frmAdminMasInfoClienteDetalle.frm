VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAdminMasInfoClienteDetalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Detalle facturación por orden"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   12750
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "Imprimir"
      Height          =   255
      Left            =   1200
      TabIndex        =   2
      Top             =   6120
      Width           =   975
   End
   Begin MSComctlLib.ListView lstDetalle 
      Height          =   4695
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   8281
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
      NumItems        =   11
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Item"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Detalle"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Cantidad"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Precio"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Total"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Fabricado"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Entregado"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Facturado"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Tot. Fact."
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "No Fact."
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Text            =   "Total No Fact."
         Object.Width           =   1940
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Volver"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   6120
      Width           =   975
   End
   Begin VB.Label lblDescuento 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   1200
      TabIndex        =   10
      Top             =   840
      Width           =   11415
   End
   Begin VB.Label lblDetalle 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   1200
      TabIndex        =   9
      Top             =   600
      Width           =   11415
   End
   Begin VB.Label lblCliente 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   1200
      TabIndex        =   8
      Top             =   360
      Width           =   11415
   End
   Begin VB.Label lbOrigen 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   1200
      TabIndex        =   7
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Descuento"
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
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Detalle"
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
      TabIndex        =   5
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cliente"
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
      TabIndex        =   4
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Origen"
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
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.Menu ver 
      Caption         =   "ver"
      Visible         =   0   'False
      Begin VB.Menu facturas 
         Caption         =   "Facturas..."
      End
   End
End
Attribute VB_Name = "frmAdminMasInfoClienteDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New Recordset
Dim clase As New classPlaneamiento
Dim vorigen As Integer
Dim votoe As Long
Dim Desc As Double
Public Property Let Origen(nOrigen)
    vorigen = nOrigen
End Property
Public Property Let otoe(nOtoe)
    votoe = nOtoe
End Property
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub facturas_Click()
    frmAdminFacturasAplicadas.Origen = -1
    frmAdminFacturasAplicadas.idOrigen = CLng(Me.lstDetalle.selectedItem.Tag)
    'frmAdminFacturasAplicadas.Frame1.Caption = "[ Nro." & Me.lstPedidos.SelectedItem & " ]"
    frmAdminFacturasAplicadas.Show

End Sub

Private Sub Form_Load()
    FormHelper.Customize Me
    Dim x As ListItem
    strsql = "select p.*,c.razon from pedidos p inner join clientes c on p.idcliente=c.id where p.id=" & votoe
    Set rs = conectar.RSFactory(strsql)

    If vorigen = 0 Then
        ort = "O/T"
    ElseIf vorigen = 1 Then
        ort = "O/E"
    End If

    If Not rs.EOF And Not rs.BOF Then
        Me.lblCliente = rs!razon
        Me.lblDescuento = rs!dto & "%"
        Me.lblDetalle = rs!descripcion
        Me.lbOrigen = ort & " " & Format(votoe, "0000")

        Desc = 1 - (rs!dto / 100)

    Else
        Exit Sub
    End If

    strsql = "select dp.id,dp.item,s.detalle,dp.cantidad, dp.precio, (dp.cantidad*dp.precio) as total,dp.cantidad_fabricados as fabricado,dp.cantidad_entregada as entregado,dp.cantidad_Facturada as facturado,(dp.cantidad_facturada*dp.precio) as total_facturado,(dp.cantidad-cantidad_facturada) as no_facturado,((dp.cantidad-cantidad_facturada)*precio) as total_no_facturado from detalles_pedidos dp inner join stock s on dp.idPieza=s.id where idPedido=" & votoe
    Set rs = conectar.RSFactory(strsql)

    While Not rs.EOF
        Set x = Me.lstDetalle.ListItems.Add(, , rs!item)
        x.SubItems(1) = rs!detalle
        x.SubItems(2) = funciones.FormatearDecimales(rs!Cantidad, 2)
        x.SubItems(3) = funciones.FormatearDecimales(rs!Precio * Desc, 2)
        x.SubItems(4) = funciones.FormatearDecimales(rs!Total * Desc, 2)
        x.SubItems(5) = funciones.FormatearDecimales(rs!fabricado, 2)
        x.SubItems(6) = funciones.FormatearDecimales(rs!entregado, 2)
        x.SubItems(7) = funciones.FormatearDecimales(rs!Facturado, 2)
        x.SubItems(8) = funciones.FormatearDecimales(rs!total_facturado * Desc, 2)
        x.SubItems(9) = funciones.FormatearDecimales(rs!no_Facturado, 2)
        x.SubItems(10) = funciones.FormatearDecimales(rs!total_no_facturado * Desc, 2)
        x.Tag = rs!Id
        rs.MoveNext
    Wend
    Set rs = Nothing
End Sub

Private Sub lstDetalle_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Me.lstDetalle.ListItems.count > 0 Then
        If Button = 2 Then
            Me.PopupMenu Me.Ver
        End If
    End If
End Sub


