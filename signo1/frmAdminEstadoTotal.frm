VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmAdminResumenEstadoTotal 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Estado total de Situación"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12540
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   12540
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
      Height          =   1335
      Left            =   11040
      TabIndex        =   1
      Top             =   5160
      Width           =   1335
      Begin VB.CommandButton Command5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Imprimir"
         Height          =   255
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
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
         TabIndex        =   3
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Exportar"
         CausesValidation=   0   'False
         Height          =   255
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   600
         Width           =   1095
      End
   End
   Begin MSComctlLib.ListView lstTotal 
      Height          =   5055
      Left            =   0
      TabIndex        =   0
      Top             =   0
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Razón Social"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Moneda"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Total"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Facturado"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "No Facturado"
         Object.Width           =   3528
      EndProperty
   End
End
Attribute VB_Name = "frmAdminResumenEstadoTotal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strsql As String
'Dim claseA As New classAdministracion

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    FormHelper.Customize Me
    Dim rs As New Recordset
    'strsql = "select c.razon,m.nombre_corto,sum(dp.precio*dp.cantidad) as tot,sum(dp.cantidad_facturada*dp.precio) as factu, sum((dp.cantidad-dp.cantidad_facturada)*dp.precio)as nofactu from detalles_pedidos dp inner join pedidos p on dp.idPedido=p.id inner join clientes c on p.idCliente=c.id inner join AdminConfigMonedas m on p.idMoneda=m.id group by p.idCliente, p.idMoneda"
    strsql = "select c.razon,m.nombre_corto,sum(dp.precio*dp.cantidad*(1-(p.dto/100))) as tot,sum(dp.cantidad_facturada*dp.precio*(1-(p.dto/100))) as factu, sum((dp.cantidad-dp.cantidad_facturada)*dp.precio*(1-(p.dto/100)))as nofactu from detalles_pedidos dp inner join pedidos p on dp.idPedido=p.id inner join clientes c on p.idCliente=c.id inner join AdminConfigMonedas m on p.idMoneda=m.id group by p.idCliente, p.idMoneda"
    Dim x As ListItem
    Set rs = conectar.RSFactory(strsql)
    While Not rs.EOF
        Set x = Me.lstTotal.ListItems.Add(, , rs!razon)
        x.SubItems(1) = rs!Nombre_corto
        x.SubItems(2) = funciones.FormatearDecimales(rs!tot, 2)
        x.SubItems(3) = funciones.FormatearDecimales(rs!FACTU, 2)
        x.SubItems(4) = funciones.FormatearDecimales(rs!NOFACTU, 2)

        rs.MoveNext
    Wend

End Sub
