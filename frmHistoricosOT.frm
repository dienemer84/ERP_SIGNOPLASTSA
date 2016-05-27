VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPlaneamientoHistoricosOT 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Historicos"
   ClientHeight    =   4710
   ClientLeft      =   4785
   ClientTop       =   4275
   ClientWidth     =   7950
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   7950
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ Seleccione elemento ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7935
      Begin VB.CommandButton Command3 
         BackColor       =   &H00E0E0E0&
         Cancel          =   -1  'True
         Caption         =   "Salir"
         Height          =   375
         Left            =   6840
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   4200
         Width           =   975
      End
      Begin MSComctlLib.ListView lst 
         Height          =   2895
         Left            =   120
         TabIndex        =   6
         Top             =   1200
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   5106
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Presu"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Cant"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Detalle presupuesto"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Valor"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Fecha"
            Object.Width           =   3669
         EndProperty
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ver"
         Default         =   -1  'True
         Height          =   375
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0C0C0&
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
         Top             =   4320
         Width           =   7695
      End
      Begin VB.Label lblCliente 
         BackColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   1440
         TabIndex        =   4
         Top             =   840
         Width           =   3975
      End
      Begin VB.Label Label2 
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
         Left            =   600
         TabIndex        =   3
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Código pieza"
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
         TabIndex        =   1
         Top             =   480
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmPlaneamientoHistoricosOT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim idpiezaelegida As Long
Dim baseS As New classStock
Private Sub Command1_Click()
On Error Resume Next
Dim codigo As String
Dim I As Long
Dim rs As Recordset
Dim X As ListItem
codigo = Trim(Me.txtCodigo)

Set rs = baseS.CrearRS("select count(id) as cant from sp.stock where id=" & idpiezaelegida)
If rs!Cant = 1 Then

Set rs = baseS.CrearRS("select c.razon,p.fecha,dp.cantidad,dp.idPedido,dp.valor,p.detalle from sp.clientes c, sp.detalles_pedidos dp,sp.pedidos p where p.id=dp.idPedido and dp.idpieza=" & idpiezaelegida & " and p.idCliente=c.id")

Me.lst.ListItems.Clear
cantP = 0
valo = 0

  While Not rs.EOF
  Me.lblCliente = rs!razon
  Set X = Me.lst.ListItems.Add(, , Format(rs!idpedido, "0000"))
    X.SubItems(1) = rs!cantidad
    cantP = cantP + rs!cantidad
    X.SubItems(2) = rs!detalle
    X.SubItems(3) = rs!valor
    valo = valo + (rs!valor * rs!cantidad)
    
    X.SubItems(4) = rs!FEcha
  rs.MoveNext
  Wend
'Me.lblInfo = "Total Piezas: " & cantP & " unidades, total cotizado: $" & valo & ", Promedio venta: $" & Math.Round(valo / cantP, 2)
End If
Set rs = Nothing
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
txtCodigo_Change
End Sub
Private Sub txtCodigo_Change()
If Trim(Me.txtCodigo) = Empty Then
Me.Command1.Enabled = False
Else
Me.Command1.Enabled = True
End If
End Sub
Private Sub txtCodigo_DblClick()
frmListarStock_seleccion.Show 1
Me.txtCodigo = funciones.quePiezaElegidaDetalle
idpiezaelegida = funciones.quePiezaElegida
Command1_Click
End Sub
Private Sub txtCodigo_GotFocus()
foco Me.txtCodigo
End Sub
