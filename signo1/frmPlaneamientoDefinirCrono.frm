VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmPlaneamientoDefinirCrono 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Definir Entregas..."
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   12420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Agregar"
      Height          =   375
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "Volver"
      Height          =   375
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Definir"
      Default         =   -1  'True
      Height          =   375
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5520
      Width           =   1095
   End
   Begin MSComctlLib.ListView lstCronograma 
      Height          =   2535
      Left            =   120
      TabIndex        =   4
      Top             =   4080
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   4471
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Cantidad"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Fecha"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   255
      Left            =   4320
      TabIndex        =   3
      Top             =   4560
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      Format          =   58523649
      CurrentDate     =   39847
   End
   Begin VB.TextBox txtCantidad 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4320
      TabIndex        =   2
      Text            =   "0"
      Top             =   4200
      Width           =   1335
   End
   Begin MSComctlLib.ListView lstDetallePedido 
      Height          =   3855
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   6800
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
         Text            =   "Item"
         Object.Width           =   1676
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Detalle"
         Object.Width           =   11553
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Cantidad"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Cronograma"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Entrega"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Fecha"
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
      Left            =   3480
      TabIndex        =   1
      Top             =   4560
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cantidad"
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
      Left            =   3240
      TabIndex        =   0
      Top             =   4200
      Width           =   975
   End
End
Attribute VB_Name = "frmPlaneamientoDefinirCrono"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim plan As New classPlaneamiento
Dim vIdPedido As Long
Dim vIdDetallePedido As Long
Public Property Let idpedido(nIdPedido As Long)
    vIdPedido = nIdPedido
End Property
Private Sub Command2_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    FormHelper.Customize Me
    llenarDetallePedido
End Sub
Private Sub llenarDetallePedido()
    Dim x As ListItem
    Me.lstDetallePedido.ListItems.Clear
    Dim rs_pedido As Recordset
    Set rs_pedido = conectar.RSFactory("select dp.idPieza,s.conjunto,dp.fechaEntrega,dp.item,dp.id,dp.idPieza,dp.cantidad,dp.crono_definido,s.detalle from detalles_pedidos dp inner join stock s on dp.idPieza=s.id  where idPedido=" & vIdPedido)
    While Not rs_pedido.EOF And Not rs_pedido.BOF
        Set x = Me.lstDetallePedido.ListItems.Add(, , rs_pedido!item)
        x.SubItems(1) = rs_pedido!detalle
        x.SubItems(2) = rs_pedido!Cantidad
        x.ListSubItems(1).Tag = rs_pedido!conjunto

        x.ListSubItems(2).Tag = rs_pedido!idPieza


        aa = CInt(rs_pedido!crono_definido)
        If aa = 0 Then
            proceso = "Pendiente"
        ElseIf aa = 1 Then
            proceso = "Definido"
        ElseIf aa = 2 Then
            proceso = "No Definido"
        End If
        x.SubItems(3) = proceso
        x.ListSubItems(3).Tag = rs_pedido!crono_definido

        x.SubItems(4) = rs_pedido!FechaEntrega

        x.Tag = rs_pedido!Id

        rs_pedido.MoveNext
    Wend
    Set rs_pedido = Nothing
End Sub
Private Sub lstDetallePedido_ItemClick(ByVal item As MSComctlLib.ListItem)
    If Me.lstDetallePedido.ListItems.count > 0 Then
        vIdDetallePedido = CLng(Me.lstDetallePedido.selectedItem.Tag)
        llenarLstCrono
    End If
End Sub
Private Sub txtCantidad_Validate(Cancel As Boolean)
    If Not IsNumeric(Me.txtCantidad) Then Cancel = True Else Cancel = False
End Sub
Private Sub llenarLstCrono()
    Dim rs As Recordset
    Dim x As ListItem
    Me.lstCronograma.ListItems.Clear
    Set rs = conectar.RSFactory("select * from detalles_pedidos_cronograma where idDetallePedido=" & vIdDetallePedido)
    While Not rs.EOF And Not rs.BOF
        Set x = Me.lstCronograma.ListItems.Add(, , rs!Cantidad)
        x.SubItems(1) = Format(rs!FEcha, "dd-mm-yyyy")
        x.Tag = rs!Id
        rs.MoveNext
    Wend

    Set rs = Nothing
    Set x = Nothing
End Sub
