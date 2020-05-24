VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEntregasMultiples 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cerrar..."
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   10380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Cancel          =   -1  'True
      Caption         =   "Command3"
      Height          =   255
      Left            =   5520
      TabIndex        =   8
      Top             =   8280
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ Orden de trabajo ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10335
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "[ Entregas ]"
         Height          =   2655
         Left            =   120
         TabIndex        =   4
         Top             =   3840
         Width           =   10095
         Begin MSComctlLib.ListView lstEntregas 
            Height          =   2295
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   4048
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
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "Cantidad"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Text            =   "Remito"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Text            =   "Fecha Remito"
               Object.Width           =   3881
            EndProperty
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "[ Detalle ]"
         Height          =   3375
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   10095
         Begin VB.CommandButton Command2 
            Caption         =   "Remitar"
            Height          =   255
            Left            =   6840
            TabIndex        =   7
            Top             =   3000
            Width           =   1455
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Cerrar"
            Height          =   255
            Left            =   8400
            TabIndex        =   5
            Top             =   3000
            Width           =   1575
         End
         Begin MSComctlLib.ListView lstDetallePedido 
            Height          =   2535
            Left            =   120
            TabIndex        =   2
            Top             =   360
            Width           =   9855
            _ExtentX        =   17383
            _ExtentY        =   4471
            View            =   3
            LabelEdit       =   1
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
      End
   End
   Begin VB.Label lblIdOT 
      Caption         =   "Label1"
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   8160
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "frmEntregasMultiples"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim claseP As New classPlaneamiento
Dim idOT As Long

Private Sub Command1_Click()
Dim error1 As Boolean
Dim idP As Long
'verifico que esten todos los ítems fabricados
error1 = False
'verifico q no haya alguno entregado completamente
error2 = False
idP = CLng(Me.lblIdOT)
If Not claseP.estaTodoEntregado(idP) Then

    If Not claseP.estaCerrado(idP) Then
        For nn = 1 To Me.lstDetallePedido.ListItems.count
            cantidad_pedida = CLng(Me.lstDetallePedido.ListItems(nn).ListSubItems(3))
            cantidad_fabricada = CLng(Me.lstDetallePedido.ListItems(nn).ListSubItems(4))
            cantidad_deStock = CLng(Me.lstDetallePedido.ListItems(nn).ListSubItems(6))
            resto = cantidad_fabricada + cantidad_deStock
            claseP.ejecutar_consulta "select estado from pedidos where id=" & idP
            estado = claseP.estadoOT
            If cantidad_pedida > resto Or estado = 4 Then
                error1 = True
            End If
        Next nn

        If Not error1 Then
            'si todo lo pedido esta fabricado o proveniente de stock, proceso a realizar la entrega.
                frmEntregaTotal.Show 1
        End If

    Else
        MsgBox "El pedido se encuentra cerrado", vbInformation, "Información"
        End If
Else
'el pedido ya se entrego, falta cerrar.
    If claseP.CerrarPedido(idP) Then
        MsgBox "El pedido " & idP & " se cerro correctamente.", vbInformation, "Información"
        'Unload Me
    End If
End If



If error1 Then
MsgBox "Para cerrar el pedido debe tener todo fabricado o proveniente de stock.", vbCritical, "Error"
End If
End Sub

Private Sub Command2_Click()
Me.realizaEntrega
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Activate()
idOT = CLng(Me.lblIdOT)
If claseP.ExistePedido(idOT) Then
claseP.llenar_lista_detalle Me.lstDetallePedido, idOT, 1
iditem = Me.lstDetallePedido.SelectedItem
llenarLstEntregas iditem
If claseP.estaCerrado(idOT) Then
 Me.Command1.Enabled = False
Else
 Me.Command1.Enabled = True
End If
End If


End Sub

Private Sub lstDetallePedido_Click()
'iditem = Me.lstDetallePedido.SelectedItem
'llenarLstEntregas iditem

End Sub


Function realizaEntrega()
fabricados = CLng(Me.lstDetallePedido.SelectedItem.ListSubItems(4))
entregados = CLng(Me.lstDetallePedido.SelectedItem.ListSubItems(5))
pedidos = CLng(Me.lstDetallePedido.SelectedItem.ListSubItems(3))
deStock = CLng(Me.lstDetallePedido.SelectedItem.ListSubItems(6))
If Me.lstDetallePedido.ListItems.count > 0 Then
'si no hay elementos del ítem fabricados
If fabricados + deStock = 0 Then
    MsgBox "Para entregar este item debería tenerlo fabricado", vbCritical, "Error"
Else
If pedidos = entregados Then
    MsgBox "El ítem está completamente entregado. No se permiten más entregas", vbCritical, "Error"
Else

    'veo si selecciono un item o varios
  c = 0
  For q = 1 To Me.lstDetallePedido.ListItems.count
   If Me.lstDetallePedido.ListItems(q).Selected = True Then c = c + 1
   Next q
    
    
If c = 1 Then 'se selecciono solo un item

    frmRealizarEntrega.lblIdPieza = Me.lstDetallePedido.SelectedItem
    frmRealizarEntrega.lblPieza = Me.lstDetallePedido.SelectedItem.ListSubItems(2)
    frmRealizarEntrega.lblPedidos = Me.lstDetallePedido.SelectedItem.ListSubItems(3)
    frmRealizarEntrega.lblFabricados = Me.lstDetallePedido.SelectedItem.ListSubItems(4)
    frmRealizarEntrega.lblEntregados = Me.lstDetallePedido.SelectedItem.ListSubItems(5)
    frmRealizarEntrega.lblDeStock = Me.lstDetallePedido.SelectedItem.ListSubItems(6)
    frmRealizarEntrega.lblOt = Me.lblIdOT
    frmRealizarEntrega.lblItem = Me.lstDetallePedido.SelectedItem.ListSubItems(1)
    frmRealizarEntrega.Show 1
Else
    frmEntregasMultiples
End If
End If
End If
End Function


Public Sub llenarLstEntregas(iditem)
On Error GoTo err44
Dim r As Recordset
Set r = claseP.listaRS("select id,cantidad,remito,fecha from entregas where idDetallePedido=" & iditem & " and origen=1") 'origen 1 es de OT
Dim X As ListItem
Me.lstEntregas.ListItems.Clear
While Not r.EOF

Set X = Me.lstEntregas.ListItems.Add(, , r!id)
       X.SubItems(1) = r!cantidad
       X.SubItems(2) = r!remito
       X.SubItems(3) = r!FEcha
r.MoveNext
Wend
Exit Sub
err44:
MsgBox Err.Description

End Sub

Private Sub lstDetallePedido_DblClick()
Me.realizaEntrega
End Sub
