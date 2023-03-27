VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmRemitar 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Remitar entrega..."
   ClientHeight    =   6045
   ClientLeft      =   -75
   ClientTop       =   3360
   ClientWidth     =   8235
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   8235
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   255
      Left            =   6720
      TabIndex        =   15
      Top             =   5640
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ Realizar entrega ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   4080
      TabIndex        =   4
      Top             =   3480
      Width           =   4095
      Begin VB.CommandButton Command3 
         Caption         =   "Cerrar"
         Height          =   255
         Left            =   1440
         TabIndex        =   16
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox txtRemito 
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   11
         ToolTipText     =   "Doble click para seleccionar remito"
         Top             =   1560
         Width           =   3015
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Agregar"
         Default         =   -1  'True
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox txtCantEntregar 
         Height          =   285
         Left            =   960
         TabIndex        =   8
         Top             =   1200
         Width           =   3015
      End
      Begin VB.Label idPe 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Label1"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   2160
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lblPedidas 
         BackColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   960
         TabIndex        =   13
         Top             =   840
         Width           =   2895
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pedidas"
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
         TabIndex        =   12
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Remito"
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
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label2 
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
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label lblDetalle 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Label1"
         Height          =   615
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   3855
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ Entregas realizadas ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   0
      TabIndex        =   3
      Top             =   3480
      Width           =   3975
      Begin MSComctlLib.ListView lstEntregas 
         Height          =   2175
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   3836
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Cant"
            Object.Width           =   1094
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Remito"
            Object.Width           =   1217
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Fecha"
            Object.Width           =   3669
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8175
      Begin MSComctlLib.ListView lstDetalleEntrega 
         Height          =   2895
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   5106
         View            =   3
         Arrange         =   1
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Pieza"
            Object.Width           =   8414
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Cantidad"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Entregadas"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Valor"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "idDetalleEntrega"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Label idPedidoEntrega 
      Caption         =   "Label1"
      Height          =   255
      Left            =   7560
      TabIndex        =   2
      Top             =   3600
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "frmRemitar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim claseP As New classPlaneamiento
Dim pedidos As Long, Entregados As Long
Attribute Entregados.VB_VarUserMemId = 1073938433

Public Sub listaOE()
    Dim rs As Recordset
    Dim x As ListItem
    Dim idPEntrega As Long
    idPEntrega = CLng(Me.idPedidoEntrega)
    Set rs = conectar.RSFactory("select s.detalle,pe.id,pe.cantidad,pe.entregados,pe.vale from stock s inner join detallesPedidosEntregas pe on pe.idPieza=s.id and pe.idPedidoEntrega=" & idPEntrega)
    Me.lstDetalleEntrega.ListItems.Clear
    While Not rs.EOF
        Set x = Me.lstDetalleEntrega.ListItems.Add(, , rs!detalle)
        x.SubItems(1) = rs!Cantidad
        x.SubItems(2) = rs!Entregados
        x.SubItems(3) = rs!vale
        x.Tag = rs!Id
        rs.MoveNext
    Wend
End Sub

Private Sub Command1_Click()
    If Trim(Me.txtRemito) = Empty Then Exit Sub
    CANTIDAD_items = 0
    For x = 1 To Me.lstDetalleEntrega.ListItems.count
        If Me.lstDetalleEntrega.ListItems(x).Selected Then
            CANTIDAD_items = CANTIDAD_items + 1
        End If
    Next x

    If CANTIDAD_items = 1 Then
        If pedidos > Entregados Then
            If MsgBox("¿Está seguro de remitar " & Trim(Me.txtCantEntregar) & " unidades de este ítem?", vbYesNo, "Confirmación") = vbYes Then
                modo = 1

                alfa = claseP.RealizarEntrega(modo, CLng(Me.txtRemito), CLng(Me.txtCantEntregar), CLng(Me.idPe), CLng(Me.idPedidoEntrega), 2)
            End If
        End If

    Else
        modo = 3    'entrega multiple per ono total


        Dim v() As Long

        c = 1
        For o = 1 To Me.lstDetalleEntrega.ListItems.count
            ReDim Preserve v(c) As Long
            If Me.lstDetalleEntrega.ListItems(o).Selected Then
                v(c) = Me.lstDetalleEntrega.ListItems(o).Tag
                c = c + 1
            End If
        Next o
        If MsgBox("¿Desea realizar la entrega de los items seleccionados en el remito " & CLng(Me.txtRemito) & "?", vbYesNo, "Confirmación") Then
            alfa = claseP.RealizarEntrega(modo, CLng(Me.txtRemito), CLng(Me.txtCantEntregar), CLng(Me.idPe), CLng(Me.idPedidoEntrega), 2, v)
        End If
    End If






    If alfa Then
        'verificar si esta entregado parcial o totalmente para cambiar estado
        'en el pedido.
        If claseP.verificarOrdenEntrega(CLng(Me.idPedidoEntrega), True) Then
            'devuelve verdadero si está todo lo pedido, Remitado.
            If MsgBox("Orden completamente remitada ¿Proceder con la entrega?", vbYesNo, "Confirmación") = vbYes Then
                'si procede con la entrega, cambio el estado del pedido a 3 que es estado cerrado.

            End If



        End If



        verMarcado
    End If

    'Else
    '   MsgBox "Item completamente entregado", vbExclamation, "Error"
    'End If
End Sub

Private Sub Command2_Click()
    If MsgBox("¿Está seguro de salir?", vbYesNo, "Confirmación") = vbYes Then
        Unload Me
    End If

End Sub

Private Sub Command4_Click()

End Sub

Private Sub Form_Activate()
    listaOE
    verMarcado
    validar
End Sub
Private Sub verMarcado()
    can = 0
    For x = 1 To Me.lstDetalleEntrega.ListItems.count
        If Me.lstDetalleEntrega.ListItems(x).Selected Then
            can = can + 1
        End If
    Next x
    If can = 1 Then
        If Me.lstDetalleEntrega.ListItems.count > 0 Then
            Me.lblPedidas = Me.lstDetalleEntrega.selectedItem.ListSubItems(1)
            Me.lblDetalle = Me.lstDetalleEntrega.selectedItem
            Me.idPe = Me.lstDetalleEntrega.selectedItem.Tag
            Dim rs As Recordset
            pedidos = CLng(Me.lstDetalleEntrega.selectedItem.ListSubItems(1))
            Entregados = CLng(Me.lstDetalleEntrega.selectedItem.ListSubItems(2))

            saldo = pedidos - Entregados
            Me.txtCantEntregar = saldo
            If Entregados < pedidos Then
                'si todavia no se entrego todo lo pedido
                Me.Command1.Enabled = True
            Else
                Me.Command1.Enabled = False
            End If
            llenarLstEntregas
        End If
    Else
        Me.Command1.Enabled = True
    End If
End Sub

Private Sub Form_Load()
    FormHelper.Customize Me
End Sub

Private Sub lstDetalleEntrega_ItemClick(ByVal item As MSComctlLib.ListItem)
    verMarcado
End Sub

Private Sub txtCantEntregar_Change()
    validar
End Sub

Private Sub txtCantEntregar_Validate(Cancel As Boolean)
    If Not IsNumeric(Me.txtCantEntregar) Then Cancel = True
End Sub

Private Sub txtRemito_Change()
    validar

End Sub

Private Sub validar()

    If Trim(Me.txtCantEntregar) = Empty Or Trim(Me.txtRemito) = Empty Or Entregados >= pedidos Then
        Me.Command1.Enabled = False
    Else
        Me.Command1.Enabled = True
    End If

End Sub
Public Sub llenarLstEntregas()
    Dim rs As Recordset
    Dim x As ListItem
    Dim idPEntrega As Long
    idPEntrega = CLng(Me.lstDetalleEntrega.selectedItem.Tag)
    Set rs = conectar.RSFactory("select r.estado,e.cantidad,e.remito,e.fecha from entregas e inner join remitos r on e.remito=r.id where origen=2 and e.idDetallePedido=" & idPEntrega)
    Me.lstEntregas.ListItems.Clear
    While Not rs.EOF
        Set x = Me.lstEntregas.ListItems.Add(, , rs!Cantidad)
        If rs!estado = 3 Then    'anulado
            rto = "*" & rs!Remito
        Else
            rto = rs!Remito
        End If

        x.SubItems(1) = rto
        x.SubItems(2) = rs!FEcha

        If rs!estado = 3 Then x.ListSubItems(1).ForeColor = vbRed
        rs.MoveNext
    Wend

End Sub

Private Sub txtRemito_DblClick()
    Dim strsql As String
    Dim idRto As Long
    Dim r As Recordset

    frmPlaneamientoRemitosListaProceso.mostrar = 0
    frmPlaneamientoRemitosListaProceso.Show 1
    If funciones.queRemitoElegido <> -1 Then
        Me.txtRemito = funciones.queRemitoElegido
    Else
        Me.txtRemito = Empty
    End If
End Sub
