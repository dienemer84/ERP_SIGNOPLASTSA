VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmPlaneamientoSeguimiento 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Seguimiento..."
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   13275
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   13275
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   11760
      TabIndex        =   22
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ Datos ]"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   0
      TabIndex        =   10
      Top             =   1800
      Width           =   13215
      Begin VB.CommandButton Command2 
         Caption         =   "Todo fabricado"
         Height          =   255
         Left            =   3840
         TabIndex        =   21
         Top             =   3960
         Width           =   1335
      End
      Begin VB.CommandButton Ac 
         Caption         =   "Actualizar"
         Height          =   255
         Left            =   2880
         TabIndex        =   18
         Top             =   3960
         Width           =   855
      End
      Begin MSComctlLib.ListView lstDetallePedidos 
         Height          =   2655
         Left            =   120
         TabIndex        =   17
         Top             =   480
         Width           =   12935
         _ExtentX        =   22807
         _ExtentY        =   4683
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
         NumItems        =   11
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "id"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Item"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Detalle"
            Object.Width           =   9701
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Cant"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Fecha"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Fabricados"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Entregas"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Stock"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "de stock"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "% Tarea"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Prom "
            Object.Width           =   1411
         EndProperty
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1320
         TabIndex        =   16
         ToolTipText     =   "Puede utilizar un número negativo para restar al seguimiento"
         Top             =   3960
         Width           =   1455
      End
      Begin VB.Label lblItem 
         BackColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   9720
         TabIndex        =   20
         Top             =   3960
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblIdPedido 
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   11040
         TabIndex        =   19
         Top             =   3840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fabricados"
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
         Top             =   3960
         Width           =   1095
      End
      Begin VB.Label lblTerminados 
         BackColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   1320
         TabIndex        =   14
         Top             =   3600
         Width           =   1335
      End
      Begin VB.Label lblPedidos 
         BackColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   1320
         TabIndex        =   13
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pedidos"
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
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Terminados"
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
         TabIndex        =   11
         Top             =   3600
         Width           =   1215
      End
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
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13215
      Begin VB.CommandButton Command1 
         Caption         =   "Ver"
         Default         =   -1  'True
         Height          =   285
         Left            =   3120
         TabIndex        =   3
         Top             =   360
         Width           =   750
      End
      Begin VB.TextBox txtOT 
         Height          =   285
         Left            =   1920
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblfechaEntrega 
         BackColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   1920
         TabIndex        =   9
         Top             =   1320
         Width           =   7965
      End
      Begin VB.Label lblDescripcion 
         BackColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   1440
         TabIndex        =   8
         Top             =   1080
         Width           =   10725
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha de entrega"
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
         Left            =   240
         TabIndex        =   7
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción"
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
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label3 
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
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lblCliente 
         BackColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   1080
         TabIndex        =   4
         Top             =   840
         Width           =   7965
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Número de Orden"
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
         Left            =   300
         TabIndex        =   1
         Top             =   390
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmPlaneamientoSeguimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim claseP As New classPlaneamiento
Dim claseS As New classStock

Private Sub Ac_Click()
    Dim idP As Long
    Dim iditem As Long
    Dim Cant As Double
    If IsNumeric(Me.Text1) Then
        'si la cant fabricad es numerica, procedo.
        Cant = CDbl(Me.Text1)
        'If Cant > 0 Then
        'si fabrico número positivo de piezas, continuo.
        idP = CLng(Me.lblIdPedido)
        iditem = CLng(Me.lblItem)
        h = MsgBox("¿Está seguro de agregar " & Cant & " piezas fabricadas?", vbYesNo, "Confirmación")
        If h = vbYes Then
            'actualizo el campo fabricados

            claseP.ejecutarComando "update detalles_pedidos set cantidad_fabricados=cantidad_fabricados+" & Cant & " where idPedido=" & idP & " and id= " & iditem

            Dim detalle_pedido As New DetalleOrdenTrabajo

            Set detalle_pedido = DAODetalleOrdenTrabajo.FindById(iditem)
            claseP.ejecutarComando "update stock set ya_fabricado=" & conectar.Escape(True) & " where id= " & detalle_pedido.Pieza.id

            DAODetalleOrdenTrabajo.SaveCantidad iditem, Cant, CantidadFabricada_, 0, 9, 0, 0, 0




            'veo si es conjunto para mandar a fabricar las partes internas




            claseP.ejecutar_consulta "select cantidad_fabricados, cantidad, reserva_stock from detalles_pedidos where idPedido=" & idP
            pos = Me.lstDetallePedidos.selectedItem
            fab = claseP.Fabricados
            ped = claseP.pedidos
            stock = claseP.reservados
            If claseP.estaTodoFabricado(idP) Then Me.Command1.Enabled = False

            llenar_lista_detalle Me.lstDetallePedidos, idP, pos



        End If
        'Else
        ' MsgBox "Debe fabricar un número positivo de elementos", vbCritical, "Error"
        'End If
    End If

End Sub

Private Sub Command1_Click()
    On Error GoTo er1

    Dim idP As Long
    If Trim(Me.txtOT) <> Empty Then
        idP = CLng(Me.txtOT)
        A = claseP.ExistePedido(idP)
        If A = 0 Or A = -1 Then
            MsgBox "Dato inválido.", vbCritical, "Error"
            Me.lstDetallePedidos.ListItems.Clear
            Me.lblTerminados = Empty
            Me.lblPedidos = Empty
            Frame2.Enabled = False
            Me.lblIdPedido = Empty
        Else
            claseP.ejecutar_consulta "select c.razon as cliente,p.tipo_orden,p.estado,p.descripcion,p.fechaEntrega from pedidos p,clientes c where p.id=" & idP & " and c.id=p.idCliente"
            estado = claseP.estadoOT
            If claseP.TipoOrden <> OT_TRADICIONAL Then
                MsgBox "OT Invalida para hacer el seguimiento!", vbInformation
                Exit Sub
            End If
            


            '    If estado = 2 Then 'esta en proceso
            lblCliente = claseP.cliente
            lblDescripcion = claseP.descripcion
            lblfechaEntrega = claseP.FechaEntrega
            llenar_lista_detalle Me.lstDetallePedidos, idP, 1
            Frame2.Enabled = True
            Me.lblIdPedido = idP
            Me.lblItem = Me.lstDetallePedidos.selectedItem
            verSeleccionado
            'Else
            If estado = 3 Or estado = 4 Or estado = 6 Then
                'deshabilito la opcion de modificar cantidades
                MsgBox "La OT no está actualmente en proceso.", vbCritical, "Error"
                Me.Ac.Enabled = False
                Me.Text1.Enabled = False
                Me.Command2.Enabled = False
            End If
            'End If
        End If
    End If

    Exit Sub
er1:

    Me.txtOT = Empty
End Sub
Private Sub verSeleccionado()
    Me.lblPedidos = Me.lstDetallePedidos.selectedItem.ListSubItems(3) & " unidades"
    Me.lblTerminados = Me.lstDetallePedidos.selectedItem.ListSubItems(5) & " unidades"
    Me.lblItem = Me.lstDetallePedidos.selectedItem
End Sub
Private Sub Command2_Click()
    Dim idP As Long
    Dim Cant As Double
    Dim lbitem As Long
    If MsgBox("¿Está seguro de fabricar toda la orden?", vbYesNo, "Confirmación") = vbYes Then
        For P = 1 To Me.lstDetallePedidos.ListItems.count
            pedidos = CDbl(Me.lstDetallePedidos.ListItems(P).ListSubItems(3))
            term = CDbl(Me.lstDetallePedidos.ListItems(P).ListSubItems(5))

            If term < pedidos Then
                faltantes = pedidos - term
            Else
                faltantes = pedidos
            End If
            If faltantes > 0 Then
                Cant = faltantes    '+ (pedidos - faltantes)
                If pedidos > term Then
                    'Me.lstDetallePedidos.ListItems(P).ListSubItems(4).Text = faltantes + (pedidos - faltantes)
                    lbPedidos = Me.lstDetallePedidos.ListItems(P).ListSubItems(3)
                    lbTerminados = Me.lstDetallePedidos.ListItems(P).ListSubItems(5)
                    lbitem = Me.lstDetallePedidos.ListItems(P)
                    idP = CLng(Me.lblIdPedido)
                    'actualizo el campo fabricados
                    claseP.ejecutar_consulta "update detalles_pedidos set cantidad_fabricados=cantidad_fabricados+" & Cant & " where idPedido=" & idP & " and id= " & lbitem
                    DAODetalleOrdenTrabajo.SaveCantidad lbitem, Cant, CantidadFabricada_, 0, 9, 0, 0, 0

                    claseP.ejecutar_consulta "select cantidad_fabricados, cantidad, reserva_stock from detalles_pedidos where idPedido=" & idP
                    fab = claseP.Fabricados
                    ped = claseP.pedidos
                    stock = claseP.reservados
                End If
            End If
        Next P

        idP = CLng(Me.lblIdPedido)

        If claseP.estaTodoFabricado(idP) Then Me.Command1.Enabled = False
        llenar_lista_detalle Me.lstDetallePedidos, idP, pos
    End If

End Sub
Private Sub Command3_Click()
    If MsgBox("¿Está seguro de salir?", vbYesNo, "Confirmación") = vbYes Then
        Unload Me
    End If
End Sub



Private Sub Form_Load()
    FormHelper.Customize Me
End Sub

Private Sub lstDetallePedidos_DblClick()
    pedidos = CLng(Me.lstDetallePedidos.selectedItem.ListSubItems(3))
    term = CLng(Me.lstDetallePedidos.selectedItem.ListSubItems(5))

    If term < pedidos Then
        faltantes = pedidos - term
    Else
        faltantes = pedidos
    End If
    If faltantes > 0 Then
        Me.Text1 = faltantes
    End If
End Sub
Private Sub lstDetallePedidos_ItemClick(ByVal item As MSComctlLib.ListItem)
    verSeleccionado
End Sub
Private Sub Text1_GotFocus()
    Me.Command1.default = True
    foco Me.Text1
End Sub
Private Sub txtOT_GotFocus()
    foco Me.txtOT
End Sub
Private Sub llenar_lista_detalle(lst As ListView, idpedido, Optional pos)
    On Error GoTo eja
    Dim rs As Recordset
    'trae los detalles del pedido con su pieza
    Set rs = conectar.RSFactory("select dp.fechaEntrega,if(dp.retirado=1,dp.reserva_stock,0) as deStock,dp.nota,dp.id,dp.idpieza,dp.item,dp.cantidad,dp.cantidad_entregada as entregados,dp.cantidad_fabricados as fabricados,s.detalle from detalles_pedidos dp,stock s where s.id=dp.idpieza and idPedido=" & idpedido & " group by dp.id")
    lst.ListItems.Clear
    Dim x As ListItem
    conta = 0
    While Not rs.EOF
        conta = conta + 1
        Set x = lst.ListItems.Add(, , Format(rs!id, "000"))    'si no anda, dejar id solo
        x.SubItems(1) = rs!item
        If (rs!Nota) = Empty Then
            x.SubItems(2) = rs!detalle
        Else
            x.SubItems(2) = rs!detalle & ", " & rs!Nota
        End If

        x.SubItems(3) = rs!Cantidad
        x.SubItems(4) = rs!FechaEntrega
        x.SubItems(5) = rs!Fabricados
        x.SubItems(6) = rs!Entregados
        x.SubItems(7) = rs!deStock
        x.SubItems(8) = conta
        Dim pp As Double
        Dim prom_fab As Double
        prom_fab = 0
        'de cada detalle del pedido se calcula su avance
        'claseP.AvanceProceso rs!id, pp, prom_fab
        'DAODetalleOrdenTrabajo.CalcularPorcentajeAvanceYPromedioFabricado rs!id, pp, prom_fab

        If pp > -1 Then
            x.SubItems(9) = pp & "%"
        Else
            x.SubItems(9) = "S/D"
        End If

        x.SubItems(10) = prom_fab
        x.Tag = Format(rs!id, "0000")
        If IsNumeric(pos) Then
            If x = pos Then
                x.Selected = True
                x.EnsureVisible
            End If
        End If
        A = DateDiff("d", Date, rs!FechaEntrega)
        If A < 0 Then
            x.ListSubItems(1).ForeColor = vbRed
            x.ListSubItems(4).ForeColor = vbRed
        ElseIf A = 0 Then
            x.ListSubItems(1).ForeColor = vbGreen
            x.ListSubItems(4).ForeColor = vbGreen
        End If
        rs.MoveNext
    Wend
    Exit Sub
eja:
    MsgBox Err.Description
End Sub



