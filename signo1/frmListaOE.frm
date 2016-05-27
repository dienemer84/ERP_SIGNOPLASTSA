VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPlaneamientoOELista 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ordenes de entrega..."
   ClientHeight    =   5175
   ClientLeft      =   1800
   ClientTop       =   2985
   ClientWidth     =   10980
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   10980
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "Volver"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4680
      Width           =   1095
   End
   Begin MSComctlLib.ListView lstOE 
      Height          =   4575
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   8070
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Número"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Cliente"
         Object.Width           =   3704
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Referencia"
         Object.Width           =   6315
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Fecha"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Usuario"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Estado"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Menu entrega 
      Caption         =   "entregas"
      Visible         =   0   'False
      Begin VB.Menu OENumero 
         Caption         =   "Numero"
         Enabled         =   0   'False
      End
      Begin VB.Menu vereditar 
         Caption         =   "vereditar"
      End
      Begin VB.Menu as234 
         Caption         =   "-"
      End
      Begin VB.Menu AprobarOE 
         Caption         =   "Aprobar"
      End
      Begin VB.Menu remitar 
         Caption         =   "Remitar..."
      End
      Begin VB.Menu cerrarOE 
         Caption         =   "Cerrar..."
      End
      Begin VB.Menu verHistorialOE 
         Caption         =   "Ver historial..."
      End
      Begin VB.Menu nada 
         Caption         =   "-"
      End
      Begin VB.Menu RtosEntregados 
         Caption         =   "Remitos entregados...."
      End
      Begin VB.Menu printOrder 
         Caption         =   "Imprimir..."
      End
   End
End
Attribute VB_Name = "frmPlaneamientoOELista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim claseP As New classPlaneamiento
Dim rs As Recordset
Dim vereditarOE As Integer
Public Sub listaOE()
    Dim x As ListItem
    Me.lstOE.ListItems.Clear

    Set rs = conectar.RSFactory("select pe.estado,pe.id,pe.referencia,pe.fecha,u.usuario,c.razon from PedidosEntregas pe inner join clientes c on pe.idCliente=c.id inner join usuarios u on pe.usuario=u.id")
    While Not rs.EOF
        Set x = Me.lstOE.ListItems.Add(, , Format(rs!id, "0000"))
        x.SubItems(1) = rs!razon
        x.SubItems(2) = rs!referencia
        x.SubItems(3) = rs!FEcha
        x.SubItems(4) = rs!usuario
        x.SubItems(5) = funciones.estado_OE(rs!estado)
        If rs!estado = 1 Then
            x.ListSubItems(5).ForeColor = vbRed
            x.ForeColor = vbRed
        ElseIf rs!estado = 2 Then
            x.ListSubItems(5).ForeColor = vbBlue
            x.ForeColor = vbBlue
        ElseIf rs!estado = 3 Then
            x.ListSubItems(5).ForeColor = vbGreen
            x.ForeColor = vbGreen
        ElseIf rs!estado = 4 Then
            x.ListSubItems(5).ForeColor = vbMagenta
            x.ForeColor = vbMagenta
        End If
        rs.MoveNext
    Wend
End Sub


Private Sub AprobarOE_Click()
    Dim vidOe As Long
    vidOe = CLng(Me.lstOE.selectedItem)
    If MsgBox("¿Está seguro de aprobar la O/E?", vbYesNo, "Confirmación") = vbYes Then
        If claseP.AprobarOrdenEntrega(vidOe) Then
            MsgBox "Orden de Entrega aprobada con éxito!", vbInformation, "Información"
        Else
            MsgBox "Se produjo un error al aprobar la OE!", vbCritical, "Error"
        End If
    End If

End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Activate()

    Me.Refresh
    Me.lstOE.Refresh
End Sub

Private Sub Form_Load()
    FormHelper.Customize Me
    listaOE
End Sub

Private Sub lstOE_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    funciones.LstOrdenar Me.lstOE, CInt(ColumnHeader.index)
End Sub

Private Sub lstOE_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Me.lstOE.ListItems.count > 0 Then
        If Button = 2 Then
            IDOE = Me.lstOE.selectedItem
            Me.OENumero.caption = "[ Nro. " & IDOE & " ]"
            Set rs = conectar.RSFactory("select estado from PedidosEntregas where id=" & IDOE)

            If rs!estado = 1 Then    'pendiente
                Me.vereditar.caption = "Editar..."
                Me.remitar.Enabled = False
                Me.cerrarOE.Enabled = False
                Me.RtosEntregados.Enabled = False
                vereditarOE = 1
                Me.printOrder.Enabled = False
                Me.AprobarOE.Enabled = True
            ElseIf rs!estado = 2 Then    'aprobado
                Me.remitar.Enabled = True
                Me.vereditar.Enabled = True
                Me.cerrarOE.Enabled = False
                Me.RtosEntregados.Enabled = False
                Me.vereditar.caption = "Ver..."
                vereditarOE = 3
                Me.printOrder.Enabled = True
                Me.AprobarOE.Enabled = False
            ElseIf rs!estado = 4 Then    'entregada
                Me.remitar.Enabled = False
                Me.cerrarOE.Enabled = True
                Me.RtosEntregados.Enabled = False
                Me.vereditar.caption = "Ver..."
                vereditarOE = 3
                Me.printOrder.Enabled = False
                Me.AprobarOE.Enabled = False
            ElseIf rs!estado = 3 Then    'finalizada
                Me.vereditar.caption = "Ver..."
                vereditarOE = 3
                Me.remitar.Enabled = False
                Me.cerrarOE.Enabled = False
                Me.RtosEntregados.Enabled = True
                Me.printOrder.Enabled = False
                Me.AprobarOE.Enabled = False
            End If

            If Not Permisos.planOEaprobaciones Then Me.AprobarOE.Enabled = False
            Me.PopupMenu entrega
        End If
    End If

End Sub

Private Sub printOrder_Click()
    If Me.lstOE.ListItems.count > 0 Then
        claseP.imprimirOrdenEntrega (CLng(Me.lstOE.selectedItem))
    End If


End Sub

Private Sub remitar_Click()
    If Me.lstOE.ListItems.count > 0 Then
        frmRemitar.idPedidoEntrega = CLng(Me.lstOE.selectedItem)
        frmRemitar.idPe = CLng(Me.lstOE.selectedItem)
        frmRemitar.Frame1.caption = "[ Nro." & Me.lstOE.selectedItem & " ]"
        frmRemitar.Show
    End If
End Sub


Private Sub RtosEntregados_Click()
    If Me.lstOE.ListItems.count > 0 Then
        frmRemitosEntregados.Origen = 2
        frmRemitosEntregados.idPedidoEntrega = Me.lstOE.selectedItem
        frmRemitosEntregados.caption = "Nro." & Me.lstOE.selectedItem
        frmRemitosEntregados.Show
    End If
End Sub

Private Sub vereditar_Click()
    If vereditarOE = 3 Then    'ver (porque la oe esta finalziada)
        frmPlaneamientoOEVer.IDOE = CLng(Me.lstOE.selectedItem)
        frmPlaneamientoOEVer.Show
    ElseIf vereditarOE = 1 Then    'editar (porq la oe esta en proceso)
        frmPlaneamientoOEEditar.IDOE = CLng(Me.lstOE.selectedItem)
        frmPlaneamientoOEEditar.Show
    End If
End Sub
