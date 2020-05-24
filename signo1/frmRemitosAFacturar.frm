VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmAdminRemitosAFacturar 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Remitos a facturar..."
   ClientHeight    =   3510
   ClientLeft      =   2610
   ClientTop       =   1350
   ClientWidth     =   11430
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   11430
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Command1"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   5160
      Width           =   975
   End
   Begin MSComctlLib.ListView lstRemitosPendientes 
      Height          =   3495
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   6165
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
         Text            =   "Rto"
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Detalle"
         Object.Width           =   6350
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Fecha"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Estado"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Cliente"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Facturado"
         Object.Width           =   2117
      EndProperty
   End
   Begin VB.Menu mnuRemitos 
      Caption         =   "menu"
      Visible         =   0   'False
      Begin VB.Menu numero 
         Caption         =   "nro"
         Enabled         =   0   'False
      End
      Begin VB.Menu facturable 
         Caption         =   "No facturable"
      End
      Begin VB.Menu valorizar 
         Caption         =   "Valorizar Remito"
      End
      Begin VB.Menu f 
         Caption         =   "-"
      End
      Begin VB.Menu verRemito 
         Caption         =   "Ver remito..."
      End
      Begin VB.Menu facturacion 
         Caption         =   "Detalle Facturas..."
      End
   End
End
Attribute VB_Name = "frmAdminRemitosAFacturar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim claseP As New classPlaneamiento
Dim claseA As New classAdministracion
Dim rs As recordset
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub facturable_Click()
    Dim r As recordset
    IdRemito = CLng(Me.lstRemitosPendientes.SelectedItem)
    Set r = conectar.RSFactory("select estado, estadoFacturado from remitos where id=" & IdRemito)
    If Not r.EOF And Not r.BOF Then
        estado = r!estado
        EstadoFacturado = r!EstadoFacturado
        If estado <> 3 Then    'si no esta anulado
            If EstadoFacturado = 0 Then    ' no facturado
                If MsgBox("¿Está seguro de marcar este remito como No Facturable?", vbYesNo, "Confirmar") = vbYes Then
                    claseA.ejecutarComando "update remitos set estadoFacturado=3 where id=" & IdRemito
                End If
            ElseIf EstadoFacturado = 3 Then
                If MsgBox("¿Está seguro de marcar este remito como Facturable?", vbYesNo, "Confirmar") = vbYes Then
                    claseA.ejecutarComando "update remitos set estadoFacturado=0 where id=" & IdRemito
                End If
            End If
        End If
    End If
End Sub
Private Sub facturacion_Click()
    If Me.lstRemitosPendientes.ListItems.count > 0 Then
        '  frmAdminFacturacion.Frame1.Caption = "[ Rto Nro. " & Me.lstRemitosPendientes.SelectedItem & " ]"
        '  frmFacturacion.remito = CLng(Me.lstRemitosPendientes.SelectedItem)
        '  frmFacturacion.Show
    End If
End Sub
Private Sub llenarLST()
    Set rs = conectar.RSFactory("select r.id,r.detalle,r.fecha,r.estado,r.estadoFacturado,c.razon from remitos r inner join clientes c on r.idCliente=c.id ORDER BY r.id DESC")
    Me.lstRemitosPendientes.ListItems.Clear
    While Not rs.EOF
        Set x = Me.lstRemitosPendientes.ListItems.Add(, , Format(rs!id, "0000"))
        x.SubItems(1) = rs!detalle
        x.SubItems(2) = Format(rs!FEcha, "dd/mm/yyyy")
        x.SubItems(3) = funciones.estado_rto(rs!estado)
        If rs!estado = 1 Then
            x.ListSubItems(3).ForeColor = vbMagenta
        ElseIf rs!estado = 2 Then
            x.ListSubItems(3).ForeColor = vbGreen
        End If
        x.SubItems(4) = rs!Razon
        If rs!estado = 3 Then    'anulado
            x.SubItems(5) = "Anulado"
        Else
            x.SubItems(5) = funciones.estado_remitos_facturas(rs!EstadoFacturado)
        End If
        If rs!estado = 3 Then
            x.ListSubItems(5).ForeColor = vbBlack
        Else
            If rs!EstadoFacturado = 0 Then
                x.ListSubItems(5).ForeColor = vbRed
            ElseIf rs!EstadoFacturado = 1 Then
                x.ListSubItems(5).ForeColor = vbCyan
            ElseIf rs!EstadoFacturado = 2 Then
                x.ListSubItems(5).ForeColor = vbBlue
            End If
        End If
        If x = marca Then
            x.Selected = True
            x.EnsureVisible
        End If
        rs.MoveNext
    Wend
End Sub
Private Sub Form_Activate()
    llenarLST
End Sub

Private Sub Form_Load()
FormHelper.Customize Me
End Sub

Private Sub lstRemitosPendientes_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    funciones.LstOrdenar Me.lstRemitosPendientes, CInt(ColumnHeader.index)
End Sub
Private Sub lstRemitosPendientes_DblClick()

    verRemito_Click
End Sub
Private Sub lstRemitosPendientes_ItemClick(ByVal Item As MSComctlLib.ListItem)
    marca = Item
End Sub
Private Sub lstRemitosPendientes_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim r As recordset
    If Me.lstRemitosPendientes.ListItems.count > 0 Then
        If Button = 2 Then
            Me.numero.caption = "[ Nro. " & Me.lstRemitosPendientes.SelectedItem & " ]"
            Set r = conectar.RSFactory("select estadoFacturado,estado from remitos where id=" & CLng(Me.lstRemitosPendientes.SelectedItem))
            If Not r.EOF And Not r.BOF Then
                esta = r!EstadoFacturado
                estado = r!estado
            End If
            Set r = Nothing
            If estado = 3 Then    'anulado
                Me.facturable.caption = "No Facturable"
                Me.facturable.Enabled = False
            Else
                'si no esta anulado, se analiza
                If esta = 3 Then
                    Me.facturable.caption = "Facturable"
                    Me.facturable.Enabled = True
                    Me.valorizar.Enabled = False

                ElseIf esta = 1 Then    'PARCIAL
                    Me.facturable.Enabled = False
                    Me.valorizar.Enabled = True
                ElseIf esta = 0 Then    'no facturado
                    Me.facturable.Enabled = True
                    Me.facturable.caption = "No Facturable"
                    Me.valorizar.Enabled = True
                ElseIf esta = 2 Then    'completo
                    Me.facturable.Enabled = False
                    Me.valorizar.Enabled = False
                End If
            End If
            Me.PopupMenu Me.mnuRemitos
        End If
    End If
End Sub

Private Sub valorizar_Click()
    If Me.lstRemitosPendientes.ListItems.count > 0 Then

        frmPlaneamientoRemitosDetalle.valorizable = True
        frmPlaneamientoRemitosDetalle.Editable = False
        frmPlaneamientoRemitosDetalle.usable = False
        frmPlaneamientoRemitosDetalle.cmdUsarItemFactura = False
        frmPlaneamientoRemitosDetalle.cmdUsarItem = False
        frmPlaneamientoRemitosDetalle.rtoNro = CLng(Me.lstRemitosPendientes.SelectedItem)
        frmPlaneamientoRemitosDetalle.Show 1
    End If
End Sub

Private Sub verRemito_Click()
    If Me.lstRemitosPendientes.ListItems.count > 0 Then
        Dim idRto As Long
        idRto = CLng(Me.lstRemitosPendientes.SelectedItem)
        frmPlaneamientoRemitosDetalle.valorizable = False
        frmPlaneamientoRemitosDetalle.rtoNro = idRto
        frmPlaneamientoRemitosDetalle.Command2.Visible = False
        frmPlaneamientoRemitosDetalle.usable = True
        frmPlaneamientoRemitosDetalle.Editable = False


        frmPlaneamientoRemitosDetalle.caption = "Remito Nro. " & idRto
        frmPlaneamientoRemitosDetalle.Show 1

    End If


End Sub
