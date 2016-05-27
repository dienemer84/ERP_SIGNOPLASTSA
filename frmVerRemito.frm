VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmPlaneamientoRemitosDetalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ver Remito..."
   ClientHeight    =   7020
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   7950
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   7950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Marcar item como transporte / Sin Cargo"
      Height          =   255
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6720
      Width           =   3015
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Quitar"
      Height          =   195
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6720
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdUsarItemFactura 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Usar item fc"
      Height          =   255
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdUsarItem 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Usar item"
      Height          =   255
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1215
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Editar"
      Height          =   255
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Valorizar"
      Height          =   255
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Facturar"
      Default         =   -1  'True
      Height          =   255
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "Volver"
      Height          =   255
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtFecha 
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   120
      Width           =   2175
   End
   Begin VB.TextBox txtCliente 
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   480
      Width           =   6615
   End
   Begin VB.TextBox txtReferencia 
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   840
      Width           =   6615
   End
   Begin VB.CommandButton cmdEntrega 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ver entrega"
      Height          =   255
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin MSComctlLib.ListView lstRemito 
      Height          =   5055
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   8916
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Item"
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Cant"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Detalle"
         Object.Width           =   7320
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Origen"
         Object.Width           =   1623
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Valor"
         Object.Width           =   1411
      EndProperty
   End
   Begin VB.Label Label1 
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
      TabIndex        =   6
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Referencia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Fecha de creación"
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
      Top             =   120
      Width           =   1815
   End
   Begin VB.Menu menuRemito 
      Caption         =   "MENU REMITOS"
      Visible         =   0   'False
      Begin VB.Menu nroRemito 
         Caption         =   "nro"
         Enabled         =   0   'False
      End
      Begin VB.Menu marcar 
         Caption         =   "marcar"
      End
   End
End
Attribute VB_Name = "frmPlaneamientoRemitosDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim idContacto As Long
Dim vusable As Boolean
Dim vvalorizable As Boolean
Dim claseP As New classPlaneamiento
Dim vRemito As Long
Dim vEditable As Boolean
Dim vUsarItem As Boolean
Dim vUsarItemFactura As Boolean

Property Let usarItem(nUsarItem As Boolean)
    vUsarItem = nUsarItem
End Property

Property Let usarItemFactura(nUsarItemFactura As Boolean)
    vUsarItemFactura = nUsarItemFactura
End Property

Property Let Editable(nEditable As Boolean)
    vEditable = nEditable
End Property





Property Let Usable(nusable As Boolean)
    vusable = nusable
End Property
Property Let valorizable(nvalorizable As Boolean)
    vvalorizable = nvalorizable
End Property



Property Let rtoNro(remito As Long) 'id remito en realidad
    vRemito = remito
End Property
Private Sub verRemito()
    On Error GoTo err93
    Dim rs As Recordset
    Me.lstRemito.ListItems.Clear
    Set rs = conectar.RSFactory("select r.idContacto,r.id,r.detalle,r.fecha,r.estado,c.razon from remitos r inner join clientes c on r.idcliente=c.id where r.id=" & vRemito)
    While Not rs.EOF
        Me.txtCliente = rs!Razon
        Me.txtReferencia = rs!detalle
        Me.txtFecha = rs!FEcha
        idContacto = rs!idContacto
        rs.MoveNext
    Wend
    If idContacto > 0 Then
        Me.cmdEntrega.Enabled = True
    Else
        Me.cmdEntrega.Enabled = False
    End If

    Set rs = conectar.RSFactory("select dp.item,e.facturable,e.facturado,e.id,e.valor,e.idPedido,sum(e.cantidad) AS CANTIDAD,s.detalle,e.origen from entregas e,detalles_pedidos dp,stock s where e.idDetallePedido=dp.id and dp.idpieza=s.id and e.remito =" & vRemito & " and e.origen=1 group by e.id union all select '000' as item,e.facturable,e.facturado,e.id,e.valor,e.idPedido,sum(e.cantidad) as cantidad,s.detalle,e.origen from entregas e,detallesPedidosEntregas dp,stock s where e.idDetallePedido=dp.id and dp.idpieza=s.id and e.remito=" & vRemito & " and e.origen=2 group by e.id union all select '000' as item,e.facturable,e.facturado,e.id,e.valor,e.idPedido,e.cantidad,e.concepto as detalle,e.origen from entregas e  where  e.remito=" & vRemito & " and (e.origen=3 or e.origen=4) ")
    it = 0
    While Not rs.EOF
        it = it + 1
        Dim x As ListItem
        Set x = Me.lstRemito.ListItems.Add(, , Format(it, "000"))
        x.Tag = rs!Id
        x.SubItems(1) = funciones.FormatearDecimales(rs!cantidad, 2)
        x.SubItems(2) = Format(rs!Item, "000") & "  " & rs!detalle
        If rs!origen = 1 Then
            ori = "O/T"
        ElseIf rs!origen = 2 Then
            ori = "O/E"
        ElseIf rs!origen = 3 Then
            ori = "Concepto"

        ElseIf rs!origen = 4 Then    'concepto aplicado
            ori = "OTA"

        End If
        If rs!origen = 3 Then
            x.SubItems(3) = ori
        Else
            x.SubItems(3) = ori & " " & Format(rs!idpedido, "0000")
        End If
        If Permisos.sistemaVerPrecios Then
            x.SubItems(4) = funciones.FormatearDecimales(rs!valor, 2)
        Else
            x.SubItems(4) = funciones.FormatearDecimales(0, 2)
        End If
        'x.SubItems(5) = rs!id

        x.ListSubItems(4).Tag = rs!facturable

        If vusable = True Then    'si se puede usar para facturar, pinto de orjo loq esta facturado
            If rs!facturable = 0 Then
                x.ForeColor = vbBlue
                x.ListSubItems(1).ForeColor = vbBlue
                x.ListSubItems(2).ForeColor = vbBlue
                x.ListSubItems(3).ForeColor = vbBlue
                x.ListSubItems(4).ForeColor = vbBlue
                'x.ListSubItems(5).ForeColor = vbRed
                x.Tag = rs!Id
                x.ListSubItems(1).Tag = 0    'no eliminado

            End If

            If rs!Facturado = 1 Then
                x.ForeColor = vbRed
                x.ListSubItems(1).ForeColor = vbRed
                x.ListSubItems(2).ForeColor = vbRed
                x.ListSubItems(3).ForeColor = vbRed
                x.ListSubItems(4).ForeColor = vbRed
                'x.ListSubItems(5).ForeColor = vbRed
                x.ListSubItems(1).Tag = 0    'no eliminado
                x.Tag = rs!Id
            End If
        End If
        rs.MoveNext
    Wend
    Exit Sub
err93:
    MsgBox Err.Description
End Sub
Private Sub cmdEntrega_Click()
    frmPlaneamientoRemitosEntrega.idContacto = idContacto
    frmPlaneamientoRemitosEntrega.Show 1
End Sub

Private Sub cmdUsarItem_Click()
    Dim rs1 As Recordset
    ide = Me.lstRemito.selectedItem.Tag

    Set rs1 = conectar.RSFactory("select * from entregas where id=" & ide)
    If Not rs1.EOF And Not rs1.BOF Then
        If rs1!idpedido <= 0 And rs1!idDetallePedido = 0 And rs1!origen = 3 Then  'ojo q idPedido estab en -1 solo (30-11-09)
            'si es de concepto y no está asignado entonces ss
            funciones.itemRemito = rs1!Id
            Unload Me
        Else
            MsgBox "Debería elegir un item que sea de concepto y no esté asignado!", vbInformation, "Error"
            funciones.itemRemito = -1
        End If
    Else
        Unload Me
        funciones.itemRemito = -1    'devuelvo un-1 si hay error
    End If

    Set rs1 = Nothing

End Sub

Private Sub cmdUsarItemFactura_Click()
    ide = Me.lstRemito.selectedItem.Tag
    Set rs1 = conectar.RSFactory("select facturado,cantidad,idpedido from entregas where id=" & ide)
    If Not rs1.EOF And Not rs1.BOF Then
        If rs1!Facturado = 0 Then    'And rs1!idpedido = -1 Then
            'si no está facturado, se puede aplicar
            funciones.itemRemito = ide    'rs1!id

            Unload Me
        Else
            MsgBox "Debería elegir un item que No esté facturado!", vbInformation, "Error"
            funciones.itemRemito = -1
        End If
    Else
        Unload Me
        funciones.itemRemito = -1    'devuelvo un-1 si hay error
    End If

End Sub

Private Sub Command1_Click()
    col = Empty
    Set col = Nothing
    funciones.idEntrega = col
    Unload Me
End Sub
Private Sub Command2_Click()
    Dim col As New Collection
    Dim algunosFacturados As Boolean
    Dim algunosNoFacturables As Boolean
    Dim ide As Long
    If vusable Then
        For x = 1 To Me.lstRemito.ListItems.count
            If Me.lstRemito.ListItems(x).Selected = True Then
                ide = Me.lstRemito.ListItems(x).Tag

                If Not claseP.itemFacturado(ide) Then
                    If Not claseP.itemNoFacturable(ide) Then
                        col.Add ide
                    Else
                        algunosNoFacturables = True

                    End If
                Else
                    algunosFacturados = True
                End If
            End If
        Next x
    End If
    funciones.idEntrega = col
    msg1 = Empty
    MSG2 = Empty
    If algunosNoFacturables Then MSG2 = "Algunos items no son facturables"
    If algunosFacturados Then msg1 = "Algunos items ya estaban facturados "

    If algunosNoFacturables Or algunosFacturados Then
        MsgBox "Se produjeron los siguientes errores," & Chr(10) & msg1 & Chr(10) & MSG2 & Chr(10) & msgextra & "No se agregaran a la factura", vbCritical, "Error"

    End If
    Unload Me
End Sub

Private Sub Command3_Click()

    Dim ide As Long
    If vvalorizable Then
        If MsgBox("¿Desea utilizar estos valores para el remito actual?", vbYesNo, "Confirmación") = vbYes Then
            For x = 1 To Me.lstRemito.ListItems.count
                'If Me.lstRemito.ListItems(X).Selected = True Then
                If Me.lstRemito.ListItems(x).ForeColor = vbRed Then
                    MsgBox "Este item ya está facturado, no se puede revalorizar", vbInformation, "Información"
                Else
                    ide = Me.lstRemito.ListItems(x).Tag
                    vale = funciones.FormatearDecimales(CDbl(Me.lstRemito.ListItems(x).ListSubItems(4)), 2)
                    claseP.ejecutarComando "update entregas set ModifValor=1, valor=" & vale & " where id=" & ide
                End If


                'End If
            Next x
        End If
        Unload Me
    End If

End Sub



Private Sub Command4_Click()


'los que están marcados para eliminar hay quye analizarlos
'si son de concepto no hay problema
'si son desde algun otro origen hay que
'volverlos a su lugar en el origen (para poder volver a entregarlos)

    For x = 1 To Me.lstRemito.ListItems.count
        Id = Me.lstRemito.ListItems(x).Tag
        borrado = Me.lstRemito.ListItems(x).ListSubItems(1).Tag

        If borrado = 1 Then

        End If


    Next x
End Sub

Private Sub Command5_Click()
    For i = Me.lstRemito.ListItems.count To 1 Step -1
        'If Me.lstRemito.ListItems(I).Checked = True Then
        ' Me.lstRemito.ListItems.Remove (I)
        'End If

        'marco el registro como eliminado
        a = Me.lstRemito.ListItems(x).ListSubItems(1).Tag
        'no eliminado

        If a = 0 Then a = 1 Else a = 0
        Me.lstRemito.ListItems(x).ListSubItems(1).Tag = a
    Next i
End Sub

Private Sub Command6_Click()
'    If MsgBox("¿Está seguro de cambiar el estado de facturación del item?", vbYesNo, "Confirmación") = vbYes Then
'        id_Entrega = Me.lstRemito.SelectedItem.Tag
'        id_Entrega2 = Me.lstRemito.SelectedItem.ListSubItems(4).Tag
'        If claseP.cambiarEstadoItemFacturable(id_Entrega, nuevo_estado) Then
'            MsgBox "Cambio exitoso!", vbInformation, "Información"
'            Me.lstRemito.SelectedItem.ListSubItems(4).Tag = nuevo_estado
'        End If
'    End If

End Sub

Private Sub Form_Activate()
'verRemito
End Sub

Private Sub Form_Load()
FormHelper.Customize Me
Me.caption = vRemito
    Me.lstRemito.CheckBoxes = False
    Me.Command3.Visible = False
    Me.Command2.Visible = False
    Me.Command4.Visible = False
    Me.cmdEntrega.Visible = False
    cmdUsarItemFactura.Visible = False
    Me.Command5.Visible = False
    Me.cmdUsarItem.Visible = False


    Set col = Nothing

    If vUsarItemFactura Then
        Me.cmdUsarItemFactura.Visible = True
        Me.lstRemito.CheckBoxes = False
    End If


    If vUsarItem Then
        Me.cmdUsarItem.Visible = True
        Me.cmdEntrega.Default = True
        Me.lstRemito.CheckBoxes = False

    End If

    If vusable Then
        Me.Command2.Visible = True
        Me.Command2.Default = True

    End If
    If vvalorizable Then

        Me.Command3.Visible = True
        Me.Command3.Default = True
        Me.lstRemito.CheckBoxes = False
    End If

    If vEditable Then
        Me.Command4.Visible = True
        Me.Command4.Default = True
        Me.Command5.Visible = True
        Me.lstRemito.CheckBoxes = True
    End If

    'vEditable = False
    'vusable = False
    'vvalorizable = False
    'vUsarItem = False
    'vUsarItemFactura = False

    verRemito
End Sub



Private Sub lstRemito_DblClick()
    If vvalorizable Then    'si es para valorizar, muestro la opcion. Sino no.
        frmAdminRemitosValorizarNuevo.idEntrega = CLng(Me.lstRemito.selectedItem.Tag)
        frmAdminRemitosValorizarNuevo.remito = vRemito
        frmAdminRemitosValorizarNuevo.Show 1
    End If
End Sub

Private Sub lstRemito_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        If Me.lstRemito.ListItems.count > 0 Then
            estadoActual = Me.lstRemito.selectedItem.ListSubItems(4).Tag
            Me.nroRemito.caption = "[ Remito " & Format(vRemito, "0000") & " ]"
            If estadoActual = 0 Then mensaje = "Hacer facturable..."
            If estadoActual = 1 Then mensaje = "Hacer no facturable"
            Me.marcar.caption = mensaje
            Me.marcar.Tag = estadoActual
            If Permisos.PlanRemitosControl Then
                Me.marcar.Enabled = True
            Else
                Me.marcar.Enabled = False
            End If


            Me.PopupMenu Me.menuRemito
        End If
    End If
End Sub

