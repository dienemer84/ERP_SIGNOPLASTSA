VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmPlaneamientoOEEditar 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Editar OE"
   ClientHeight    =   8190
   ClientLeft      =   3720
   ClientTop       =   1995
   ClientWidth     =   9780
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   9780
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ Orígen stock ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   9735
      Begin VB.ComboBox cboClientes 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   360
         Width           =   7455
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ver Stock"
         Default         =   -1  'True
         Height          =   375
         Left            =   8520
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   360
         Width           =   1095
      End
      Begin MSComctlLib.ListView lstStockPositivo 
         Height          =   2175
         Left            =   120
         TabIndex        =   17
         Top             =   840
         Width           =   9495
         _ExtentX        =   16748
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Detalle"
            Object.Width           =   12577
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Cantidad"
            Object.Width           =   2434
         EndProperty
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label idCliente 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Label6"
         Height          =   255
         Left            =   8760
         TabIndex        =   20
         Top             =   600
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ Entrega ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   3240
      Width           =   9735
      Begin VB.CommandButton Command4 
         BackColor       =   &H00E0E0E0&
         Cancel          =   -1  'True
         Caption         =   "Salir"
         Height          =   375
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Guadar"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   960
         Width           =   975
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "[ Detalle ]"
         Height          =   3375
         Left            =   120
         TabIndex        =   3
         Top             =   1440
         Width           =   9495
         Begin VB.CommandButton q 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Quitar"
            Height          =   255
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   2640
            Width           =   735
         End
         Begin VB.ComboBox cboMonedas 
            Height          =   315
            Left            =   6000
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   3000
            Width           =   1215
         End
         Begin VB.ComboBox cboClientesDestino 
            Height          =   315
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   3000
            Width           =   4215
         End
         Begin VB.TextBox txrRefe 
            Height          =   285
            Left            =   1080
            TabIndex        =   4
            Top             =   360
            Width           =   8295
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   255
            Left            =   8160
            TabIndex        =   5
            Top             =   3000
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            _Version        =   393216
            Format          =   58982401
            CurrentDate     =   38923
         End
         Begin MSComctlLib.ListView lstOE 
            Height          =   1815
            Left            =   120
            TabIndex        =   7
            Top             =   720
            Width           =   9255
            _ExtentX        =   16325
            _ExtentY        =   3201
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
            NumItems        =   6
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Detalle"
               Object.Width           =   6703
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Cantidad"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Valor"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Cliente"
               Object.Width           =   3792
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "idCliente"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Stock"
               Object.Width           =   1693
            EndProperty
         End
         Begin VB.Label Label8 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Moneda"
            Height          =   255
            Left            =   5160
            TabIndex        =   24
            Top             =   3000
            Width           =   735
         End
         Begin VB.Label Label4 
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
            TabIndex        =   10
            Top             =   3000
            Width           =   735
         End
         Begin VB.Label Label6 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Entrega"
            Height          =   255
            Left            =   7440
            TabIndex        =   9
            Top             =   3000
            Width           =   615
         End
         Begin VB.Label Label7 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Referencia"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.TextBox txtCantidad 
         Height          =   285
         Left            =   4920
         TabIndex        =   2
         Text            =   "0"
         Top             =   600
         Width           =   3375
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Agregar"
         Height          =   375
         Left            =   8520
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label2 
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
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblDetalle 
         BackColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   960
         TabIndex        =   14
         Top             =   360
         Width           =   7575
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cantidad Disponible"
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
         TabIndex        =   13
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label lblCantDispo 
         BackColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   2160
         TabIndex        =   12
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cantidad Requerida"
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
         Left            =   3120
         TabIndex        =   11
         Top             =   600
         Width           =   1815
      End
   End
   Begin VB.Label idPieza 
      Height          =   255
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "frmPlaneamientoOEEditar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim clasea As New classAdministracion
Dim grabado As Boolean
Dim claseC As New classConfigurar
Dim vidOe As Long
Dim rss As Recordset
Dim rs As Recordset
Dim claseSP As New classSignoplast
Dim IdMoneda As Long
Dim claseS As New classStock

Dim claseP As New classPlaneamiento
Dim Cantidad As Long
Dim detalle As String
Dim idStock As Long
Dim vValor As Double
Dim c As Long
Public Property Let IDOE(nidoe As Long)
    vidOe = nidoe
End Property
Private Sub llenarLstClientes(rs As Recordset)
    Dim x As ListItem
    lstStockPositivo.ListItems.Clear
    While Not rs.EOF
        Set x = Me.lstStockPositivo.ListItems.Add(, , rs!detalle)
        x.SubItems(1) = rs!Cantidad
        x.SubItems(2) = rs!razon
        x.SubItems(3) = rs!id_cliente
        x.Tag = rs!idPieza
        rs.MoveNext
    Wend
End Sub
Private Sub llenarListaStock()
    Dim rs As Recordset
    Dim strsql As String
    cla = Me.cboClientes.ItemData(cboClientes.ListIndex)
    Me.idCliente = cla
    If Me.cboClientes.ItemData(cboClientes.ListIndex) = -1 Then
        strsql = "select id ,detalle,cantidad from stock where cantidad>0 order by detalle "
    Else
        strsql = "select id,detalle,cantidad from stock where cantidad>0 and id_cliente=" & cla & " order by detalle"
    End If
    Set rs = conectar.RSFactory(strsql)
    Me.lstStockPositivo.ListItems.Clear

    While Not rs.EOF
        Set x = Me.lstStockPositivo.ListItems.Add(, , rs!detalle)
        x.SubItems(1) = rs!Cantidad
        x.Tag = rs!Id

        rs.MoveNext
    Wend




    verMarcado
End Sub

Private Sub cboMonedas_Click()


    If IdMoneda = CInt(Me.cboMonedas.ItemData(Me.cboMonedas.ListIndex)) Then

    Else
        IdMoneda = CInt(Me.cboMonedas.ItemData(Me.cboMonedas.ListIndex))
        cambiarPrecios IdMoneda
    End If
End Sub

Private Sub Command1_Click()
'If MsgBox("¿Desea crear una nueva O/E?", vbYesNo, "Confirmación") = vbYes Then
'    Me.lstOE.ListItems.Clear
'    Me.lstStockPositivo.ListItems.Clear
    llenarListaStock
    'End If

End Sub

Private Sub Command2_Click()
    Dim cantpedida As Double
    Dim esta As Boolean
    If CLng(Me.txtCantidad) > 0 Then

        Dim valorr As Double
        idStock = CLng(Me.idPieza)
        Dim idMoneda_pieza As Long
        cantpedida = CLng(Me.txtCantidad)
        If cantpedida <= Cantidad Then
            'si la cantidad que piden es menor ue la cantidad en stock real opero
            'y agrego datos a la lista a procesar como nueva orden de entrega
            esta = False
            'tengo que fijarme que no exista la pieza en la OE, si existe tengo que sumarla
            For y = 1 To Me.lstOE.ListItems.count
                If Me.lstOE.ListItems(y).Tag = idStock Then
                    esta = True
                    aponer = funciones.FormatearDecimales(CDbl(Me.lstOE.ListItems(y).ListSubItems(1)) + CDbl(cantpedida), 2)
                    'controlo que haya stock
                    If aponer > Cantidad Then
                        MsgBox "No hay disponibilidad de stock!", vbCritical, "Error"
                    Else
                        'si hay stock disponible
                        Me.lstOE.ListItems(y).ListSubItems(1) = aponer
                    End If
                End If
            Next y


            valorr = claseP.precio_pieza2(idStock, idMoneda_pieza)    '0  'elegir valor más alto vendido de la pieza

            If IdMoneda <> idMoneda_pieza Then
                'si no es la misma moneda convierto a lo necesario
                'subitem2 de la lista
                valorr = clasea.realizaCambio(valorr, idMoneda_pieza, IdMoneda)
            End If



            Dim x As ListItem
            If Not esta Then
                Dim Pieza As Pieza
                'claseP.ejecutar_consulta "select s.id as idpieza,c.razon,c.id as idCliente from clientes c,stock s where c.id=s.id_cliente and s.id=" & CLng(Me.idPieza)
                Set Pieza = DAOPieza.FindById(CLng(Me.idPieza), FL_0)

                Set x = Me.lstOE.ListItems.Add(, , detalle)
                x.SubItems(1) = funciones.FormatearDecimales(cantpedida, 2)
                x.SubItems(2) = funciones.FormatearDecimales(valorr, 2)
                x.SubItems(3) = Pieza.cliente.razon
                x.SubItems(4) = Pieza.cliente.Id


                x.Tag = idStock

            End If




            verMarcado
        Else
            MsgBox "No hay stock suficiente de esta pieza para la cantidad solicitada.", vbCritical, "Error"
        End If
    End If
End Sub
Private Sub Command3_Click()
'On Error Resume Next
    Dim refe As String
    Dim nroOEGenerada As Long
    Dim clie As Long
    Dim IdMoneda As Integer
    clie = Me.cboClientesDestino.ItemData(cboClientesDestino.ListIndex)
    refe = normaliza(Me.txrRefe)
    If MsgBox("¿Desea guardar los cambios?", vbYesNo, "Confirmacion") = vbYes Then
        IdMoneda = CInt(Me.cboMonedas.ItemData(Me.cboMonedas.ListIndex))
        If claseP.editarOE(Me.lstOE, Me.DTPicker1, refe, clie, vidOe, IdMoneda) Then
            MsgBox "Cambios guardados correctamente", vbInformation, "Información"
            grabado = True
        Else
            MsgBox "Cambios no guardados", vbInformation, "Información"
            grabado = False
        End If

    End If
End Sub

Private Sub Command4_Click()
    If grabado Then
        Unload Me
    Else
        If MsgBox("¿Está seguro de Salir?", vbYesNo, "Confirmación") = vbYes Then
            Unload Me
        End If
    End If
End Sub


Private Sub Form_Load()
    FormHelper.Customize Me
    grabado = True

    DAOCliente.LlenarCombo Me.cboClientes
    DAOCliente.LlenarCombo Me.cboClientesDestino, True
    DAOMoneda.LlenarCombo Me.cboMonedas
    Me.DTPicker1 = Now
    llenarDatosOE
    'lleno el combo de cliente destino
    'lleno la lista de OE armada
    'lleno la fecha de entrega
    'lleno el campo descripción
End Sub
Public Sub llenarDatosOE()
    Dim x As ListItem
    On Error GoTo err551
    Set rs = conectar.RSFactory("select fecha,referencia,idmoneda,idCliente from PedidosEntregas where id=" & vidOe)
    While Not rs.EOF
        Me.DTPicker1 = rs!FEcha
        Me.txrRefe = rs!referencia
        Me.cboMonedas.ListIndex = rs!IdMoneda
        IdMoneda = CInt(Me.cboMonedas.ItemData(Me.cboMonedas.ListIndex))
        Me.cboClientesDestino.ListIndex = funciones.PosIndexCbo(rs!idCliente, Me.cboClientesDestino)
        rs.MoveNext
    Wend

    Set rs = conectar.RSFactory("Select s.cantidad as cantStock,s.id as idpieza,s.detalle,dp.cantidad,dp.vale,c.razon,c.id as idcliente from stock s,detallesPedidosEntregas dp, PedidosEntregas p, clientes c  where idPedidoEntrega=" & vidOe & " And s.id_cliente = c.id And p.id = dp.idPedidoEntrega And dp.idPieza = s.id")
    While Not rs.EOF
        Set x = Me.lstOE.ListItems.Add(, , rs!detalle)
        x.SubItems(1) = funciones.FormatearDecimales(rs!Cantidad, 2)
        x.SubItems(2) = funciones.FormatearDecimales(rs!vale, 2)
        x.SubItems(3) = rs!razon
        x.SubItems(4) = rs!idCliente
        x.SubItems(5) = rs!cantStock
        x.Tag = rs!idPieza


        If rs!cantStock < rs!Cantidad Then
            x.ForeColor = vbRed
            x.ListSubItems(1).ForeColor = vbRed
            x.ListSubItems(2).ForeColor = vbRed
            x.ListSubItems(3).ForeColor = vbRed
            x.ListSubItems(4).ForeColor = vbRed
            x.ListSubItems(5).ForeColor = vbRed


        End If



        rs.MoveNext
    Wend



    Exit Sub
err551:
    MsgBox Err.Description


End Sub
Private Sub verMarcado()
    If Me.lstStockPositivo.ListItems.count > 0 Then
        idStock = CLng(Me.lstStockPositivo.selectedItem.Tag)
        detalle = Me.lstStockPositivo.selectedItem
        Cantidad = CLng(Me.lstStockPositivo.selectedItem.ListSubItems(1))
        Me.lblCantDispo = Cantidad
        Me.lblDetalle = detalle
        Me.idPieza = idStock
    End If
End Sub
Private Sub Form_Terminate()
    Set rss = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set rss = Nothing
End Sub

Private Sub lstOE_DblClick()
    If Me.lstOE.ListItems.count > 0 Then
        funciones.valorOE = funciones.FormatearDecimales(Me.lstOE.selectedItem.ListSubItems(2), 2)
        funciones.cantOE = funciones.FormatearDecimales(Me.lstOE.selectedItem.ListSubItems(1), 2)
        frmPlaneamientoOEModificarCantidad.Show 1
        Me.lstOE.selectedItem.ListSubItems(2) = FormatearDecimales(funciones.valorOE, 2)
        Me.lstOE.selectedItem.ListSubItems(1) = FormatearDecimales(funciones.cantOE, 2)
        grabado = False
    End If
End Sub

Private Sub lstStockPositivo_ItemClick(ByVal item As MSComctlLib.ListItem)
    verMarcado
End Sub

Private Sub q_Click()

    If MsgBox("¿Está seguro de eliminar los items seleecionados?", vbYesNo, "Confirmacion") = vbYes Then
        For i = Me.lstOE.ListItems.count To 1 Step -1
            If Me.lstOE.ListItems(i).Checked = True Then
                Me.lstOE.ListItems.remove (i)
                grabado = False
            End If
        Next i
    End If

End Sub

Private Sub txtCantidad_GotFocus()
    foco Me.txtCantidad
End Sub
Private Sub txtCantidad_Validate(Cancel As Boolean)
    If Not IsNumeric(Me.txtCantidad) Then Cancel = True
End Sub


Private Sub cambiarPrecios(IdMoneda)
    Dim vale As Double
    Dim idMoneda_pieza As Long



    For x = 1 To Me.lstOE.ListItems.count
        idStock = Me.lstOE.ListItems(x).Tag
        'vale = claseP.precio_pieza(idStock, idMoneda_pieza) '0  'elegir valor más alto vendido de la pieza
        vale = CLng(Me.lstOE.ListItems(x).ListSubItems(2))
        'If idMoneda <> idMoneda_pieza Then
        vale = clasea.realizaCambio(vale, idMoneda_pieza, IdMoneda)
        'si no es la misma moneda convierto a lo necesario
        'subitem2 de la lista
        Me.lstOE.ListItems(x).ListSubItems(2) = funciones.FormatearDecimales(vale, 2)
        'End If
    Next



End Sub
