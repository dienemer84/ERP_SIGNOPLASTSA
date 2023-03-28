VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPlaneamientoResumenProduccion 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Historico fabricación..."
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7995
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   7995
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
      Height          =   5175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7935
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Top             =   480
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ver"
         Default         =   -1  'True
         Height          =   375
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00E0E0E0&
         Cancel          =   -1  'True
         Caption         =   "Salir"
         Height          =   375
         Left            =   6840
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   4200
         Width           =   975
      End
      Begin MSComctlLib.ListView lst 
         Height          =   2895
         Left            =   120
         TabIndex        =   2
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
            Text            =   "O/T"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Cant"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Detalle O/T"
            Object.Width           =   4939
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
            Object.Width           =   2646
         EndProperty
      End
      Begin VB.Label lblCotizado 
         BackColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   2040
         TabIndex        =   13
         Top             =   4440
         Width           =   1575
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pendientes"
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
         TabIndex        =   12
         Top             =   4440
         Width           =   1695
      End
      Begin VB.Label lblPromedio 
         BackColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   2040
         TabIndex        =   11
         Top             =   4680
         Width           =   1695
      End
      Begin VB.Label lblTotal 
         BackColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   2040
         TabIndex        =   10
         Top             =   4200
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Promedio de venta"
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
         TabIndex        =   9
         Top             =   4680
         Width           =   1695
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Total de piezas"
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
         TabIndex        =   8
         Top             =   4200
         Width           =   1695
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
         TabIndex        =   7
         Top             =   480
         Width           =   1335
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
         TabIndex        =   6
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lblCliente 
         BackColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   1440
         TabIndex        =   5
         Top             =   840
         Width           =   3975
      End
   End
End
Attribute VB_Name = "frmPlaneamientoResumenProduccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim idpiezaelegida As Long
Dim baseS As New classStock
Private Sub Command1_Click()
    On Error Resume Next
    Dim codigo As String
    Dim i As Long
    Dim rs As Recordset
    Dim x As ListItem
    codigo = Trim(Me.txtCodigo)
    Dim deta As DetalleOrdenTrabajo

    Set rs = conectar.RSFactory("select count(id) as cant from stock where id=" & idpiezaelegida)
    If rs!Cant = 1 Then

        Set rs = conectar.RSFactory("select  dp.id,dp.item,c.razon,p.fechaCreado,dp.cantidad,dp.idPedido,dp.precio,p.descripcion from clientes c, detalles_pedidos dp,sp.pedidos p where p.id=dp.idPedido and dp.idpieza=" & idpiezaelegida & " and p.idCliente=c.id")

        Me.lst.ListItems.Clear
        cantp = 0
        valo = 0
        Dim falta As Double
        Dim totfalta As Double
        totfalta = 0
        While Not rs.EOF


            Set deta = DAODetalleOrdenTrabajo.FindById(rs!Id, True)


            Me.lblCliente = rs!razon
            Set x = Me.lst.ListItems.Add(, , Format(rs!idpedido, "0000") & " | " & rs!item)
            falta = deta.CantidadPedida - deta.Cantidad_Entregada
            totfalta = totfalta + falta
            x.SubItems(1) = rs!Cantidad & " (" & falta & ")"
            cantp = cantp + rs!Cantidad
            x.SubItems(2) = rs!descripcion



            x.SubItems(3) = funciones.RedondearDecimales(rs!Precio)
            valo = valo + (rs!Precio * rs!Cantidad)

            x.SubItems(4) = rs!fechaCreado
            rs.MoveNext
        Wend


        '        Dim ia As Long
        '
        '        ia = idpiezaelegida
        '
        '
        '        Dim colpieza As New Collection
        '       Dim pie As Pieza
        '
        '
        '       Set pie = DAOPieza.FindById(idpiezaelegida, FL_0)
        '       colpieza.Add pie
        '
        '
        '
        '        Dim cola As New Collection
        '        Set cola = DAODetalleOrdenTrabajo.FindAllByPieza(colpieza)
        '
        '        For Each deta In cola
        '        Set deta = DAODetalleOrdenTrabajo.FindById(deta.Id, True)
        '        Set deta.OrdenTrabajo = DAOOrdenTrabajo.FindById(deta.OrdenTrabajo.Id)
        '         Me.lblCliente = deta.OrdenTrabajo.ClienteFacturar.razon
        '            Set x = Me.lst.ListItems.Add(, , Format(deta.OrdenTrabajo.Id, "0000") & " | " & deta.item)
        '            x.SubItems(1) = deta.CantidadPedida & " (" & (deta.CantidadPedida - deta.CantidadEntregada) & ")"
        '
        '            cantp = cantp + deta.CantidadPedida
        '            x.SubItems(2) = deta.OrdenTrabajo.Descripcion
        '            x.SubItems(3) = deta.Precio
        '            valo = valo + (deta.Precio * deta.CantidadPedida)
        '
        '            x.SubItems(4) = rs!FechaCreado
        '
        '
        '        Next deta




        Me.lblTotal = cantp
        Me.lblPromedio = Math.Round(valo / cantp, 2)
        Me.lblCotizado = totfalta
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
    FormHelper.Customize Me
    txtCodigo_Change
End Sub

Private Sub Label6_Click()

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

