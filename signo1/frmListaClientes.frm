VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmVentasClientesLista22 
   BackColor       =   &H00FF8080&
   Caption         =   " Lista de Clientes..."
   ClientHeight    =   8235
   ClientLeft      =   60
   ClientTop       =   270
   ClientWidth     =   15480
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8235
   ScaleWidth      =   15480
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFC0C0&
      Height          =   315
      ItemData        =   "frmListaClientes.frx":0000
      Left            =   240
      List            =   "frmListaClientes.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   4680
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Filtrar"
      Default         =   -1  'True
      Height          =   375
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4680
      Width           =   1095
   End
   Begin VB.TextBox txtFiltro 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3480
      TabIndex        =   2
      Top             =   4680
      Width           =   5655
   End
   Begin GridEX20.GridEX grilla 
      Height          =   4575
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   8070
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      MethodHoldFields=   -1  'True
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      BackColorHeader =   16761024
      DataMode        =   99
      HeaderFontBold  =   -1  'True
      HeaderFontWeight=   700
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   15
      Column(1)       =   "frmListaClientes.frx":0022
      Column(2)       =   "frmListaClientes.frx":012E
      Column(3)       =   "frmListaClientes.frx":020A
      Column(4)       =   "frmListaClientes.frx":02DE
      Column(5)       =   "frmListaClientes.frx":03B2
      Column(6)       =   "frmListaClientes.frx":047E
      Column(7)       =   "frmListaClientes.frx":0552
      Column(8)       =   "frmListaClientes.frx":061A
      Column(9)       =   "frmListaClientes.frx":06EA
      Column(10)      =   "frmListaClientes.frx":07B6
      Column(11)      =   "frmListaClientes.frx":087E
      Column(12)      =   "frmListaClientes.frx":0952
      Column(13)      =   "frmListaClientes.frx":0A1E
      Column(14)      =   "frmListaClientes.frx":0AEE
      Column(15)      =   "frmListaClientes.frx":0BB6
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmListaClientes.frx":0C8E
      FormatStyle(2)  =   "frmListaClientes.frx":0DC6
      FormatStyle(3)  =   "frmListaClientes.frx":0E76
      FormatStyle(4)  =   "frmListaClientes.frx":0F2A
      FormatStyle(5)  =   "frmListaClientes.frx":1002
      FormatStyle(6)  =   "frmListaClientes.frx":10BA
      ImageCount      =   0
      PrinterProperties=   "frmListaClientes.frx":119A
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF8080&
      Caption         =   "Razón"
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
      Left            =   2640
      TabIndex        =   5
      Top             =   4680
      Width           =   735
   End
   Begin VB.Menu m3 
      Caption         =   "m3"
      Visible         =   0   'False
      Begin VB.Menu numero 
         Caption         =   "numero"
         Enabled         =   0   'False
      End
      Begin VB.Menu verDetalle 
         Caption         =   "Editar..."
      End
      Begin VB.Menu masContacto 
         Caption         =   "Contáctos..."
      End
      Begin VB.Menu n4 
         Caption         =   "-"
      End
      Begin VB.Menu CambiarEstado 
         Caption         =   "Actmel"
      End
   End
End
Attribute VB_Name = "frmVentasClientesLista22"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Implements ISuscriber


Dim id_suscriber As String
Dim rows As Long
Dim rectemp As clsCliente
Dim est As EstadoCliente
Dim clientes As Collection
Private Sub Combo1_Click()
    Command1_Click
End Sub
Private Sub Command1_Click()
    llenar_Grilla
End Sub
Private Sub Command2_Click()
    Unload Me
End Sub


Private Sub Form_Activate()
    If rows = 0 Then Exit Sub
    grilla.RefreshRowIndex rows
    Me.grilla.Refresh
End Sub

Private Sub Form_Deactivate()
    Channel.RemoverSuscripcionTotal Me
End Sub

Private Sub Form_Load()
    FormHelper.Customize Me
    GridEXHelper.CustomizeGrid Me.grilla

    Combo1.ListIndex = 0
    llenar_Grilla
    rows = 1
    id_suscriber = funciones.CreateGUID
    Channel.AgregarSuscriptor Me, Clientes_


End Sub
Private Sub Form_Resize()
    On Error Resume Next
    Me.grilla.Width = Me.ScaleWidth
    Me.grilla.Height = Me.Height - (Me.Combo1.Height + (1000 - Me.Combo1.Height))
    Me.grilla.ColumnAutoResize = True
    Me.Combo1.Top = Me.Height - 900
    Me.txtFiltro.Top = Me.Combo1.Top
    Me.Command1.Top = Me.Combo1.Top
    Me.Command2.Top = Me.Combo1.Top

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Channel.RemoverSuscripcionTotal Me
End Sub


Private Sub grilla_ColumnHeaderClick(ByVal Column As GridEX20.JSColumn)
    ordenar_grilla Column, Me.grilla
End Sub
Private Sub grilla_DblClick()
    verDeta
End Sub

Private Sub grilla_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Set rectemp = clientes(grilla.RowIndex(Me.grilla.row))
        Me.numero.caption = "Nro." & Format(rectemp.id, "0000")
        If rectemp.estado = 0 Then
            Me.CambiarEstado.caption = "Activar..."
        ElseIf rectemp.estado = 1 Then
            Me.CambiarEstado.caption = "Desactivar..."
        End If
        frmVentasClientesLista.PopupMenu m3
    End If
End Sub
Private Sub grilla_SelectionChange()
    rows = grilla.RowIndex(grilla.row)
End Sub

Private Sub grilla_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Set rectemp = clientes.item(RowIndex)
    With rectemp
        Values(1) = Format(.id, "0000")
        Values(2) = UCase(.razon)
        Values(3) = .Domicilio
        Values(4) = .localidad.nombre
        Values(5) = .localidad.cp
        Values(6) = .telefono
        Values(7) = .Fax
        Values(8) = .email
        Values(9) = .Cuit
        If .TipoIVA Is Nothing Then
            Values(10) = Empty
        Else
            Values(10) = .TipoIVA.detalle
        End If

        Values(11) = .provincia.nombre
        Values(12) = .provincia.pais.nombre
        Values(13) = .estado
        Values(14) = .FP
        Values(15) = .exLocalidad
    End With
End Sub

Private Property Get ISuscriber_id() As String
    ISuscriber_id = id_suscriber
End Property

Private Function ISuscriber_Notificarse(EVENTO As clsEventoObserver) As Variant
    If EVENTO.EVENTO = agregar_ Then
        clientes.Add EVENTO.Elemento
        grilla.ItemCount = clientes.count
    End If
End Function

Private Sub masContacto_Click()
    If grilla.rowcount > 0 Then
        Set rectemp = clientes(grilla.RowIndex(grilla.row))
        frmVentasClientesNuevoContacto.Cliente = rectemp
        frmVentasClientesNuevoContacto.Show

    End If
End Sub

Private Sub txtFiltro_GotFocus()
    foco Me.txtFiltro
End Sub
Private Sub verDetalle_Click()
    verDeta
End Sub
Private Sub llenar_Grilla()
    est = Me.Combo1.ItemData(Me.Combo1.ListIndex)
    'Set clientes = DAOCliente.GetAll(Trim(Me.txtFiltro), est)

    Dim filter As String

    filter = "{cliente}.{estado} = " & est

    If LenB(Trim(Me.txtFiltro.text)) > 0 Then
        filter = filter & " AND {cliente}.{razon} LIKE '%{value}%'"
        filter = Replace$(filter, "{razon}", DAOCliente.CAMPO_RAZON_SOCIAL)
        filter = Replace$(filter, "{value}", Me.txtFiltro.text)
    End If

    filter = Replace$(filter, "{estado}", DAOCliente.CAMPO_ESTADO)
    filter = Replace$(filter, "{cliente}", DAOCliente.TABLA_CLIENTE)

    Set clientes = DAOCliente.FindAll(filter)

    grilla.ItemCount = clientes.count
    grilla.ReBind
End Sub

Private Sub verDeta()
    If grilla.rowcount Then
        Set rectemp = clientes(grilla.RowIndex(grilla.row))
        frmVentasClienteNuevo.Cliente = rectemp
        frmVentasClienteNuevo.Show
    End If
End Sub

