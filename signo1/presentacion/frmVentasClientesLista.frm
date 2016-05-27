VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmVentasClientesLista 
   Caption         =   "Clientes"
   ClientHeight    =   5220
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   13470
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5220
   ScaleWidth      =   13470
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   4
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
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Filtrar"
      Default         =   -1  'True
      Height          =   375
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4680
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFC0C0&
      Height          =   315
      ItemData        =   "frmVentasClientesLista.frx":0000
      Left            =   240
      List            =   "frmVentasClientesLista.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   4680
      Width           =   2055
   End
   Begin GridEX20.GridEX grilla 
      Height          =   4575
      Left            =   0
      TabIndex        =   3
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
      Column(1)       =   "frmVentasClientesLista.frx":0022
      Column(2)       =   "frmVentasClientesLista.frx":012E
      Column(3)       =   "frmVentasClientesLista.frx":022A
      Column(4)       =   "frmVentasClientesLista.frx":0326
      Column(5)       =   "frmVentasClientesLista.frx":0422
      Column(6)       =   "frmVentasClientesLista.frx":050A
      Column(7)       =   "frmVentasClientesLista.frx":0602
      Column(8)       =   "frmVentasClientesLista.frx":06EE
      Column(9)       =   "frmVentasClientesLista.frx":07DE
      Column(10)      =   "frmVentasClientesLista.frx":08C2
      Column(11)      =   "frmVentasClientesLista.frx":09B2
      Column(12)      =   "frmVentasClientesLista.frx":0A9E
      Column(13)      =   "frmVentasClientesLista.frx":0B82
      Column(14)      =   "frmVentasClientesLista.frx":0C76
      Column(15)      =   "frmVentasClientesLista.frx":0D5A
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmVentasClientesLista.frx":0E5A
      FormatStyle(2)  =   "frmVentasClientesLista.frx":0F92
      FormatStyle(3)  =   "frmVentasClientesLista.frx":1042
      FormatStyle(4)  =   "frmVentasClientesLista.frx":10F6
      FormatStyle(5)  =   "frmVentasClientesLista.frx":11CE
      FormatStyle(6)  =   "frmVentasClientesLista.frx":1286
      ImageCount      =   0
      PrinterProperties=   "frmVentasClientesLista.frx":1366
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
      Begin VB.Menu numero 
         Caption         =   "Numero"
      End
      Begin VB.Menu VerDetalle 
         Caption         =   "Editar"
      End
      Begin VB.Menu masContactos 
         Caption         =   "Contáctos"
      End
      Begin VB.Menu df 
         Caption         =   "-"
      End
      Begin VB.Menu CambiarEstado 
         Caption         =   "Cambiar Estado"
      End
   End
End
Attribute VB_Name = "frmVentasClientesLista"
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
    Me.Combo1.Top = Me.Height - 950
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
    On Error Resume Next
    Set rectemp = clientes.item(RowIndex)
    With rectemp
        Values(1) = Format(.id, "0000")
        Values(2) = UCase(.razon)
        Values(3) = .Domicilio
        Values(4) = .localidad.nombre
        Values(5) = .localidad.cp
        Values(6) = .provincia.nombre
        Values(7) = .provincia.pais.nombre

        Values(8) = .telefono
        Values(9) = .Fax
        Values(10) = .email
        Values(11) = .Cuit
        If .TipoIVA Is Nothing Then
            Values(12) = Empty
        Else
            Values(12) = .TipoIVA.detalle
        End If

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

Private Sub masContactos_Click()
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


