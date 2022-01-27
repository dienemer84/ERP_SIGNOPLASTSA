VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmSistemaAgendaGlobal 
   BackColor       =   &H00FF8080&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Agenda"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   13530
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   13530
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Filtrar"
      Default         =   -1  'True
      Height          =   255
      Left            =   11160
      TabIndex        =   3
      Top             =   4320
      Width           =   1095
   End
   Begin VB.TextBox txtFiltro 
      Height          =   285
      Left            =   960
      TabIndex        =   2
      Top             =   4320
      Width           =   10095
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   255
      Left            =   12360
      TabIndex        =   0
      Top             =   4320
      Width           =   1095
   End
   Begin MSComctlLib.ListView lstAgenda 
      Height          =   4215
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   7435
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Cód"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Razón Social"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Domicilio"
         Object.Width           =   5821
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Localidad"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "E-Mail"
         Object.Width           =   4939
      EndProperty
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF8080&
      Caption         =   "Búsqueda"
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
      Left            =   0
      TabIndex        =   4
      Top             =   4320
      Width           =   975
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
      Begin VB.Menu percepcionar 
         Caption         =   "Definir percepciones..."
      End
      Begin VB.Menu pedCur 
         Caption         =   "Pedidos en curso..."
      End
      Begin VB.Menu n4 
         Caption         =   "-"
      End
      Begin VB.Menu CambiarEstado 
         Caption         =   "Actmel"
      End
   End
End
Attribute VB_Name = "frmSistemaAgendaGlobal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim base As New classSignoplast
Dim c As New classPlaneamiento

Private Sub Command1_Click()
    Me.LlenarListaAgenda Trim(Me.txtFiltro)
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    FormHelper.Customize Me
    Me.LlenarListaAgenda Trim(Me.txtFiltro)
    
        Me.caption = caption & " (" & Name & ")"
        
        
End Sub

Private Sub lstAgenda_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    funciones.LstOrdenar Me.lstAgenda, CInt(ColumnHeader.Index)
End Sub

Private Sub lstAgenda_DblClick()
    If Me.lstAgenda.ListItems.count > 0 Then
        idcon = CLng(Me.lstAgenda.selectedItem)
        frmSistemaAgendaVerContacto.idContacto = idcon
        frmSistemaAgendaVerContacto.Show


    End If

End Sub


Private Sub txtFiltro_GotFocus()
    foco Me.txtFiltro
End Sub



Public Function LlenarListaAgenda(Optional filtro As String = Empty)
    Dim rs As New Recordset
    Me.lstAgenda.ListItems.Clear
    strsql = Empty

    strsql = "select * from agenda"





    If Len(Trim(filtro)) > 0 Then
        'strsql = strsql & " where empresa like '%" & filtro & "%'"
        strsql = "SELECT a.*,da.detalle FROM agenda a  inner join datos_agenda da on da.id_agenda=a.id where da.detalle LIKE '%" & filtro & "%' or  a.empresa LIKE '%" & filtro & "%' group by a.id"
    End If

    Set rs = conectar.RSFactory(strsql)

    While Not rs.EOF
        Set x = Me.lstAgenda.ListItems.Add(, , Format(rs!Id, "0000"))
        x.SubItems(1) = rs!empresa
        x.SubItems(2) = rs!direccion
        x.SubItems(3) = rs!localidad
        x.SubItems(4) = rs!email

        x.Tag = rs!Id
        rs.MoveNext
    Wend


End Function

