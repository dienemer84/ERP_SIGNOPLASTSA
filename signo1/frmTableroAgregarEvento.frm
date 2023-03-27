VERSION 5.00
Begin VB.Form frmSistemaTableroAgregarEvento 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Agregar Evento"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7695
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ Complete los datos ]"
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7695
      Begin VB.TextBox Text1 
         Height          =   1485
         Left            =   1560
         TabIndex        =   8
         Top             =   840
         Width           =   6015
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Cancel          =   -1  'True
         Caption         =   "Salir"
         Height          =   375
         Left            =   6360
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2520
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Evento modificable"
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
         Left            =   1440
         TabIndex        =   6
         Top             =   2520
         Width           =   2055
      End
      Begin VB.CommandButton cmdAgregar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Agregar"
         Default         =   -1  'True
         Height          =   375
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2520
         Width           =   1335
      End
      Begin VB.ComboBox cboEventos 
         Height          =   315
         ItemData        =   "frmTableroAgregarEvento.frx":0000
         Left            =   1560
         List            =   "frmTableroAgregarEvento.frx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   360
         Width           =   6015
      End
      Begin VB.CommandButton cmdModificar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Modificar"
         Height          =   375
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción "
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
         TabIndex        =   2
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo de Evento "
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
         TabIndex        =   1
         Top             =   360
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmSistemaTableroAgregarEvento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim claseP As New classPlaneamiento
Dim vGrupoEvento As Long
Dim grabado As Boolean
Dim vIdEVento As Long
Dim vEventoTipo As Long
Dim vDetalleEvento As String
Dim vFechaEvento As Date

Public Property Let idEvento(nIdEvento As Long)
    vIdEVento = nIdEvento
End Property
Public Property Let GrupoEvento(nGrupoEvento As Long)
    vGrupoEvento = nGrupoEvento
End Property
Public Property Let EventoTipo(nEventoTipo As Integer)
    vEventoTipo = nEventoTipo
End Property
Public Property Let DetalleEvento(nDetalleEvento As String)
    vDetalleEvento = nDetalleEvento
End Property
Public Property Let FechaEvento(nFechaEvento As Date)
    vFechaEvento = nFechaEvento
End Property

Public Property Get EventoTipo() As Integer
    EventoTipo = vEventoTipo
End Property
Public Property Get DetalleEvento() As String
    DetalleEvento = vDetalleEvento
End Property
Public Property Get FechaEvento() As Date
    FechaEvento = vFechaEvento
End Property

Private Sub cboEventos_Change()
    grabado = False
End Sub

Private Sub Check1_Click()
    grabado = False
End Sub

Private Sub cmdAgregar_Click()
    Dim strsql As String
    If MsgBox("¿Está seguro de agregar este evento?", vbYesNo, "Confirmación") = vbYes Then
        idUsuario = funciones.getUser
        FechaDia = Day(FechaEvento)
        fechaMes = Month(FechaEvento)
        fechaAnio = Year(FechaEvento)
        Tipo = Me.cboEventos.ListIndex
        mem = Trim(UCase(Me.Text1))
        Modificable = Me.Check1.value
        fechaCreado = funciones.datetimeFormateada(Now)
        strsql = "insert into usuariosTablero (idUsuario, FechaDia, FechaMes, FechaAnio, tipo,titulo,EventoModificable,GrupoUsuarios,fechaCreado) values (" & idUsuario & "," & FechaDia & "," & fechaMes & "," & fechaAnio & "," & Tipo & ",'" & mem & "'," & Modificable & "," & vGrupoEvento & ",'" & fechaCreado & "' )"
        Unload Me
        If Not claseP.ejecutarComando(strsql) Then MsgBox "Se produjo algun error al agregar el dato!", vbCritical, "Error"


    End If
End Sub

Private Sub cmdModificar_Click()
    If MsgBox("¿Está seguro de modificar este evento?", vbYesNo, "Confirmación") = vbYes Then
        Unload Me
    End If

End Sub

Private Sub Command1_Click()
    If Not grabado Then
        If MsgBox("¿Desea perder los cambios?", vbYesNo, "Confirmación") = vbYes Then
            Unload Me
        End If
    Else
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    FormHelper.Customize Me
    Me.cboEventos.ListIndex = 0
    If vIdEVento = -1 Then
        grabado = True
        Me.cmdAgregar.Visible = True
        Me.cmdModificar.Visible = False
    Else
        grabado = False
        Me.cmdAgregar.Visible = False
        Me.cmdModificar.Visible = True
    End If
    Me.Frame1.caption = "[ " & funciones.FEcha(vFechaEvento) & " ]"
End Sub

Private Sub Text1_Change()
    grabado = False
End Sub
