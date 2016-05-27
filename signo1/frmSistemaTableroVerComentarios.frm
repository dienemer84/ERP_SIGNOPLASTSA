VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSistemaTableroVerComentarios 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ver comentarios"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10290
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   10290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ Detalle de comentarios ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   10215
      Begin VB.TextBox txtComentario 
         Height          =   1935
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   3240
         Width           =   9975
      End
      Begin MSComctlLib.ListView lstComentarios 
         Height          =   2775
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   4895
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Fecha"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Autor"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Comentario"
            Object.Width           =   10583
         EndProperty
      End
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "salir"
      Height          =   375
      Left            =   9000
      TabIndex        =   0
      Top             =   6960
      Width           =   1095
   End
   Begin VB.Label lblTitulo 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   960
      TabIndex        =   8
      Top             =   600
      Width           =   9255
   End
   Begin VB.Label lblTipo 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   960
      TabIndex        =   7
      Top             =   360
      Width           =   9255
   End
   Begin VB.Label lblFecha 
      BackColor       =   &H00C0C0C0&
      Caption         =   " "
      Height          =   255
      Left            =   960
      TabIndex        =   6
      Top             =   120
      Width           =   9255
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Titulo"
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
      TabIndex        =   5
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tipo"
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
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Fecha"
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
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmSistemaTableroVerComentarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim r As Recordset
Dim claseSP As New classSignoplast
Dim vIdEVento As Long
Public Property Let idEvento(nIdEvento As Long)
    vIdEVento = nIdEvento
End Property
Private Sub Command1_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    FormHelper.Customize Me
    'muestro el encabezado
    Set r = conectar.RSFactory("select tipo,fechaCreado,titulo from usuariosTablero where id=" & vIdEVento)
    If Not r.EOF And Not r.BOF Then
        Me.lblFecha = r!FechaCreado
        Me.lblTitulo = r!titulo
        Me.lblTipo = funciones.tipoEvento(r!Tipo)
    End If
    Dim x As ListItem
    'traigo los comentarios
    Set r = conectar.RSFactory("select id,idUsuario,fecha,comentario from usuariosTableroComentarios where idTablero=" & vIdEVento)
    While Not r.EOF
        Set x = Me.lstComentarios.ListItems.Add(, , r!FEcha)
        x.SubItems(1) = claseSP.queUsuario(r!idUsuario)
        x.SubItems(2) = r!Comentario
        x.Tag = r!id
        r.MoveNext
    Wend
    lstComentarios_ItemClick Me.lstComentarios.selectedItem
End Sub

Private Sub lstComentarios_ItemClick(ByVal item As MSComctlLib.ListItem)

    If Me.lstComentarios.ListItems.count > 0 Then
        id = CLng(Me.lstComentarios.selectedItem.Tag)

        Set r = conectar.RSFactory("select comentario from usuariosTableroComentarios where id=" & id)
        If Not r.EOF And Not r.BOF Then
            Me.txtComentario = r!Comentario
        End If

    End If
End Sub
