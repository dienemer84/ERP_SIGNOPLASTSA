VERSION 5.00
Begin VB.Form frmSistemaTableroAgregarComentario 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Agregar Comentario..."
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7665
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   7665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ Nuevo comentario ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7575
      Begin VB.TextBox Text1 
         Height          =   1575
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   600
         Width           =   7335
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Cancel          =   -1  'True
         Caption         =   "Volver"
         Height          =   375
         Left            =   6240
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   2400
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Agregar"
         Default         =   -1  'True
         Height          =   375
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Comentario"
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
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmSistemaTableroAgregarComentario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim claseSP As New classSignoplast
Dim vIdTablero As Long
Public Property Let idTablero(nIdTablero As Long)
    vIdTablero = nIdTablero
End Property
Private Sub Command1_Click()
    If MsgBox("¿Desea abandonar el comentario?", vbYesNo, "Confirmación") = vbYes Then
        Unload Me
    End If
End Sub
Private Sub Command2_Click()
    If Trim(Me.Text1) <> Empty Then
        If MsgBox("¿Está seguro de agregar el comentario?", vbYesNo, "Confirmación") = vbYes Then
            fech = funciones.datetimeFormateada(Now)
            idUsuario = funciones.getUser
            Comentario = Trim(UCase(Me.Text1))
            If claseSP.ejecutarComando("insert into usuariosTableroComentarios (idTablero, idUsuario, fecha, comentario) values  (" & vIdTablero & "," & idUsuario & ",'" & fech & "','" & Comentario & "')") Then Me.Text1 = Empty

        End If
    End If
End Sub

Private Sub Form_Load()
    FormHelper.Customize Me
End Sub
