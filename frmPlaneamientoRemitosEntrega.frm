VERSION 5.00
Begin VB.Form frmPlaneamientoRemitosEntrega 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Datos de la entrega"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7350
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   7350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Imprimir"
      Height          =   255
      Left            =   4920
      TabIndex        =   15
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Volver"
      Height          =   255
      Left            =   6120
      TabIndex        =   14
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label detalle 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   1320
      TabIndex        =   13
      Top             =   1560
      Width           =   6015
   End
   Begin VB.Label pais 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   1320
      TabIndex        =   12
      Top             =   1200
      Width           =   4575
   End
   Begin VB.Label provincia 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   1320
      TabIndex        =   11
      Top             =   1080
      Width           =   4575
   End
   Begin VB.Label localidad 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   1320
      TabIndex        =   10
      Top             =   840
      Width           =   4575
   End
   Begin VB.Label direccion 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   1320
      TabIndex        =   9
      Top             =   600
      Width           =   4575
   End
   Begin VB.Label telefono 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   1320
      TabIndex        =   8
      Top             =   360
      Width           =   4575
   End
   Begin VB.Label nombre 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   1320
      TabIndex        =   7
      Top             =   120
      Width           =   4575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Dirección "
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
      TabIndex        =   6
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Nombre "
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
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "País "
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
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Localidad "
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
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Provincia "
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
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Detalle "
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
      TabIndex        =   1
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Teléfono "
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
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "frmPlaneamientoRemitosEntrega"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim claseP As New classPlaneamiento
Dim vIdContacto As Long
Public Property Let idContacto(nIdContacto As Long)
    vIdContacto = nIdContacto
End Property

Private Sub Command1_Click()
    Unload Me
End Sub


Private Sub Form_Load()
    FormHelper.Customize Me
    Dim rs As Recordset
    Set rs = conectar.RSFactory("select * from clientesContactos where id=" & vIdContacto)
    If Not rs.EOF And Not rs.BOF Then
        Me.nombre = UCase(rs!nombre)
        Me.telefono = UCase(rs!tel)
        Me.direccion = UCase(rs!direccion)
        Me.localidad = UCase(rs!localidad)
        Me.detalle = UCase(rs!detalle)
        Me.pais = UCase(rs!país)
        Me.provincia = UCase(rs!provincia)

    End If
End Sub
