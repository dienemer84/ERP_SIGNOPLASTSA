VERSION 5.00
Begin VB.Form frmSistemaIngresar 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ingrese valor"
   ClientHeight    =   1215
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   5820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1845
      TabIndex        =   2
      Top             =   195
      Width           =   3855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   1845
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   645
      Width           =   1140
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3045
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   645
      Width           =   1140
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ingrese el Nombre"
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
      Left            =   165
      TabIndex        =   3
      Top             =   195
      Width           =   1575
   End
End
Attribute VB_Name = "frmSistemaIngresar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vNombre
Public Property Let nombre(nombre)
    vNombre = nombre
End Property
Public Property Get nombre()
    nombre = vNombre
End Property

Private Sub Command1_Click()
    vNombre = Trim(Me.Text1)
    Unload Me
End Sub

Private Sub Command2_Click()

    If MsgBox("¿Está seguro de volver?", vbYesNo, "Confirmación") = vbYes Then
        vNombre = Empty
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    FormHelper.Customize Me
    Me.Text1 = vNombre
End Sub
