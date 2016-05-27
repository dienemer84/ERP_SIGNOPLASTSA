VERSION 5.00
Begin VB.Form frmCapaCaras 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Capas y caras de pintura"
   ClientHeight    =   1530
   ClientLeft      =   60
   ClientTop       =   225
   ClientWidth     =   4035
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   4035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCapa 
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Top             =   480
      Width           =   2295
   End
   Begin VB.TextBox txtCaras 
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Default         =   -1  'True
      Height          =   375
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cantidad de capas"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cantidad de caras"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmCapaCaras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    frmConfigurarTerminacion.lstPiezas.selectedItem.ListSubItems(6) = Trim(Me.txtCapa)
    frmConfigurarTerminacion.lstPiezas.selectedItem.ListSubItems(5) = Trim(Me.txtCaras)
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    FormHelper.Customize Me
End Sub

Private Sub txtCaras_Change()
    foco Me.txtCapa
End Sub

Private Sub txtCaras_GotFocus()
    foco Me.txtCaras
End Sub
