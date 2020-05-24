VERSION 5.00
Begin VB.Form frmModificarCantidadOE 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Valor"
   ClientHeight    =   1350
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   2880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ Modificar valor ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2895
      Begin VB.TextBox txtValor 
         Height          =   285
         Left            =   600
         TabIndex        =   3
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton frmModificaCantidadOE 
         Caption         =   "Aceptar"
         Default         =   -1  'True
         Height          =   255
         Left            =   840
         TabIndex        =   2
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Valor"
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
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Command2"
      Height          =   255
      Left            =   840
      TabIndex        =   0
      Top             =   1920
      Width           =   975
   End
End
Attribute VB_Name = "frmModificarCantidadOE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub frmModificaCantidadOE_Click()

End Sub
