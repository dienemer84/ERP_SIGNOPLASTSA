VERSION 5.00
Begin VB.Form frmCincado 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Terminación..."
   ClientHeight    =   1860
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   4590
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1860
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ Cincado ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Cancel          =   -1  'True
         Caption         =   "Salir"
         Height          =   375
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cuenta"
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
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Definir cuenta para KG de material cincado"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   3615
      End
   End
End
Attribute VB_Name = "frmCincado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub
