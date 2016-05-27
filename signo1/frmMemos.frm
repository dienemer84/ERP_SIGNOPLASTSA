VERSION 5.00
Begin VB.Form frmMemos 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Memos..."
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   8805
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   8805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ Memo ]"
      Height          =   5175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8775
      Begin VB.CommandButton Command2 
         Cancel          =   -1  'True
         Caption         =   "Salir"
         Height          =   375
         Left            =   7560
         TabIndex        =   4
         Top             =   4680
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Actualizar"
         Default         =   -1  'True
         Height          =   375
         Left            =   6360
         TabIndex        =   2
         Top             =   4680
         Width           =   1095
      End
      Begin VB.TextBox txtMemo 
         Height          =   4335
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   240
         Width           =   8535
      End
   End
   Begin VB.Label idpresu 
      Caption         =   "Label1"
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   5760
      Width           =   735
   End
End
Attribute VB_Name = "frmMemos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim V As New classVentas
V.addMemo Me.txtMemo, CLng(Me.idpresu)
Set V = Nothing
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

