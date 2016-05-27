VERSION 5.00
Begin VB.Form frmSistemaEscanear 
   Caption         =   "Form1"
   ClientHeight    =   6705
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6705
   ScaleWidth      =   7440
   Begin VB.PictureBox Picture1 
      Height          =   2655
      Left            =   1440
      ScaleHeight     =   2595
      ScaleWidth      =   5355
      TabIndex        =   1
      Top             =   3720
      Width           =   5415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   960
      TabIndex        =   0
      Top             =   840
      Width           =   1695
   End
End
Attribute VB_Name = "frmSistemaEscanear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim newdoc As Long

Private Sub Form_Load()
    FormHelper.Customize Me
End Sub
