VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1110
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   2700
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1110
   ScaleWidth      =   2700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   915
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   2385
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         Caption         =   "Espere..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1500
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
    'lblVersion.caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    'lblProductName.caption = App.Title
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub
