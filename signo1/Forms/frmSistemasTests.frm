VERSION 5.00
Begin VB.Form frmSistemasTests 
   Caption         =   "Tests"
   ClientHeight    =   6675
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11025
   Icon            =   "frmSistemasTests.frx":0000
   MDIChild        =   -1  'True
   ScaleHeight     =   6675
   ScaleWidth      =   11025
   Begin VB.CommandButton Command 
      Caption         =   "frmAdminExtrasReporteIVACompras"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   4935
   End
   Begin VB.CommandButton btnPrueba_04_Click 
      Caption         =   "frmAdminPagosCrearOrdenPagoNew"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   4935
   End
End
Attribute VB_Name = "frmSistemasTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnPrueba_04_Click_Click()
    Dim f226 As New frmAdminPagosCrearOrdenPagoNew
    f226.Show
End Sub

Private Sub Command_Click()
    Dim f125 As New frmAdminExtrasReporteIVACompras
    f125.Show
End Sub
