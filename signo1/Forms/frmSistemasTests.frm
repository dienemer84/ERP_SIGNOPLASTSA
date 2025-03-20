VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
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
   Begin XtremeSuiteControls.PushButton PushButton 
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   2280
      Width           =   4935
      _Version        =   786432
      _ExtentX        =   8705
      _ExtentY        =   1296
      _StockProps     =   79
      Caption         =   "Agenda nueva"
      Appearance      =   6
   End
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
      Index           =   0
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

Private Sub Form_Load()
    FormHelper.Customize Me
End Sub

Private Sub PushButton_Click()
    Dim f227 As New frmAgendaNueva
    f227.Show
End Sub
