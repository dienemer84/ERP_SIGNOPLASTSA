VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmAviso 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1410
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   240
   End
   Begin XtremeSuiteControls.Label lblMsg 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      _Version        =   786432
      _ExtentX        =   7646
      _ExtentY        =   1931
      _StockProps     =   79
      Caption         =   "Label1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
   End
End
Attribute VB_Name = "frmAviso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub mostrar(msj As String)
    Me.lblMsg.caption = msj
    
    'centrar en pantalla
    Me.Left = (Screen.Width - Me.Width) \ 2
    Me.Top = (Screen.Height - Me.Height) \ 2
    
    Me.Show vbModeless
    Timer1.Enabled = True
    
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    Me.Hide
End Sub
