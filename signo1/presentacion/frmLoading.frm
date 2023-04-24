VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmLoading 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cargando .."
   ClientHeight    =   615
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9270
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   615
   ScaleWidth      =   9270
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.ProgressBar ProgressBar 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9015
      _Version        =   786432
      _ExtentX        =   15901
      _ExtentY        =   661
      _StockProps     =   93
      Appearance      =   6
   End
End
Attribute VB_Name = "frmLoading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Me.Height = 1215
    Me.Width = 9345
    
    Me.Left = frmPrincipal.ScaleWidth / 3
    Me.Top = frmPrincipal.ScaleHeight / 3
    
  
End Sub


