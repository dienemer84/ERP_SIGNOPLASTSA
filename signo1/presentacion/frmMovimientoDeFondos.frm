VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmMovimientoDeFondos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Movimiento de fontos"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   9870
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   2310
      Left            =   105
      TabIndex        =   0
      Top             =   3585
      Width           =   9615
      _Version        =   786432
      _ExtentX        =   16960
      _ExtentY        =   4075
      _StockProps     =   79
      Caption         =   "Destino"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.GroupBox GroupBox2 
      Height          =   3045
      Left            =   75
      TabIndex        =   1
      Top             =   135
      Width           =   9615
      _Version        =   786432
      _ExtentX        =   16960
      _ExtentY        =   5371
      _StockProps     =   79
      Caption         =   "Origen"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.TabControl TabControl1 
         Height          =   2520
         Left            =   150
         TabIndex        =   2
         Top             =   300
         Width           =   9300
         _Version        =   786432
         _ExtentX        =   16404
         _ExtentY        =   4445
         _StockProps     =   68
         ItemCount       =   3
         Item(0).Caption =   "Item"
         Item(0).ControlCount=   0
         Item(1).Caption =   "Item"
         Item(1).ControlCount=   0
         Item(2).Caption =   "Item"
         Item(2).ControlCount=   0
      End
   End
End
Attribute VB_Name = "frmMovimientoDeFondos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Customize Me
End Sub
