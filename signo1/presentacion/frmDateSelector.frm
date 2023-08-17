VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmDateSelector 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Seleccione período"
   ClientHeight    =   1515
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4245
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   4245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.PushButton btnSeleccionar 
      Height          =   390
      Left            =   2550
      TabIndex        =   4
      Top             =   1005
      Width           =   1575
      _Version        =   786432
      _ExtentX        =   2778
      _ExtentY        =   688
      _StockProps     =   79
      Caption         =   "Seleccionar"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.ComboBox cboMes 
      Height          =   315
      Left            =   780
      TabIndex        =   2
      Top             =   225
      Width           =   1845
      _Version        =   786432
      _ExtentX        =   3254
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Style           =   2
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboAnio 
      Height          =   315
      Left            =   780
      TabIndex        =   3
      Top             =   615
      Width           =   930
      _Version        =   786432
      _ExtentX        =   1640
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Style           =   2
      Text            =   "ComboBox2"
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Año:"
      Height          =   195
      Left            =   300
      TabIndex        =   1
      Top             =   630
      Width           =   330
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Mes:"
      Height          =   195
      Left            =   300
      TabIndex        =   0
      Top             =   255
      Width           =   345
   End
End
Attribute VB_Name = "frmDateSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public DateSelected As Boolean


Private Sub btnSeleccionar_Click()

    If Me.cboAnio.ListIndex = -1 Or Me.cboMes.ListIndex = -1 Then Exit Sub

    DateSelected = True
    Me.Hide
End Sub


Private Sub Form_Load()
    Customize Me

    Me.cboMes.Clear
    Dim i As Integer
    For i = 1 To 12
        Me.cboMes.AddItem MonthName(i)
    Next i

    For i = 2008 To Year(Date) + 1
        Me.cboAnio.AddItem i
    Next i


End Sub
