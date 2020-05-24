VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVerRubros 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ver rubros..."
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   4125
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   4125
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView lstRubros 
      Height          =   2175
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   3836
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label lblid 
      Caption         =   "Label1"
      Height          =   255
      Left            =   1080
      TabIndex        =   0
      Top             =   3000
      Width           =   1095
   End
End
Attribute VB_Name = "frmVerRubros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim claseP As New classCompras
Private Sub Form_Activate()
claseP.llenar_lista_rubros Me.lstRubros, -1, 3750, False, lblid
End Sub

Private Sub Form_Load()
'Set claseP = New classProveedor
lblid = 0
End Sub

