VERSION 5.00
Begin VB.Form frmAgregarTareaTiempoProceso 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agregar Tarea"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4875
   Icon            =   "frmAgregarTareaTiempoProceso.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4875
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboSectores 
      Height          =   315
      Left            =   1275
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   3450
   End
   Begin VB.TextBox txtObservaciones 
      Height          =   1140
      Left            =   1275
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1125
      Width           =   3450
   End
   Begin VB.ComboBox cboTareas 
      Height          =   315
      Left            =   1275
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   585
      Width           =   3450
   End
   Begin VB.Label lblSector 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sector"
      Height          =   195
      Left            =   645
      TabIndex        =   5
      Top             =   150
      Width           =   465
   End
   Begin VB.Label lblObservaciones 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Observaciones"
      Height          =   195
      Left            =   90
      TabIndex        =   4
      Top             =   1125
      Width           =   1065
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tarea"
      Height          =   195
      Left            =   690
      TabIndex        =   3
      Top             =   615
      Width           =   420
   End
End
Attribute VB_Name = "frmAgregarTareaTiempoProceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Customize Me
End Sub
