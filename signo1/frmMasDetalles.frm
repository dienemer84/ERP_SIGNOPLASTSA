VERSION 5.00
Begin VB.Form frmVentasPresupuestoMasDetalles 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Presupuesto..."
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   3315
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   3315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Cancel          =   -1  'True
      Caption         =   "Volver"
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label lblmas15 
      BackColor       =   &H00FFC0C0&
      Caption         =   "lblMas15"
      Height          =   255
      Left            =   1560
      TabIndex        =   1
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      Caption         =   "MarkUp MDO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      Caption         =   "Gastos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      Caption         =   "Menos 10Kg "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      Caption         =   "Menos 15Kg"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      Caption         =   "Mas 15Kg"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label lblGastos 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Label6"
      Height          =   255
      Left            =   1560
      TabIndex        =   5
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblMuMDO 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Label7"
      Height          =   255
      Left            =   1560
      TabIndex        =   4
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label lblMen10 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Label8"
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label lblmen15 
      BackColor       =   &H00FFC0C0&
      Caption         =   "lblMen15"
      Height          =   255
      Left            =   1560
      TabIndex        =   2
      Top             =   1200
      Width           =   1455
   End
End
Attribute VB_Name = "frmVentasPresupuestoMasDetalles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vpresu As clsPresupuesto

Public Property Let presu(T As clsPresupuesto)
    Set vpresu = T
End Property

Private Sub Command1_Click()
    Unload Me
End Sub


Private Sub Form_Load()
    FormHelper.Customize Me
    Me.lblGastos = vpresu.Gastos & "%"
    Me.lblmas15 = vpresu.PorcMas15 & "%"
    Me.lblMen10 = vpresu.PorcMen10 & "%"
    Me.lblmen15 = vpresu.PorcMen15 & "%"
    Me.lblMuMDO = vpresu.PorcMDO & "%"

End Sub
