VERSION 5.00
Begin VB.Form frmFacturacion 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Detalles de facturación..."
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8700
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   8700
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3720
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8655
   End
End
Attribute VB_Name = "frmFacturacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rtoNro As Long
Public Property Let remito(vremito As Long)
rtoNro = vrto
End Property

Private Sub Command1_Click()
Unload Me
End Sub
