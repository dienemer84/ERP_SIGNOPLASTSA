VERSION 5.00
Begin VB.Form frmPlaneamientoOEModificarCantidad 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Valor"
   ClientHeight    =   1710
   ClientLeft      =   5310
   ClientTop       =   5520
   ClientWidth     =   2880
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   2880
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ Modificar valor ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   2895
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1080
         TabIndex        =   5
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txtValor 
         Height          =   285
         Left            =   1080
         TabIndex        =   0
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton frmModificaCantidadOE 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Aceptar"
         Default         =   -1  'True
         Height          =   375
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cantidad "
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
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Valor "
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
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Command2"
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   2280
      Width           =   975
   End
End
Attribute VB_Name = "frmPlaneamientoOEModificarCantidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    FormHelper.Customize Me
    Me.txtValor = funciones.valorOE
    Me.Text1 = funciones.cantOE
End Sub

Private Sub frmModificaCantidadOE_Click()
    If IsNumeric(Me.txtValor) Then
        If CLng(Me.txtValor) > 0 Then
            funciones.valorOE = funciones.FormatearDecimales(CDbl(Me.txtValor), 2)
            funciones.cantOE = funciones.FormatearDecimales(CDbl(Me.Text1), 2)

        Else
            funciones.valorOE = 0
            funciones.cantOE = 0

        End If
    Else
        funciones.valorOE = 0
        funciones.cantOE = 0

    End If

    Unload Me
End Sub


Private Sub Text1_GotFocus()
    foco Me.Text1
End Sub

Private Sub txtValor_GotFocus()
    foco Me.txtValor
End Sub
