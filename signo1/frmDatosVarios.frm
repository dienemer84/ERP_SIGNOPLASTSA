VERSION 5.00
Begin VB.Form frmDatosVarios 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Datos Varios..."
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   5595
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   5595
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ Datos Varios ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5535
      Begin VB.TextBox txtMOM 
         Height          =   285
         Left            =   2520
         TabIndex        =   16
         Top             =   1560
         Width           =   2895
      End
      Begin VB.TextBox txtManteOferta 
         Height          =   285
         Left            =   2520
         TabIndex        =   14
         Top             =   480
         Width           =   2895
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Volver"
         Default         =   -1  'True
         Height          =   375
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   4800
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Modificar"
         Height          =   375
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   4800
         Width           =   975
      End
      Begin VB.TextBox txtPintura 
         Height          =   285
         Left            =   1680
         TabIndex        =   11
         Top             =   4200
         Width           =   3735
      End
      Begin VB.TextBox txtMas15 
         Height          =   285
         Left            =   2400
         TabIndex        =   9
         Top             =   3360
         Width           =   3015
      End
      Begin VB.TextBox txtMenos15 
         Height          =   285
         Left            =   2400
         TabIndex        =   8
         Top             =   3000
         Width           =   3015
      End
      Begin VB.TextBox txtMenos10 
         Height          =   285
         Left            =   2400
         TabIndex        =   7
         Top             =   2640
         Width           =   3015
      End
      Begin VB.TextBox txtMDO 
         Height          =   285
         Left            =   2520
         TabIndex        =   2
         Top             =   1080
         Width           =   2895
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Porcentual MDO Muerta"
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
         TabIndex        =   17
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mantenimiento de oferta"
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
         TabIndex        =   15
         Top             =   480
         Width           =   2175
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         X1              =   120
         X2              =   5400
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   120
         X2              =   5400
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pintura por M2"
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
         Top             =   4200
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Piezas > 15 Kg"
         Height          =   255
         Left            =   1080
         TabIndex        =   6
         Top             =   3360
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Piezas < 15 Kg"
         Height          =   255
         Left            =   1080
         TabIndex        =   5
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Piezas < 10 Kg"
         Height          =   255
         Left            =   1080
         TabIndex        =   4
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Porcentuales sobre materiales"
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
         TabIndex        =   3
         Top             =   2280
         Width           =   2775
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Porcentual sobre MDO"
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
         TabIndex        =   1
         Top             =   1080
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmDatosVarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim base As New classConfigurar

Private Sub Command1_Click()
    g = MsgBox("¿Seguro de actualizar datos?", vbYesNo, "Advertencia")
    If g = 6 Then
        strsql = "update configuracion set ManteOferta=" & CDbl(Me.txtManteOferta) & " , mano_obra_muerta=" & CDbl(Me.txtMOM) & ", PorcMO=" & CDbl(Me.txtMDO) & ",PintM2=" & CDbl(Me.txtPintura) & ", porMAMenos10=" & Me.txtMenos10 & ",porMaMas15=" & CDbl(Me.txtMas15) & ",porMaMenos15=" & CDbl(Me.txtMenos15)
        base.ejecutar_consulta (strsql)
    End If
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Me.txtManteOferta = Configurar.manteOferta
    Me.txtMas15 = Configurar.PorMaMas15
    Me.txtMDO = Configurar.PorcMO
    Me.txtMenos10 = Configurar.PorMAMenos10
    Me.txtMenos15 = Configurar.PorMAMenos15
    Me.txtMOM = Configurar.Mano_obra_muerta
    Me.txtPintura = Configurar.PintM2


End Sub

Private Sub Form_Load()
    FormHelper.Customize Me
    If Trim(Me.txtMDO) = Empty Or Trim(Me.txtMas15) = Empty Or Trim(Me.txtMenos10) = Empty Or Trim(Me.txtMenos15) = Empty Or Trim(Me.txtPintura) = Empty Then
        Command1.Enabled = False
    Else
        Command1.Enabled = True
    End If

End Sub
Private Sub txtManteOferta_Validate(Cancel As Boolean)
    funciones.ValidarTextBox txtManteOferta, Cancel
End Sub

Private Sub txtMas15_Change()
    If Trim(Me.txtMDO) = Empty Or Trim(Me.txtMas15) = Empty Or Trim(Me.txtMenos10) = Empty Or Trim(Me.txtMenos15) = Empty Or Trim(Me.txtPintura) = Empty Then
        Command1.Enabled = False
    Else
        Command1.Enabled = True
    End If

End Sub

Private Sub txtMas15_GotFocus()
    foco Me.txtMas15
End Sub

Private Sub txtMas15_Validate(Cancel As Boolean)
    funciones.ValidarTextBox txtMas15, Cancel
End Sub

Private Sub txtMDO_Change()
    If Trim(Me.txtMDO) = Empty Or Trim(Me.txtMas15) = Empty Or Trim(Me.txtMenos10) = Empty Or Trim(Me.txtMenos15) = Empty Or Trim(Me.txtPintura) = Empty Then
        Command1.Enabled = False
    Else
        Command1.Enabled = True
    End If


End Sub

Private Sub txtMDO_GotFocus()
    foco Me.txtMDO
End Sub

Private Sub txtMDO_Validate(Cancel As Boolean)
    funciones.ValidarTextBox txtMDO, Cancel
End Sub

Private Sub txtMenos10_Change()
    If Trim(Me.txtMDO) = Empty Or Trim(Me.txtMas15) = Empty Or Trim(Me.txtMenos10) = Empty Or Trim(Me.txtMenos15) = Empty Or Trim(Me.txtPintura) = Empty Then
        Command1.Enabled = False
    Else
        Command1.Enabled = True
    End If

End Sub

Private Sub txtMenos10_GotFocus()
    foco Me.txtMenos10
End Sub

Private Sub txtMenos10_Validate(Cancel As Boolean)
    funciones.ValidarTextBox txtMenos10, Cancel
End Sub

Private Sub txtMenos15_Change()
    If Trim(Me.txtMDO) = Empty Or Trim(Me.txtMas15) = Empty Or Trim(Me.txtMenos10) = Empty Or Trim(Me.txtMenos15) = Empty Or Trim(Me.txtPintura) = Empty Then
        Command1.Enabled = False
    Else
        Command1.Enabled = True
    End If

End Sub

Private Sub txtMenos15_GotFocus()
    foco Me.txtMenos15
End Sub

Private Sub txtMenos15_Validate(Cancel As Boolean)
    If Not IsNumeric(Me.txtMenos15) Then Cancel = True Else Cancel = False
End Sub

Private Sub txtMOM_Validate(Cancel As Boolean)
    funciones.ValidarTextBox txtMOM, Cancel
End Sub

Private Sub txtPintura_Change()
    If Trim(Me.txtMDO) = Empty Or Trim(Me.txtMas15) = Empty Or Trim(Me.txtMenos10) = Empty Or Trim(Me.txtMenos15) = Empty Or Trim(Me.txtPintura) = Empty Then
        Command1.Enabled = False
    Else
        Command1.Enabled = True
    End If

End Sub

Private Sub txtPintura_GotFocus()
    foco Me.txtPintura
End Sub

Private Sub txtPintura_Validate(Cancel As Boolean)
    funciones.ValidarTextBox txtPintura, Cancel
End Sub

Public Sub foco(ByRef texto As TextBox)
    texto.SelStart = 0
    texto.SelLength = Len(texto)
End Sub
