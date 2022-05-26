VERSION 5.00
Begin VB.Form frmConfigurarPintura 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configurar pintura"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   4710
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ Datos Necesarios ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4695
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Cancel          =   -1  'True
         Caption         =   "Salir"
         Height          =   375
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   3600
         Width           =   1095
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   2760
         TabIndex        =   16
         Text            =   "Text6"
         Top             =   3120
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Actualizar"
         Default         =   -1  'True
         Height          =   375
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   3600
         Width           =   1095
      End
      Begin VB.TextBox Textg 
         Height          =   285
         Left            =   2760
         TabIndex        =   7
         Text            =   "Text7"
         Top             =   2760
         Width           =   1575
      End
      Begin VB.TextBox Textf 
         Height          =   285
         Left            =   2760
         TabIndex        =   6
         Text            =   "Text6"
         Top             =   2400
         Width           =   1575
      End
      Begin VB.TextBox Texta 
         Height          =   285
         Left            =   2760
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox Textb 
         Height          =   285
         Left            =   2760
         TabIndex        =   1
         Text            =   "Text2"
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox Textc 
         Height          =   285
         Left            =   2760
         TabIndex        =   2
         Text            =   "Text3"
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox Textd 
         Height          =   285
         Left            =   2760
         TabIndex        =   4
         Text            =   "Text4"
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox Texte 
         Height          =   285
         Left            =   2760
         TabIndex        =   5
         Text            =   "Text5"
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Espesor de la pintura"
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
         Top             =   3120
         Width           =   2535
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Indice de aumento"
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
         Top             =   2760
         Width           =   2535
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Indice de aumento MDO"
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
         TabIndex        =   14
         Top             =   2400
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cantidad de Pintura por M2"
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
         TabIndex        =   13
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fosfatos"
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
         TabIndex        =   12
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tiempo preparacion por M2"
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
         TabIndex        =   11
         Top             =   1200
         Width           =   2535
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tiempo pintura por M2"
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
         TabIndex        =   10
         Top             =   1560
         Width           =   2535
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tiempo horneado por M2"
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
         TabIndex        =   9
         Top             =   1920
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frmConfigurarPintura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim base As New classConfigurar

Private Sub Command1_Click()
    base.ejecutar "update Terminacion set CantPintM2=" & CDbl(Texta) & ", CantForfatosM2=" & CDbl(Textb) & ",TpoPrepSupM2=" & CDbl(Textc) & ", TpoPinturaM2=" & CDbl(Textd) & ", TpoHorneado=" & CDbl(Texte) & ", CantKg=0, factorAumentoMat=" & CDbl(Textf) & ", factotAumentoMDO=" & CDbl(Textg) & ", espesor=" & CDbl(Text6)
End Sub

Private Sub Command2_Click()
    If MsgBox("¿Seguro de salir?", vbYesNo, "Confirmación") = vbYes Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    FormHelper.Customize Me
    comprobar
    base.ver_datos_pintura A, B, c, d, E, F, g, h, i
    Me.Texta = A    'cantpintm2
    Me.Textb = B    'cantfosfatos
    Me.Textc = c    'tpo prrp sup
    Me.Textd = d    'tpo pint m2
    Me.Texte = E    'tpo horno
    Me.Textf = h    'factor mdo
    Me.Textg = g    'factor mat
    Me.Text6 = i    'espesor pintura
    
        Me.caption = caption & " (" & Name & ")"
        
        
End Sub

Private Sub comprobar()
    If Trim(Texta) = Empty Or Trim(Textb) = Empty Or Trim(Textc) = Empty Or Trim(Textd) = Empty Or Trim(Texte) = Empty Or Trim(Textf) = Empty Or Trim(Textg) = Empty Or Trim(Text6) = Empty Then
        Command1.Enabled = False
    Else
        Command1.Enabled = True
    End If
End Sub

Private Sub Text6_Change()
    comprobar
End Sub

Private Sub Text6_GotFocus()
    foco Me.Text6
End Sub

Private Sub Texta_Change()
    comprobar
End Sub

Private Sub Texta_GotFocus()
    foco Me.Texta
End Sub

Private Sub Textb_Change()
    comprobar
End Sub

Private Sub Textb_GotFocus()
    foco Me.Textb
End Sub

Private Sub Textc_Change()
    comprobar
End Sub

Private Sub Textc_GotFocus()
    foco Me.Textc
End Sub

Private Sub Textd_Change()
    comprobar
End Sub

Private Sub Textd_GotFocus()
    foco Me.Textd
End Sub

Private Sub Texte_Change()
    comprobar
End Sub

Private Sub Texte_GotFocus()
    foco Me.Texte
End Sub

Private Sub Textf_Change()
    comprobar
End Sub

Private Sub Textf_GotFocus()
    foco Me.Textf
End Sub

Private Sub Textg_Change()
    comprobar
End Sub

Private Sub Textg_GotFocus()
    foco Me.Textg
End Sub
