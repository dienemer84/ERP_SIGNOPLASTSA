VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPlaneamientoOTModificarCantidad 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Modificar..."
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4620
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ Nota a producción ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   0
      TabIndex        =   15
      Top             =   4080
      Width           =   4575
      Begin VB.CommandButton Command3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ok"
         Height          =   255
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox notaProduccion 
         Height          =   1215
         Left            =   120
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   16
         Top             =   240
         Width           =   4335
      End
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Command1"
      Height          =   735
      Left            =   3840
      TabIndex        =   14
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ Valor ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   12
      Top             =   2040
      Width           =   4575
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ok"
         Height          =   255
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox txtValor 
         Height          =   285
         Left            =   1080
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Valor"
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
         Left            =   480
         TabIndex        =   13
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[  Fecha de Entrega  ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   10
      Top             =   3000
      Width           =   4575
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   255
         Left            =   1080
         TabIndex        =   6
         Top             =   360
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   450
         _Version        =   393216
         Format          =   75956225
         CurrentDate     =   39167
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ok"
         Height          =   255
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha"
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
         Left            =   360
         TabIndex        =   11
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ Nueva cantidad ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4575
      Begin VB.TextBox txtCantidad 
         Height          =   285
         Left            =   1080
         TabIndex        =   0
         Top             =   480
         Width           =   2415
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ok"
         Default         =   -1  'True
         Height          =   255
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtDetalle 
         Height          =   765
         Left            =   1080
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cantidad"
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
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Detalle"
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
         Left            =   360
         TabIndex        =   8
         Top             =   1080
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmPlaneamientoOTModificarCantidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    frmPlaneamientoOTNueva.lstDetalleOT.SelectedItem.ListSubItems(4).text = funciones.formatearDecimales(CDbl(Trim(Me.txtValor)), 2)
    Unload Me
    
End Sub

Private Sub Command3_Click()
    If Trim(Me.notaProduccion) <> Empty Then
        frmPlaneamientoOTNueva.lstDetalleOT.SelectedItem.ListSubItems(7).Tag = UCase(Me.notaProduccion)
        Unload Me
    End If

End Sub

Private Sub Command4_Click()
    If IsNumeric(Me.txtCantidad) Then
        frmPlaneamientoOTNueva.lstDetalleOT.SelectedItem.ListSubItems(1).text = CDbl(Me.txtCantidad)
        frmPlaneamientoOTNueva.lstDetalleOT.SelectedItem.ListSubItems(7).text = UCase(Me.txtDetalle)
        Unload Me
    End If
End Sub
Private Sub Command5_Click()
    frmPlaneamientoOTNueva.lstDetalleOT.SelectedItem.ListSubItems(6).text = Me.DTPicker1
    Unload Me
End Sub

Private Sub DTPicker1_GotFocus()
    Me.Command5.Default = True
End Sub

Private Sub Form_Load()
    Me.DTPicker1 = Now
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
    If Not IsNumeric(Text1) Then
        Cancel = True
    Else
        Cancel = False
    End If
End Sub

Private Sub Text4_Change()
    If Not IsNumeric(Text3) Then
        Cancel = True
    Else
        Cancel = False
    End If
End Sub

Private Sub notaProduccion_GotFocus()
    foco Me.notaProduccion
    Me.Command3.Default = True
End Sub

Private Sub txtCantidad_GotFocus()
    foco Me.txtCantidad
    Me.Command4.Default = True
End Sub
Private Sub txtDetalle_GotFocus()
    foco Me.txtDetalle
    Me.Command4.Default = True
End Sub
Private Sub txtValor_GotFocus()
    foco Me.txtValor
    Me.Command2.Default = True
End Sub
