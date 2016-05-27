VERSION 5.00
Begin VB.Form FrmAdminFacturarConcepto 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Facturar Concepto..."
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   6180
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   6180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ Detalles del concepto ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2265
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      Begin VB.TextBox txtDescuento 
         Height          =   285
         Left            =   1080
         TabIndex        =   9
         Top             =   1440
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Volver"
         Height          =   375
         Left            =   4890
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1785
         Width           =   1140
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Cancel          =   -1  'True
         Caption         =   "Agregar"
         Default         =   -1  'True
         Height          =   375
         Left            =   3630
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1785
         Width           =   1140
      End
      Begin VB.TextBox txtValor 
         Height          =   285
         Left            =   1080
         TabIndex        =   6
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox txtCantidad 
         Height          =   285
         Left            =   1080
         TabIndex        =   5
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox txtConcepto 
         Height          =   285
         Left            =   1080
         TabIndex        =   4
         Top             =   360
         Width           =   4935
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descuento"
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
         Left            =   75
         TabIndex        =   10
         Top             =   1470
         Width           =   975
      End
      Begin VB.Label Label3 
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
         TabIndex        =   3
         Top             =   1095
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
         TabIndex        =   2
         Top             =   735
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Concepto "
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
         TabIndex        =   1
         Top             =   375
         Width           =   975
      End
   End
End
Attribute VB_Name = "FrmAdminFacturarConcepto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    On Error Resume Next
    Dim Errores As Boolean
    If Not IsNumeric(Me.txtCantidad) Or Not IsNumeric(Me.txtValor) Then Errores = True
    If Not Errores Then
        funciones.ConcCantidad = funciones.FormatearDecimales(CDbl(Me.txtCantidad), 2)
        funciones.ConcValor = funciones.FormatearDecimales(CDbl(Me.txtValor), 2)
        funciones.ConcConc = UCase(Me.txtConcepto)
        funciones.DescuentoDetalleFactura = Val(Me.txtDescuento.text)
        Unload Me
    Else
        MsgBox "Por favor, ingrese datos válidos.", vbCritical, "Error"
    End If

End Sub

Private Sub Command2_Click()
    If MsgBox("¿Seguro de volver?", vbYesNo, "Confirmación") = vbYes Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    FormHelper.Customize Me
    ValidarTXT
End Sub

Private Sub ValidarTXT()
    If Trim(Me.txtCantidad) = Empty Or Trim(Me.txtConcepto) = Empty Or Trim(Me.txtValor) = Empty Then
        Me.Command1.Enabled = False
    Else
        Me.Command1.Enabled = True
    End If
End Sub

Private Sub txtCantidad_Change()
    ValidarTXT
End Sub

Private Sub txtConcepto_Change()
    ValidarTXT
End Sub

Private Sub txtValor_Change()
    ValidarTXT
End Sub
