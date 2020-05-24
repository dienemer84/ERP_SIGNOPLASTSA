VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAdminFacturasAdeudadasRegistrar 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Registrar fecha de pago propuesta..."
   ClientHeight    =   2445
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   6090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   6090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Volver"
      Height          =   375
      Left            =   3240
      TabIndex        =   9
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   1440
      TabIndex        =   8
      Top             =   1920
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   1320
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   450
      _Version        =   393216
      Format          =   72351745
      CurrentDate     =   39777
   End
   Begin VB.Label lblCliente 
      BackColor       =   &H00C0C0C0&
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
      Left            =   1800
      TabIndex        =   7
      Top             =   600
      Width           =   4095
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Comprobante"
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
      Width           =   1575
   End
   Begin VB.Label lblMonto 
      BackColor       =   &H00FF8080&
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
      Left            =   1800
      TabIndex        =   4
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label lblComprobante 
      BackColor       =   &H00C0C0C0&
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
      Left            =   1800
      TabIndex        =   3
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Fecha Propuesta"
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
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Monto"
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
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Comprobante"
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
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmAdminFacturasAdeudadasRegistrar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim admin As New classAdministracion
Dim vIdFactura As Long
Public Property Let idFactura(nIdFactura As Long)
    vIdFactura = nIdFactura
End Property


Private Sub Command1_Click()
    Dim fechada As Date

    hoy = Now
    fechada = CDate(Me.DTPicker1)
    a = DateDiff("d", Now, fechada)
    If a < 0 Then
        MsgBox "Error, fecha incorrecta!", vbCritical, "Error"
        Me.DTPicker1 = Now
        Exit Sub
    End If

    If MsgBox("¿Desea asumir " & Format(DTPicker1, "dd/mm/yyyy") & " como fecha de pago de comprobante?", vbYesNo, "Confirmación") = vbYes Then

        If admin.definirFechaPago(vIdFactura, fechada) Then
            MsgBox "Se modificó la fecha propuesta de pago!", vbInformation, "Información"
        Else
            MsgBox "Se produjo algun error, no se realizan cambios!", vbCritical, "Error"
        End If
    End If


End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
FormHelper.Customize Me
    Me.DTPicker1 = Now
    Dim Total As Double
    Dim IdMoneda As Integer
    Dim Razon As String
    Dim idCliente As Long
    Dim nro As Long
    admin.TotalFactura vIdFactura, Total, IdMoneda, Razon, idCliente, True, nro
    Me.lblMonto = funciones.queMoneda(IdMoneda) & " " & funciones.FormatearDecimales(Total, 2)
    Me.lblComprobante = nro '& "-" & admin.queTipoFactura(vIdFactura)
    Me.lblCliente = Razon
End Sub
