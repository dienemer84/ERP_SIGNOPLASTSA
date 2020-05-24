VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPlaneamientoOENueva 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Nueva Orden de entrega..."
   ClientHeight    =   1920
   ClientLeft      =   1050
   ClientTop       =   1845
   ClientWidth     =   8160
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   8160
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Top             =   840
      Width           =   7095
   End
   Begin VB.CommandButton Command5 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Generar"
      Default         =   -1  'True
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   1440
      Width           =   1095
   End
   Begin VB.ComboBox cboClientesDestino 
      Height          =   315
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   480
      Width           =   7080
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   255
      Left            =   960
      TabIndex        =   4
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      Format          =   55508993
      CurrentDate     =   38923
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
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
      Left            =   0
      TabIndex        =   7
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cliente"
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
      Left            =   0
      TabIndex        =   6
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Entrega"
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
      Left            =   0
      TabIndex        =   5
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmPlaneamientoOENueva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rss As Recordset
Dim claseS As New classStock

Dim claseC As New classConfigurar
Dim claseP As New classPlaneamiento
Dim Cantidad As Long
Dim detalle As String
Dim idStock As Long
Dim vValor As Double
Dim c As Long
Public Property Let Valor(nValor As Double)
    vValor = nValor
End Property
Private Sub Command3_Click()
    On Error Resume Next
    Dim refe As String
    Dim nroOEGenerada As Long
    Dim clie As Long
    clie = Me.cboClientesDestino.ItemData(cboClientesDestino.ListIndex)
    refe = normaliza(Me.Text1)
    If MsgBox("¿Está seguro de crear una nueva Orden de entrega?", vbYesNo, "Confirmación") = vbYes Then
        If claseP.generarOE(Me.DTPicker1, refe, clie, nroOEGenerada) Then
            MsgBox "Orden de entrega creada correctamente con el número " & nroOEGenerada, vbInformation, "Información"
            Unload Me
        Else
            MsgBox "Error en la creación de la orden de entrega.", vbCritical, "Error"
        End If
    End If
End Sub
Private Sub Command5_Click()
    If MsgBox("¿Está seguro de salir?", vbYesNo, "Confirmación") = vbYes Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    FormHelper.Customize Me
    DAOCliente.LlenarCombo Me.cboClientesDestino
    Me.DTPicker1 = Now
End Sub

Private Sub Form_Terminate()
    Set rss = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set rss = Nothing
End Sub


