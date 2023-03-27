VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmSistemaAgregarNotasActualizacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Carga de detalle de Actualización"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10665
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   10665
   Begin XtremeSuiteControls.PushButton PushButtonCerrar 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   4320
      Width           =   1815
      _Version        =   786432
      _ExtentX        =   3201
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Cerrar"
      Appearance      =   1
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10335
      _Version        =   786432
      _ExtentX        =   18230
      _ExtentY        =   7011
      _StockProps     =   79
      Caption         =   "Datos"
      Appearance      =   6
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   1080
         TabIndex        =   7
         Text            =   "Text2"
         Top             =   1020
         Width           =   3495
      End
      Begin VB.TextBox Text1 
         Height          =   1695
         Left            =   1080
         TabIndex        =   3
         Top             =   1560
         Width           =   9015
      End
      Begin XtremeSuiteControls.DateTimePicker DateTimePicker1 
         Height          =   375
         Left            =   1080
         TabIndex        =   2
         Top             =   480
         Width           =   1335
         _Version        =   786432
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   68
         Format          =   1
         CurrentDate     =   44613.3355208333
      End
      Begin VB.Label Label3 
         Caption         =   "Módulo:"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha:"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   540
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Detalle:"
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   5
         Top             =   2220
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Fecha:"
         Height          =   255
         Index           =   0
         Left            =   12600
         TabIndex        =   4
         Top             =   7080
         Width           =   615
      End
   End
   Begin XtremeSuiteControls.PushButton PushButtonCargar 
      Height          =   375
      Index           =   0
      Left            =   8640
      TabIndex        =   1
      Top             =   4320
      Width           =   1815
      _Version        =   786432
      _ExtentX        =   3201
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Cargar"
      Appearance      =   1
   End
End
Attribute VB_Name = "frmSistemaAgregarNotasActualizacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Nota As clsNotas


Private Sub PushButtonCargar_Click(Index As Integer)
    CargarDetalleNuevo
End Sub

Private Sub CargarDetalleNuevo()

    On Error GoTo E

    If Nota Is Nothing Then Set Nota = New clsNotas


    Nota.FechaD_ = Now
    Nota.TextoD_ = Me.Text1
    Nota.Modulo_ = Me.Text2


    If DAOActualizar.CargarNuevoDetalle(Nota) Then

        MsgBox "Nueva nota ingresada con éxito!", vbInformation, "Información"

        Me.Text1 = ""
        Me.Text2 = ""

    Else

        MsgBox "Se produjo algún error, no se guardó la nota!", vbCritical, "Error"

    End If

    Exit Sub
E:
    MsgBox Err.Description, vbCritical

End Sub


Private Sub PushButtonCerrar_Click()
    Unload Me

End Sub
