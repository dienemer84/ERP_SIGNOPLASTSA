VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~3.OCX"
Begin VB.Form frmContactoInterno 
   Caption         =   "Contacto"
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8115
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6510
   ScaleWidth      =   8115
   Begin XtremeSuiteControls.PushButton cmdGuardar 
      Height          =   450
      Left            =   150
      TabIndex        =   1
      Top             =   5040
      Width           =   1260
      _Version        =   786432
      _ExtentX        =   2222
      _ExtentY        =   794
      _StockProps     =   79
      Caption         =   "Guardar"
      UseVisualStyle  =   -1  'True
   End
   Begin VB.TextBox txtContacto 
      Height          =   4725
      Left            =   150
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   7890
   End
   Begin XtremeSuiteControls.PushButton cmdRefresh 
      Height          =   450
      Left            =   1545
      TabIndex        =   2
      Top             =   5055
      Width           =   1260
      _Version        =   786432
      _ExtentX        =   2222
      _ExtentY        =   794
      _StockProps     =   79
      Caption         =   "Actualizar"
      UseVisualStyle  =   -1  'True
   End
End
Attribute VB_Name = "frmContactoInterno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim usrKar As clsUsuario
Private Sub cmdGuardar_Click()
    usrKar.Memo = Me.txtContacto
    DAOUsuarios.SaveMemo usrKar
End Sub

Private Sub cmdRefresh_Click()
    MsgBox DateDiff("D", CDate("2010-11-01"), Now)
    refrescar
End Sub

Private Sub Form_Load()
    Customize Me
    refrescar
End Sub
Private Sub refrescar()
    Set usrKar = DAOUsuarios.GetById(24)
    Me.txtContacto = usrKar.Memo
End Sub
Private Sub Form_Resize()
On Error Resume Next
    Me.txtContacto.Height = Me.ScaleHeight - Me.cmdGuardar.Height - 320
    Me.txtContacto.Width = Me.ScaleWidth - 320
    Me.cmdGuardar.Top = Me.txtContacto.Height + 250
    Me.cmdRefresh.Top = Me.cmdGuardar.Top
End Sub
