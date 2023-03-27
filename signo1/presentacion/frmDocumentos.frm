VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmDocumentos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Documentos"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   8295
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.PushButton cmVer 
      Height          =   405
      Left            =   6675
      TabIndex        =   2
      Top             =   705
      Width           =   960
      _Version        =   786432
      _ExtentX        =   1693
      _ExtentY        =   714
      _StockProps     =   79
      Caption         =   "Ver..."
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.ComboBox cboDocumentos 
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Top             =   750
      Width           =   5055
      _Version        =   786432
      _ExtentX        =   8916
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.PushButton cmdNuevo 
      Height          =   405
      Left            =   6660
      TabIndex        =   3
      Top             =   1215
      Width           =   960
      _Version        =   786432
      _ExtentX        =   1693
      _ExtentY        =   714
      _StockProps     =   79
      Caption         =   "Nuevo"
      UseVisualStyle  =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "Documento"
      Height          =   285
      Left            =   540
      TabIndex        =   0
      Top             =   795
      Width           =   1425
   End
End
Attribute VB_Name = "frmDocumentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements ISuscriber
Private Id_ As String
Dim Doc As documento

Private Sub cmdNuevo_Click()
    Dim frm As New frmConfigurarDocumentos
    frm.Show

End Sub

Private Sub cmVer_Click()
    If Me.cboDocumentos.ListIndex <> -1 Then
        Set Doc = DAODocumentos.FindById(Me.cboDocumentos.ItemData(Me.cboDocumentos.ListIndex))
    Else
        Set Doc = Nothing
    End If
    If IsSomething(Doc) Then
        Dim frm As New frmConfigurarDocumentos
        frm.Document Doc
        frm.Show
    End If

End Sub

Private Sub Form_Load()
    Customize Me
    Id_ = funciones.CreateGUID
    Channel.AgregarSuscriptor Me, Documentos_
    LlenarCbo
End Sub

Private Sub LlenarCbo()
    DAODocumentos.llenarComboXtremeSuite Me.cboDocumentos
End Sub

Private Sub Form_Terminate()
    Channel.RemoverSuscripcionTotal Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Channel.RemoverSuscripcionTotal Me
End Sub

Private Property Get ISuscriber_id() As String
    ISuscriber_id = Id_
End Property

Private Function ISuscriber_Notificarse(EVENTO As clsEventoObserver) As Variant
    LlenarCbo
End Function
