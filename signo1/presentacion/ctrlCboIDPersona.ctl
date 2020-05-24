VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~3.OCX"
Begin VB.UserControl ctrlCboIDPersona 
   BackColor       =   &H8000000D&
   ClientHeight    =   405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8940
   ScaleHeight     =   405
   ScaleWidth      =   8940
   Begin VB.TextBox txtId 
      Height          =   285
      Left            =   15
      TabIndex        =   0
      Top             =   20
      Width           =   1440
   End
   Begin XtremeSuiteControls.ComboBox cboPersonas 
      Height          =   315
      Left            =   1485
      TabIndex        =   1
      Top             =   0
      Width           =   7470
      _Version        =   786432
      _ExtentX        =   13176
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Sorted          =   -1  'True
      Style           =   2
      UseVisualStyle  =   -1  'True
      Text            =   "ComboBox1"
   End
End
Attribute VB_Name = "ctrlCboIDPersona"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim vId As Long
Dim col As Collection
Private Sub cboPersonas_Click()
    On Error Resume Next
    txtId = cboPersonas.ItemData(cboPersonas.ListIndex)
    vId = Val(txtId)
End Sub

Public Property Let Personas(vcol As Collection)
    Set col = vcol
    Dim cli As clsCliente
    If IsSomething(col) Then
        For Each cli In col
            cboPersonas.AddItem cli.razon
            cboPersonas.ItemData(cboPersonas.NewIndex) = cli.id
        Next
        If cboPersonas.ListCount > 0 Then
            cboPersonas.ListIndex = 0
        End If
    End If
End Property
Private Sub txtId_Change()
    cboPersonas.ListIndex = funciones.PosIndexCbo(Val(txtId), cboPersonas)
    vId = Val(txtId)
End Sub



Public Property Get id() As Long
    id = vId
End Property



Private Sub UserControl_Initialize()
    BackColor = FormHelper.FondoCeleste
End Sub

Private Sub UserControl_Resize()
    cboPersonas.Width = UserControl.Width - 1600

End Sub
