VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAdminCCResumenSaldos 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Resúmen Saldos"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   5115
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   5115
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "Salir"
      CausesValidation=   0   'False
      Height          =   255
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Imprimir"
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3360
      Width           =   975
   End
   Begin MSComctlLib.ListView lstSaldos 
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5085
      _ExtentX        =   8969
      _ExtentY        =   5741
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Cliente"
         Object.Width           =   5821
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Saldo"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmAdminCCResumenSaldos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub llenarLST()
    Dim rs As Recordset
    Dim strsql As String
    Dim x As ListItem
    strsql = "select c.id,c.razon, sum(if(a.operacion=0,round(a.debehaber,2),0)-if(a.operacion=1,round(a.debehaber,2),0)) as saldo from AdminClientesCC a inner join clientes c on c.id=a.idcliente group by idcliente"
    Set rs = conectar.RSFactory(strsql)
    Me.lstSaldos.ListItems.Clear
    While Not rs.EOF
        client = Format(rs!Id, "0000") & " - " & rs!razon
        Set x = Me.lstSaldos.ListItems.Add(, , client)
        x.SubItems(1) = funciones.FormatearDecimales(rs!saldo, 2)
        If rs!saldo < 0 Then
            x.ListSubItems(1).ForeColor = vbRed
        End If
        rs.MoveNext
    Wend
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    FormHelper.Customize Me

    llenarLST
    Me.Refresh
End Sub

