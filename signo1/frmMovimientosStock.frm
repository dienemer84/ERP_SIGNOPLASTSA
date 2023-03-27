VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmMovimientosStock 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Historico de movimientos..."
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   5925
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   255
      Left            =   4920
      TabIndex        =   5
      Top             =   3720
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Frame1"
      Height          =   4095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      Begin MSComctlLib.ListView lstHistoricoStock 
         Height          =   3015
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   5318
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "id"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Cantidad"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Operación"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Nota"
            Object.Width           =   3316
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Fecha"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lblCliente 
         BackColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   840
         TabIndex        =   2
         Top             =   360
         Width           =   4935
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente:"
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
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Label lblid 
      Caption         =   "Label2"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   4080
      Width           =   975
   End
End
Attribute VB_Name = "frmMovimientosStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim baseS As classStock

Private Sub Command1_Click()
    Unload Me

End Sub

Private Sub Form_Activate()
    Set baseS = New classStock
    baseS.llenar_historico_stock CInt(Me.lblid), Me.lstHistoricoStock
End Sub

Private Sub Form_Load()
    FormHelper.Customize Me

    ''Me.caption = caption & " (" & Name & ")"
End Sub
