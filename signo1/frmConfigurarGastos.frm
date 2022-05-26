VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmConfigurarGastos 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configurar gastos..."
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   4110
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   4110
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   255
      Left            =   2640
      TabIndex        =   12
      Top             =   6360
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ Modificar ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   0
      TabIndex        =   6
      Top             =   3720
      Width           =   4095
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   960
         TabIndex        =   9
         Top             =   360
         Width           =   3015
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   960
         TabIndex        =   8
         Top             =   720
         Width           =   3015
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ok!"
         Height          =   255
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Concepto"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Porcentual"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ Altas ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   2160
      Width           =   4095
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ok!"
         Height          =   255
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   960
         TabIndex        =   4
         Top             =   720
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Porcentual"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Concepto"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
   End
   Begin MSComctlLib.ListView lstGastos 
      Height          =   2055
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   3625
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "id"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Detalle"
         Object.Width           =   4851
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Valor"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "valor_Real"
         Object.Width           =   0
      EndProperty
   End
End
Attribute VB_Name = "frmConfigurarGastos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim base As classConfigurar
Private Sub Command1_Click()
    If Not IsNumeric(Me.Text2) Then
        MsgBox "Debe introducir un valor porcentual ", vbCritical, "Error"
        Me.Text2.SetFocus
    Else
        strsql = "insert into gastos (concepto,porcentual) VALUES ('" & normaliza(Trim(Text1)) & "'," & CDbl(Trim(Text2)) & ")"
        base.ejecutar_consulta (strsql)
        base.LlenarListaGastos Me.lstGastos
    End If
End Sub
Private Sub Command2_Click()
    Concepto = normaliza(Trim(Text4))
    porcentual = CDbl(Trim(Text3))
    If Not IsNumeric(Me.Text3) Then
        MsgBox "Debe introducir un valor porcentual ", vbCritical, "Error"
        Me.Text3.SetFocus
    Else
        strsql = "UPDATE gastos set concepto='" & Concepto & "', porcentual=" & porcentual
        base.ejecutar_consulta (strsql)
        base.LlenarListaGastos Me.lstGastos
    End If
End Sub
Private Sub Command3_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    FormHelper.Customize Me
    Set base = New classConfigurar
    base.LlenarListaGastos Me.lstGastos
    If Trim(Text1) = Empty Or Trim(Text2) = Empty Then
        Command1.Enabled = False
    Else
        Command1.Enabled = True
    End If
    If Trim(Text3) = Empty Or Trim(Text4) = Empty Then
        Command2.Enabled = False
    Else
        Command2.Enabled = True
    End If
    
        Me.caption = caption & " (" & Name & ")"
        
        
End Sub

Private Sub lstGastos_Click()
    If Me.lstGastos.ListItems.count > 0 Then
        base.VerParaModificar Me.lstGastos.selectedItem, Me
    End If
End Sub

Private Sub Text1_Change()
    If Trim(Text1) = Empty Or Trim(Text2) = Empty Then
        Command1.Enabled = False
    Else
        Command1.Enabled = True
    End If
End Sub

Private Sub Text2_Change()
    If Trim(Text1) = Empty Or Trim(Text2) = Empty Then
        Command1.Enabled = False
    Else
        Command1.Enabled = True
    End If
End Sub

Private Sub Text3_Change()
    If Trim(Text3) = Empty Or Trim(Text4) = Empty Then
        Command2.Enabled = False
    Else
        Command2.Enabled = True
    End If
End Sub
Private Sub Text4_Change()
    If Trim(Text3) = Empty Or Trim(Text4) = Empty Then
        Command2.Enabled = False
    Else
        Command2.Enabled = True
    End If
End Sub
