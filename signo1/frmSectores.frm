VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmSectores 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Sectores..."
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   4170
   ClipControls    =   0   'False
   DrawMode        =   14  'Copy Pen
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   4170
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ Alta ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   12
      Top             =   2160
      Width           =   4095
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   840
         TabIndex        =   14
         Top             =   360
         Width           =   2535
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Ok!"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3480
         TabIndex        =   13
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sector"
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   465
      End
   End
   Begin VB.Frame Frame4 
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
      Height          =   1095
      Left            =   0
      TabIndex        =   7
      Top             =   3240
      Width           =   4095
      Begin VB.CommandButton Command3 
         Caption         =   "Ok!"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3480
         TabIndex        =   9
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   840
         TabIndex        =   8
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label Label4 
         Caption         =   "Label3"
         Height          =   135
         Left            =   840
         TabIndex        =   11
         Top             =   480
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sector"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.CommandButton Command4 
      Cancel          =   -1  'True
      Caption         =   "Command4"
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Top             =   7080
      Width           =   855
   End
   Begin VB.Frame Frame3 
      Caption         =   "[ Baja ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   5280
      Width           =   4095
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   480
         Width           =   2535
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Ok!"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3480
         TabIndex        =   2
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Sector"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Label3"
         Height          =   135
         Left            =   840
         TabIndex        =   4
         Top             =   480
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin MSComctlLib.ListView lstSectores 
      Height          =   2055
      Left            =   0
      TabIndex        =   16
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "id"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Sector"
         Object.Width           =   6667
      EndProperty
   End
   Begin VB.Label lblsectorviejo 
      Caption         =   "Label6"
      Height          =   735
      Left            =   1800
      TabIndex        =   0
      Top             =   6480
      Width           =   615
   End
End
Attribute VB_Name = "frmSectores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim base As classSectores

Private Sub Command1_Click()
    strsql = "insert into sectores (sector) VALUES ('" & normaliza(Text1) & "')"
    base.ejecutar (strsql)
    base.LlenarListaSectores -1
    Text1 = Empty
    elegirMarcado
End Sub

Private Sub Command2_Click()

    A = MsgBox("¿Desea eliminar el sector " & Text2 & "?", vbOKCancel, "Advertencia")

    If A = 1 Then
        base.ejecutar ("delete from sectores where id=" & CInt(Label3) & " limit 1")
        Text2 = Empty
        Label3 = Empty
        'base.LlenarListaSectores
        elegirMarcado
    End If

End Sub

Private Sub Command3_Click()
    A = MsgBox("¿Desea modificar el sector " & Me.lblsectorviejo & "?", vbOKCancel, "Advertencia")
    If A = 1 Then
        base.ejecutar ("update sectores set sector='" & normaliza(Me.Text3) & "' where id=" & CInt(Label3))
        Text2 = Empty
        Label3 = Empty
        elegirMarcado
        base.LlenarListaSectores CInt(Me.Label3)
    End If
End Sub

Private Sub Command4_Click()
    If MsgBox("¿Está seguro de salir?", vbYesNo, "Confirmación") = vbYes Then
        Unload Me
    End If
End Sub

Private Sub Form_Activate()
    base.LlenarListaSectores CInt(Me.Label3)
End Sub
Private Sub Form_Load()
    FormHelper.Customize Me
    Set base = New classSectores
    Text1 = Empty
    Text2 = Empty
    Text3 = Empty

    Label3 = -1
    base.LlenarListaSectores -1
    elegirMarcado
    
        Me.caption = caption & " (" & Name & ")"

End Sub

Private Sub lstSectores_Click()
    elegirMarcado
End Sub

Private Sub Text1_Change()
    If Trim(Text1) = Empty Then
        Command1.Enabled = False
    Else
        Command1.Enabled = True
    End If
End Sub


Function elegirMarcado()
    If lstSectores.ListItems.count > 0 Then
        Text2 = lstSectores.selectedItem.ListSubItems(1)
        Label3 = lstSectores.selectedItem
        Text3 = Text2
        lblsectorviejo = Trim(Text3)
    End If
End Function

Private Sub Text1_GotFocus()
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1)
End Sub

Private Sub Text2_Change()
    If Trim(Text2) = Empty Then
        Command2.Enabled = False
    Else
        Command2.Enabled = True
    End If
End Sub

Private Sub Text3_Change()
    If Trim(Text3) = Empty Then
        Command3.Enabled = False
    Else
        Command3.Enabled = True
    End If

End Sub















