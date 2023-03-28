VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSectorizar 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sectorizar"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   255
   ClientWidth     =   5625
   ClipControls    =   0   'False
   Icon            =   "frmSectorizar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   5625
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ Sectores ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   0
      TabIndex        =   1
      Top             =   960
      Width           =   5535
      Begin VB.CommandButton Command5 
         Cancel          =   -1  'True
         Caption         =   "Salir"
         Height          =   375
         Left            =   2880
         TabIndex        =   12
         Top             =   3720
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Actualizar"
         Height          =   375
         Left            =   1440
         TabIndex        =   11
         Top             =   3720
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   2640
         TabIndex        =   9
         Top             =   1800
         Width           =   255
      End
      Begin VB.CommandButton Command2 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   2640
         TabIndex        =   8
         Top             =   360
         Width           =   255
      End
      Begin MSComctlLib.ListView lstSectores 
         Height          =   3135
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   5530
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Sector"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Sector"
            Object.Width           =   0
         EndProperty
      End
      Begin MSComctlLib.ListView lstSectoresEmpleados 
         Height          =   3135
         Left            =   2880
         TabIndex        =   7
         Top             =   360
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   5530
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Sector"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Sector"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ Empleado ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5535
      Begin VB.CommandButton Command1 
         Caption         =   "Ver"
         Default         =   -1  'True
         Height          =   255
         Left            =   1800
         TabIndex        =   4
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   960
         TabIndex        =   3
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblLegajo 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Label3"
         Height          =   255
         Left            =   5160
         TabIndex        =   10
         Top             =   960
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lblApellido 
         BackColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   2520
         TabIndex        =   5
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Legajo"
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
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmSectorizar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim claseSP As New classSignoplast
Dim claseS As New classSectores
Dim rec As Recordset
Dim IdEmpleado As Long


Public Sub Command1_Click()
    Dim errando1
    If Trim(Me.Text1) <> Empty Then
        If Not IsNumeric(Me.Text1) Then
            errando1 = "Ingrese un legajo válido."
            Me.Frame2.Enabled = False
        Else
            If IsSomething(DAOEmpleados.GetByLegajo(CLng(Me.Text1))) Then
                Me.Frame2.Enabled = True
                llenar
            Else
                MsgBox "El legajo indicado no existe.", vbCritical, "Error"
                llenarLstSectores
                limpiar
                Me.Frame2.Enabled = False
            End If
        End If
    End If
End Sub
Private Sub limpiar()
    Me.lblApellido = Empty
    'Me.lblNombre = Empty
    Me.lstSectoresEmpleados.ListItems.Clear
End Sub
Private Sub llenar()
    Dim emple As clsEmpleado

    Set emple = DAOEmpleados.GetByLegajo(Me.Text1)
    If IsSomething(emple) Then
        IdEmpleado = emple.Id
        Me.lblApellido = emple.NombreCompleto
        Me.lblLegajo = emple.legajo
        llenarLstSectoresEmpleados
    End If
End Sub
Private Sub llenarLstSectores()
    Me.lstSectores.ListItems.Clear
    Dim x As ListItem
    Set rec = claseSP.ListaSectores
    While Not rec.EOF
        Set x = Me.lstSectores.ListItems.Add(, , rec!Sector)
        x.SubItems(1) = rec!Id
        rec.MoveNext
    Wend
End Sub
Private Sub llenarLstSectoresEmpleados()
    Me.lstSectoresEmpleados.ListItems.Clear
    Dim x As ListItem
    '    Set rec = claseP.ListaSectoresEmpleados(CLng(Me.Text1))
    Dim Sector As clsSector
    Dim sectores As Collection
    Set sectores = DAOSectores.GetByIdEmpleado(IdEmpleado)


    For Each Sector In sectores
        Set x = Me.lstSectoresEmpleados.ListItems.Add(, , Sector.Sector)
        x.SubItems(1) = Sector.Id
    Next Sector


    '    While Not rec.EOF
    '        Set x = Me.lstSectoresEmpleados.ListItems.Add(, , rec!sector)
    '        x.SubItems(1) = rec!id
    '        rec.MoveNext
    '    Wend
End Sub

Private Sub Command2_Click()
    For x = 1 To Me.lstSectores.ListItems.count
        If Me.lstSectores.ListItems(x).Checked = True Then
            esta = False
            For i = 1 To Me.lstSectoresEmpleados.ListItems.count
                If Me.lstSectoresEmpleados.ListItems(i) = Me.lstSectores.ListItems(x) Then esta = True
            Next i
            If Not esta Then
                Dim h As ListItem
                Set h = Me.lstSectoresEmpleados.ListItems.Add(, , Me.lstSectores.ListItems(x))
                h.SubItems(1) = Me.lstSectores.ListItems(x).ListSubItems(1)
            End If
        End If
    Next x
End Sub

Private Sub Command3_Click()
    For i = Me.lstSectoresEmpleados.ListItems.count To 1 Step -1
        If Me.lstSectoresEmpleados.ListItems(i).Checked = True Then
            Me.lstSectoresEmpleados.ListItems.remove (i)
        End If
    Next i

End Sub

Private Sub Command4_Click()
    If MsgBox("¿Seguro de actualizar?", vbYesNo, "Confirmación") = vbYes Then

        Dim li As ListItem
        Dim q As String
        For Each li In Me.lstSectoresEmpleados.ListItems
            q = "INSERT INTO sectorizacion (idEmpleado, idSector) VALUES (" & IdEmpleado & ", " & li.SubItems(1) & ") ON DUPLICATE KEY UPDATE idSector = VALUES(idSector)"
            Debug.Print "ejecucion:", conectar.execute(q)
        Next li

        'claseP.sectorizar Me.lstSectoresEmpleados, idEmpleado
    End If
End Sub

Private Sub Command5_Click()
    If MsgBox("¿Está seguro de salir?", vbYesNo, "Confirmación") = vbYes Then
        Unload Me
    End If
End Sub

Private Sub Form_Activate()
    llenarLstSectores
    valida
End Sub
Private Sub valida()
    If Trim(Text1) = Empty Then
        Command1.Enabled = False
        Me.Frame2.Enabled = False
    Else
        Command1.Enabled = True
        Me.Frame2.Enabled = True
    End If
End Sub

Private Sub Form_Load()
    FormHelper.Customize Me
End Sub

Private Sub Text1_Change()
    If Trim(Me.Text1) = Empty Then
        Command1.Enabled = False
    Else
        Command1.Enabled = True
    End If
End Sub

Private Sub Text1_GotFocus()
    foco Me.Text1
End Sub
