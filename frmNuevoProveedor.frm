VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmComprasProveedoresNuevo 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Nuevo proveedor..."
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   7950
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   7950
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
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
      Height          =   7095
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   7935
      Begin VB.ComboBox CboIVA 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   3960
         Width           =   6495
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "[ Rubros ]"
         Height          =   2295
         Left            =   120
         TabIndex        =   28
         Top             =   4680
         Width           =   7695
         Begin VB.CommandButton Command5 
            BackColor       =   &H00E0E0E0&
            Cancel          =   -1  'True
            Caption         =   "Salir"
            Height          =   375
            Left            =   6600
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   1800
            Width           =   975
         End
         Begin VB.CommandButton Command4 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Agregar"
            Default         =   -1  'True
            Height          =   375
            Left            =   5520
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   1800
            Width           =   975
         End
         Begin VB.CommandButton Command3 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "<"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   3720
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   960
            Width           =   255
         End
         Begin VB.CommandButton Command2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   ">"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   3720
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   240
            Width           =   255
         End
         Begin MSComctlLib.ListView lstRubros 
            Height          =   1455
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   2566
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
            NumItems        =   0
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1455
            Left            =   3960
            TabIndex        =   13
            Top             =   240
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   2566
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
               Object.Width           =   529
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "Rubro habilitado"
               Object.Width           =   5292
            EndProperty
         End
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Dólares"
         Height          =   255
         Left            =   2640
         TabIndex        =   10
         Top             =   4440
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pago contra entrega"
         Height          =   255
         Left            =   4080
         TabIndex        =   11
         Top             =   4440
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   1320
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   360
         Width           =   6495
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   1320
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   720
         Width           =   6495
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   1320
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   1080
         Width           =   6495
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   3
         Left            =   1320
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   1440
         Width           =   6495
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   4
         Left            =   1320
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   1800
         Width           =   6495
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   5
         Left            =   1320
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   2160
         Width           =   6495
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   6
         Left            =   1320
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   2520
         Width           =   6495
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   7
         Left            =   1320
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   2880
         Width           =   6495
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   9
         Left            =   1320
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   3600
         Width           =   6495
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   8
         Left            =   1320
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   3240
         Width           =   6495
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Iva "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   31
         Top             =   3960
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Razón Social "
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
         TabIndex        =   27
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Domicilio "
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
         TabIndex        =   26
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ciudad "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "CP "
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
         TabIndex        =   24
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Teléfonos "
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
         TabIndex        =   23
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fax "
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
         TabIndex        =   22
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "E-Mail "
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
         TabIndex        =   21
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Contácto "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pago  "
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
         TabIndex        =   19
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Bonificación "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   3600
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmComprasProveedoresNuevo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim baseA As New classAdministracion
Dim baseP As New classCompras

Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()
For x = 1 To Me.lstRubros.ListItems.count
If Me.lstRubros.ListItems(x).Checked = True Then
 esta = False
 For I = 1 To Me.ListView1.ListItems.count
   If Me.ListView1.ListItems(I) = Me.lstRubros.ListItems(x) Then esta = True
 Next I
 If Not esta Then
    Set h = Me.ListView1.ListItems.Add(, , Me.lstRubros.ListItems(x))
     h.SubItems(1) = Me.lstRubros.ListItems(x).ListSubItems(1)
 End If
End If
Next x
End Sub

Private Sub Command3_Click()
For I = Me.ListView1.ListItems.count To 1 Step -1
If Me.ListView1.ListItems(I).Checked = True Then
 Me.ListView1.ListItems.Remove (I)
End If
  
Next I
End Sub

Public Sub foco(ByRef texto As TextBox)
texto.SelStart = 0
texto.SelLength = Len(texto)
End Sub

Private Sub Command4_Click()
If Trim(Text1(9)) = Empty Then Text1(9) = 0
baseP.cargar_datos Me.ListView1, 0
limpiar
Unload Me

End Sub

Private Sub Command5_Click()
If MsgBox("¿Está seguro de salir?", vbYesNo, "Confirmación") = vbYes Then
    Unload Me
End If
End Sub
Private Sub Form_Load()
'Set basep = New classProveedor
baseP.llenar_lista_rubros Me.lstRubros, -1, 3000, True, 0
baseA.llenarComboIVA Me.CboIVA, 1, True
limpiar
verificar Me.Command4
End Sub
Function limpiar()
For x = 0 To 8
Text1(x) = Empty
Next x
Text1(9) = 0
End Function
Function verificar(boton As CommandButton)
If Trim(Me.Text1(0)) = Empty Then
boton.Enabled = False
Else
boton.Enabled = True
End If
End Function

Private Sub Text1_Change(Index As Integer)
verificar Me.Command4
End Sub

Private Sub Text1_GotFocus(Index As Integer)
foco Text1(Index)
End Sub
