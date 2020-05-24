VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmComprasPreciosPorRubro 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Administrador de precios..."
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   14175
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   14175
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command5 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   255
      Left            =   12960
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   6240
      Width           =   975
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ Modificar valor ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   0
      TabIndex        =   4
      Top             =   4320
      Width           =   14175
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "[ Por Artículo ]"
         Height          =   2055
         Left            =   7560
         TabIndex        =   7
         Top             =   240
         Width           =   6495
         Begin VB.TextBox txtPorMatValor 
            Height          =   285
            Left            =   1320
            TabIndex        =   27
            Top             =   1080
            Width           =   3855
         End
         Begin VB.CommandButton Command4 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Modificar"
            Height          =   255
            Left            =   5400
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   1080
            Width           =   855
         End
         Begin VB.CommandButton Command3 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Modificar"
            Height          =   255
            Left            =   5400
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox txtPorMat 
            Height          =   285
            Left            =   1320
            TabIndex        =   20
            Top             =   720
            Width           =   3855
         End
         Begin VB.Label Label7 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Valor"
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
            TabIndex        =   28
            Top             =   1080
            Width           =   1095
         End
         Begin VB.Label IdMaterial 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Label7"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   2280
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label Label5 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Incremental"
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
            TabIndex        =   21
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label6 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Material"
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
            TabIndex        =   19
            Top             =   360
            Width           =   855
         End
         Begin VB.Label lblmat 
            BackColor       =   &H00E0E0E0&
            Height          =   255
            Left            =   1080
            TabIndex        =   18
            Top             =   360
            Width           =   5295
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "[ Por Grupo ]"
         Height          =   2055
         Left            =   3840
         TabIndex        =   6
         Top             =   240
         Width           =   3615
         Begin VB.CommandButton Command2 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Modificar"
            Height          =   375
            Left            =   1440
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   1440
            Width           =   855
         End
         Begin VB.TextBox txtPorGrupo 
            Height          =   285
            Left            =   1200
            TabIndex        =   12
            Top             =   720
            Width           =   2175
         End
         Begin VB.Label idGrupo 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Label7"
            Height          =   255
            Left            =   240
            TabIndex        =   24
            Top             =   2280
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label lblGrupo 
            BackColor       =   &H00E0E0E0&
            Height          =   255
            Left            =   840
            TabIndex        =   17
            Top             =   360
            Width           =   2535
         End
         Begin VB.Label Label4 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Grupo"
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
            TabIndex        =   15
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Incremental"
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
            TabIndex        =   11
            Top             =   720
            Width           =   1095
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "[ Por Rubro ]"
         Height          =   2055
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   3615
         Begin VB.CommandButton Command1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Modificar"
            Height          =   375
            Left            =   1440
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   1200
            Width           =   855
         End
         Begin VB.TextBox txtPorRubro 
            Height          =   285
            Left            =   1320
            TabIndex        =   9
            Top             =   720
            Width           =   2055
         End
         Begin VB.Label idRubro 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Label7"
            Height          =   375
            Left            =   120
            TabIndex        =   23
            Top             =   2280
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label lblRubro 
            BackColor       =   &H00E0E0E0&
            Height          =   255
            Left            =   720
            TabIndex        =   16
            Top             =   360
            Width           =   2655
         End
         Begin VB.Label Label3 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Rubro"
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
            TabIndex        =   14
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Incremental"
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
            TabIndex        =   8
            Top             =   720
            Width           =   1095
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ Seleccione nivel ]"
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
      TabIndex        =   0
      Top             =   0
      Width           =   14175
      Begin MSComctlLib.ListView lstMateriales 
         Height          =   3735
         Left            =   7560
         TabIndex        =   3
         Top             =   240
         Width           =   6450
         _ExtentX        =   11377
         _ExtentY        =   6588
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lstGrupos 
         Height          =   3735
         Left            =   3840
         TabIndex        =   2
         Top             =   240
         Width           =   3600
         _ExtentX        =   6350
         _ExtentY        =   6588
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lstRubros 
         Height          =   3735
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3600
         _ExtentX        =   6350
         _ExtentY        =   6588
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
End
Attribute VB_Name = "frmComprasPreciosPorRubro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim baseSP As New classSignoplast
Dim baseC As New classCompras
Private Sub Command1_Click()
    If Me.txtPorRubro <> 0 Then
        If baseC.cambiarPrecios(CInt(Me.idRubro), CDbl(Me.txtPorRubro), 1) Then
            MsgBox "Cambio exitoso", vbInformation, "Información"
            Me.txtPorRubro = 0
            llenarListas
        Else
            MsgBox "Se produjo un error", vbCritical, "Error"
        End If
    End If
End Sub
Private Sub Command2_Click()
    If Me.txtPorGrupo <> 0 Then
        If baseC.cambiarPrecios(CInt(Me.idGrupo), CDbl(Me.txtPorGrupo), 2) Then
            MsgBox "Cambio exitoso", vbInformation, "Información"
            Me.txtPorGrupo = 0
            llenarListas
        Else
            MsgBox "Se produjo un error", vbCritical, "Error"
        End If
    End If
End Sub
Private Sub Command3_Click()
    If Trim(Me.txtPorMat) <> Empty Then
        If baseC.cambiarPrecios(CInt(Me.IdMaterial), CDbl(Me.txtPorMat), 3) Then
            MsgBox "Cambio exitoso", vbInformation, "Información"
            Me.txtPorMat = 0
            llenarListas
        Else
            MsgBox "Se produjo un error", vbCritical, "Error"
        End If
    End If
End Sub
Private Sub Command4_Click()
    If Me.txtPorMatValor <> 0 Then
        If baseC.cambiarPrecios(CInt(Me.IdMaterial), CDbl(Me.txtPorMatValor), 3, 1) Then
            MsgBox "Cambio exitoso", vbInformation, "Información"
            Me.txtPorMatValor = 0
            llenarListas
        Else
            MsgBox "Se produjo un error", vbCritical, "Error"
        End If
    End If
End Sub
Private Sub Command5_Click()
    If MsgBox("¿Está seguro de salir?", vbYesNo, "Confirmación") = vbYes Then
        Unload Me
    End If
End Sub

Private Sub Form_Activate()
    frame3.Enabled = True
    Frame4.Enabled = False
    Frame5.Enabled = False
    llenarListas
    Command1.Default = True
    Me.txtPorRubro.SetFocus

End Sub
Function llenarListas()
    baseSP.llenarLstRubros Me.lstRubros
    baseSP.llenarLstGrupos CInt(Me.lstRubros.selectedItem), Me.lstGrupos
    baseSP.llenarLstmateriales CInt(Me.lstRubros.selectedItem), CInt(Me.lstGrupos.selectedItem), Me.lstMateriales
    leerListas
End Function
Private Sub Form_Load()
    FormHelper.Customize Me
    Me.txtPorGrupo = 0
    Me.txtPorMat = 0
    Me.txtPorRubro = 0
    Me.txtPorMatValor = 0
End Sub

Private Sub lstGrupos_Click()
    frame3.Enabled = False
    Frame4.Enabled = True
    Frame5.Enabled = False
    baseSP.llenarLstmateriales CInt(Me.lstRubros.selectedItem), CInt(Me.lstGrupos.selectedItem), Me.lstMateriales
    leerListas
    Me.Command2.Default = True
    Me.txtPorGrupo.SetFocus
End Sub
Private Sub lstMateriales_Click()
    frame3.Enabled = False
    Frame4.Enabled = False
    Frame5.Enabled = True
    leerListas
    Me.Command3.Default = True
    Me.txtPorMat.SetFocus
End Sub
Private Sub lstRubros_Click()
    On Error Resume Next
    frame3.Enabled = True
    Frame4.Enabled = False
    Frame5.Enabled = False
    Me.lstGrupos.ListItems.Clear
    Me.lstMateriales.ListItems.Clear
    baseSP.llenarLstGrupos CInt(Me.lstRubros.selectedItem), Me.lstGrupos
    baseSP.llenarLstmateriales CInt(Me.lstRubros.selectedItem), CInt(Me.lstGrupos.selectedItem), Me.lstMateriales
    leerListas
    Command1.Default = True

    Me.txtPorRubro.SetFocus
End Sub


Private Sub leerListas()
    'grupos
    If Me.lstGrupos.ListItems.count > 0 Then
        Me.lblGrupo = Me.lstGrupos.selectedItem.ListSubItems(1)
        Me.idGrupo = Me.lstGrupos.selectedItem
    Else
        Me.lblGrupo = Empty
        Me.idGrupo = Empty
    End If
    'rubros
    If Me.lstRubros.ListItems.count > 0 Then
        Me.lblRubro = Me.lstRubros.selectedItem.ListSubItems(1)
        Me.idRubro = Me.lstRubros.selectedItem
    Else
        Me.lblRubro = Empty
        Me.idRubro = Empty
    End If
    If Me.lstMateriales.ListItems.count > 0 Then
        Me.lblmat = Me.lstMateriales.selectedItem.ListSubItems(1)
        Me.IdMaterial = Me.lstMateriales.selectedItem
    Else
        Me.lblmat = Empty
        Me.IdMaterial = Empty
    End If

End Sub
