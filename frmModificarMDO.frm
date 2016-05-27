VERSION 5.00
Begin VB.Form frmModificarMDO 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Modifiicar"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   4815
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ Tarea activa ]"
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      Begin VB.TextBox txtDetalle 
         Height          =   285
         Left            =   1200
         TabIndex        =   15
         Text            =   "Text2"
         Top             =   2280
         Width           =   3375
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00E0E0E0&
         Cancel          =   -1  'True
         Caption         =   "Salir"
         Height          =   375
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   2880
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Actualizar"
         Default         =   -1  'True
         Height          =   375
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2880
         Width           =   1335
      End
      Begin VB.TextBox txtTiempo 
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Text            =   "Text2"
         Top             =   1920
         Width           =   3375
      End
      Begin VB.TextBox txtCantOp 
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   1560
         Width           =   3375
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Detalle"
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
         TabIndex        =   16
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sector"
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
         TabIndex        =   12
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblSector 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Label2"
         Height          =   255
         Left            =   960
         TabIndex        =   11
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tarea"
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
         TabIndex        =   10
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblDescripcion 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Label2"
         Height          =   255
         Left            =   1440
         TabIndex        =   9
         Top             =   1080
         Width           =   3015
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tiempo"
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
         TabIndex        =   5
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Cant 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cant Oper"
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
         TabIndex        =   3
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label lblTarea 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Label2"
         Height          =   255
         Left            =   960
         TabIndex        =   2
         Top             =   720
         Width           =   3615
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripcion"
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
         TabIndex        =   1
         Top             =   1080
         Width           =   1095
      End
   End
   Begin VB.Label idDesMDO 
      Caption         =   "Label3"
      Height          =   255
      Left            =   600
      TabIndex        =   13
      Top             =   5400
      Width           =   855
   End
   Begin VB.Label lblCPP 
      Caption         =   "lblCPP"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   2880
      Width           =   615
   End
End
Attribute VB_Name = "frmModificarMDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim claseS As New classSignoplast
Dim formu As Form
Dim Modificado As Boolean
Public Property Let nuevo_form(frm As Variant)
    Set formu = frm
End Property


Private Sub Command1_Click()
    On Error GoTo err441
    If MsgBox("¿Desea actualizar datos?", vbYesNo, "Confirmacion") = vbYes Then
        idDM = CLng(Me.idDesMDO)
        claseS.ejecutar "select v.valor from tareas t inner join valores_MDO v on t.id=v.id_tarea and t.id=" & idDM
        Valor = claseS.valorMDO
        Tiempo = CDbl(Me.txtTiempo)
        cpp = CInt(Me.lblCPP)
        cantop = CDbl(Me.txtCantOp)
        detalle = UCase(Trim(Me.txtDetalle))

        If cpp > 0 Then    '(cpp variable)
            totmin = cantop * Tiempo / cpp
            totplata = totmin * Valor
        Else
            totmin = cantop * Tiempo
            totplata = totmin * Valor
        End If

        formu.ListView2.selectedItem.SubItems(9) = Format(Math.Round(totmin, 2), "0.00")
        formu.ListView2.selectedItem.ListSubItems(2).text = Me.txtCantOp
        formu.ListView2.selectedItem.ListSubItems(3).text = Me.txtTiempo
        formu.ListView2.selectedItem.ListSubItems(9).text = Math.Round(totmin, 2)
        formu.ListView2.selectedItem.ListSubItems(10).text = Math.Round(totplata, 2)
        formu.ListView2.selectedItem.Tag = detalle








        Unload Me
    End If
    Exit Sub
err441:
    MsgBox "Error: " & Err.Number
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
    If MsgBox("¿Está seguro de salir?", vbYesNo, "Confirmación") = vbYes Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    FormHelper.Customize Me
    Me.txtCantOp = Empty
    Me.txtTiempo = Empty
    Me.txtDetalle = Empty
    valida
End Sub

Private Sub txtCantOp_Change()
    valida
End Sub

Private Sub txtCantOp_GotFocus()
    foco Me.txtCantOp
End Sub

Private Sub txtCantOp_Validate(Cancel As Boolean)
    If Not IsNumeric(Me.txtCantOp) Then
        Cancel = True
    Else
        Cancel = False
    End If
End Sub

Private Sub txtTiempo_Change()
    valida
End Sub

Private Sub txtTiempo_GotFocus()
    foco Me.txtTiempo
End Sub

Private Sub valida()

    If Trim(Me.txtCantOp) = Empty Or Trim(Me.txtTiempo) = Empty Then
        Me.Command1.Enabled = False
    Else
        Me.Command1.Enabled = True
    End If
End Sub

Private Sub txtTiempo_Validate(Cancel As Boolean)
    If Not IsNumeric(Me.txtTiempo) Then
        Cancel = True
    Else
        Cancel = False
    End If
End Sub
