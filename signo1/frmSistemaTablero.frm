VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmSistemaTablero 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FF8080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tablero"
   ClientHeight    =   7770
   ClientLeft      =   570
   ClientTop       =   1065
   ClientWidth     =   14730
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   14730
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   120
      Top             =   8160
   End
   Begin MSComctlLib.ListView lstEventos 
      Height          =   1695
      Left            =   9720
      TabIndex        =   31
      Top             =   4800
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   2990
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Titulo"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Autor"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Estado"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Fecha"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.ComboBox cboGrupos 
      BackColor       =   &H00FFC0C0&
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   11895
   End
   Begin VB.Frame FRAME_ELEGIDO 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   7935
      Left            =   9720
      TabIndex        =   3
      Top             =   120
      Width           =   5055
      Begin VB.TextBox txtVer 
         BackColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   34
         Top             =   6480
         Width           =   4935
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Ver Tablero"
         Height          =   375
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   0
         Width           =   1335
      End
      Begin VB.CommandButton cmdAgregarEvento 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Command1"
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
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   4080
         Width           =   4935
      End
      Begin nucleo.stCalendar stCalendar1 
         Height          =   3375
         Left            =   0
         TabIndex        =   5
         Top             =   720
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   5953
         BorderStyle     =   1
         ViewHeaderLang  =   3
         cDay            =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblQueDia 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "Descripción"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   30
         Top             =   4440
         Width           =   4935
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   4
         Top             =   480
         Width           =   4935
      End
   End
   Begin VB.Frame FRAME_ALMANAQUE 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   7695
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   9735
      Begin nucleo.stCalendar cal1 
         Height          =   2055
         Index           =   1
         Left            =   2520
         TabIndex        =   9
         Top             =   360
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   3625
         BorderStyle     =   1
         ViewSelCell     =   5
         ViewHeaderLang  =   3
         cDay            =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin nucleo.stCalendar cal1 
         Height          =   2055
         Index           =   2
         Left            =   4920
         TabIndex        =   7
         Top             =   360
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   3625
         BorderStyle     =   1
         ViewSelCell     =   0
         ViewHeaderLang  =   3
         cDay            =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin nucleo.stCalendar cal1 
         Height          =   2055
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   3625
         BorderStyle     =   1
         ViewSelCell     =   0
         ViewHeaderLang  =   3
         cDay            =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin nucleo.stCalendar cal1 
         Height          =   2055
         Index           =   3
         Left            =   7320
         TabIndex        =   12
         Top             =   360
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   3625
         BorderStyle     =   1
         ViewSelCell     =   0
         ViewHeaderLang  =   3
         cDay            =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin nucleo.stCalendar cal1 
         Height          =   2055
         Index           =   4
         Left            =   120
         TabIndex        =   14
         Top             =   2760
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   3625
         BorderStyle     =   1
         ViewSelCell     =   0
         ViewHeaderLang  =   3
         cDay            =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin nucleo.stCalendar cal1 
         Height          =   2055
         Index           =   5
         Left            =   2520
         TabIndex        =   15
         Top             =   2760
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   3625
         BorderStyle     =   1
         ViewSelCell     =   0
         ViewHeaderLang  =   3
         cDay            =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin nucleo.stCalendar cal1 
         Height          =   2055
         Index           =   6
         Left            =   4920
         TabIndex        =   16
         Top             =   2760
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   3625
         BorderStyle     =   1
         ViewSelCell     =   0
         ViewHeaderLang  =   3
         cDay            =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin nucleo.stCalendar cal1 
         Height          =   2055
         Index           =   7
         Left            =   7320
         TabIndex        =   17
         Top             =   2760
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   3625
         BorderStyle     =   1
         ViewSelCell     =   0
         ViewHeaderLang  =   3
         cDay            =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin nucleo.stCalendar cal1 
         Height          =   2055
         Index           =   8
         Left            =   120
         TabIndex        =   18
         Top             =   5160
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   3625
         BorderStyle     =   1
         ViewSelCell     =   0
         ViewHeaderLang  =   3
         cDay            =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin nucleo.stCalendar cal1 
         Height          =   2055
         Index           =   10
         Left            =   4920
         TabIndex        =   19
         Top             =   5160
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   3625
         BorderStyle     =   1
         ViewSelCell     =   0
         ViewHeaderLang  =   3
         cDay            =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin nucleo.stCalendar cal1 
         Height          =   2055
         Index           =   11
         Left            =   7320
         TabIndex        =   20
         Top             =   5160
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   3625
         BorderStyle     =   1
         ViewSelCell     =   0
         ViewHeaderLang  =   3
         cDay            =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin nucleo.stCalendar cal1 
         Height          =   2055
         Index           =   9
         Left            =   2520
         TabIndex        =   21
         Top             =   5160
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   3625
         BorderStyle     =   1
         ViewSelCell     =   0
         ViewHeaderLang  =   3
         cDay            =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ENERO 2008"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   11
         Left            =   7320
         TabIndex        =   29
         Top             =   4920
         Width           =   2295
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ENERO 2008"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   10
         Left            =   4920
         TabIndex        =   28
         Top             =   4920
         Width           =   2295
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ENERO 2008"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   9
         Left            =   2520
         TabIndex        =   27
         Top             =   4920
         Width           =   2295
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ENERO 2008"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   26
         Top             =   4920
         Width           =   2295
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ENERO 2008"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   7
         Left            =   7320
         TabIndex        =   25
         Top             =   2520
         Width           =   2295
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ENERO 2008"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   6
         Left            =   4920
         TabIndex        =   24
         Top             =   2520
         Width           =   2295
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ENERO 2008"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   5
         Left            =   2520
         TabIndex        =   23
         Top             =   2520
         Width           =   2295
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ENERO 2008"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   22
         Top             =   2520
         Width           =   2295
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ENERO 2008"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   7320
         TabIndex        =   13
         Top             =   120
         Width           =   2295
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ENERO 2008"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   4920
         TabIndex        =   11
         Top             =   120
         Width           =   2295
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ENERO 2008"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   2520
         TabIndex        =   8
         Top             =   120
         Width           =   2295
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ENERO 2008"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   2295
      End
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF8080&
      Caption         =   "GRUPO"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Menu events 
      Caption         =   "menu_eventos"
      Visible         =   0   'False
      Begin VB.Menu ver_comentarios 
         Caption         =   "Ver comentarios"
      End
      Begin VB.Menu editar_evento 
         Caption         =   "Finalizar evento..."
      End
      Begin VB.Menu agregar_comentario 
         Caption         =   "Agregar comentario"
      End
   End
End
Attribute VB_Name = "frmSistemaTablero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim marcado
Dim claseSP As New classSignoplast
Dim Grupo As Long
Dim fechaH As Date

Private Sub agregar_comentario_Click()
    If Me.lstEventos.ListItems.count > 0 Then
        frmSistemaTableroAgregarComentario.idTablero = CLng(Me.lstEventos.selectedItem.Tag)
        frmSistemaTableroAgregarComentario.Show 1
        ddia = Me.stCalendar1.cDay
        stCalendar1_DayClicked 1, 1, ddia, False
    End If
End Sub

Private Sub cal1_DayClicked(Index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal iDay As Integer, Cancel As Boolean)
'mostrarEvento cal1(Index).cDay, cal1(Index).cMonth, cal1(Index).cYear
    mess = Me.cal1(Index).cMonth
    anioo = Me.cal1(Index).cYear
    ddia = iDay
    verMes Grupo, mess, anioo, ddia
    stCalendar1_DayClicked 1, 1, ddia, False
    Label3_Click Index
End Sub


Private Sub cboGrupos_Click()
    Command1_Click
End Sub

Private Sub cmdAgregarEvento_Click()
    frmSistemaTableroAgregarEvento.FechaEvento = fechaH
    frmSistemaTableroAgregarEvento.GrupoEvento = Grupo
    frmSistemaTableroAgregarEvento.idEvento = -1  '-1 indica que es un evento nuevo
    frmSistemaTableroAgregarEvento.Show 1

    verSeleccionTablero (Grupo)
    stCalendar1_DayClicked 1, 1, Day(fechaH), False
    verEventos Month(fechaH), Year(fechaH), Grupo
    stCalendar1.CalendarRedraw
End Sub
Private Sub cmdModificarEvento_Click()
    idEvento = 12000000
    frmSistemaTableroAgregarEvento.FechaEvento = fechaH
    frmSistemaTableroAgregarEvento.idEvento = idEvento  '-1 indica que es un evento nuevo
    frmSistemaTableroAgregarEvento.Show 1

End Sub

Private Sub Command1_Click()
    For x = 0 To 11
        Me.Label3(x).ForeColor = vbBlack
    Next
    Grupo = Me.cboGrupos.ItemData(Me.cboGrupos.ListIndex)
    verTablero Grupo, Empty, Empty
End Sub

Private Sub editar_evento_Click()
    Dim s As New Recordset
    If Me.lstEventos.ListItems.count > 0 Then
        If MsgBox("¿Está seguro de finalizar este evento?", vbYesNo, "Confirmación") = vbYes Then
            idt = CLng(Me.lstEventos.selectedItem.Tag)
            Set s = conectar.RSFactory("Select estado from usuariosTablero where id=" & idt)
            If Not s.EOF And Not s.BOF Then
                If s!estado = 1 Then
                    MsgBox "Este evento se encuentra finalizado!", vbExclamation, "Información"
                ElseIf s!estado = 0 Then
                    claseSP.ejecutarComando "update usuariosTablero set estado=1 where id= " & idt
                Else
                    MsgBox "Se produjo un error. Todos lso cambios abortados!", vbCritical, "Error"
                End If


            End If
        End If
    End If
    ddia = Me.stCalendar1.cDay
    stCalendar1_DayClicked 1, 1, ddia, False
End Sub

Private Sub Form_Load()
    FormHelper.Customize Me
    'lleno la lista de grupos
    Dim grupoDef As Long
    claseSP.llenarComboRubrosGrupos Me.cboGrupos, funciones.getUser


    Grupo = Permisos.sistemaGrupoDefault


    If Grupo > 0 Then    'si tiene definido un grupo como default
        Me.cboGrupos.ListIndex = funciones.PosIndexCbo(Grupo, Me.cboGrupos)
    Else
        MsgBox "No existe ningun grupo por default!", vbInformation, "Información"
    End If

    Me.Left = 0
    Me.Top = 0

    usuario = funciones.getUser
    'lleno la lista de grupos que puede ver

    'armo la grilla de calendarios
    'mes actual
    DIAACTUAL = Day(Now)
    MESACTUAL = Month(Now)
    ANIOACTUAL = Year(Now)
    Me.Label3(1).caption = mesAnio(MESACTUAL, ANIOACTUAL)
    Me.cal1(1).cMonth = MESACTUAL
    Me.cal1(1).cYear = ANIOACTUAL
    Me.cal1(1).cDay = DIAACTUAL

    'mes anterior
    calcularMesDesplazado Now, 1, True, anioNuevo, mesnuevo
    mesAnterior = mesnuevo
    ANIOAnterior = anioNuevo
    Me.Label3(0).caption = mesAnio(mesAnterior, ANIOAnterior)
    Me.cal1(0).cMonth = mesAnterior
    Me.cal1(0).cYear = ANIOAnterior

    For x = 2 To 11
        calcularMesDesplazado Now, x - 1, False, A, m
        Me.Label3(x).caption = mesAnio(m, A)
        Me.cal1(x).cMonth = m
        Me.cal1(x).cYear = A
    Next x
    verTablero Grupo, Empty, Empty    'muestra el mes en curso al comienzo (onload)
End Sub
Private Sub verTablero(Grupo, mes, anio)
    verSeleccionTablero (Grupo)
    verMes Grupo, mes, anio
End Sub
Private Sub verSeleccionTablero(Grupo)
    Dim rs As Recordset
    For i = 0 To 11    'recorro los 12 meses
        For x = 0 To Me.cal1(i).DayCount  'borro las marcas que haya
            Me.cal1(i).DayMarking x, 0, False
            Me.cal1(i).DayMarking x, 1, False
            Me.cal1(i).DayMarking x, 2, False
            Me.cal1(i).DayMarking i, 3, False
        Next x
        anio = cal1(i).cYear
        mes = cal1(i).cMonth
        strsql = "Select tipo,FechaDia from usuariosTablero where FechaMes=" & mes & " and  FechaAnio=" & anio & " and grupoUsuarios=" & Grupo
        Set rs = conectar.RSFactory(strsql)
        While Not rs.EOF
            Me.cal1(i).DayMarking rs!FechaDia, rs!Tipo, True
            rs.MoveNext
        Wend
        Me.cal1(i).CalendarRedraw
    Next
End Sub
Private Sub verMes(Grupo, Optional mes = Empty, Optional anio = Empty, Optional dia = Empty)
    If mes = Empty And anio = Empty Then
        'ESTO ES PARA Q MUESTRE TODO APENAS ABRE EL TABLERO
        Me.stCalendar1.cDay = Day(Now)
        Me.stCalendar1.cMonth = Month(Now)
        Me.stCalendar1.cYear = Year(Now)
        mes = stCalendar1.cMonth
        anio = stCalendar1.cYear
        verEventos mes, anio, Grupo
        Me.Label1.caption = mesAnio(Me.stCalendar1.cMonth, Me.stCalendar1.cYear)
        fechaH = CDate(Format(stCalendar1.cDay) & "/" & Format(stCalendar1.cMonth) & "/" & Format(stCalendar1.cYear))
        Me.cmdAgregarEvento.caption = funciones.FEcha(fechaH)
        Label3(1).ForeColor = vbRed
        verEventos mes, anio, Grupo
        ddia = Me.stCalendar1.cDay
        stCalendar1_DayClicked 1, 1, ddia, False
    Else
        'mes = stCalendar1.cMonth
        'anio = stCalendar1.cYear
        'Me.stCalendar1.cDay = Day()
        Me.stCalendar1.cMonth = mes
        Me.stCalendar1.cYear = anio

        If dia <> Empty Then
            Me.stCalendar1.cDay = dia
        End If
        verEventos mes, anio, Grupo
        Me.Label1.caption = mesAnio(mes, anio)

    End If
    Me.stCalendar1.CalendarRedraw
End Sub

Public Sub mostrarEvento(dia, mes, anio)
    MsgBox "evento para " & dia & "-" & mes & "-" & anio
End Sub
Private Sub calcularMesDesplazado(actual As Date, desplazamiento, atras As Boolean, ByRef anioDesplazado, ByRef mesDesplazado)
    MsgBox ("Función Desactivada")

    '    desplazado = desplazamiento
    '    If atras Then desplazado = desplazamiento * -1
    '    Dim m As Integer
    '    m = DateAdd("m", desplazado, actual)
    '    mesDesplazado = Month(m)    'a
    '    anioDesplazado = Year(m)    'anio
End Sub
Private Function mesAnio(mesA, anio) As String
    Dim mes(1 To 12) As String
    mes(1) = "ENE"
    mes(2) = "FEB"
    mes(3) = "MAR"
    mes(4) = "ABR"
    mes(5) = "MAY"
    mes(6) = "JUN"
    mes(7) = "JUL"
    mes(8) = "AGO"
    mes(9) = "SEP"
    mes(10) = "OCT"
    mes(11) = "NOV"
    mes(12) = "DIC"

    mesAnio = mes(mesA) & " " & anio
End Function

Private Sub verEventos(mes, anio, Grupo)
    Dim rs As Recordset

    'muestro los eventos del mes seleccionado
    'borro todsas las marcas
    For i = 1 To Me.stCalendar1.DayCount
        Me.stCalendar1.DayMarking i, 0, False
        Me.stCalendar1.DayMarking i, 1, False
        Me.stCalendar1.DayMarking i, 2, False
        Me.stCalendar1.DayMarking i, 3, False
    Next i
    'marco este mes
    strsql = "Select tipo,FechaDia from usuariosTablero where FechaMes=" & mes & " and  FechaAnio=" & anio & " and grupoUsuarios=" & Grupo
    Set rs = conectar.RSFactory(strsql)
    While Not rs.EOF
        Me.stCalendar1.DayMarking rs!FechaDia, rs!Tipo, True
        rs.MoveNext
    Wend

End Sub

Private Sub Form_Unload(Cancel As Integer)
'    frmPrincipal.SmartMenuXP1.MenuItems.value(131) = smiUnchecked
End Sub



Private Sub Label3_Click(Index As Integer)
    mess = Me.cal1(Index).cMonth
    anioo = Me.cal1(Index).cYear
    verMes Grupo, mess, anioo
    A = vbBlack
    For i = 0 To 11
        Label3(i).ForeColor = A
    Next
    Label3(Index).ForeColor = vbRed
End Sub

Private Sub lstEventos_ItemClick(ByVal item As MSComctlLib.ListItem)
    Me.txtVer = item
End Sub

Private Sub lstEventos_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim rs As New Recordset
    If Me.lstEventos.ListItems.count > 0 Then


        If Button = 2 Then


            idEvento = Me.lstEventos.selectedItem.Tag
            Set rs = conectar.RSFactory("select estado,idUsuario, eventoModificable from usuariosTablero where id=" & idEvento)
            If Not rs.EOF And Not rs.BOF Then
                estado = rs!estado
                idUsuario = rs!idUsuario
                Modificable = rs!eventoModificable
            End If

            If Modificable = 1 Then  'si se puede modificar, activo la opcion
                Me.agregar_comentario.Enabled = True
                Me.editar_evento.Enabled = True
            Else
                If idUsuario = funciones.getUser Then    'si no se puede modificar, solo lo hará el usuario q lo creo
                    Me.agregar_comentario.Enabled = True
                    Me.editar_evento.Enabled = True
                Else
                    Me.agregar_comentario.Enabled = True
                    Me.editar_evento.Enabled = False
                End If
            End If

            If estado = 1 Then
                Me.editar_evento.Enabled = False
                Me.agregar_comentario.Enabled = False
            End If
            Me.PopupMenu Me.events
        End If
    End If
End Sub
Private Sub stCalendar1_DayClicked(ByVal Button As Integer, ByVal Shift As Integer, ByVal iDay As Integer, Cancel As Boolean)
    fechaH = CDate(Format(iDay) & "/" & Format(stCalendar1.cMonth) & "/" & Format(stCalendar1.cYear))
    Me.cmdAgregarEvento.caption = funciones.FEcha(fechaH)
    Me.cmdAgregarEvento.Tag = fechaH
    diah = iDay
    mesh = Me.stCalendar1.cMonth
    anioh = Me.stCalendar1.cYear
    Me.lstEventos.ListItems.Clear
    Dim r As Recordset

    Set r = conectar.RSFactory("select distinct(select count(0) from usuariosTableroComentarios where idTablero=ut.id)as cantidad,idUsuario,id,titulo,estado,fechaCreado from usuariosTablero ut where fechaMes=" & mesh & " and fechaDia=" & diah & " and fechaAnio=" & anioh & " and GrupoUsuarios=" & Grupo)
    c = 0
    While Not r.EOF
        Dim x As ListItem
        Set x = Me.lstEventos.ListItems.Add(, , r!titulo & "  (" & r!Cantidad & ")")
        x.SubItems(1) = claseSP.queUsuario(r!idUsuario)

        est = r!estado
        If est = 0 Then
            estado = "En curso"
        ElseIf est = 1 Then
            estado = "Finalizado"
        End If
        x.SubItems(2) = estado
        x.SubItems(3) = r!fechaCreado
        x.Tag = r!Id
        If marcado = x Then
            x.Selected = True
            x.EnsureVisible
        End If
        r.MoveNext
        c = c + 1
    Wend
    If c > 0 Then
        Me.txtVer = Me.lstEventos.selectedItem
    Else
        Me.txtVer = Empty
    End If

End Sub

Private Sub Timer1_Timer()

    verSeleccionTablero (Grupo)
    stCalendar1_DayClicked 1, 1, Day(fechaH), False
    verEventos Month(fechaH), Year(fechaH), Grupo

    stCalendar1.CalendarRedraw

End Sub

Private Sub txtVer_Change()
    If Me.lstEventos.ListItems.count > 0 Then marcado = Me.lstEventos.selectedItem
End Sub

Private Sub ver_comentarios_Click()
    If Me.lstEventos.ListItems.count > 0 Then
        frmSistemaTableroVerComentarios.idEvento = CLng(Me.lstEventos.selectedItem.Tag)
        frmSistemaTableroVerComentarios.Show 1
    End If
End Sub
