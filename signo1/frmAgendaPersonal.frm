VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUsuariosAgendaPersonal 
   BackColor       =   &H00FF8080&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Agenda Personal..."
   ClientHeight    =   5175
   ClientLeft      =   810
   ClientTop       =   930
   ClientWidth     =   11625
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   11625
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbMonth 
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   120
      Width           =   3375
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FF8080&
      Caption         =   "[ Nuevo Evento ]"
      Height          =   2175
      Left            =   4680
      TabIndex        =   2
      Top             =   2880
      Width           =   6735
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Agregar"
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   840
         TabIndex        =   7
         Top             =   1080
         Width           =   5655
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmAgendaPersonal.frx":0000
         Left            =   840
         List            =   "frmAgendaPersonal.frx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label lblFechaNuevoEvento 
         BackColor       =   &H00FF8080&
         Height          =   255
         Left            =   840
         TabIndex        =   12
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FF8080&
         Caption         =   "Fecha"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FF8080&
         Caption         =   "Evento"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF8080&
         Caption         =   "Color"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Cancel          =   -1  'True
      Caption         =   "Salir"
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   375
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4680
      Width           =   975
   End
   Begin VB.ComboBox cmbYear 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   1080
   End
   Begin nucleo.stCalendar stCalendar1 
      Height          =   3975
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   7011
      BorderStyle     =   1
      ViewHeaderCell  =   5
      ViewSelCell     =   5
      ViewEmptyCell   =   4
      DayCount        =   30
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ListView lstEventos 
      Height          =   2535
      Left            =   4680
      TabIndex        =   5
      Top             =   120
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   4471
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nro."
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Color"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Evento"
         Object.Width           =   7761
      EndProperty
   End
   Begin VB.Menu eventos 
      Caption         =   "menuEvent"
      Visible         =   0   'False
      Begin VB.Menu eventnumber 
         Caption         =   "[ Evento ]"
         Enabled         =   0   'False
      End
      Begin VB.Menu delEvent 
         Caption         =   "Eliminar Evento..."
      End
   End
End
Attribute VB_Name = "frmUsuariosAgendaPersonal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As Recordset
Dim diahoy As Integer
'Dim clsPE As New ClassPersonal
Dim fechaHoy
Private Sub cmbMonth_Click()
    stCalendar1.cMonth = cmbMonth.ListIndex + 1
    clsPE.marcarCalendario Me.stCalendar1, cmbMonth.ListIndex + 1
    Me.lblFechaNuevoEvento.caption = Format(stCalendar1.cDay) & "/" & Format(stCalendar1.cMonth) & "/" & Format(stCalendar1.cYear)
    verEventos
End Sub

Private Sub cmbYear_Click()
    stCalendar1.cYear = cmbYear.list(cmbYear.ListIndex)
    clsPE.marcarCalendario Me.stCalendar1, cmbMonth.ListIndex + 1
    Me.lblFechaNuevoEvento.caption = Format(stCalendar1.cDay) & "/" & Format(stCalendar1.cMonth) & "/" & Format(stCalendar1.cYear)
    verEventos
End Sub
Private Sub verEventos()

    Me.lstEventos.ListItems.Clear
    Dim x As ListItem
    Set rs = conectar.RSFactory("select id, tipo, memo from usuariosAgenda where month(fecha)=" & Me.stCalendar1.cMonth & " and day(fecha)=" & diahoy & " and idUsuario=" & funciones.getUser)
    While Not rs.EOF
        Set x = Me.lstEventos.ListItems.Add(, , Format(rs!id, "0000"))

        If rs!Tipo = 0 Then
            Tipo = "Violeta"
        ElseIf rs!Tipo = 1 Then
            Tipo = "Rojo"
        ElseIf rs!Tipo = 2 Then
            Tipo = "Fucsia"
        ElseIf rs!Tipo = 3 Then
            Tipo = "Verde"
        End If
        x.SubItems(1) = Tipo
        x.SubItems(2) = rs!Memo
        rs.MoveNext
    Wend
    llenarCalendario


End Sub

Private Sub Command1_Click()
    If MsgBox("¿Agregar evento?", vbYesNo, "Confirmación") = vbYes Then
        If Trim(Me.Text1) <> Empty Then
            fechaHoyNuevo = Format(CDate(Format(stCalendar1.cDay) & "/" & Format(stCalendar1.cMonth) & "/" & Format(stCalendar1.cYear)), "YYYY/MM/DD")
            Tipo = Me.Combo1.ItemData(Me.Combo1.ListIndex)
            Memo = normalizaVieja(Me.Text1)
            clsPE.ejecutar "insert into usuariosAgenda (idUsuario, fecha, tipo, memo) values (" & funciones.getUser & ",'" & fechaHoyNuevo & "'," & Tipo & ",'" & Memo & "')"
        End If
        clsPE.marcarCalendario Me.stCalendar1, cmbMonth.ListIndex + 1
    End If
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub delEvent_Click()
    If MsgBox("¿Desea eliminar el evento?", vbYesNo, "Confirmación") = vbYes Then
        clsPE.ejecutar ("delete from usuariosAgenda where id=" & CLng(Me.lstEventos.selectedItem))
        verEventos


    End If

End Sub
Private Sub Form_Load()
    FormHelper.Customize Me


    Dim emple As clsUsuario
    Set emple = DAOUsuarios.GetById(funciones.getUser)


    Me.caption = "[ " & emple.Empleado.NombreCompleto & " ]"
    Dim i As Long
    Dim Yr As Long

    Yr = Val(Format(Now, "yyyy"))
    For i = Yr - 6 To Yr + 6
        cmbYear.AddItem str(i)
    Next i
    cmbYear.ListIndex = 6

    For i = 1 To 12
        cmbMonth.AddItem Format(CDate("1/" & i & "/1999"), "mmmm")
    Next i
    cmbMonth.ListIndex = Month(Now) - 1

    Me.lblFechaNuevoEvento.caption = Format(stCalendar1.cDay) & "/" & Format(stCalendar1.cMonth) & "/" & Format(stCalendar1.cYear)

    Me.Combo1.ListIndex = 0
    verEventos
    Me.stCalendar1.CalendarRedraw

End Sub

Private Sub lstEventos_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Me.lstEventos.ListItems.count > 0 Then
        If Button = 2 Then
            Me.PopupMenu eventos
            Me.eventnumber.caption = "[ " & Format(Me.lstEventos.selectedItem, "0000") & " ]"
        End If
    End If
End Sub

Private Sub stCalendar1_DayClicked(ByVal Button As Integer, ByVal Shift As Integer, ByVal iDay As Integer, Cancel As Boolean)
    fechaHoy = CDate(Format(iDay) & "/" & Format(stCalendar1.cMonth) & "/" & Format(stCalendar1.cYear))
    diahoy = iDay
    Me.lblFechaNuevoEvento.caption = fechaHoy

    verEventos

End Sub



Private Sub llenarCalendario()

    mes = Me.stCalendar1.cMonth
    anio = Me.stCalendar1.cYear
    dia = Me.stCalendar1.cDay
    clsPE.ejecutar "select day(fecha) as dia,tipo from usuariosAgenda where month(fecha) = " & mes & " and year(fecha) = " & anio & " and idUsuario=" & funciones.getUser & " order by fecha desc"
    While Not rs.EOF
        Me.stCalendar1.DayMarking rs!dia, rs!Tipo, True

        rs.MoveNext
    Wend

End Sub
