VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmEventos 
   Caption         =   "Eventos"
   ClientHeight    =   6315
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8820
   Icon            =   "frmEventos.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6315
   ScaleWidth      =   8820
   Begin XtremeSuiteControls.PushButton btnQuitarUsuario 
      Height          =   285
      Left            =   8370
      TabIndex        =   13
      Top             =   270
      Width           =   255
      _Version        =   786432
      _ExtentX        =   450
      _ExtentY        =   503
      _StockProps     =   79
      Caption         =   "X"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton btnMarcarNoLeidosComoLeidos 
      Height          =   405
      Left            =   5625
      TabIndex        =   1
      Top             =   5865
      Width           =   3135
      _Version        =   786432
      _ExtentX        =   5530
      _ExtentY        =   714
      _StockProps     =   79
      Caption         =   "Marcar seleccionados como leídos"
      UseVisualStyle  =   -1  'True
   End
   Begin VB.Timer tmrMarcarLeido 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   5850
      Top             =   1290
   End
   Begin GridEX20.GridEX grilla 
      Height          =   4620
      Left            =   75
      TabIndex        =   0
      Top             =   1185
      Width           =   8670
      _ExtentX        =   15293
      _ExtentY        =   8149
      Version         =   "2.0"
      PreviewRowIndent=   200
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      PreviewColumn   =   "descripcion"
      PreviewRowLines =   1
      MultiSelect     =   -1  'True
      MethodHoldFields=   -1  'True
      AllowEdit       =   0   'False
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   5
      Column(1)       =   "frmEventos.frx":000C
      Column(2)       =   "frmEventos.frx":014C
      Column(3)       =   "frmEventos.frx":0268
      Column(4)       =   "frmEventos.frx":0388
      Column(5)       =   "frmEventos.frx":048C
      FormatStylesCount=   7
      FormatStyle(1)  =   "frmEventos.frx":05B8
      FormatStyle(2)  =   "frmEventos.frx":06F0
      FormatStyle(3)  =   "frmEventos.frx":07A0
      FormatStyle(4)  =   "frmEventos.frx":0854
      FormatStyle(5)  =   "frmEventos.frx":092C
      FormatStyle(6)  =   "frmEventos.frx":09E4
      FormatStyle(7)  =   "frmEventos.frx":0AC4
      ImageCount      =   0
      PrinterProperties=   "frmEventos.frx":0C0C
   End
   Begin XtremeSuiteControls.GroupBox grpFecha 
      Height          =   1035
      Left            =   90
      TabIndex        =   2
      Top             =   75
      Width           =   4560
      _Version        =   786432
      _ExtentX        =   8043
      _ExtentY        =   1826
      _StockProps     =   79
      Caption         =   "Fecha de Evento"
      BackColor       =   12632256
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.ComboBox cboRangos 
         Height          =   315
         Left            =   795
         TabIndex        =   5
         Top             =   210
         Width           =   3645
         _Version        =   786432
         _ExtentX        =   6429
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.DateTimePicker dtpHasta 
         Height          =   315
         Left            =   2970
         TabIndex        =   4
         Top             =   585
         Width           =   1470
         _Version        =   786432
         _ExtentX        =   2593
         _ExtentY        =   556
         _StockProps     =   68
         CheckBox        =   -1  'True
         Format          =   1
      End
      Begin XtremeSuiteControls.DateTimePicker dtpDesde 
         Height          =   315
         Left            =   795
         TabIndex        =   3
         Top             =   600
         Width           =   1470
         _Version        =   786432
         _ExtentX        =   2593
         _ExtentY        =   556
         _StockProps     =   68
         CheckBox        =   -1  'True
         Format          =   1
      End
      Begin XtremeSuiteControls.Label Label4 
         Height          =   195
         Left            =   225
         TabIndex        =   8
         Top             =   255
         Width           =   480
         _Version        =   786432
         _ExtentX        =   847
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Rango"
         BackColor       =   12632256
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label5 
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   630
         Width           =   465
         _Version        =   786432
         _ExtentX        =   820
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Desde"
         BackColor       =   12632256
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label6 
         Height          =   195
         Left            =   2415
         TabIndex        =   6
         Top             =   645
         Width           =   420
         _Version        =   786432
         _ExtentX        =   741
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Hasta"
         BackColor       =   12632256
         AutoSize        =   -1  'True
      End
   End
   Begin XtremeSuiteControls.ComboBox cboUsuarios 
      Height          =   315
      Left            =   5460
      TabIndex        =   10
      Top             =   255
      Width           =   2880
      _Version        =   786432
      _ExtentX        =   5080
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Sorted          =   -1  'True
      Style           =   2
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboTipoEvento 
      Height          =   315
      Left            =   5805
      TabIndex        =   12
      Top             =   645
      Width           =   2535
      _Version        =   786432
      _ExtentX        =   4471
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Sorted          =   -1  'True
      Style           =   2
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.PushButton btnQuitarTipoEvento 
      Height          =   285
      Left            =   8370
      TabIndex        =   14
      Top             =   660
      Width           =   255
      _Version        =   786432
      _ExtentX        =   450
      _ExtentY        =   503
      _StockProps     =   79
      Caption         =   "X"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   195
      Left            =   4800
      TabIndex        =   11
      Top             =   690
      Width           =   915
      _Version        =   786432
      _ExtentX        =   1614
      _ExtentY        =   344
      _StockProps     =   79
      Caption         =   "Tipo Evento:"
      Alignment       =   1
      AutoSize        =   -1  'True
   End
   Begin XtremeSuiteControls.Label lblUsuario 
      Height          =   195
      Left            =   4800
      TabIndex        =   9
      Top             =   300
      Width           =   585
      _Version        =   786432
      _ExtentX        =   1032
      _ExtentY        =   344
      _StockProps     =   79
      Caption         =   "Usuario:"
      Alignment       =   1
      AutoSize        =   -1  'True
   End
End
Attribute VB_Name = "frmEventos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private eventos As Collection
Private EVENTO As EVENTO

Private formLoaded As Boolean

Private Sub btnMarcarNoLeidosComoLeidos_Click()
    Dim ev As EVENTO
    Dim it As JSSelectedItem
    For Each it In Me.grilla.SelectedItems
        If it.RowIndex > 0 Then
            Set ev = eventos.item(it.RowIndex)

            If Not funciones.BuscarEnColeccion(ev.Lecturas, CStr(funciones.GetUserObj.Id)) Then
                If DAOEvento.Read(ev.Id) Then
                    Dim lectura As New LecturaEvento
                    lectura.FechaLectura = Now
                    lectura.idUsuario = funciones.GetUserObj.Id
                    ev.Lecturas.Add lectura, CStr(funciones.GetUserObj.Id)
                End If
            End If
        End If
    Next it

    Me.grilla.row = -1
    UpdateCaption
    Me.grilla.ReBind
End Sub

Private Sub btnQuitarTipoEvento_Click()
    Me.cboTipoEvento.ListIndex = -1
End Sub

Private Sub btnQuitarUsuario_Click()
    Me.cboUsuarios.ListIndex = -1
End Sub

Private Sub cboRangos_Click()
    funciones.CalculateDateRange Me.cboRangos, Me.dtpDesde, Me.dtpHasta
    llenar
End Sub

Private Sub cboTipoEvento_Click()
    llenar
End Sub

Private Sub cboUsuarios_Click()
    llenar
End Sub

Private Sub dtpDesde_Change()
    llenar
End Sub

Private Sub dtpHasta_Change()
    llenar
End Sub

Private Sub Form_Load()
    Customize Me
    GridEXHelper.CustomizeGrid Me.grilla, True
    Me.grilla.ItemCount = 0

    Me.cboUsuarios.Clear
    Dim usuario As clsUsuario
    Dim nameToShow As String
    For Each usuario In DAOUsuarios.FindAll()
        nameToShow = usuario.usuario
        If IsSomething(usuario.Empleado) Then nameToShow = nameToShow & " (" & usuario.Empleado.NombreCompleto & ")"
        Me.cboUsuarios.AddItem nameToShow
        Me.cboUsuarios.ItemData(Me.cboUsuarios.NewIndex) = usuario.Id
    Next usuario
    Me.cboUsuarios.ListIndex = -1


    Me.cboTipoEvento.Clear
    Dim tipoEvento As Collection
    For Each tipoEvento In DAOEvento.GetEventBroadCastTypes()
        Me.cboTipoEvento.AddItem tipoEvento(2)
        Me.cboTipoEvento.ItemData(Me.cboTipoEvento.NewIndex) = tipoEvento(1)
    Next tipoEvento
    Me.cboTipoEvento.ListIndex = -1

    funciones.FillComboBoxDateRanges Me.cboRangos
    formLoaded = True
End Sub

Public Sub llenar()
    If formLoaded Then
        Dim F As String: F = "1 = 1"
        Dim tmpFecha As Date
        If Me.cboUsuarios.ListIndex <> -1 Then
            F = F & " AND e.id_usuario_involucrado = " & Me.cboUsuarios.ItemData(Me.cboUsuarios.ListIndex)
        End If

        If Me.cboTipoEvento.ListIndex <> -1 Then
            F = F & " AND e.id_tipo_evento = " & Me.cboTipoEvento.ItemData(Me.cboTipoEvento.ListIndex)
        End If

        If Not IsNull(Me.dtpDesde.value) Then
            tmpFecha = CDate(Fix(CDbl(Me.dtpDesde.value)))
            F = F & " AND e.fecha_creacion >= " & conectar.Escape(tmpFecha)    'solo parte de fecha, no numerica
        End If

        If Not IsNull(Me.dtpHasta.value) Then
            tmpFecha = CDate(Fix(CDbl(Me.dtpHasta.value)))
            tmpFecha = tmpFecha + TimeSerial(23, 59, 59)    'para que llegue al final del dia

            F = F & " AND e.fecha_creacion <= " & conectar.Escape(tmpFecha)
        End If

        Set eventos = DAOEvento.FindAllByUser(funciones.GetUserObj.Id, False, F)
        Me.grilla.ItemCount = 0
        Me.grilla.ItemCount = eventos.count
        UpdateCaption
        Me.grilla.row = -1
    End If
End Sub

Private Sub UpdateCaption()
    Dim leidos As Long
    Dim noLeidos As Long
    Dim ev As EVENTO
    For Each ev In eventos
        If funciones.BuscarEnColeccion(ev.Lecturas, CStr(funciones.GetUserObj.Id)) Then
            leidos = leidos + 1
        Else
            noLeidos = noLeidos + 1
        End If
    Next ev

    Me.caption = "Eventos (" & eventos.count & ") [Leidos (" & leidos & ") | No Leidos (" & noLeidos & ")]"
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Me.grilla.Width = Me.ScaleWidth - 150
    Me.grilla.Height = Me.ScaleHeight - Me.grilla.Top - 500
    Me.btnMarcarNoLeidosComoLeidos.Left = Me.ScaleWidth - Me.btnMarcarNoLeidosComoLeidos.Width - 80
    Me.btnMarcarNoLeidosComoLeidos.Top = Me.ScaleHeight - Me.btnMarcarNoLeidosComoLeidos.Height - 50
End Sub

Private Sub grilla_RowFormat(RowBuffer As GridEX20.JSRowData)
    If RowBuffer.RowIndex > 0 And RowBuffer.RowIndex <= eventos.count And eventos.count > 0 Then
        Set EVENTO = eventos.item(RowBuffer.RowIndex)

        If Not funciones.BuscarEnColeccion(EVENTO.Lecturas, CStr(funciones.GetUserObj.Id)) Then
            RowBuffer.RowStyle = "noleido"
        End If
    End If
End Sub

Private Sub grilla_SelectionChange()
    Me.tmrMarcarLeido.Enabled = False
    If Me.grilla.RowIndex(Me.grilla.row) > 0 And eventos.count > 0 Then
        Set EVENTO = eventos.item(Me.grilla.RowIndex(Me.grilla.row))
        Me.tmrMarcarLeido.Enabled = Not funciones.BuscarEnColeccion(EVENTO.Lecturas, CStr(funciones.GetUserObj.Id))
    End If
End Sub

Private Sub grilla_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex > 0 And RowIndex <= eventos.count And eventos.count > 0 Then
        Set EVENTO = eventos.item(RowIndex)
        Values(1) = EVENTO.FechaCreacion
        If IsSomething(EVENTO.UsuarioInvolucrado.Empleado) Then
            Values(2) = EVENTO.UsuarioInvolucrado.Empleado.NombreCompleto
        Else
            Values(2) = EVENTO.UsuarioInvolucrado.usuario
        End If
        Values(3) = EVENTO.descripcion
        If EVENTO.tipoEvento > 0 Then
            Values(4) = DAOEvento.GetEventBroadCastTypes(CStr(EVENTO.tipoEvento))(2)    'optimizar....
        End If
        If EVENTO.Lecturas.count > 0 Then
            Values(5) = EVENTO.Lecturas.item(1).FechaLectura
        End If

    End If
End Sub

Private Sub tmrMarcarLeido_Timer()
    If IsSomething(EVENTO) Then
        If Not funciones.BuscarEnColeccion(EVENTO.Lecturas, CStr(funciones.GetUserObj.Id)) Then
            If DAOEvento.Read(EVENTO.Id) Then
                Dim lectura As New LecturaEvento
                lectura.FechaLectura = Now
                lectura.idUsuario = funciones.GetUserObj.Id
                EVENTO.Lecturas.Add lectura, CStr(funciones.GetUserObj.Id)
                Dim row As Long
                row = Me.grilla.row
                Me.grilla.ReBind
                Me.grilla.row = row

                UpdateCaption
            End If
        End If
    End If
End Sub
