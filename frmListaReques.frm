VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmComprasRequesLista 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Requerimientos"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10365
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   10365
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   6015
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   10610
      _Version        =   393216
      TabOrientation  =   3
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Todos"
      TabPicture(0)   =   "frmListaReques.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lstTodos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "En proceso"
      TabPicture(1)   =   "frmListaReques.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Pendientes"
      TabPicture(2)   =   "frmListaReques.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "Procesados"
      TabPicture(3)   =   "frmListaReques.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      Begin MSComctlLib.ListView lstTodos 
         Height          =   5655
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   9975
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Número"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Sector"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Destino"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Fecha "
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Estado"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Usuario"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Default         =   -1  'True
      Height          =   375
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ Listado completo ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10335
   End
   Begin VB.Menu menu_1 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu nroReq 
         Caption         =   "[ REQ ]"
         Enabled         =   0   'False
      End
      Begin VB.Menu editReq 
         Caption         =   "Editar..."
      End
      Begin VB.Menu aprobar 
         Caption         =   "Aprobar..."
      End
      Begin VB.Menu df 
         Caption         =   "-"
      End
      Begin VB.Menu requeDeta 
         Caption         =   "Ver Detalles.."
      End
      Begin VB.Menu verHistorial 
         Caption         =   "Ver Historial"
      End
   End
End
Attribute VB_Name = "frmComprasRequesLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim classSP As New classSignoplast
Dim pos As Long
Dim rs As Recordset
Dim classC As New classCompras

Private Sub definir_Click()
'frmRequeCompraDefinir.idr = CLng(Me.lstReques.SelectedItem)
'frmRequeCompraDefinir.Show

End Sub

Private Sub aprobar_Click()
If MsgBox("¿Seguro de aprobar este requerimiento?", vbYesNo, "Confirmación") = vbYes Then
'    If classC.aprobarReque(CLng(Me.lstReques.SelectedItem)) Then
'        MsgBox "Requerimiento aprobado con éxito!", vbInformation, "Información"
'    End If
End If
    
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub editReq_Click()
If Me.lstTodos.ListItems.count > 0 Then
idReque = CLng(Me.lstTodos.SelectedItem)
frmComprasRequesNuevo.accion = idReque
frmComprasRequesNuevo.Show
End If
frmComprasRequesNuevo.Refresh
Me.Refresh
Me.lstTodos.Refresh
End Sub



Private Sub Form_Load()
llenarLSTTodos
'llenarPendientes
'llenarLSTPROCESO
'llenarProcesados
End Sub
Private Sub llenarLSTTodos()
On Error GoTo err44
    Set rs = classC.CrearRS("select r.idUsuarioCreador,r.id,s.sector,r.idPedido,r.fechaCreado,r.estado from sp.ComprasRequerimientos r inner join sp.sectores s on r.idSector=s.id ")
Me.lstTodos.ListItems.Clear
Dim x As ListItem
While Not rs.EOF
Set x = Me.lstTodos.ListItems.Add(, , Format(rs!id, "0000"))
    x.SubItems(1) = rs!sector
    If rs!idpedido < 1 Then para = "Stock" Else para = Format(rs!idpedido, "0000")
    x.SubItems(2) = para
    x.SubItems(3) = rs!FechaCreado
    Dim es As Integer
    es = rs!estado
    x.SubItems(4) = funciones.estado_reque(es)
    x.SubItems(5) = classSP.queUsuario(rs!idUsuarioCreador)
    If x = pos Then
        x.Selected = True
        x.EnsureVisible
    End If
    If es = 0 Then
     x.ListSubItems(4).ForeColor = vbMagenta
    ElseIf es = 1 Then
     x.ListSubItems(4).ForeColor = vbBlue
    End If
    rs.MoveNext


Wend
Set rs = Nothing

Exit Sub
err44:
MsgBox Err.Description
End Sub
Private Sub lstReques_ItemClick(ByVal Item As MSComctlLib.ListItem)
pos = Item
End Sub
Private Sub lstTodos_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Me.lstTodos.ListItems.count > 0 Then
idr = Me.lstTodos.SelectedItem

    If Button = 2 Then
        Dim r As Recordset
        
        Set r = classC.CrearRS("select estado from ComprasRequerimientos where id=" & idr)
        If Not r.EOF And Not r.BOF Then est = CInt(r!estado)
        
        Me.nroReq.Caption = "[ Nro. " & idr & " ]"
        If est = 0 Then 'en proceso
         Me.editReq.Enabled = True
         If permisos.AprobadorReques Then
            Me.aprobar = True
         Else
            Me.aprobar = False
         End If
        ElseIf est = 1 Then 'finalizado
         Me.editReq.Enabled = False
         Me.aprobar = False
        End If
        Me.PopupMenu menu_1
    End If
End If
End Sub
Private Sub requeDeta_Click()
frmComprasRequeDetalles.idReque = Me.lstTodos.SelectedItem
frmComprasRequeDetalles.Show
End Sub

Private Sub verHistorial_Click()
'If Me.lstReques.ListItems.count > 0 Then
'idReque = CLng(Me.lstReques.SelectedItem)
'frmComprasRequesHistorial.idReque = idReque
''frmComprasRequesHistorial.Show
'End If
End Sub



Private Sub llenarLSTPROCESO()
On Error GoTo err44




     Set rs = classC.CrearRS("select r.id,s.sector,r.idPedido,r.fechaCreado,r.estado from sp.ComprasRequerimientos r inner join sp.sectores s on r.idSector=s.id where r.estado=0")
Me.lstPRoceso.ListItems.Clear
Dim x As ListItem
While Not rs.EOF
Set x = Me.lstPRoceso.ListItems.Add(, , Format(rs!id, "0000"))
    x.SubItems(1) = rs!sector
    If rs!idpedido < 1 Then para = "Stock" Else para = Format(rs!idpedido, "0000")
    x.SubItems(2) = para
    x.SubItems(3) = rs!FechaCreado
    Dim es As Integer
    es = rs!estado
    x.SubItems(4) = funciones.estado_reque(es)
    If x = pos Then
        x.Selected = True
        x.EnsureVisible
    End If
    If es = 0 Then
     x.ListSubItems(4).ForeColor = vbMagenta
    ElseIf es = 1 Then
     x.ListSubItems(4).ForeColor = vbBlue
    End If
    rs.MoveNext


Wend
Set rs = Nothing

Exit Sub
err44:
MsgBox Err.Description

End Sub



Private Sub llenarPendientes()
On Error GoTo err44



'If Me.opTodos Then
'    Set rs = classC.CrearRS("select r.id,s.sector,r.idPedido,r.fechaCreado,r.estado from sp.ComprasRequerimientos r inner join sp.sectores s on r.idSector=s.id ")
'ElseIf Me.opPendientes Then
    Set rs = classC.CrearRS("select r.id,s.sector,r.idPedido,r.fechaCreado,r.estado from sp.ComprasRequerimientos r inner join sp.sectores s on r.idSector=s.id where r.estado=1")
'ElseIf Me.Option1 Then
'     Set rs = classC.CrearRS("select r.id,s.sector,r.idPedido,r.fechaCreado,r.estado from sp.ComprasRequerimientos r inner join sp.sectores s on r.idSector=s.id where r.estado=2")
'ElseIf Me.enProceso Then
'     Set rs = classC.CrearRS("select r.id,s.sector,r.idPedido,r.fechaCreado,r.estado from sp.ComprasRequerimientos r inner join sp.sectores s on r.idSector=s.id where r.estado=0")
'End If
Me.lstPendientes.ListItems.Clear
Dim x As ListItem
While Not rs.EOF
Set x = Me.lstPendientes.ListItems.Add(, , Format(rs!id, "0000"))
    x.SubItems(1) = rs!sector
    If rs!idpedido < 1 Then para = "Stock" Else para = Format(rs!idpedido, "0000")
    x.SubItems(2) = para
    x.SubItems(3) = rs!FechaCreado
    Dim es As Integer
    es = rs!estado
    x.SubItems(4) = funciones.estado_reque(es)
    If x = pos Then
        x.Selected = True
        x.EnsureVisible
    End If
    If es = 0 Then
     x.ListSubItems(4).ForeColor = vbMagenta
    ElseIf es = 1 Then
     x.ListSubItems(4).ForeColor = vbBlue
    End If
    rs.MoveNext


Wend
Set rs = Nothing

Exit Sub
err44:
MsgBox Err.Description

End Sub


Private Sub llenarProcesados()
On Error GoTo err44



'If Me.opTodos Then
'    Set rs = classC.CrearRS("select r.id,s.sector,r.idPedido,r.fechaCreado,r.estado from sp.ComprasRequerimientos r inner join sp.sectores s on r.idSector=s.id ")
'ElseIf Me.opPendientes Then
'    Set rs = classC.CrearRS("select r.id,s.sector,r.idPedido,r.fechaCreado,r.estado from sp.ComprasRequerimientos r inner join sp.sectores s on r.idSector=s.id where r.estado=1")
'ElseIf Me.Option1 Then
     Set rs = classC.CrearRS("select r.id,s.sector,r.idPedido,r.fechaCreado,r.estado from sp.ComprasRequerimientos r inner join sp.sectores s on r.idSector=s.id where r.estado=2")
'ElseIf Me.enProceso Then
'     Set rs = classC.CrearRS("select r.id,s.sector,r.idPedido,r.fechaCreado,r.estado from sp.ComprasRequerimientos r inner join sp.sectores s on r.idSector=s.id where r.estado=0")
'End If
Me.lstProcesados.ListItems.Clear
Dim x As ListItem
While Not rs.EOF
Set x = Me.lstProcesados.ListItems.Add(, , Format(rs!id, "0000"))
    x.SubItems(1) = rs!sector
    If rs!idpedido < 1 Then para = "Stock" Else para = Format(rs!idpedido, "0000")
    x.SubItems(2) = para
    x.SubItems(3) = rs!FechaCreado
    Dim es As Integer
    es = rs!estado
    x.SubItems(4) = funciones.estado_reque(es)
    If x = pos Then
        x.Selected = True
        x.EnsureVisible
    End If
    If es = 0 Then
     x.ListSubItems(4).ForeColor = vbMagenta
    ElseIf es = 1 Then
     x.ListSubItems(4).ForeColor = vbBlue
    End If
    rs.MoveNext


Wend
Set rs = Nothing

Exit Sub
err44:
MsgBox Err.Description

End Sub


