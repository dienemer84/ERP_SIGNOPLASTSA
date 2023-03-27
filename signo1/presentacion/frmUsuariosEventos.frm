VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmUsuariosEventos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asignación de Eventos de Usuarios"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7545
   Icon            =   "frmUsuariosEventos.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   7545
   Begin XtremeSuiteControls.PushButton btnGuardar 
      Height          =   465
      Left            =   5235
      TabIndex        =   2
      Top             =   5385
      Width           =   2250
      _Version        =   786432
      _ExtentX        =   3969
      _ExtentY        =   820
      _StockProps     =   79
      Caption         =   "Guardar Eventos de Usuario"
      UseVisualStyle  =   -1  'True
   End
   Begin GridEX20.GridEX gridUsuarios 
      Height          =   5775
      Left            =   60
      TabIndex        =   0
      Top             =   75
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   10186
      Version         =   "2.0"
      HoldSortSettings=   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      MethodHoldFields=   -1  'True
      GroupByBoxVisible=   0   'False
      DataMode        =   99
      HeaderFontName  =   "Tahoma"
      FontName        =   "Tahoma"
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   1
      Column(1)       =   "frmUsuariosEventos.frx":000C
      SortKeysCount   =   1
      SortKey(1)      =   "frmUsuariosEventos.frx":0124
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmUsuariosEventos.frx":018C
      FormatStyle(2)  =   "frmUsuariosEventos.frx":02B4
      FormatStyle(3)  =   "frmUsuariosEventos.frx":0364
      FormatStyle(4)  =   "frmUsuariosEventos.frx":0418
      FormatStyle(5)  =   "frmUsuariosEventos.frx":04F0
      FormatStyle(6)  =   "frmUsuariosEventos.frx":05A8
      ImageCount      =   0
      PrinterProperties=   "frmUsuariosEventos.frx":0688
   End
   Begin GridEX20.GridEX gridEventos 
      Height          =   5220
      Left            =   2235
      TabIndex        =   1
      Top             =   75
      Width           =   5235
      _ExtentX        =   9234
      _ExtentY        =   9208
      Version         =   "2.0"
      HoldSortSettings=   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      MethodHoldFields=   -1  'True
      GroupByBoxVisible=   0   'False
      DataMode        =   99
      HeaderFontName  =   "Tahoma"
      FontName        =   "Tahoma"
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   2
      Column(1)       =   "frmUsuariosEventos.frx":0858
      Column(2)       =   "frmUsuariosEventos.frx":0988
      SortKeysCount   =   1
      SortKey(1)      =   "frmUsuariosEventos.frx":0AB4
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmUsuariosEventos.frx":0B1C
      FormatStyle(2)  =   "frmUsuariosEventos.frx":0C44
      FormatStyle(3)  =   "frmUsuariosEventos.frx":0CF4
      FormatStyle(4)  =   "frmUsuariosEventos.frx":0DA8
      FormatStyle(5)  =   "frmUsuariosEventos.frx":0E80
      FormatStyle(6)  =   "frmUsuariosEventos.frx":0F38
      ImageCount      =   0
      PrinterProperties=   "frmUsuariosEventos.frx":1018
   End
   Begin VB.Line Line 
      BorderColor     =   &H00FFDBBF&
      DrawMode        =   9  'Not Mask Pen
      X1              =   2145
      X2              =   2145
      Y1              =   75
      Y2              =   5850
   End
End
Attribute VB_Name = "frmUsuariosEventos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private usuarios As Collection
Private usuario As clsUsuario
Private tipoEventos As Collection

Private Sub btnGuardar_Click()

    If Me.gridUsuarios.row = -1 Then Exit Sub

    'para forzar el lostfoucs del ultimo item clickeado cuando doy guardar
    If Me.gridEventos.row = 1 Then
        Me.gridEventos.row = 2
    Else
        Me.gridEventos.row = 1
    End If

    If DAOEvento.AddBroadCastTypesSuscribedForUser(usuario.Id, usuario.EventosSuscriptos) Then
        MsgBox "Eventos para el usuario [" & usuario.usuario & "] actualizados.", vbInformation + vbOKOnly
    Else
        MsgBox "No se pudo guardar los datos.", vbOKOnly + vbCritical
    End If
End Sub

Private Sub Form_Load()
    Customize Me
    GridEXHelper.CustomizeGrid Me.gridUsuarios
    GridEXHelper.CustomizeGrid Me.gridEventos, , True

    Set usuarios = DAOUsuarios.FindAll()
    Me.gridUsuarios.ItemCount = 0
    Me.gridUsuarios.ItemCount = usuarios.count

    Set tipoEventos = DAOEvento.GetEventBroadCastTypes()
    Me.gridEventos.ItemCount = 0

    Me.gridUsuarios.row = -1
End Sub

Private Sub gridEventos_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex > 0 And tipoEventos.count > 0 Then
        Values(1) = usuario.EventosSuscriptos.Exists(CStr(tipoEventos.item(RowIndex)(1)))
        Values(2) = tipoEventos.item(RowIndex)(2)
    End If
End Sub

Private Sub gridEventos_UnboundUpdate(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex > 0 Then
        If Values(1) Then
            If Not usuario.EventosSuscriptos.Exists(CStr(tipoEventos.item(RowIndex)(1))) Then
                usuario.EventosSuscriptos.Add CStr(tipoEventos.item(RowIndex)(1)), tipoEventos.item(RowIndex)(1)
            End If
        Else
            If usuario.EventosSuscriptos.Exists(CStr(tipoEventos.item(RowIndex)(1))) Then
                usuario.EventosSuscriptos.remove CStr(tipoEventos.item(RowIndex)(1))
            End If
        End If
    End If
End Sub

Private Sub gridUsuarios_SelectionChange()
    If Me.gridUsuarios.row > 0 And IsSomething(tipoEventos) Then
        Set usuario = usuarios.item(Me.gridUsuarios.RowIndex(Me.gridUsuarios.row))
        Set usuario.EventosSuscriptos = Nothing
        Me.gridEventos.ItemCount = 0
        Me.gridEventos.ItemCount = tipoEventos.count
        GridEXHelper.AutoSizeColumns Me.gridEventos
    End If
End Sub

Private Sub gridUsuarios_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex > 0 And usuarios.count > 0 Then
        Set usuario = usuarios.item(RowIndex)
        Values(1) = usuario.usuario
    End If
End Sub
