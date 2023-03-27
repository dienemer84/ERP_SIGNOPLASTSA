VERSION 5.00
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#12.0#0"; "CODEJO~3.OCX"
Begin VB.Form frmListaTareas 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Tareas"
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   270
   ClientWidth     =   11835
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmListaTareas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   11835
   Begin XtremeReportControl.ReportControl ReportControl 
      Height          =   6270
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Width           =   11835
      _Version        =   786432
      _ExtentX        =   20876
      _ExtentY        =   11060
      _StockProps     =   64
      BorderStyle     =   3
      MultipleSelection=   0   'False
   End
End
Attribute VB_Name = "frmListaTareas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim registro As Long
Implements ISuscriber
Private m_tareas As Collection
Private suscriber_id As String
Private Sub CargarTareas()
    Dim m_tarea As clsTarea
    Me.ReportControl.Records.DeleteAll
    Set m_tareas = DAOTareas.FindAll()
    Dim rec As ReportRecord

    Dim cantxproc As String

    For Each m_tarea In m_tareas
        Set rec = Me.ReportControl.Records.Add
        rec.Tag = m_tarea.Id
        rec.AddItem m_tarea.Id
        rec.AddItem m_tarea.Sector.Sector

        If m_tarea.CantPorProc = 0 Then
            cantxproc = "Fijo"
        ElseIf m_tarea.CantPorProc = -1 Then
            cantxproc = "Cambio"
        Else
            cantxproc = m_tarea.CantPorProc
        End If

        rec.AddItem cantxproc
        rec.AddItem m_tarea.Tarea
        rec.AddItem m_tarea.descripcion
        If m_tarea.CategoriaSueldo Is Nothing Then
            rec.AddItem vbNullString
        Else
            rec.AddItem m_tarea.CategoriaSueldo.nombre
        End If

        If m_tarea.CategoriaSueldo Is Nothing Then
            rec.AddItem 0
        Else
            rec.AddItem m_tarea.CategoriaSueldo.Valor
        End If

        rec.AddItem m_tarea.FEcha
    Next m_tarea

    Me.ReportControl.Populate
    ReportControl.SelectedRows.DeleteAll
    Me.ReportControl.FocusedRow = Me.ReportControl.rows(registro)

    ReportControl.SelectedRows.Add Me.ReportControl.rows(registro)
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CargarTareas
    End If
End Sub
Private Sub Form_Load()
    FormHelper.Customize Me
    suscriber_id = funciones.CreateGUID
    Channel.AgregarSuscriptor Me, Tareas_
    Me.ReportControl.PaintManager.HorizontalGridStyle = xtpGridSmallDots
    Me.ReportControl.PaintManager.VerticalGridStyle = xtpGridSmallDots
    Me.ReportControl.Columns.DeleteAll
    AddColumn "Codigo", xtpAlignmentRight
    AddColumn "Sector"
    AddColumn "Cant x Proc", xtpAlignmentRight
    AddColumn "Tarea"
    AddColumn "Descripcion"
    AddColumn "Categoria"
    AddColumn "Valor", xtpAlignmentRight
    AddColumn "Fecha", xtpAlignmentRight
    Me.ReportControl.AutoColumnSizing = True
    CargarTareas

    ''Me.caption = caption & " (" & Name & ")"


End Sub

Private Sub AddColumn(ByVal caption As String, Optional align As XTPColumnAlignment = xtpAlignmentLeft)
    Static pos As Long
    Dim col As ReportColumn
    Set col = Me.ReportControl.Columns.Add(pos, caption, 25, True)
    col.Icon = 0
    col.Sortable = True
    col.AllowDrag = False
    col.AllowRemove = False
    col.Alignment = align
    pos = pos + 1
End Sub
Private Sub Form_Resize()
    Me.ReportControl.Width = Me.ScaleWidth
    Me.ReportControl.Height = Me.ScaleHeight
End Sub
Private Sub Form_Terminate()
    Channel.RemoverSuscripcionTotal Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Channel.RemoverSuscripcionTotal Me
End Sub
Private Property Get ISuscriber_id() As String
    ISuscriber_id = suscriber_id
End Property
Private Function ISuscriber_Notificarse(EVENTO As clsEventoObserver) As Variant

    CargarTareas
End Function

Private Sub ReportControl_RowDblClick(ByVal row As XtremeReportControl.IReportRow, ByVal item As XtremeReportControl.IReportRecordItem)
    Dim F As New frmNuevaMDO
    F.tareaId = row.record.Tag
    F.Show
End Sub

Private Sub ReportControl_SelectionChanged()
    If Me.ReportControl.SelectedRows.count > 0 Then
        registro = Me.ReportControl.SelectedRows(0).Index
    End If



End Sub
