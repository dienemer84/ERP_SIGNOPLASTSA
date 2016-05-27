VERSION 5.00
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#12.0#0"; "CODEJO~1.OCX"
Begin VB.Form frmEmpleadosTareas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tareas posibles para empleado"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4650
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEmpleadosTareas.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   4650
   Begin XtremeReportControl.ReportControl ReportControl 
      Height          =   4650
      Left            =   120
      TabIndex        =   0
      Top             =   765
      Width           =   4395
      _Version        =   786432
      _ExtentX        =   7752
      _ExtentY        =   8202
      _StockProps     =   64
      BorderStyle     =   3
   End
   Begin VB.Label lblEmpleado 
      AutoSize        =   -1  'True
      Caption         =   "Empleado: "
      Height          =   195
      Left            =   150
      TabIndex        =   2
      Tag             =   "Empleado: "
      Top             =   420
      Width           =   795
   End
   Begin VB.Label lblLegajo 
      AutoSize        =   -1  'True
      Caption         =   "Legajo: "
      Height          =   195
      Left            =   150
      TabIndex        =   1
      Tag             =   "Legajo: "
      Top             =   105
      Width           =   585
   End
End
Attribute VB_Name = "frmEmpleadosTareas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private personal_id As Long

Public Property Let personalId(value As Long)
    personal_id = value

    Dim emp As clsEmpleado
    Set emp = DAOEmpleados.GetById(personal_id)
    Me.lblLegajo.caption = Me.lblLegajo.Tag & emp.legajo
    Me.lblEmpleado.caption = Me.lblEmpleado.Tag & emp.Apellido & " " & emp.nombre

    Me.caption = "Tareas posibles para [" & emp.Apellido & " " & emp.nombre & "]"





    Dim rec As ReportRecord

    Dim tar As clsTarea
    For Each tar In DAOTareas.FindAll()
        If funciones.BuscarEnColeccion(emp.sectores, CStr(tar.Sector.id)) Then    'solo mostrar las tareas que pertenezcan a un sector donde este asignado el empleado
            Set rec = ReportControl.Records.Add
            rec.Tag = tar.id
            rec.AddItem tar.Sector.Sector
            rec.AddItem tar.Tarea
            rec.item(1).HasCheckbox = True
        End If
    Next tar

    Me.ReportControl.SortOrder.Add Me.ReportControl.Columns(0)

    Me.ReportControl.Populate




    Dim dic As Dictionary
    Set dic = DAOEmpleados.GetTareasIdAsignadasByPersonalId(personal_id)
    For Each rec In Me.ReportControl.Records
        rec.item(1).Checked = dic.Exists(rec.Tag)
    Next rec
End Property


Private Sub Form_Load()
    FormHelper.Customize Me
    Me.ReportControl.PaintManager.NoItemsText = "No hay sectores, debe sectorizar al empleado"

    Me.ReportControl.Columns.DeleteAll

    Dim Column As ReportColumn

    Set Column = Me.ReportControl.Columns.Add(0, "Sector", 20, True)
    Column.Icon = 0
    Column.Sortable = True
    Column.AllowDrag = False
    Column.AllowRemove = False

    Set Column = Me.ReportControl.Columns.Add(1, "Tarea", 25, True)
    Column.Icon = 0
    Column.Sortable = True
    Column.AllowDrag = False
    Column.AllowRemove = False

    Me.ReportControl.PaintManager.HorizontalGridStyle = xtpGridSmallDots
    Me.ReportControl.PaintManager.VerticalGridStyle = xtpGridSmallDots

    Me.ReportControl.Records.DeleteAll


End Sub


Private Sub ReportControl_ItemCheck(ByVal row As XtremeReportControl.IReportRow, ByVal item As XtremeReportControl.IReportRecordItem)
    SetTareaAsignada personal_id, item.record.Tag, Not item.Checked
End Sub
