VERSION 5.00
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#12.0#0"; "CODEJO~1.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~3.OCX"
Begin VB.Form frmUbicaciones 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ubicaciones"
   ClientHeight    =   7365
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   12195
   ShowInTaskbar   =   0   'False
   Begin XtremeReportControl.ReportControl repLocalidades 
      Height          =   5820
      Left            =   6885
      TabIndex        =   2
      Top             =   15
      Width           =   5265
      _Version        =   786432
      _ExtentX        =   9287
      _ExtentY        =   10266
      _StockProps     =   64
   End
   Begin XtremeReportControl.ReportControl repProvincias 
      Height          =   5820
      Left            =   3495
      TabIndex        =   1
      Top             =   0
      Width           =   3360
      _Version        =   786432
      _ExtentX        =   5927
      _ExtentY        =   10266
      _StockProps     =   64
   End
   Begin XtremeReportControl.ReportControl repPaises 
      Height          =   5820
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   3360
      _Version        =   786432
      _ExtentX        =   5927
      _ExtentY        =   10266
      _StockProps     =   64
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1395
      Left            =   60
      TabIndex        =   3
      Top             =   5895
      Width           =   3390
      _Version        =   786432
      _ExtentX        =   5980
      _ExtentY        =   2461
      _StockProps     =   79
      Caption         =   "País"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   330
         Left            =   495
         TabIndex        =   6
         Top             =   930
         Width           =   1155
         _Version        =   786432
         _ExtentX        =   2037
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "Guardar"
         UseVisualStyle  =   -1  'True
      End
      Begin VB.TextBox txtNombrePaís 
         Height          =   285
         Left            =   720
         TabIndex        =   5
         Top             =   375
         Width           =   2400
      End
      Begin XtremeSuiteControls.PushButton PushButton2 
         Height          =   330
         Left            =   1695
         TabIndex        =   7
         Top             =   930
         Width           =   1155
         _Version        =   786432
         _ExtentX        =   2037
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "Nuevo"
         UseVisualStyle  =   -1  'True
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre"
         Height          =   225
         Left            =   90
         TabIndex        =   4
         Top             =   375
         Width           =   675
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox2 
      Height          =   1395
      Left            =   3480
      TabIndex        =   8
      Top             =   5895
      Width           =   3390
      _Version        =   786432
      _ExtentX        =   5980
      _ExtentY        =   2461
      _StockProps     =   79
      Caption         =   "Provincia"
      UseVisualStyle  =   -1  'True
      Begin VB.TextBox txtNombreProvincia 
         Height          =   285
         Left            =   720
         TabIndex        =   10
         Top             =   360
         Width           =   2400
      End
      Begin XtremeSuiteControls.PushButton PushButton3 
         Height          =   330
         Left            =   510
         TabIndex        =   9
         Top             =   930
         Width           =   1155
         _Version        =   786432
         _ExtentX        =   2037
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "Guardar"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton4 
         Height          =   330
         Left            =   1695
         TabIndex        =   11
         Top             =   930
         Width           =   1155
         _Version        =   786432
         _ExtentX        =   2037
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "Nuevo"
         UseVisualStyle  =   -1  'True
      End
      Begin VB.Label Label2 
         Caption         =   "Nombre"
         Height          =   225
         Left            =   90
         TabIndex        =   12
         Top             =   375
         Width           =   675
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox3 
      Height          =   1395
      Left            =   6930
      TabIndex        =   13
      Top             =   5895
      Width           =   5220
      _Version        =   786432
      _ExtentX        =   9208
      _ExtentY        =   2461
      _StockProps     =   79
      Caption         =   "Localidad"
      UseVisualStyle  =   -1  'True
      Begin VB.TextBox txtCP 
         Height          =   285
         Left            =   1185
         TabIndex        =   18
         Top             =   435
         Width           =   825
      End
      Begin VB.TextBox txtNombreLocalidad 
         Height          =   285
         Left            =   2040
         TabIndex        =   14
         Top             =   435
         Width           =   3015
      End
      Begin XtremeSuiteControls.PushButton PushButton5 
         Height          =   330
         Left            =   1440
         TabIndex        =   15
         Top             =   945
         Width           =   1155
         _Version        =   786432
         _ExtentX        =   2037
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "Guardar"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton6 
         Height          =   330
         Left            =   2655
         TabIndex        =   16
         Top             =   930
         Width           =   1155
         _Version        =   786432
         _ExtentX        =   2037
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "Nuevo"
         UseVisualStyle  =   -1  'True
      End
      Begin VB.Label Label4 
         Caption         =   "CP | Nombre"
         Height          =   225
         Left            =   180
         TabIndex        =   17
         Top             =   435
         Width           =   1110
      End
   End
End
Attribute VB_Name = "frmUbicaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private paises As New Collection
Private pais As pais
Private provincias As New Collection
Private provincia As provincia
Private localidades As New Collection
Private localidad As localidad


Private Sub Form_Load()
    Customize Me
    ArmarColumnasPaises
    ArmarColumnasProvincias
    ArmarColumnasLocalidades
    CargarPaises
    repPaises_SelectionChanged
    repProvincias_SelectionChanged

End Sub
Private Sub ArmarColumnasPaises()
    Me.repPaises.Columns.DeleteAll
    ReportControlAddColumn Me.repPaises, 0, "Pais"
    Me.repPaises.PaintManager.HorizontalGridStyle = xtpGridSmallDots
    Me.repPaises.PaintManager.VerticalGridStyle = xtpGridSmallDots
End Sub

Private Sub ArmarColumnasProvincias()
    Me.repProvincias.Columns.DeleteAll
    ReportControlAddColumn Me.repProvincias, 0, "Provincia"
    Me.repProvincias.PaintManager.HorizontalGridStyle = xtpGridSmallDots
    Me.repProvincias.PaintManager.VerticalGridStyle = xtpGridSmallDots
End Sub

Private Sub ArmarColumnasLocalidades()
    Me.repLocalidades.Columns.DeleteAll
    ReportControlAddColumn Me.repLocalidades, 0, "C.P.", , , 10
    ReportControlAddColumn Me.repLocalidades, 1, "Localidad"
    Me.repLocalidades.PaintManager.HorizontalGridStyle = xtpGridSmallDots
    Me.repLocalidades.PaintManager.VerticalGridStyle = xtpGridSmallDots
End Sub

Private Sub CargarPaises()
    Dim P As pais
    Dim rec As ReportRecord
    Set paises = DAOPais.FindAll
    Me.repPaises.Records.DeleteAll
    For Each P In paises
        Set rec = Me.repPaises.Records.Add
        rec.AddItem P.nombre
        rec.Tag = P.id
    Next P
    Me.repPaises.Populate
End Sub
Private Sub CargarProvincias()
    Dim P As provincia
    Dim rec As ReportRecord
    Set provincias = DAOProvincias.FindAllByPais(Me.repPaises.SelectedRows(0).record.Tag)
    Me.repProvincias.Records.DeleteAll
    For Each P In provincias
        Set rec = Me.repProvincias.Records.Add
        rec.AddItem P.nombre
        rec.Tag = P.id
    Next P
    Me.repProvincias.Populate
End Sub
Private Sub CargarLocalidades()
    On Error Resume Next
    Dim l As localidad
    Dim rec As ReportRecord
    Set localidades = DAOLocalidades.FindAllByProvincia(Me.repProvincias.SelectedRows(0).record.Tag)
    Me.repLocalidades.Records.DeleteAll
    For Each l In localidades
        Set rec = Me.repLocalidades.Records.Add
        rec.AddItem l.cp
        rec.AddItem l.nombre
        rec.Tag = l.id
    Next l
    Me.repLocalidades.Populate


End Sub


Private Sub PushButton1_Click()
    On Error Resume Next
    If Not IsSomething(pais) Then pais = New pais
    pais.nombre = Me.txtNombrePaís
    DAOPais.Save pais
    CargarPaises
    CargarProvincias
    CargarLocalidades

End Sub

Private Sub PushButton2_Click()
    Set pais = New pais
    Me.txtNombrePaís = vbNullString
End Sub

Private Sub PushButton3_Click()
    If Not IsSomething(provincia) Then provincia = New provincia
    provincia.nombre = Me.txtNombreProvincia
    Set provincia.pais = pais
    DAOProvincias.Save provincia
    CargarProvincias
    CargarLocalidades
End Sub
Private Sub PushButton4_Click()
    Set provincia = New provincia
    Me.txtNombreProvincia = vbNullString
End Sub

Private Sub PushButton5_Click()
    If Not IsSomething(localidad) Then Set localidad = New localidad

    localidad.cp = Me.txtCP
    localidad.nombre = Me.txtNombreLocalidad
    Set localidad.provincia = provincia
    DAOLocalidades.Save localidad
    Set localidad = New localidad
    Me.txtCP = vbNullString
    Me.txtNombreLocalidad = vbNullString

    CargarLocalidades


End Sub

Private Sub PushButton6_Click()
    Set localidad = New localidad
    Me.txtCP = vbNullString
    Me.txtNombreLocalidad = vbNullString
    CargarLocalidades
End Sub



Private Sub repLocalidades_SelectionChanged()
    Set localidad = DAOLocalidades.FindById(repLocalidades.SelectedRows(0).record.Tag)
    Me.txtCP = localidad.cp
    Me.txtNombreLocalidad = localidad.nombre

End Sub

Private Sub repPaises_SelectionChanged()
    Me.repLocalidades.Records.DeleteAll
    CargarProvincias
    CargarLocalidades
    Set pais = DAOPais.FindById(repPaises.SelectedRows(0).record.Tag)
    Me.txtNombrePaís = pais.nombre
    repProvincias_SelectionChanged
End Sub
Private Sub repProvincias_SelectionChanged()

    CargarLocalidades
    Set provincia = DAOProvincias.FindById(repProvincias.SelectedRows(0).record.Tag)
    Me.txtNombreProvincia = provincia.nombre
End Sub
