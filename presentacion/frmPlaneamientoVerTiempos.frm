VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#12.0#0"; "CODEJO~1.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~3.OCX"
Begin VB.Form frmPlaneamientoVerTiempos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ver Tiempos"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12270
   Icon            =   "frmPlaneamientoVerTiempos.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   12270
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   6600
      Left            =   15
      TabIndex        =   0
      Top             =   0
      Width           =   12255
      _Version        =   786432
      _ExtentX        =   21616
      _ExtentY        =   11642
      _StockProps     =   68
      Appearance      =   10
      Color           =   32
      PaintManager.BoldSelected=   -1  'True
      PaintManager.ShowIcons=   -1  'True
      PaintManager.FixedTabWidth=   0
      ItemCount       =   4
      Item(0).Caption =   "Por Legajo"
      Item(0).ControlCount=   7
      Item(0).Control(0)=   "Label2"
      Item(0).Control(1)=   "txtLegajo"
      Item(0).Control(2)=   "Command1"
      Item(0).Control(3)=   "lblEmpleado"
      Item(0).Control(4)=   "grilla_por_legajo"
      Item(0).Control(5)=   "lblTotalHorasPorLegajo"
      Item(0).Control(6)=   "GroupBox1"
      Item(1).Caption =   "Por Período"
      Item(1).ControlCount=   4
      Item(1).Control(0)=   "GroupBox2"
      Item(1).Control(1)=   "grilla_por_periodo"
      Item(1).Control(2)=   "cmdBuscarPorPeriodo"
      Item(1).Control(3)=   "lblTotalPorPeriodo"
      Item(2).Caption =   "Por OT"
      Item(2).ControlCount=   4
      Item(2).Control(0)=   "txtNroOt"
      Item(2).Control(1)=   "Command2"
      Item(2).Control(2)=   "Label3"
      Item(2).Control(3)=   "TabControl2"
      Item(3).Caption =   "Por Pieza"
      Item(3).ControlCount=   4
      Item(3).Control(0)=   "reportPorPieza"
      Item(3).Control(1)=   "txtPieza"
      Item(3).Control(2)=   "PushButton1"
      Item(3).Control(3)=   "Label1"
      Begin XtremeReportControl.ReportControl reportPorPieza 
         Height          =   5100
         Left            =   -69895
         TabIndex        =   14
         Top             =   1230
         Visible         =   0   'False
         Width           =   11985
         _Version        =   786432
         _ExtentX        =   21140
         _ExtentY        =   8996
         _StockProps     =   64
         BorderStyle     =   3
      End
      Begin XtremeSuiteControls.TabControl TabControl2 
         Height          =   5265
         Left            =   -69955
         TabIndex        =   30
         Top             =   1245
         Visible         =   0   'False
         Width           =   12120
         _Version        =   786432
         _ExtentX        =   21378
         _ExtentY        =   9287
         _StockProps     =   68
         Appearance      =   10
         Color           =   32
         ItemCount       =   2
         Item(0).Caption =   "Detallado"
         Item(0).ControlCount=   1
         Item(0).Control(0)=   "reportControl"
         Item(1).Caption =   "Global"
         Item(1).ControlCount=   0
         Begin XtremeReportControl.ReportControl reportControl 
            Height          =   4650
            Left            =   165
            TabIndex        =   31
            Top             =   480
            Width           =   11760
            _Version        =   786432
            _ExtentX        =   20743
            _ExtentY        =   8202
            _StockProps     =   64
            BorderStyle     =   3
         End
      End
      Begin VB.TextBox txtNroOt 
         Height          =   285
         Left            =   -69175
         TabIndex        =   18
         Top             =   840
         Visible         =   0   'False
         Width           =   1530
      End
      Begin VB.TextBox txtPieza 
         Height          =   285
         Left            =   -69175
         TabIndex        =   15
         Top             =   840
         Visible         =   0   'False
         Width           =   2715
      End
      Begin VB.TextBox txtLegajo 
         Height          =   285
         Left            =   750
         TabIndex        =   1
         Top             =   1305
         Width           =   1080
      End
      Begin XtremeSuiteControls.PushButton Command1 
         Height          =   375
         Left            =   1935
         TabIndex        =   2
         Top             =   1260
         Width           =   1380
         _Version        =   786432
         _ExtentX        =   2434
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Buscar"
         BackColor       =   16777215
         Appearance      =   6
      End
      Begin GridEX20.GridEX grilla_por_legajo 
         Height          =   4470
         Left            =   180
         TabIndex        =   3
         Top             =   1695
         Width           =   11895
         _ExtentX        =   20981
         _ExtentY        =   7885
         Version         =   "2.0"
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         ColumnAutoResize=   -1  'True
         MethodHoldFields=   -1  'True
         DataMode        =   99
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   7
         Column(1)       =   "frmPlaneamientoVerTiempos.frx":000C
         Column(2)       =   "frmPlaneamientoVerTiempos.frx":00F8
         Column(3)       =   "frmPlaneamientoVerTiempos.frx":01E8
         Column(4)       =   "frmPlaneamientoVerTiempos.frx":02EC
         Column(5)       =   "frmPlaneamientoVerTiempos.frx":03E4
         Column(6)       =   "frmPlaneamientoVerTiempos.frx":0514
         Column(7)       =   "frmPlaneamientoVerTiempos.frx":063C
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmPlaneamientoVerTiempos.frx":0740
         FormatStyle(2)  =   "frmPlaneamientoVerTiempos.frx":0878
         FormatStyle(3)  =   "frmPlaneamientoVerTiempos.frx":0928
         FormatStyle(4)  =   "frmPlaneamientoVerTiempos.frx":09DC
         FormatStyle(5)  =   "frmPlaneamientoVerTiempos.frx":0AB4
         FormatStyle(6)  =   "frmPlaneamientoVerTiempos.frx":0B6C
         ImageCount      =   0
         PrinterProperties=   "frmPlaneamientoVerTiempos.frx":0C4C
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   1230
         Left            =   7395
         TabIndex        =   7
         Top             =   375
         Width           =   4680
         _Version        =   786432
         _ExtentX        =   8255
         _ExtentY        =   2170
         _StockProps     =   79
         Caption         =   "Período"
         BackColor       =   16777152
         Appearance      =   4
         Begin XtremeSuiteControls.RadioButton rbFecha 
            Height          =   270
            Left            =   225
            TabIndex        =   9
            Top             =   240
            Width           =   765
            _Version        =   786432
            _ExtentX        =   1349
            _ExtentY        =   476
            _StockProps     =   79
            Caption         =   "Fecha"
            UseVisualStyle  =   -1  'True
            Value           =   -1  'True
         End
         Begin XtremeSuiteControls.DateTimePicker DateTimeFecha 
            Height          =   270
            Left            =   1020
            TabIndex        =   8
            Top             =   225
            Width           =   3495
            _Version        =   786432
            _ExtentX        =   6165
            _ExtentY        =   476
            _StockProps     =   68
            CurrentDate     =   40136.5580671296
         End
         Begin XtremeSuiteControls.RadioButton rbMes 
            Height          =   270
            Left            =   225
            TabIndex        =   10
            Top             =   540
            Width           =   765
            _Version        =   786432
            _ExtentX        =   1349
            _ExtentY        =   476
            _StockProps     =   79
            Caption         =   "Mes"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton rbAno 
            Height          =   270
            Left            =   225
            TabIndex        =   11
            Top             =   840
            Width           =   765
            _Version        =   786432
            _ExtentX        =   1349
            _ExtentY        =   476
            _StockProps     =   79
            Caption         =   "Año"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.ComboBox cboAnios 
            Height          =   315
            Left            =   1020
            TabIndex        =   13
            Top             =   825
            Width           =   3495
            _Version        =   786432
            _ExtentX        =   6165
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Style           =   2
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.ComboBox cboMeses 
            Height          =   315
            Left            =   1020
            TabIndex        =   12
            Top             =   510
            Width           =   3495
            _Version        =   786432
            _ExtentX        =   6165
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Style           =   2
            Text            =   "ComboBox1"
         End
      End
      Begin XtremeSuiteControls.PushButton PushButton1 
         Default         =   -1  'True
         Height          =   375
         Left            =   -66250
         TabIndex        =   16
         Top             =   780
         Visible         =   0   'False
         Width           =   1380
         _Version        =   786432
         _ExtentX        =   2434
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Buscar"
         BackColor       =   16777215
         Appearance      =   6
      End
      Begin XtremeSuiteControls.GroupBox GroupBox2 
         Height          =   1230
         Left            =   -69820
         TabIndex        =   20
         Top             =   390
         Visible         =   0   'False
         Width           =   4620
         _Version        =   786432
         _ExtentX        =   8149
         _ExtentY        =   2170
         _StockProps     =   79
         Caption         =   "Período"
         BackColor       =   16777152
         Appearance      =   4
         Begin XtremeSuiteControls.RadioButton rbFecha2 
            Height          =   270
            Left            =   210
            TabIndex        =   21
            Top             =   240
            Width           =   765
            _Version        =   786432
            _ExtentX        =   1349
            _ExtentY        =   476
            _StockProps     =   79
            Caption         =   "Fecha"
            UseVisualStyle  =   -1  'True
            Value           =   -1  'True
         End
         Begin XtremeSuiteControls.DateTimePicker dtFecha2 
            Height          =   270
            Left            =   1005
            TabIndex        =   22
            Top             =   210
            Width           =   3495
            _Version        =   786432
            _ExtentX        =   6165
            _ExtentY        =   476
            _StockProps     =   68
            CurrentDate     =   40136.5580671296
         End
         Begin XtremeSuiteControls.RadioButton rbMes2 
            Height          =   270
            Left            =   210
            TabIndex        =   23
            Top             =   540
            Width           =   765
            _Version        =   786432
            _ExtentX        =   1349
            _ExtentY        =   476
            _StockProps     =   79
            Caption         =   "Mes"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton rbAño2 
            Height          =   270
            Left            =   210
            TabIndex        =   24
            Top             =   840
            Width           =   765
            _Version        =   786432
            _ExtentX        =   1349
            _ExtentY        =   476
            _StockProps     =   79
            Caption         =   "Año"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.ComboBox cboAño2 
            Height          =   315
            Left            =   1020
            TabIndex        =   25
            Top             =   825
            Width           =   3495
            _Version        =   786432
            _ExtentX        =   6165
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Style           =   2
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.ComboBox cboMeses2 
            Height          =   315
            Left            =   1020
            TabIndex        =   26
            Top             =   510
            Width           =   3495
            _Version        =   786432
            _ExtentX        =   6165
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Style           =   2
            Text            =   "ComboBox1"
         End
      End
      Begin GridEX20.GridEX grilla_por_periodo 
         Height          =   4680
         Left            =   -69820
         TabIndex        =   27
         Top             =   1695
         Visible         =   0   'False
         Width           =   11895
         _ExtentX        =   20981
         _ExtentY        =   8255
         Version         =   "2.0"
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         ColumnAutoResize=   -1  'True
         MethodHoldFields=   -1  'True
         DataMode        =   99
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   2
         Column(1)       =   "frmPlaneamientoVerTiempos.frx":0E24
         Column(2)       =   "frmPlaneamientoVerTiempos.frx":0F3C
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmPlaneamientoVerTiempos.frx":1064
         FormatStyle(2)  =   "frmPlaneamientoVerTiempos.frx":119C
         FormatStyle(3)  =   "frmPlaneamientoVerTiempos.frx":124C
         FormatStyle(4)  =   "frmPlaneamientoVerTiempos.frx":1300
         FormatStyle(5)  =   "frmPlaneamientoVerTiempos.frx":13D8
         FormatStyle(6)  =   "frmPlaneamientoVerTiempos.frx":1490
         ImageCount      =   0
         PrinterProperties=   "frmPlaneamientoVerTiempos.frx":1570
      End
      Begin XtremeSuiteControls.PushButton cmdBuscarPorPeriodo 
         Height          =   450
         Left            =   -65035
         TabIndex        =   28
         Top             =   1125
         Visible         =   0   'False
         Width           =   1305
         _Version        =   786432
         _ExtentX        =   2302
         _ExtentY        =   794
         _StockProps     =   79
         Caption         =   "Buscar"
         BackColor       =   16777215
         Appearance      =   6
      End
      Begin XtremeSuiteControls.PushButton Command2 
         Height          =   375
         Left            =   -67420
         TabIndex        =   32
         Top             =   690
         Visible         =   0   'False
         Width           =   1380
         _Version        =   786432
         _ExtentX        =   2434
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Buscar"
         BackColor       =   16777215
         Appearance      =   6
      End
      Begin VB.Label lblTotalPorPeriodo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   -58015
         TabIndex        =   29
         Top             =   1380
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "O/T"
         Height          =   195
         Left            =   -69745
         TabIndex        =   19
         Top             =   855
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pieza"
         Height          =   195
         Left            =   -69730
         TabIndex        =   17
         Top             =   855
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.Label lblEmpleado 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Height          =   210
         Left            =   9180
         TabIndex        =   6
         Top             =   1395
         Width           =   2910
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Legajo"
         Height          =   195
         Left            =   180
         TabIndex        =   5
         Top             =   1320
         Width           =   480
      End
      Begin VB.Label lblTotalHorasPorLegajo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   11970
         TabIndex        =   4
         Top             =   6255
         Width           =   45
      End
   End
End
Attribute VB_Name = "frmPlaneamientoVerTiempos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim piezas As Collection
Dim id_pieza_elegida As Long
Dim colAnio As Collection
Dim colMeses As Collection
Dim colAnio2 As Collection
Dim colMeses2 As Collection
Dim Ot As OrdenTrabajo
Dim detalles_leg As New Collection
Dim detalles_per As New Collection
Dim ptp_detalles_ot As New Collection

Dim emple As clsEmpleado
Dim deta As PlaneamientoTiempoProcesoDetalle  'DetalleOrdenTrabajo
Dim deta2 As DetalleOrdenTrabajo
Dim ptpd As Collection
Dim OtId As Long
Dim colptp As Collection    'PlaneamientoTiempoProceso
Dim ptp As PlaneamientoTiempoProceso
Private Sub cboAnios_Click()
    Me.rbAno.value = True
End Sub

Private Sub cboAnios_GotFocus()
    Me.rbAno.value = True
End Sub

Private Sub cboAño2_Click()
    Me.rbAño2.value = True
End Sub
Private Sub cboMeses_Click()
    Me.rbMes.value = True
End Sub

Private Sub cboMeses_GotFocus()
    Me.rbMes.value = True
End Sub

Private Sub cboMeses2_Click()
    Me.rbMes2.value = True
End Sub
Private Sub cmdBuscarPorPeriodo_Click()
    LlenarListaPorPeriodo
End Sub
Private Sub Command1_Click()
    LlenarListaPorLegajo
End Sub
Private Sub LimpiarPorLegajo()
    Me.grilla_por_legajo.ItemCount = 0
    Me.lblEmpleado = Empty
End Sub
Private Sub Command2_Click()
    Me.ReportControl.PaintManager.NoItemsText = "No hay piezas"
    LlenarListaPorOT
End Sub

Private Sub LlenarListaPorPiezas()
    Dim Pieza As Pieza
    Set Pieza = DAOPieza.FindById(id_pieza_elegida, FL_4, True, False, True)
    Me.reportPorPieza.Records.DeleteAll

    If IsSomething(Pieza) Then
        AgregarPieza Pieza, Nothing
    End If

    Me.reportPorPieza.Populate

End Sub
Private Sub AgregarPieza(Pieza As Pieza, parent As ReportRecord)
    Dim pieza_hija As Pieza
    Dim rr As ReportRecord
    Dim rr2 As ReportRecord
    Dim tar As DesarrolloManoObra
    If IsSomething(parent) Then
        Set rr = parent.Childs.Add
    Else
        Set rr = Me.reportPorPieza.Records.Add
    End If
    rr.Expanded = True
    rr.AddItem Pieza.nombre
    For Each tar In Pieza.desarrollosManoObra
        Set rr2 = rr.Childs.Add
        rr2.AddItem "Tarea: " & tar.Tarea.id & " - " & tar.Tarea.Tarea
        rr2.AddItem tar.Cantidad
        rr2.AddItem funciones.FormatearDecimales(tar.Tiempo)
        rr2.AddItem 0
        rr2.AddItem funciones.FormatearDecimales(tar.TiempoPromedioHistorico)
    Next

    For Each pieza_hija In Pieza.PiezasHijas
        AgregarPieza pieza_hija, rr
    Next


End Sub

Private Sub LlenarListaPorOT()
    On Error GoTo err1
    Dim deta As DetalleOrdenTrabajo
    Dim filtro As String
    If LenB(Me.txtNroOt.text) > 0 And IsNumeric(Me.txtNroOt.text) Then
        OtId = CLng(Me.txtNroOt.text)
        filtro = " ptp.idPedido=" & OtId
        ReportControl.Records.DeleteAll
        Set Ot = DAOOrdenTrabajo.FindById(OtId)
        Set Ot.Detalles = DAODetalleOrdenTrabajo.FindAllByOrdenTrabajo(Ot.id)
        CargarDetallesOT
    End If

    Exit Sub
err1:

End Sub
Private Sub AddTareas(ByRef rec As ReportRecord, ByRef idDetallePedido As Long, Optional ByRef idDetallePedidoConjunto As Long = 0)
    Dim rechijo As ReportRecord
    Dim item As ReportRecordItem
    Dim SUM As Double
    Dim op As String
    'For Each ptp In DAOTiemposProceso.FindAllByDetallePedidoIdAndPiezaId(idDetallePedido, P.Id, True )
    For Each ptp In DAOTiemposProceso.FindAllByDetallePedidoId(idDetallePedido, idDetallePedidoConjunto, True)
        Set rechijo = rec.Childs.Add
        rechijo.Tag = (ptp.id * -1)
        rechijo.AddItem vbNullString
        SUM = 0
        For Each deta In ptp.Detalles
            SUM = SUM + deta.DiferenciaTiempos
        Next
        Set item = rechijo.AddItem("Tarea: " & ptp.Tarea.id & " - " & ptp.Tarea.Tarea)

        rechijo.AddItem vbNullString
        rechijo.AddItem vbNullString
        rechijo.AddItem funciones.FormatearDecimales(SUM)
        rechijo.AddItem ptp.TiempoCotizado & " Min (" & ptp.OperariosCotizado & " Oper)"
    Next ptp
    If rec.Childs.count > 0 Then rec.Expanded = False    'me cierra los nodos que tienen tareas dentro
End Sub
Private Sub CargarDetallesOT()
    Dim porc As Double, prom As Double
    Dim P As Pieza
    Dim tmpdeta2 As DetalleOTConjuntoDTO
    Dim tmpdeta3 As DetalleOTConjuntoDTO
    Dim tmpdeta4 As DetalleOTConjuntoDTO
    Me.ReportControl.Records.DeleteAll
    Dim record As ReportRecord
    Dim Record2 As ReportRecord
    Dim Record3 As ReportRecord
    Dim Record4 As ReportRecord
    Dim item As ReportRecordItem
    For Each deta2 In Ot.Detalles
        Set record = Me.ReportControl.Records.Add
        record.Tag = deta2.id
        record.AddItem deta2.item
        Set item = record.AddItem(deta2.Pieza.nombre)
        record.PreviewText = deta2.Nota
        record.AddItem deta2.CantidadPedida
        record.AddItem deta2.FechaEntrega
        DAODetalleOrdenTrabajo.CalcularPorcentajeAvanceYPromedioFabricado deta2.id, porc, prom

        record.AddItem funciones.FormatearDecimales(porc) & "% avance"

        AddTareas record, deta2.id

        If deta2.Pieza.EsConjunto Then
            For Each tmpdeta2 In DAODetalleOrdenTrabajo.FindAllConjunto(deta2.id, deta2.Pieza.id)
                Set Record2 = record.Childs.Add()
                Record2.Tag = tmpdeta2.id
                Record2.AddItem vbNullString
                Set item = Record2.AddItem(tmpdeta2.Pieza.nombre)
                Record2.Expanded = True
                Record2.AddItem tmpdeta2.Cantidad * record.item(2).value
                Record2.AddItem vbNullString

                AddTareas Record2, deta2.id, tmpdeta2.id

                If tmpdeta2.Pieza.EsConjunto Then
                    For Each tmpdeta3 In DAODetalleOrdenTrabajo.FindAllConjunto(deta2.id, tmpdeta2.Pieza.id)
                        Set Record3 = Record2.Childs.Add
                        Record3.Tag = tmpdeta3.id
                        Record3.AddItem vbNullString
                        Set item = Record3.AddItem(tmpdeta3.Pieza.nombre)
                        Record3.Expanded = True
                        'Item.HasCheckbox = True
                        Record3.AddItem tmpdeta3.Cantidad * Record2.item(2).value
                        Record3.AddItem vbNullString
                        AddTareas Record3, deta2.id, tmpdeta3.id
                        If tmpdeta3.Pieza.EsConjunto Then
                            For Each tmpdeta4 In DAODetalleOrdenTrabajo.FindAllConjunto(deta2.id, tmpdeta3.Pieza.id)
                                Set Record4 = Record3.Childs.Add
                                Record4.Tag = tmpdeta4.id
                                Record4.AddItem vbNullString
                                Set item = Record4.AddItem(tmpdeta4.Pieza.nombre)
                                Record4.Expanded = True
                                'Item.HasCheckbox = True
                                Record4.AddItem tmpdeta4.Cantidad * Record3.item(2).value
                                Record4.AddItem vbNullString
                                AddTareas Record4, deta2.id, tmpdeta4.id
                            Next tmpdeta4
                        End If
                    Next tmpdeta3
                End If
            Next tmpdeta2
        End If
    Next
    Me.ReportControl.Populate
End Sub
Public Sub AgregarTareas(deta As DetalleOrdenTrabajo, Optional ByVal parent As ReportRecord = Nothing)
    Dim rec As ReportRecord
    'Set colptp = DAOTiemposProceso.FindAllByDetallePedidoIdAndPiezaId(deta.Id, pieza.Id)
    Set colptp = DAOTiemposProceso.FindAllByDetallePedidoId(deta.id)
    If colptp.count = 0 Then Exit Sub

    For Each ptp In colptp
        Set rec = parent.Childs.Add
        rec.AddItem ptp.Tarea.Tarea
    Next
End Sub

Private Sub DateTimeFecha_Change()
    Me.rbFecha.value = True
End Sub

Private Sub DateTimeFecha_GotFocus()
    Me.rbFecha.value = True
End Sub

Private Sub dtFecha2_Change()
    Me.rbFecha2.value = True
End Sub

Private Sub Form_Load()
    Me.DateTimeFecha.value = Now
    FormHelper.Customize Me
    GridEXHelper.CustomizeGrid Me.grilla_por_legajo, False, False
    GridEXHelper.CustomizeGrid Me.grilla_por_periodo, False, False
    Me.grilla_por_legajo.ItemCount = 0
    Me.grilla_por_periodo.ItemCount = 0
    LlenarPeriodoMeses
    LlenarPeriodoAño
    Me.rbFecha.value = True
    Me.rbFecha2.value = True
    ArmarColumnasPorOt
    ArmarColumnasPorPieza
    Me.TabControl1.selectedItem = 0

    dtFecha2.value = Date
End Sub
Private Sub LlenarPeriodoAño()
    Dim dto As DTOTiempoProcesoDetalle
    Set colAnio = DAOTiemposProcesosDetalles.FindAllPeriodosConProceso(TipoPeriodoAño)
    Set colAnio2 = colAnio
    cboAnios.Clear
    cboAño2.Clear
    For Each dto In colAnio

        cboAño2.AddItem dto.mostrar
        cboAño2.ItemData(cboAño2.NewIndex) = dto.indice
        cboAnios.AddItem dto.mostrar
        cboAnios.ItemData(cboAnios.NewIndex) = dto.indice
    Next
    If cboAnios.ListCount > 0 Then cboAnios.ListIndex = 0
    If cboAño2.ListCount > 0 Then cboAño2.ListIndex = 0
End Sub
Private Sub LlenarPeriodoMeses()
    Me.cboMeses.Clear
    Me.cboMeses2.Clear

    Dim dto As DTOTiempoProcesoDetalle
    Set colMeses = DAOTiemposProcesosDetalles.FindAllPeriodosConProceso(TipoPeriodoMes)
    Set colMeses2 = colMeses
    For Each dto In colMeses
        cboMeses.AddItem dto.mostrar
        cboMeses.ItemData(cboMeses.NewIndex) = dto.indice
        cboMeses2.AddItem dto.mostrar
        cboMeses2.ItemData(cboMeses.NewIndex) = dto.indice
    Next
    If cboMeses.ListCount > 0 Then cboMeses.ListIndex = 0
    If cboMeses2.ListCount > 0 Then cboMeses2.ListIndex = 0
End Sub
Private Sub LlenarListaPorPeriodo()
    Dim fecha_elegida
    Dim filtro As String


    If Me.rbFecha2 Then
        fecha_elegida = funciones.dateFormateada(CDate(Me.dtFecha2.value))

        filtro = "DATE(" & DAOTiemposProcesosDetalles.CAMPO_INICIO & ") = " & conectar.Escape(fecha_elegida) & " AND DATE(" & DAOTiemposProcesosDetalles.CAMPO_FIN & ") = " & conectar.Escape(fecha_elegida)
    ElseIf Me.rbMes2 Then
        Dim mes_ As Integer
        Dim anio_ As Long
        mes_ = colMeses2.item(cboMeses2.ItemData(cboMeses2.ListIndex)).mes
        anio_ = colMeses2.item(cboMeses2.ItemData(cboMeses2.ListIndex)).Año
        filtro = "MONTH(" & DAOTiemposProcesosDetalles.CAMPO_INICIO & " )='" & mes_ & "' AND MONTH(" & DAOTiemposProcesosDetalles.CAMPO_FIN & ")='" & mes_ & "'" _
                 & " AND YEAR(" & DAOTiemposProcesosDetalles.CAMPO_INICIO & " )='" & anio_ & "' AND YEAR(" & DAOTiemposProcesosDetalles.CAMPO_FIN & ")='" & anio_ & "'"
    ElseIf Me.rbAño2 Then
        anio_ = colAnio2.item(Me.cboAño2.ItemData(Me.cboAño2.ListIndex)).Año
        filtro = "YEAR(" & DAOTiemposProcesosDetalles.CAMPO_INICIO & " )='" & anio_ & "' AND YEAR(" & DAOTiemposProcesosDetalles.CAMPO_FIN & ")='" & anio_ & "'"
    End If


    Set detalles_per = DAOTiemposProcesosDetalles.FindAllPorPeriodoAgrupado(filtro)
    Me.grilla_por_periodo.ItemCount = 0
    Me.grilla_por_periodo.ItemCount = detalles_per.count

    Dim c As Double
    c = 0
    For Each deta In detalles_per
        c = c + deta.PlaneamientoTiempoProceso.TiempoTotalReal
    Next


    lblTotalPorPeriodo = "Total: " & funciones.FormatearDecimales(c) & " horas"
End Sub
Private Sub LlenarListaPorLegajo()
    Dim fecha_elegida
    Dim filtro As String
    If LenB(Me.txtLegajo.text) > 0 And IsNumeric(Me.txtLegajo.text) Then
        Set emple = DAOEmpleados.GetByLegajo(CLng(Me.txtLegajo.text))
        If emple Is Nothing Then
            LimpiarPorLegajo
            Exit Sub
        End If
        Me.lblEmpleado.caption = emple.Apellido & ", " & emple.nombre
        fecha_elegida = funciones.dateFormateada(CDate(DateTimeFecha.value))

        If Me.rbFecha Then
            filtro = "per.legajo=" & CLng(Me.txtLegajo.text) & " AND DATE(" & DAOTiemposProcesosDetalles.CAMPO_INICIO & " )='" & fecha_elegida & "' AND DATE(" & DAOTiemposProcesosDetalles.CAMPO_FIN & ")='" & fecha_elegida & "'"
        ElseIf Me.rbMes Then
            Dim mes_ As Integer
            Dim anio_ As Long

            mes_ = colMeses.item(cboMeses.ItemData(cboMeses.ListIndex)).mes
            anio_ = colMeses.item(cboMeses.ItemData(cboMeses.ListIndex)).Año




            filtro = "per.legajo=" & CLng(Me.txtLegajo.text) _
                     & " AND MONTH(" & DAOTiemposProcesosDetalles.CAMPO_INICIO & " )='" & mes_ & "' AND MONTH(" & DAOTiemposProcesosDetalles.CAMPO_FIN & ")='" & mes_ & "'" _
                     & " AND YEAR(" & DAOTiemposProcesosDetalles.CAMPO_INICIO & " )='" & anio_ & "' AND YEAR(" & DAOTiemposProcesosDetalles.CAMPO_FIN & ")='" & anio_ & "'"
        ElseIf Me.rbAno Then
            filtro = "per.legajo=" & CLng(Me.txtLegajo.text) & " AND YEAR(" & DAOTiemposProcesosDetalles.CAMPO_INICIO & " )='" & fecha_elegida & "' AND YEAR(" & DAOTiemposProcesosDetalles.CAMPO_FIN & ")='" & fecha_elegida & "'"
        End If


        Set detalles_leg = DAOTiemposProcesosDetalles.FindAll(filtro)
        Me.grilla_por_legajo.ItemCount = 0
        Me.grilla_por_legajo.ItemCount = detalles_leg.count
        CalcularTotalTiempos
    End If
End Sub
Public Sub CalcularTotalTiempos()
    Dim tot As Double
    For Each deta In detalles_leg
        tot = tot + deta.DiferenciaTiempos
    Next
    Me.lblTotalHorasPorLegajo = tot & " horas"
End Sub

Private Sub grilla_por_legajo_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)

    If detalles_leg.count > 0 Then
        Set deta = detalles_leg(RowIndex)
        With Values
            .value(1) = deta.PlaneamientoTiempoProceso.Tarea.id & " - " & deta.PlaneamientoTiempoProceso.Tarea.Tarea
            .value(2) = deta.PlaneamientoTiempoProceso.idpedido & "/" & deta.PlaneamientoTiempoProceso.item
            .value(3) = deta.FechaInicioTarea
            .value(4) = deta.FechaFinTarea
            '.value(5) = deta.DiferenciaTiempos
            .value(5) = deta.DiferenciaTiempoHorasMinutos
            .value(6) = deta.CantidadProcesada
            .value(7) = deta.FechaCarga
        End With
    End If
End Sub
Private Sub ArmarColumnasPorPieza()
    Me.reportPorPieza.Columns.DeleteAll
    AddColumnReportControl Me.reportPorPieza, 0, "Pieza", , True, 150
    AddColumnReportControl Me.reportPorPieza, 1, "Operarios", xtpAlignmentRight, , 50
    AddColumnReportControl Me.reportPorPieza, 2, "Tiempo Cotizado", xtpAlignmentRight, , 50
    AddColumnReportControl Me.reportPorPieza, 3, "Operarios", xtpAlignmentRight, , 50
    AddColumnReportControl Me.reportPorPieza, 4, "Promedio Tiempo", xtpAlignmentRight, , 50
    Me.reportPorPieza.PaintManager.HorizontalGridStyle = xtpGridSmallDots
    Me.reportPorPieza.PaintManager.VerticalGridStyle = xtpGridSmallDots
End Sub

Private Sub ArmarColumnasPorOt()
    Me.ReportControl.Columns.DeleteAll
    AddColumnReportControl Me.ReportControl, 0, "Item", , True, 15
    AddColumnReportControl Me.ReportControl, 1, "Detalle", xtpAlignmentLeft, , 100
    AddColumnReportControl Me.ReportControl, 2, "Cantidad", xtpAlignmentRight, , 25
    AddColumnReportControl Me.ReportControl, 3, "F.Entrega", xtpAlignmentRight, , 25
    AddColumnReportControl Me.ReportControl, 4, "Tiempo Tot", xtpAlignmentRight, , 25
    AddColumnReportControl Me.ReportControl, 5, "Tiempo Cot", xtpAlignmentRight, , 25

    Me.ReportControl.PaintManager.HorizontalGridStyle = xtpGridSmallDots
    Me.ReportControl.PaintManager.VerticalGridStyle = xtpGridSmallDots
End Sub


Private Sub grilla_por_periodo_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    On Error Resume Next
    Set deta = detalles_per.item(RowIndex)
    Values(1) = deta.Empleado.legajo & " - " & deta.Empleado.NombreCompleto
    Values(2) = funciones.FormatearDecimales(deta.PlaneamientoTiempoProceso.TiempoTotalReal) & " hs"
End Sub

Private Sub PushButton1_Click()
    Me.reportPorPieza.PaintManager.NoItemsText = "No hay piezas"
    LlenarListaPorPiezas
End Sub

Private Sub TabControl1_SelectedChanged(ByVal item As Xtremesuitecontrols.ITabControlItem)
    Me.Command1.Default = (item.index = 0)
    cmdBuscarPorPeriodo.Default = (item.index = 1)
    Me.Command2.Default = (item.index = 2)
    PushButton1.Default = (item.index = 3)

End Sub

Private Sub AgregarPieza2(ByVal Pieza As Pieza, Optional ByVal parent As ReportRecord = Nothing, Optional deta As DetalleOrdenTrabajo)
    Dim rec As ReportRecord
    If parent Is Nothing Then
        Set rec = Me.ReportControl.Records.Add
    Else
        Set rec = parent.Childs.Add
    End If

    rec.AddItem Pieza.nombre
    rec.AddItem Pieza.Cantidad
    rec.Tag = Pieza.id
    rec.Expanded = False
    AgregarTareas deta, rec

    Dim piezaHija As Pieza
    For Each piezaHija In Pieza.PiezasHijas
        AgregarPieza2 piezaHija, rec, deta
    Next piezaHija

End Sub

Private Sub txtPieza_DblClick()
    frmListarStock_seleccion.Text1 = funciones.quePiezaElegidabusqueda
    frmListarStock_seleccion.Show 1
    Me.txtPieza = funciones.quePiezaElegidaDetalle
    id_pieza_elegida = funciones.quePiezaElegida

End Sub
