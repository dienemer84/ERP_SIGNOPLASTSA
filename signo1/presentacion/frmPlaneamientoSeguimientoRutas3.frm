VERSION 5.00
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#12.0#0"; "CODEJO~3.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmPlaneamientoSeguimientoRutas3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seguimiento de rutas"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   16920
   Icon            =   "frmPlaneamientoSeguimientoRutas3.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   16920
   Begin XtremeSuiteControls.GroupBox GroupBox2 
      Height          =   6660
      Left            =   8160
      TabIndex        =   10
      Top             =   1035
      Width           =   8685
      _Version        =   786432
      _ExtentX        =   15319
      _ExtentY        =   11747
      _StockProps     =   79
      Caption         =   "Tiempos de la tarea"
      UseVisualStyle  =   -1  'True
      Begin XtremeReportControl.ReportControl ReportControlDetalles 
         Height          =   6345
         Left            =   75
         TabIndex        =   11
         Top             =   225
         Width           =   8535
         _Version        =   786432
         _ExtentX        =   15055
         _ExtentY        =   11192
         _StockProps     =   64
         BorderStyle     =   3
         PreviewMode     =   -1  'True
         AllowColumnRemove=   0   'False
         AllowColumnReorder=   0   'False
         AllowColumnSort =   0   'False
         MultipleSelection=   0   'False
         ShowHeaderRows  =   -1  'True
      End
   End
   Begin XtremeSuiteControls.GroupBox fraDetalles 
      Height          =   6660
      Left            =   60
      TabIndex        =   8
      Top             =   1035
      Width           =   8070
      _Version        =   786432
      _ExtentX        =   14235
      _ExtentY        =   11747
      _StockProps     =   79
      Caption         =   "Detalles de la Orden de Trabajo"
      UseVisualStyle  =   -1  'True
      Begin XtremeReportControl.ReportControl ReportControl 
         Height          =   6345
         Left            =   75
         TabIndex        =   9
         Top             =   225
         Width           =   7920
         _Version        =   786432
         _ExtentX        =   13970
         _ExtentY        =   11192
         _StockProps     =   64
         BorderStyle     =   3
         PreviewMode     =   -1  'True
         AllowColumnRemove=   0   'False
         AllowColumnReorder=   0   'False
         AllowColumnSort =   0   'False
         MultipleSelection=   0   'False
         ShowHeaderRows  =   -1  'True
      End
   End
   Begin XtremeSuiteControls.GroupBox fraDatosOT 
      Height          =   855
      Left            =   7755
      TabIndex        =   2
      Tag             =   "Datos de la Orden de Trabajo Nº "
      Top             =   90
      Width           =   9075
      _Version        =   786432
      _ExtentX        =   16007
      _ExtentY        =   1508
      _StockProps     =   79
      Caption         =   "Datos de la Orden de Trabajo Nº "
      UseVisualStyle  =   -1  'True
      Begin VB.Label lblCliente 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   165
         TabIndex        =   6
         Top             =   240
         Width           =   525
      End
      Begin VB.Label lblFechaCreado 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Creada:"
         Height          =   195
         Left            =   165
         TabIndex        =   5
         Top             =   525
         Width           =   1050
      End
      Begin VB.Label lblFechaEntrega 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Entrega:"
         Height          =   195
         Left            =   4305
         TabIndex        =   4
         Top             =   525
         Width           =   1095
      End
      Begin VB.Label lblEstado 
         AutoSize        =   -1  'True
         Caption         =   "Estado:"
         Height          =   195
         Left            =   4830
         TabIndex        =   3
         Top             =   240
         Width           =   540
      End
   End
   Begin XtremeSuiteControls.PushButton cmdBuscar 
      Height          =   345
      Left            =   1110
      TabIndex        =   1
      Top             =   585
      Width           =   2100
      _Version        =   786432
      _ExtentX        =   3704
      _ExtentY        =   609
      _StockProps     =   79
      Caption         =   "Buscar"
      ForeColor       =   9126421
      Appearance      =   6
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   855
      Left            =   3345
      TabIndex        =   7
      Tag             =   "Datos de la Orden de Trabajo Nº "
      Top             =   90
      Width           =   4305
      _Version        =   786432
      _ExtentX        =   7594
      _ExtentY        =   1508
      _StockProps     =   79
      Caption         =   "Filtrar por"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.PushButton btnClearTarea 
         Height          =   285
         Left            =   3915
         TabIndex        =   15
         Top             =   330
         Width           =   240
         _Version        =   786432
         _ExtentX        =   423
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "X"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cboTarea 
         Height          =   315
         Left            =   735
         TabIndex        =   14
         Top             =   330
         Width           =   3105
         _Version        =   786432
         _ExtentX        =   5477
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Sorted          =   -1  'True
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tarea:"
         Height          =   195
         Left            =   180
         TabIndex        =   13
         Top             =   360
         Width           =   465
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtOTNro 
      Height          =   300
      Left            =   2445
      TabIndex        =   12
      Top             =   210
      Width           =   750
      _Version        =   786432
      _ExtentX        =   1323
      _ExtentY        =   529
      _StockProps     =   77
      BackColor       =   -2147483643
      Alignment       =   1
   End
   Begin XtremeSuiteControls.TaskDialog taskDialog 
      Left            =   0
      Top             =   0
      _Version        =   786432
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   0
      WindowTitle     =   "TaskDialog1"
   End
   Begin VB.Image Image 
      Height          =   720
      Left            =   240
      Picture         =   "frmPlaneamientoSeguimientoRutas3.frx":000C
      Top             =   180
      Width           =   720
   End
   Begin VB.Label lblOT 
      AutoSize        =   -1  'True
      Caption         =   "Orden de Trabajo"
      Height          =   195
      Left            =   1110
      TabIndex        =   0
      Top             =   255
      Width           =   1245
   End
   Begin VB.Menu mnuTarea 
      Caption         =   "Tarea"
      Visible         =   0   'False
      Begin VB.Menu mnuFinalizarTarea 
         Caption         =   "Finalizar Tarea de la Pieza..."
      End
      Begin VB.Menu mnuNNC 
         Caption         =   "Crear Nota No Conformidad..."
      End
   End
   Begin VB.Menu mnuPieza 
      Caption         =   "Pieza"
      Visible         =   0   'False
      Begin VB.Menu mnuAgregarTarea 
         Caption         =   "Agregar Tarea a la Pieza..."
      End
      Begin VB.Menu mnuAgregarRemito 
         Caption         =   "Agregar a Remito..."
      End
      Begin VB.Menu mnuArchivoPieza 
         Caption         =   "Archivos Asociados..."
      End
      Begin VB.Menu mnuArchivoDetalleOT 
         Caption         =   "Agregar archivo al detalle de la OT..."
      End
   End
   Begin VB.Menu mnuTiempo 
      Caption         =   "Tiempo"
      Visible         =   0   'False
      Begin VB.Menu mnuIniciarTiempo 
         Caption         =   "Iniciar Tiempo Nuevo..."
      End
      Begin VB.Menu mnuInicializarFinalizarTiempo 
         Caption         =   "Iniciar y Finalizar Tiempo Nuevo..."
      End
      Begin VB.Menu mnuSeparador 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFinalizarTiempo 
         Caption         =   "Finalizar Tiempo Seleccionado..."
      End
      Begin VB.Menu mnuEditarTiempo 
         Caption         =   "Editar Tiempo Seleccionado..."
      End
   End
End
Attribute VB_Name = "frmPlaneamientoSeguimientoRutas3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Ot As OrdenTrabajo
Private detalles As Collection  'of PlaneamientoTiempoProcesoDetalle
Private tareaId As Long

Private Sub btnClearTarea_Click()
    Me.cboTarea.ListIndex = -1
    cmdBuscar_Click
End Sub

Private Sub cboTarea_Click()
    If Me.cboTarea.ListIndex <> -1 Then
        tareaId = Me.cboTarea.ItemData(Me.cboTarea.ListIndex)
        LlenarData
    Else
        tareaId = 0
    End If
End Sub

Public Sub cmdBuscar_Click()
    If Not IsNumeric(Me.txtOTNro.Text) Then Exit Sub
    Set Ot = DAOOrdenTrabajo.FindById(Me.txtOTNro.Text)
    If Ot Is Nothing Then
        MsgBox "La Orden de Trabajo Nº " & Me.txtOTNro.Text & " no existe.", vbInformation + vbOKOnly
    Else

        Me.fraDatosOT.caption = Me.fraDatosOT.Tag & Ot.Id
        Me.lblCliente.caption = "Cliente: " & Ot.cliente.razon
        Me.lblFechaCreado.caption = "Fecha Creada: " & Ot.fechaCreado
        Me.lblFechaEntrega.caption = "Fecha Entrega: " & Ot.FechaEntrega
        Me.lblEstado.caption = "Estado: " & funciones.estado_pedido(Ot.estado)

        Dim tareas As Collection
        Dim tareaFilter As String
        tareaFilter = "t.id in (SELECT DISTINCT codigoTarea FROM PlaneamientoTiemposProcesos WHERE idPedido = " & Ot.Id & ")"
        Set tareas = DAOTareas.FindAll(tareaFilter)
        Me.cboTarea.Clear
        Dim T As clsTarea
        For Each T In tareas
            Me.cboTarea.AddItem T.Description
            Me.cboTarea.ItemData(Me.cboTarea.NewIndex) = T.Id
        Next T
        Me.cboTarea.ListIndex = -1

        LlenarData
    End If
End Sub

Private Sub LlenarData()
    Set Ot.detalles = DAODetalleOrdenTrabajo.FindAllByOrdenTrabajo(Ot.Id)
    CargarDetallesOT
End Sub


Private Sub CargarDetallesOT()
    Dim deta As DetalleOrdenTrabajo
    Dim P As Pieza
    Dim tmpdeta2 As DetalleOTConjuntoDTO
    Dim tmpdeta3 As DetalleOTConjuntoDTO
    Dim tmpdeta4 As DetalleOTConjuntoDTO


    Me.ReportControl.Records.DeleteAll
    Me.ReportControl.Populate
    Me.ReportControlDetalles.Records.DeleteAll
    Me.ReportControlDetalles.Populate

    Set detalles = New Collection

    Dim record As ReportRecord
    Dim Record2 As ReportRecord
    Dim Record3 As ReportRecord
    Dim Record4 As ReportRecord

    Dim item As ReportRecordItem

    For Each deta In Ot.detalles
        Set record = Me.ReportControl.Records.Add
        record.Tag = deta.Id
        record.AddItem deta.item
        Set item = record.AddItem(deta.Pieza.nombre)
        'Item.HasCheckbox = True

        record.PreviewText = deta.Nota
        record.AddItem deta.CantidadPedida
        record.AddItem deta.FechaEntrega
        'AddTareas record, deta.Id, deta.pieza
        AddTareas record, deta.Id


        If deta.Pieza.EsConjunto Then
            record.Expanded = True
            'For Each tmpdeta2 In DAODetalleOrdenTrabajo.FindAllConjunto(deta.Id, deta.pieza.Id)
            For Each tmpdeta2 In DAODetalleOrdenTrabajo.FindAllConjunto(deta.Id, deta.Pieza.Id)

                Set Record2 = record.Childs.Add()
                Record2.Tag = tmpdeta2.Id
                Record2.AddItem vbNullString
                Set item = Record2.AddItem(tmpdeta2.Pieza.nombre)

                'Item.HasCheckbox = True
                Record2.AddItem tmpdeta2.Cantidad * record.item(2).value
                Record2.AddItem vbNullString

                'AddTareas Record2, tmpdeta2.Id, tmpdeta2.pieza
                AddTareas Record2, deta.Id, tmpdeta2.Id

                If tmpdeta2.Pieza.EsConjunto Then
                    Record2.Expanded = True
                    For Each tmpdeta3 In DAODetalleOrdenTrabajo.FindAllConjunto(deta.Id, tmpdeta2.Pieza.Id)

                        Set Record3 = Record2.Childs.Add
                        Record3.Tag = tmpdeta3.Id
                        Record3.AddItem vbNullString
                        Set item = Record3.AddItem(tmpdeta3.Pieza.nombre)

                        'Item.HasCheckbox = True
                        Record3.AddItem tmpdeta3.Cantidad * Record2.item(2).value
                        Record3.AddItem vbNullString

                        'AddTareas Record3, tmpdeta3.Id, tmpdeta3.pieza
                        AddTareas Record3, deta.Id, tmpdeta3.Id


                        If tmpdeta3.Pieza.EsConjunto Then
                            Record3.Expanded = True
                            For Each tmpdeta4 In DAODetalleOrdenTrabajo.FindAllConjunto(deta.Id, tmpdeta3.Pieza.Id)

                                Set Record4 = Record3.Childs.Add
                                Record4.Tag = tmpdeta4.Id
                                Record4.AddItem vbNullString
                                Set item = Record4.AddItem(tmpdeta4.Pieza.nombre)

                                'Item.HasCheckbox = True
                                Record4.AddItem tmpdeta4.Cantidad * Record3.item(2).value
                                Record4.AddItem vbNullString

                                'AddTareas Record4, tmpdeta4.Id, tmpdeta4.pieza
                                AddTareas Record4, deta.Id, tmpdeta4.Id
                                'Record4.Expanded = True

                            Next tmpdeta4
                        End If

                    Next tmpdeta3
                End If
            Next tmpdeta2
        End If
    Next

    Me.ReportControl.Populate
End Sub

Private Sub Form_Initialize()
    txtOTNro.SetFocus
End Sub

Private Sub Form_Load()
    Customize Me

    Me.ReportControl.Columns.DeleteAll

    Dim Column As ReportColumn
    Set Column = Me.ReportControl.Columns.Add(0, "Item", 10, True)
    Column.Icon = 0
    Column.Sortable = False
    Column.AllowDrag = False
    Column.AllowRemove = False

    Set Column = Me.ReportControl.Columns.Add(1, "Detalle", 65, True)
    Column.Icon = 0
    Column.Sortable = False
    Column.TreeColumn = True
    Column.AllowDrag = False
    Column.AllowRemove = False

    Set Column = Me.ReportControl.Columns.Add(2, "Cantidad", 12, True)
    Column.Icon = 0
    Column.Sortable = False
    Column.AllowDrag = False
    Column.AllowRemove = False
    Column.Alignment = xtpAlignmentRight


    Set Column = Me.ReportControl.Columns.Add(3, "F. Entrega", 17, True)
    Column.Icon = 0
    Column.Sortable = False
    Column.AllowDrag = False
    Column.AllowRemove = False
    Column.Alignment = xtpAlignmentRight


    Me.ReportControl.PaintManager.HorizontalGridStyle = xtpGridSmallDots
    Me.ReportControl.PaintManager.VerticalGridStyle = xtpGridSmallDots




    'columnas detalles

    Me.ReportControlDetalles.PaintManager.HorizontalGridStyle = xtpGridSmallDots
    Me.ReportControlDetalles.PaintManager.VerticalGridStyle = xtpGridSmallDots

    Set Column = Me.ReportControlDetalles.Columns.Add(0, "Leg", 3, True)
    Column.Icon = 0
    Column.Sortable = False
    Column.AllowDrag = False
    Column.AllowRemove = False
    Column.Alignment = xtpAlignmentRight

    Set Column = Me.ReportControlDetalles.Columns.Add(1, "Empleado", 10, True)
    Column.Icon = 0
    Column.Sortable = False
    Column.AllowDrag = False
    Column.AllowRemove = False

    Set Column = Me.ReportControlDetalles.Columns.Add(2, "Fecha Inicio", 13, True)
    Column.Icon = 0
    Column.Sortable = False
    Column.AllowDrag = False
    Column.AllowRemove = False
    Column.Alignment = xtpAlignmentRight

    Set Column = Me.ReportControlDetalles.Columns.Add(3, "Fecha Fin", 13, True)
    Column.Icon = 0
    Column.Sortable = False
    Column.AllowDrag = False
    Column.AllowRemove = False
    Column.Alignment = xtpAlignmentRight

    Set Column = Me.ReportControlDetalles.Columns.Add(4, "Dif Tiempo", 7, True)
    Column.Icon = 0
    Column.Sortable = False
    Column.AllowDrag = False
    Column.AllowRemove = False
    Column.Alignment = xtpAlignmentRight


    Set Column = Me.ReportControlDetalles.Columns.Add(5, "Cant Proc", 7, True)
    Column.Icon = 0
    Column.Sortable = False
    Column.AllowDrag = False
    Column.AllowRemove = False
    Column.Alignment = xtpAlignmentRight

    Set Column = Me.ReportControlDetalles.Columns.Add(6, "Fecha Carga", 13, True)
    Column.Icon = 0
    Column.Sortable = False
    Column.AllowDrag = False
    Column.AllowRemove = False
    Column.Alignment = xtpAlignmentRight

    Me.cboTarea.Clear



End Sub

Private Sub AddTareas(ByRef rec As ReportRecord, ByRef idDetallePedido As Long, Optional ByRef idDetallePedidoConj As Long = 0)    'ByRef P As pieza)
    Dim rechijo As ReportRecord
    Dim item As ReportRecordItem

    Dim ptp As PlaneamientoTiempoProceso
    Dim finalizada As String

    'For Each ptp In DAOTiemposProceso.FindAllByDetallePedidoIdAndPiezaId(idDetallePedido, P.Id)
    For Each ptp In DAOTiemposProceso.FindAllByDetallePedidoId(idDetallePedido, idDetallePedidoConj, , , tareaId)
        finalizada = vbNullString
        Set rechijo = rec.Childs.Add
        rechijo.Tag = (ptp.Id * -1)    'negativo para distinguir de las piezas
        rechijo.AddItem vbNullString


        Set item = rechijo.AddItem("Tarea: " & ptp.Tarea.Id & " - " & ptp.Tarea.Tarea & finalizada & " (" & ptp.Id & ")")
        rechijo.PreviewText = ptp.Tarea.descripcion

        'Item.HasCheckbox = True
        rechijo.AddItem vbNullString

        If ptp.FINALIZADO Then finalizada = "FINALIZADA"
        rechijo.AddItem finalizada
    Next ptp

    If rec.Childs.count > 0 Then rec.Expanded = False    'me cierra los nodos que tienen tareas dentro

End Sub


Private Sub mnuAgregarRemito_Click()
    Dim nroremito As String
    nroremito = 0
    If IsSomething(Selecciones.RemitoElegido) Then nroremito = Selecciones.RemitoElegido.numero


    nroremito = InputBox("Ingrese el Nº de remito donde agregar la pieza.", "Remitar", nroremito)
    If LenB(nroremito) = 0 Or nroremito = "0" Or Not IsNumeric(nroremito) Then
        MsgBox "Ingrese un nro de remito válido.", vbExclamation
    Else
        Dim rto As Remito
        Set rto = DAORemitoS.FindByNumero(Val(nroremito))

        If Not IsSomething(rto) Then Exit Sub

        Set rto.detalles = DAORemitoSDetalle.FindAllByRemito(rto.Id)

        If rto.detalles.count = funciones.itemsPorRemito Then
            MsgBox "El remito llego al limite de items, cree otro remito.", vbCritical
            Exit Sub
        End If


        If rto.estado = RemitoPendiente Then
            Set Selecciones.RemitoElegido = rto

            If Me.ReportControl.SelectedRows.count > 0 Then
                Dim rtoDetalle As New remitoDetalle
                rtoDetalle.Origen = OrigenRemitoConcepto
                'rtoDetalle.cantidad = CDbl(Values(4))
                rtoDetalle.facturable = True
                rtoDetalle.Facturado = False
                rtoDetalle.FEcha = Now

                Dim row As ReportRow
                Dim col As Collection
                Dim deta As DetalleOrdenTrabajo
                Dim detadto As DetalleOTConjuntoDTO
                Set row = Me.ReportControl.SelectedRows(0)

                If row.ParentRow Is Nothing Then    'es el detalle de la ot
                    Set deta = DAODetalleOrdenTrabajo.FindById(row.record.Tag)
                    If deta Is Nothing Then Exit Sub
                    rtoDetalle.Concepto = deta.Pieza.nombre
                Else    'es el detalle de algun detalle
                    Set detadto = DAODetalleOrdenTrabajo.FindConjuntoById(row.record.Tag)
                    If detadto Is Nothing Then Exit Sub
                    rtoDetalle.Concepto = detadto.Pieza.nombre
                End If



                Dim Cant As String
                Cant = InputBox("Ingrese la cantidad a remitar de la pieza elegida.", "Remitar", 0)
                If LenB(Cant) = 0 Or Cant = "0" Or Not IsNumeric(Cant) Then
                    MsgBox "Ingrese una cantidad válida.", vbExclamation
                    Exit Sub
                End If

                rtoDetalle.Cantidad = Cant

                rto.detalles.Add rtoDetalle

                If Not DAORemitoS.Save(rto, True) Then
                    MsgBox "Se produjo algun error al guardar!", vbCritical, "Error"
                Else
                    MsgBox "La pieza fue agregada al remito.", vbInformation
                End If

            End If

        Else
            MsgBox "El estado del remito no es válido.", vbExclamation
        End If
    End If
End Sub

Private Sub mnuAgregarTarea_Click()


    If Me.ReportControl.SelectedRows.count > 0 Then
        Dim F As New frmPlaneamientoAgregarTareaAProceso

        Dim row As ReportRow
        Dim col As Collection
        Dim deta As DetalleOrdenTrabajo
        Dim detadto As DetalleOTConjuntoDTO
        Set row = Me.ReportControl.SelectedRows(0)

        If row.ParentRow Is Nothing Then    'es el detalle de la ot
            Set deta = DAODetalleOrdenTrabajo.FindById(row.record.Tag)
            If deta Is Nothing Then Exit Sub

            F.PIEZA_ID = deta.Pieza.Id
            F.idDetallePedido = deta.Id
        Else    'es el detalle de algun detalle
            Set detadto = DAODetalleOrdenTrabajo.FindConjuntoById(row.record.Tag)
            If detadto Is Nothing Then Exit Sub

            F.PIEZA_ID = detadto.Pieza.Id
            F.idDetallePedidoConjunto = detadto.Id
            F.idDetallePedido = detadto.idDetallePedido
        End If

        F.pedido_id = Ot.Id
        F.Show 1
        If TareaAgregada Then CargarDetallesOT
    End If


End Sub

Private Sub mnuArchivoDetalleOT_Click()

    Dim row As ReportRow
    Dim col As Collection
    Dim deta As DetalleOrdenTrabajo
    Dim detadto As DetalleOTConjuntoDTO
    Set row = Me.ReportControl.SelectedRows(0)

    Dim frmar2 As New frmArchivos2


    If row.ParentRow Is Nothing Then        'es el detalle de la ot
        frmar2.Origen = OrigenArchivos.OA_OrdenesTrabajoDetalle
        Set deta = DAODetalleOrdenTrabajo.FindById(row.record.Tag)
        If deta Is Nothing Then Exit Sub

        frmar2.ObjetoId = deta.Id
        frmar2.caption = "OT Nº " & Ot.IdFormateado & " - Item " & deta.item & " [" & deta.Pieza.nombre & "]"

    Else        'es el detalle de algun detalle
        frmar2.Origen = OrigenArchivos.OA_OrdenesTrabajoDetalleConjunto
        Set detadto = DAODetalleOrdenTrabajo.FindConjuntoById(row.record.Tag)
        If detadto Is Nothing Then Exit Sub

        Set deta = DAODetalleOrdenTrabajo.FindById(detadto.idDetallePedido)

        frmar2.ObjetoId = detadto.Id
        frmar2.caption = "OT Nº " & Ot.IdFormateado & " - Item " & deta.item & " - Subitem " & detadto.Id & " [" & deta.Pieza.nombre & "]"
    End If

    frmar2.Show
End Sub

Private Sub mnuArchivoPieza_Click()

    Dim row As ReportRow
    Dim col As Collection
    Dim deta As DetalleOrdenTrabajo
    Dim detadto As DetalleOTConjuntoDTO
    Set row = Me.ReportControl.SelectedRows(0)

    Dim frmar1 As New frmArchivos2
    frmar1.Origen = OrigenArchivos.OA_Piezas

    If row.ParentRow Is Nothing Then        'es el detalle de la ot
        Set deta = DAODetalleOrdenTrabajo.FindById(row.record.Tag)
        If deta Is Nothing Then Exit Sub
        frmar1.ObjetoId = deta.Pieza.Id
        frmar1.caption = "Pieza " & deta.Pieza.nombre
    Else        'es el detalle de algun detalle
        Set detadto = DAODetalleOrdenTrabajo.FindConjuntoById(row.record.Tag)
        If detadto Is Nothing Then Exit Sub
        frmar1.ObjetoId = detadto.Pieza.Id
        frmar1.caption = "Pieza " & detadto.Pieza.nombre
    End If

    frmar1.Show
End Sub

Private Sub mnuEditarTiempo_Click()
    Dim F As New frmPlaneamientoEdicionTiempo
    F.TEdicion = TipoEdicion.editar
    F.PlaneamientoTiempoProcesoId = Me.ReportControl.SelectedRows(0).record.Tag * -1
    Set F.detalleEditar = DAOTiemposProcesosDetalles.FindById(Me.ReportControlDetalles.SelectedRows(0).record.Tag)
    F.Show 1
    If TareaAgregada Then
        LlenarDetalles
        ReportControl_SelectionChanged
    End If
End Sub

Private Sub mnuFinalizarTarea_Click()
'ANTES CHEQUEAR QUE SE PUEDA FINALIZAR
    Dim ret As Integer
    ret = DAOTiemposProceso.CanFinalize(CLng(Me.ReportControl.SelectedRows(0).record.Tag * -1))

    If ret = 1 Then
        If MsgBox("¿Desea finalizar la tarea?", vbQuestion + vbYesNo) = vbYes Then
            If DAOTiemposProceso.Finalize(CLng(Me.ReportControl.SelectedRows(0).record.Tag * -1)) Then
                Me.ReportControl.SelectedRows(0).record.item(3).caption = "FINALIZADA"
                MsgBox "Tarea Finalizada", vbInformation


                Dim ptp As PlaneamientoTiempoProceso
                Set ptp = DAOTiemposProceso.FindById(CLng(Me.ReportControl.SelectedRows(0).record.Tag * -1))
                If IsSomething(ptp) Then
                    If ptp.Tarea.Id = 15 Then    'archivo punzonado then
                        'preguntar si quiere agergar al detalle o a la pieza

                        Me.taskDialog.Reset
                        Me.taskDialog.MessageBoxStyle = True
                        Me.taskDialog.WindowTitle = "Subir archivo"
                        Me.taskDialog.MainInstructionText = "¿Donde desea subir el programa de radan y su correspondiente PDF?"
                        Me.taskDialog.ContentText = "Elija donde subir los archivos."
                        taskDialog.RelativePosition = False

                        Me.taskDialog.CommonButtons = 0
                        taskDialog.CommonButtons = taskDialog.CommonButtons Or xtpTaskButtonOk
                        'taskDialog.CommonButtons = taskDialog.CommonButtons Or xtpTaskButtonCancel


                        taskDialog.AddRadioButton "A la pieza", 1
                        taskDialog.AddRadioButton "Al detalle de la OT", 2
                        taskDialog.AddRadioButton "No voy a subir archivos", 3
                        taskDialog.DefaultRadioButton = 3

                        taskDialog.MainIcon = xtpTaskIconInformation
                        taskDialog.ShowDialog

                        If Me.taskDialog.DefaultRadioButton <> 3 Then
                            Dim row As ReportRow
                            Set row = Me.ReportControl.SelectedRows.row(0)
                            Me.ReportControl.SelectedRows.DeleteAll
                            Me.ReportControl.SelectedRows.Add row.ParentRow

                            If Me.taskDialog.DefaultRadioButton = 1 Then
                                mnuArchivoPieza_Click
                            ElseIf Me.taskDialog.DefaultRadioButton = 2 Then
                                mnuArchivoDetalleOT_Click
                            End If

                        End If

                    End If
                End If


            Else
                MsgBox "Hubo un error al intentar finalizar la tarea.", vbError
            End If
        End If
    Else
        If ret = -1 Then
            MsgBox "No se puede finalizar ya que tiene tiempos iniciados sin finalizar.", vbExclamation
        ElseIf ret = -2 Then
            MsgBox "No se puede finalizar ya que la cantidad procesada de todos los tiempos es igual a cero.", vbExclamation
        ElseIf ret = -3 Then
            MsgBox "La tarea ya se encuentra finalizada.", vbExclamation
        End If
    End If
End Sub

Private Sub mnuFinalizarTiempo_Click()
    Dim F As New frmPlaneamientoEdicionTiempo
    F.TEdicion = TipoEdicion.finalizar
    F.PlaneamientoTiempoProcesoId = Me.ReportControl.SelectedRows(0).record.Tag * -1
    Set F.detalle = DAOTiemposProcesosDetalles.FindById(Me.ReportControlDetalles.SelectedRows(0).record.Tag)
    F.Show 1
    If TareaAgregada Then
        LlenarDetalles
        ReportControl_SelectionChanged
    End If
End Sub

Private Sub mnuInicializarFinalizarTiempo_Click()
    Dim F As New frmPlaneamientoEdicionTiempo
    F.TEdicion = TipoEdicion.IniciarFinalizar
    F.PlaneamientoTiempoProcesoId = Me.ReportControl.SelectedRows(0).record.Tag * -1
    F.Show 1
    If TareaAgregada Then
        LlenarDetalles
        ReportControl_SelectionChanged
    End If
End Sub

Private Sub mnuIniciarTiempo_Click()
    Dim F As New frmPlaneamientoEdicionTiempo
    F.TEdicion = TipoEdicion.iniciar
    F.PlaneamientoTiempoProcesoId = Me.ReportControl.SelectedRows(0).record.Tag * -1
    F.Show 1

    If TareaAgregada Then
        LlenarDetalles
        ReportControl_SelectionChanged
    End If
End Sub


Private Sub mnuNNC_Click()
    Dim idTP As Long
    idTP = CLng(Me.ReportControl.SelectedRows(0).record.Tag * -1)    '-> negativos son tareas-> la unica forma de poder diferenciar



    Dim F As New frmNotaNoConformidad
    F.idTiempoProceso = idTP

    F.Show





End Sub





Private Sub ReportControl_MouseDown(Button As Integer, Shift As Integer, x As Long, y As Long)
    If Button = 2 Then
        Dim hitinfo As ReportHitTestInfo
        Set hitinfo = Me.ReportControl.HitTest(x, y)
        If Not hitinfo Is Nothing Then
            If Not hitinfo.row Is Nothing Then
                Me.ReportControl.SelectedRows.DeleteAll
                Me.ReportControl.SelectedRows.Add hitinfo.row
                Me.ReportControl.FocusedRow = hitinfo.row
                If CLng(hitinfo.row.record.Tag) < 0 Then    'es tarea
                    Me.PopupMenu Me.mnuTarea
                Else
                    Me.PopupMenu Me.mnuPieza
                End If
            End If
        End If
    End If
End Sub

Private Sub ReportControl_RowDblClick(ByVal row As XtremeReportControl.IReportRow, ByVal item As XtremeReportControl.IReportRecordItem)
    row.Expanded = Not row.Expanded
End Sub

Public Sub ReportControl_SelectionChanged()

    If Me.ReportControl.SelectedRows.count > 0 Then
        Dim row As ReportRow
        Set row = Me.ReportControl.SelectedRows(0)

        If CLng(row.record.Tag) < 0 Then    ' es tarea
            Set detalles = DAOTiemposProcesosDetalles.FindAllByTiempoProceso(-1 * CLng(row.record.Tag))
        Else    'es pieza
            Set detalles = New Collection
        End If

        LlenarDetalles
    End If

End Sub

Private Sub LlenarDetalles()
    Me.ReportControlDetalles.Records.DeleteAll
    Me.ReportControlDetalles.Populate

    Dim det As PlaneamientoTiempoProcesoDetalle
    Dim record As ReportRecord

    For Each det In detalles
        Set record = Me.ReportControlDetalles.Records.Add
        record.Tag = det.Id

        If det.Empleado Is Nothing Then
            record.AddItem Empty
            record.AddItem Empty
        Else
            record.AddItem det.Empleado.legajo
            record.AddItem det.Empleado.NombreAbreviado
        End If

        If CDbl(det.FechaInicioTarea) = 0 Then
            record.AddItem Empty
        Else
            record.AddItem det.FechaInicioTarea
        End If


        If CDbl(det.FechaFinTarea) = 0 Then
            record.AddItem Empty
        Else
            record.AddItem det.FechaFinTarea
        End If

        record.AddItem det.DiferenciaTiempoHorasMinutos
        record.AddItem det.CantidadProcesada
        record.AddItem det.FechaCarga
    Next det

    Me.ReportControlDetalles.Populate
End Sub

Private Sub ReportControlDetalles_MouseDown(Button As Integer, Shift As Integer, x As Long, y As Long)
    If Button = 2 Then
        If Me.ReportControl.SelectedRows.count > 0 Then

            If CLng(Me.ReportControl.SelectedRows(0).record.Tag) < 0 Then    ' es tarea then

                Dim hitinfo As ReportHitTestInfo
                Set hitinfo = Me.ReportControlDetalles.HitTest(x, y)

                Me.mnuFinalizarTiempo.Enabled = False

                If Not hitinfo Is Nothing Then
                    If Not hitinfo.row Is Nothing Then
                        Me.ReportControlDetalles.SelectedRows.DeleteAll
                        Me.ReportControlDetalles.SelectedRows.Add hitinfo.row
                        Me.ReportControlDetalles.FocusedRow = hitinfo.row

                        Dim de As PlaneamientoTiempoProcesoDetalle
                        Set de = detalles.item(CStr(Me.ReportControlDetalles.FocusedRow.record.Tag))
                        Me.mnuFinalizarTiempo.Enabled = (CDbl(de.FechaFinTarea) = 0)
                        Me.mnuEditarTiempo.Enabled = True    'poner permiso
                    End If
                End If

                Dim ptp As PlaneamientoTiempoProceso
                Set ptp = DAOTiemposProceso.FindById(CLng(Me.ReportControl.SelectedRows(0).record.Tag * -1))
                If IsSomething(ptp) Then
                    If ptp.FINALIZADO Then    'no muestro menu si ya esta finalizado
                        MsgBox "No se puede operar con los tiempos ya que la tarea esta finalizada.", vbInformation
                    Else
                        Me.PopupMenu Me.mnuTiempo
                    End If
                Else
                    Me.PopupMenu Me.mnuTiempo
                End If
            End If
        End If
    End If

End Sub

Private Sub txtOTNro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdBuscar_Click
End Sub
