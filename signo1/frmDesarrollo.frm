VERSION 5.00
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#12.0#0"; "CODEJO~3.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#12.0#0"; "CODEJO~1.OCX"
Begin VB.Form frmDesarrollo 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Desarrollo"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   18450
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDesarrollo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   18450
   Begin XtremeReportControl.ReportControl reportControlPiezas 
      Height          =   7395
      Left            =   60
      TabIndex        =   8
      Top             =   105
      Width           =   7365
      _Version        =   786432
      _ExtentX        =   12991
      _ExtentY        =   13044
      _StockProps     =   64
      BorderStyle     =   3
      MultipleSelection=   0   'False
   End
   Begin VB.Frame fraManoObra 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Mano de Obra"
      Height          =   3720
      Left            =   7515
      TabIndex        =   2
      Top             =   3765
      Width           =   10815
      Begin XtremeReportControl.ReportControl reportControlManoObra 
         Height          =   3000
         Left            =   90
         TabIndex        =   3
         Top             =   255
         Width           =   10605
         _Version        =   786432
         _ExtentX        =   18706
         _ExtentY        =   5292
         _StockProps     =   64
         BorderStyle     =   3
         MultipleSelection=   0   'False
      End
      Begin VB.Label lblCostoManoObra 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "Costo $: "
         Height          =   195
         Left            =   9360
         TabIndex        =   7
         Tag             =   "Costo $: "
         Top             =   3360
         Width           =   660
      End
   End
   Begin VB.Frame fraMateriales 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Materiales"
      Height          =   3720
      Left            =   7515
      TabIndex        =   0
      Top             =   45
      Width           =   10815
      Begin XtremeReportControl.ReportControl reportControlMateriales 
         Height          =   3000
         Left            =   105
         TabIndex        =   1
         Top             =   225
         Width           =   10605
         _Version        =   786432
         _ExtentX        =   18706
         _ExtentY        =   5292
         _StockProps     =   64
         BorderStyle     =   3
         MultipleSelection=   0   'False
      End
      Begin VB.Label lblCostoMateriales 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "Costo $: "
         Height          =   195
         Left            =   9270
         TabIndex        =   6
         Tag             =   "Costo $: "
         Top             =   3360
         Width           =   660
      End
      Begin VB.Label lblTotalM2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "Total M2/Ml: "
         Height          =   195
         Left            =   6570
         TabIndex        =   5
         Tag             =   "Total M2/Ml: "
         Top             =   3360
         Width           =   930
      End
      Begin VB.Label lblTotalKG 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "Total KG: "
         Height          =   195
         Left            =   4050
         TabIndex        =   4
         Tag             =   "Total KG: "
         Top             =   3360
         Width           =   705
      End
   End
   Begin XtremeCommandBars.ImageManager ImageManager1 
      Left            =   4680
      Top             =   0
      _Version        =   786432
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmDesarrollo.frx":000C
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuArchivos 
         Caption         =   "Ver Archivos..."
      End
      Begin VB.Menu mnuCostos 
         Caption         =   "Costos..."
      End
      Begin VB.Menu mnuCopiar 
         Caption         =   "Copiar..."
      End
      Begin VB.Menu mnuModifConj 
         Caption         =   "Modificar Conjunto"
      End
      Begin VB.Menu mnuCambiar 
         Caption         =   "Cambiar..."
      End
      Begin VB.Menu mnuVerIncidencias 
         Caption         =   "Ver Incidencias..."
      End
      Begin VB.Menu mnuImprimir 
         Caption         =   "Imprimir..."
      End
   End
   Begin VB.Menu mnuMateriales 
      Caption         =   "Materiales"
      Visible         =   0   'False
      Begin VB.Menu mnuMatVerArchivos 
         Caption         =   "Ver Archivos..."
      End
   End
End
Attribute VB_Name = "frmDesarrollo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim CantArchivos As Dictionary

'para detalle presu historico
Private detaPresuId As Long
Private detallesHistoricos As New Collection

'para pieza
Private m_pieza As Pieza

Public Enum eVisualizando
    DesarrolloHistoricoDetallePresupuesto
    DesarrolloPieza
End Enum

Private vis As eVisualizando

Public Sub CargarDetallePresupuesto(detallePresupuestoId As Long)
    vis = eVisualizando.DesarrolloHistoricoDetallePresupuesto

    detaPresuId = detallePresupuestoId
    Set detallesHistoricos = DAODetallePresupuestoHistorico.FindAllByDetallePresupuestoId(detaPresuId)

    Dim tmpDetaHist As clsPresupuestoDetalleHistorico
    For Each tmpDetaHist In detallesHistoricos
        AgregarPieza tmpDetaHist
    Next tmpDetaHist

    Me.reportControlPiezas.Populate

    If Me.reportControlPiezas.rows.count > 0 Then
        'Me.reportControlPiezas.SelectedRows.DeleteAll
        'Me.reportControlPiezas.SelectedRows.Add Me.reportControlPiezas.rows(0)
        Set Me.reportControlPiezas.FocusedRow = Me.reportControlPiezas.rows(0)
    End If

    Me.caption = "Desarrollo Historico de Detalle Presupuesto Nº " & detallePresupuestoId
End Sub

Public Sub CargarPieza(piezaId As Long)
    vis = eVisualizando.DesarrolloPieza

    Set m_pieza = DAOPieza.FindById(piezaId, FL_4, False, False)

    AgregarPieza2 m_pieza

    Me.reportControlPiezas.Populate

    If Me.reportControlPiezas.rows.count > 0 Then
        'Me.reportControlPiezas.SelectedRows.DeleteAll
        'Me.reportControlPiezas.SelectedRows.Add Me.reportControlPiezas.rows(0)
        Set Me.reportControlPiezas.FocusedRow = Me.reportControlPiezas.rows(0)
    End If

    Me.caption = "Desarrollo de Pieza " & m_pieza.nombre
End Sub


Private Sub AgregarPieza2(ByVal Pieza As Pieza, Optional ByVal parent As ReportRecord = Nothing)
    Dim rec As ReportRecord

    If parent Is Nothing Then
        Set rec = Me.reportControlPiezas.Records.Add
    Else
        Set rec = parent.Childs.Add
    End If

    Set CantArchivos = DAOArchivo.GetCantidadArchivosPorReferencia(OA_Piezas)



    rec.AddItem Pieza.nombre
    If CantArchivos.item(Pieza.Id) > 0 Then
        rec.item(0).Icon = 13
    Else
        rec.item(0).Icon = 0
    End If

    rec.AddItem Pieza.Cantidad
    rec.Tag = Pieza.Id
    rec.Expanded = True

    Dim piezaHija As Pieza
    For Each piezaHija In Pieza.PiezasHijas
        AgregarPieza2 piezaHija, rec
    Next piezaHija

End Sub


Private Sub AgregarPieza(ByVal tmpDetaHist As clsPresupuestoDetalleHistorico, Optional ByVal parent As ReportRecord = Nothing)
    Dim rec As ReportRecord

    If parent Is Nothing Then
        Set rec = Me.reportControlPiezas.Records.Add
    Else
        Set rec = parent.Childs.Add
    End If

    rec.AddItem tmpDetaHist.NombrePieza
    rec.AddItem tmpDetaHist.DetallePresupuesto.Cantidad
    rec.Tag = tmpDetaHist.Id
    rec.Expanded = True

    Dim tmpDeta As clsPresupuestoDetalleHistorico
    For Each tmpDeta In tmpDetaHist.HistoricoHijos
        AgregarPieza tmpDeta, rec
    Next tmpDeta

End Sub



Private Sub Form_Load()

    FormHelper.Customize Me
    ArmarColumnasPiezas
    ArmarColumnasMateriales
    ArmarColumnasManoObra

    Me.reportControlPiezas.PaintManager.NoItemsText = "No hay piezas"
    Me.reportControlManoObra.PaintManager.NoItemsText = "No hay mano de obra"
    Me.reportControlMateriales.PaintManager.NoItemsText = "No hay materiales"

    Me.reportControlPiezas.Records.DeleteAll
    Me.reportControlMateriales.Records.DeleteAll
    Me.reportControlManoObra.Records.DeleteAll

    Set Me.reportControlPiezas.Icons = Me.ImageManager1.Icons
End Sub

Private Sub ArmarColumnasPiezas()
    Me.reportControlPiezas.Columns.DeleteAll
    Dim c As ReportColumn

    Set c = AddColumn(Me.reportControlPiezas, 0, "Pieza", , True, 70)
    AddColumn Me.reportControlPiezas, 1, "Cantidad", xtpAlignmentRight, , 25

    Me.reportControlPiezas.PaintManager.HorizontalGridStyle = xtpGridSmallDots
    Me.reportControlPiezas.PaintManager.VerticalGridStyle = xtpGridSmallDots
End Sub

Private Sub ArmarColumnasMateriales()
    Me.reportControlMateriales.Columns.DeleteAll

    AddColumn Me.reportControlMateriales, 0, "Codigo"
    AddColumn Me.reportControlMateriales, 1, "Descripcion"
    AddColumn Me.reportControlMateriales, 2, "Dim Pieza", xtpAlignmentRight
    AddColumn Me.reportControlMateriales, 3, "Scrap %", xtpAlignmentRight
    AddColumn Me.reportControlMateriales, 4, "Kg", xtpAlignmentRight
    AddColumn Me.reportControlMateriales, 5, "M2/Ml", xtpAlignmentRight
    AddColumn Me.reportControlMateriales, 6, "Costo", xtpAlignmentRight



    Me.reportControlMateriales.PaintManager.HorizontalGridStyle = xtpGridSmallDots
    Me.reportControlMateriales.PaintManager.VerticalGridStyle = xtpGridSmallDots
End Sub

Private Function AddColumn(ReportControl As ReportControl, ByVal Index As Long, ByVal caption As String, Optional align As XTPColumnAlignment = xtpAlignmentLeft, Optional ByVal tree As Boolean = False, Optional ByVal Width As Double = 25) As ReportColumn
    Dim col As ReportColumn
    Set col = ReportControl.Columns.Add(Index, caption, Width, True)
    col.Icon = 0
    col.Sortable = True
    col.AllowDrag = False
    col.AllowRemove = False
    col.Alignment = align
    col.TreeColumn = tree
    If Width <> 25 Then
        col.autoSize = True
        col.BestFitMode = XTPColumnBestFitMode.xtpBestFitModeAllData
    End If
    Set AddColumn = col
End Function


Private Sub ArmarColumnasManoObra()
    Me.reportControlManoObra.Columns.DeleteAll

    AddColumn Me.reportControlManoObra, 0, "Codigo", xtpAlignmentRight
    AddColumn Me.reportControlManoObra, 1, "Cant Op", xtpAlignmentRight
    AddColumn Me.reportControlManoObra, 2, "Tiempo (Min)", xtpAlignmentRight
    AddColumn Me.reportControlManoObra, 3, "Sector"
    AddColumn Me.reportControlManoObra, 4, "CPP"
    AddColumn Me.reportControlManoObra, 5, "Tarea"
    AddColumn Me.reportControlManoObra, 6, "Descripcion"
    AddColumn Me.reportControlManoObra, 7, "T. Total (Min)", xtpAlignmentRight
    AddColumn Me.reportControlManoObra, 8, "Costo ($)", xtpAlignmentRight

    Me.reportControlManoObra.PaintManager.HorizontalGridStyle = xtpGridSmallDots
    Me.reportControlManoObra.PaintManager.VerticalGridStyle = xtpGridSmallDots
End Sub

Private Sub mnuArchivos_Click()
    Dim P As Pieza
    Set P = m_pieza.LocatePiezaInPiezasHijas(Me.reportControlPiezas.SelectedRows(0).record.Tag)
    If Not P Is Nothing Then
        Dim frmArchi As New frmArchivos2
        frmArchi.Origen = OrigenArchivos.OA_Piezas
        frmArchi.ObjetoId = P.Id
        frmArchi.caption = "Pieza " & P.nombre
        frmArchi.Show
    End If

End Sub

Private Sub mnuCambiar_Click()
    Dim P As Pieza
    Set P = m_pieza.LocatePiezaInPiezasHijas(Me.reportControlPiezas.SelectedRows(0).record.Tag)
    Set P = DAOPieza.FindById(P.Id, FL_0, True, True, False)
    Dim basene As New classNuevoElemento
    Dim frm2 As New frmNuevoElemento
    frm2.lblidStock = P.Id
    frm2.txtNombreElemento = P.nombre
    frm2.cboComplejidad.ListIndex = funciones.PosIndexCbo(P.Complejidad, frm2.cboComplejidad)
    If frm2.cboComplejidad.ListIndex = -1 Then
        frm2.cboComplejidad.ListIndex = 0
    End If
    frm2.cboClientes.ListIndex = funciones.PosIndexCbo(P.cliente.Id, frm2.cboClientes)
    frm2.txtIdCliente = P.cliente.Id

    basene.llenarListaMDO P.Id, frm2.ListView2
    basene.llenarLstmateriales P.Id, frm2.ListView1

    frm2.caption = "Modificar desarrollo..."
    frm2.Command5.Visible = False
    frm2.btnModificar.Visible = True
    frm2.Show

End Sub

Private Sub mnuCopiar_Click()
    Dim P As Pieza
    Set P = m_pieza.LocatePiezaInPiezasHijas(Me.reportControlPiezas.SelectedRows(0).record.Tag)

    If P.EsConjunto Then
        Dim nuevoNombre As String
        If MsgBox("¿Está seguro de copiar el conjunto?", vbYesNo + vbQuestion, "Confirmación") = vbYes Then
            nuevoNombre = funciones.ingreso(P.nombre)
            If Len(Trim(nuevoNombre)) > 0 Then
                Dim claseS As New classStock
                If claseS.copiarConjuntoV2(P.Id, nuevoNombre, 0) Then
                    MsgBox "Conjunto copiado satisfactoriamente!", vbInformation, "Información"
                Else
                    MsgBox "Error en la copia de conjuntos!", vbCritical, "Error"
                End If
            End If
        End If

    Else

        Dim A
        A = funciones.ingreso(P.nombre)

        If Not IsEmpty(A) Then
            If DAOPieza.FindAll(FL_0, DAOPieza.CAMPO_NOMBRE & " = " & conectar.Escape(A)).count = 0 Then
                If MsgBox("¿Desea proceder con la copia?", vbYesNo, "Confirmación") Then
                    Dim base As New classStock
                    If base.CopiarPieza(P.Id, A) Then MsgBox "Copia exitosa!", vbInformation, "Información"
                End If
            Else
                MsgBox "El detalle ya existe en la base de datos!", vbCritical, "Error"
            End If
        End If
    End If


End Sub

Private Sub mnuCostos_Click()
    Dim P As Pieza
    Set P = m_pieza.LocatePiezaInPiezasHijas(Me.reportControlPiezas.SelectedRows(0).record.Tag)

    frmCostosIncidencia.cliente = P.cliente.Id
    frmCostosIncidencia.idp = P.Id
    frmCostosIncidencia.Show 1
End Sub

Private Sub mnuImprimir_Click()
    Dim P As Pieza
    Set P = m_pieza.LocatePiezaInPiezasHijas(Me.reportControlPiezas.SelectedRows(0).record.Tag)

    If MsgBox("¿Desea imprimir los materiales y la mano de obra de la pieza [" & P.nombre & "]?", vbQuestion + vbYesNo) = vbYes Then

        Dim ret As VbMsgBoxResult
        If P.EsConjunto Then
            ret = MsgBox("¿Desea imprimir todas las subpiezas del conjunto seleccionado o solo la pieza principal del conjunto?" & vbNewLine & "SI = todas las subpiezas" & vbNewLine & "NO = solo la pieza principal seleccionada", vbYesNo + vbQuestion)
        Else
            ret = vbNo
        End If

        If ret = vbNo Then
            printMatMDO P
        Else
            MatManoObraFromRepRecord Me.reportControlPiezas.SelectedRows(0), P
        End If

        MsgBox "Impresión finalizada.", vbInformation + vbOKOnly
    End If

End Sub

Private Sub MatManoObraFromRepRecord(row As ReportRow, P As Pieza)
    Dim repRow As ReportRow
    Set Me.reportControlPiezas.FocusedRow = row
    printMatMDO P

    For Each repRow In row.Childs
        MatManoObraFromRepRecord repRow, P.LocatePiezaInPiezasHijas(repRow.record.Tag)
    Next repRow
End Sub


Private Sub printMatMDO(P As Pieza)
    If Me.reportControlMateriales.Records.count > 0 Then
        Me.reportControlMateriales.PrintOptions.header.TextCenter = "Materiales de pieza [" & P.nombre & "]"
        Me.reportControlMateriales.PrintOptions.Landscape = True
        Me.reportControlMateriales.PrintReport2 True
    End If

    If Me.reportControlManoObra.Records.count > 0 Then
        Me.reportControlManoObra.PrintOptions.header.TextCenter = "Mano de Obra de pieza [" & P.nombre & "]"
        Me.reportControlManoObra.PrintOptions.Landscape = True
        Me.reportControlManoObra.PrintReport2 True
    End If
End Sub


Private Sub mnuMatVerArchivos_Click()
    Dim P As Collection
    Set P = DAODesarrolloMaterial.FindAll("dm.id=" & Me.reportControlMateriales.SelectedRows(0).record.Tag)

    If P.count = 1 Then
        Dim m As DesarrolloMaterial
        Set m = P(1)
        Dim frmArchi As New frmArchivos2
        frmArchi.Origen = OrigenArchivos.OA_Materiales
        frmArchi.ObjetoId = m.Material.Id
        frmArchi.caption = "Material: " & m.Material.descripcion
        frmArchi.Show
    End If


End Sub

Private Sub mnuModifConj_Click()
    Dim P As Pieza
    Set P = m_pieza.LocatePiezaInPiezasHijas(Me.reportControlPiezas.SelectedRows(0).record.Tag)
    If P.EsConjunto Then

        Dim frm As New frmDefinirConjunto

        frm.accion = 1
        frm.idPiezaMadre = P.Id
        frm.Show
    End If

End Sub

Private Sub mnuVerIncidencias_Click()

    Dim P As Pieza
    Set P = m_pieza.LocatePiezaInPiezasHijas(Me.reportControlPiezas.SelectedRows(0).record.Tag)
    If Not P Is Nothing Then
        Dim frmArchi As New frmVerIncidencias
        frmArchi.Origen = 3
        frmArchi.referencia = P.Id
        frmArchi.caption = "Pieza " & P.nombre
        frmArchi.Show
    End If

End Sub

Private Sub reportControlMateriales_MouseDown(Button As Integer, Shift As Integer, x As Long, y As Long)
    If Button = 2 And vis = DesarrolloPieza Then
        Dim hitinfo As ReportHitTestInfo
        Set hitinfo = Me.reportControlMateriales.HitTest(x, y)

        If Not hitinfo.item Is Nothing Then


            Me.PopupMenu Me.mnuMateriales


        End If
    End If

End Sub

Private Sub reportControlPiezas_MouseDown(Button As Integer, Shift As Integer, x As Long, y As Long)
    If Button = 2 And vis = DesarrolloPieza Then
        Dim hitinfo As ReportHitTestInfo
        Set hitinfo = Me.reportControlPiezas.HitTest(x, y)

        If Not hitinfo.item Is Nothing Then
            Dim P As Pieza
            Set P = m_pieza.LocatePiezaInPiezasHijas(Me.reportControlPiezas.SelectedRows(0).record.Tag)
            If Not P Is Nothing Then
                Me.mnuModifConj.Enabled = P.EsConjunto
                Me.PopupMenu Me.mnuMain

            End If
        End If
    End If
End Sub

Private Sub reportControlPiezas_SelectionChanged()
    Me.reportControlManoObra.Records.DeleteAll
    Me.reportControlMateriales.Records.DeleteAll

    LimpiarLabels

    If Me.reportControlPiezas.SelectedRows.count > 0 Then
        If vis = DesarrolloHistoricoDetallePresupuesto Then
            Dim pdh As clsPresupuestoDetalleHistorico
            Set pdh = DAODetallePresupuestoHistorico.FindItemInCollection(detallesHistoricos, Me.reportControlPiezas.SelectedRows(0).record.Tag)

            CargarMateriales pdh.HistoricoMAT
            CargarManoObra pdh.historicoMDO

            Me.lblCostoMateriales.caption = Me.lblCostoMateriales.Tag & funciones.FormatearDecimales(pdh.TotalCostoMateriales)
            Me.lblCostoManoObra.caption = Me.lblCostoManoObra.Tag & funciones.FormatearDecimales(pdh.TotalCostoMDO)
            Me.lblTotalKg.caption = Me.lblTotalKg.Tag & pdh.TotalKGMateriales
            Me.lblTotalM2.caption = Me.lblTotalM2.Tag & pdh.TotalM2Materiales
        ElseIf vis = DesarrolloPieza Then
            Dim P As Pieza
            Set P = m_pieza.LocatePiezaInPiezasHijas(Me.reportControlPiezas.SelectedRows(0).record.Tag)
            If Not P Is Nothing Then
                Set P.desarrollosManoObra = DAODesarrolloManoObra.FindAllByPiezaId(P.Id)
                Set P.DesarrollosMaterial = DAODesarrolloMaterial.FindAllByPiezaId(P.Id)

                CargarMateriales2 P.DesarrollosMaterial
                CargarManoObra2 P.desarrollosManoObra

                Me.lblCostoMateriales.caption = Me.lblCostoMateriales.Tag & funciones.FormatearDecimales(P.TotalCostoMateriales)
                Me.lblCostoManoObra.caption = Me.lblCostoManoObra.Tag & funciones.FormatearDecimales(P.TotalCostoManoObra)
                Me.lblTotalKg.caption = Me.lblTotalKg.Tag & P.TotalKG
                Me.lblTotalM2.caption = Me.lblTotalM2.Tag & P.TotalM2

            End If
        End If
    End If
End Sub

Private Sub LimpiarLabels()
    Dim ctrl As Control
    For Each ctrl In Me.Controls
        If TypeOf ctrl Is Label Then
            ctrl.caption = ctrl.Tag
        End If
    Next
End Sub


Private Sub CargarMateriales2(desamat As Collection)
    Dim rec As ReportRecord
    Dim tmp As DesarrolloMaterial
    Dim dto As DatosMaterialDTO

    For Each tmp In desamat
        dto = tmp.CalcularDatosMaterial(tmp.Material.moneda.Id)

        Set rec = Me.reportControlMateriales.Records.Add()

        rec.AddItem tmp.Material.codigo
        rec.AddItem tmp.Material.descripcion
        rec.AddItem dto.DimensionMaterial
        rec.AddItem tmp.Scrap
        rec.AddItem tmp.Kg
        rec.AddItem tmp.m2
        rec.AddItem funciones.FormatearDecimales(MonedaConverter.Convertir(dto.costo, tmp.Material.moneda.Id, 0))
        rec.Tag = tmp.Id
    Next tmp

    Me.reportControlMateriales.Populate
End Sub

Private Sub CargarMateriales(histMAT As Collection)
    Dim rec As ReportRecord

    Dim datosMat As DatosMaterialDTO

    Dim tmp As PresupuestoDetalleHistoricoMAT
    For Each tmp In histMAT
        datosMat = tmp.CalcularDatosMaterial(tmp.moneda.Id)

        Set rec = Me.reportControlMateriales.Records.Add()
        rec.AddItem tmp.Material.codigo
        rec.AddItem tmp.Material.descripcion
        rec.AddItem datosMat.DimensionPieza
        rec.AddItem tmp.Scrap
        rec.AddItem datosMat.Kg
        rec.AddItem datosMat.m2
        rec.AddItem funciones.FormatearDecimales(MonedaConverter.Convertir(datosMat.costo, tmp.Material.moneda.Id, 0))

        rec.Tag = tmp.Id
    Next tmp

    Me.reportControlMateriales.Populate
End Sub

Private Sub CargarManoObra2(desamdo As Collection)
    Dim rec As ReportRecord
    Dim tmp As DesarrolloManoObra

    For Each tmp In desamdo
        Set rec = Me.reportControlManoObra.Records.Add()
        rec.AddItem tmp.Tarea.Id
        rec.AddItem tmp.Cantidad
        rec.AddItem funciones.RedondearDecimales(tmp.Tiempo)
        rec.AddItem tmp.Tarea.Sector.Sector
        rec.AddItem tmp.Tarea.CantPorProcSmartProperty
        rec.AddItem tmp.Tarea.Tarea
        rec.AddItem tmp.Tarea.descripcion
        rec.AddItem funciones.RedondearDecimales(tmp.Tiempo * tmp.Cantidad)
        rec.AddItem funciones.RedondearDecimales(tmp.Tiempo * tmp.Cantidad * tmp.Tarea.CategoriaSueldo.Valor)

        rec.Tag = tmp.Id
    Next tmp

    Me.reportControlManoObra.Populate
End Sub


Private Sub CargarManoObra(histManoObra As Collection)
    Dim rec As ReportRecord
    Dim tmp As PresupuestoDetalleHistoricoMDO

    For Each tmp In histManoObra
        Set rec = Me.reportControlManoObra.Records.Add()
        rec.AddItem tmp.Tarea.Id
        rec.AddItem tmp.CantOperarios
        rec.AddItem funciones.RedondearDecimales(tmp.Tiempo)
        rec.AddItem tmp.Tarea.Sector.Sector
        rec.AddItem tmp.Tarea.CantPorProcSmartProperty
        rec.AddItem tmp.Tarea.Tarea
        rec.AddItem tmp.Tarea.descripcion
        rec.AddItem funciones.RedondearDecimales(tmp.Tiempo * tmp.CantOperarios)
        rec.AddItem funciones.RedondearDecimales(tmp.Tiempo * tmp.CantOperarios * tmp.Valor)

        rec.Tag = tmp.Id
    Next tmp

    Me.reportControlManoObra.Populate
End Sub
