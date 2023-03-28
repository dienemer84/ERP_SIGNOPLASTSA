VERSION 5.00
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#12.0#0"; "CODEJO~3.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Object = "{1A9D2E18-63A4-11D3-9EC5-5C91AD000000}#2.5#0"; "phGantXControl.ocx"
Begin VB.Form frmPlanificacionTemporal 
   Caption         =   "Planificacion Temporal"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17040
   Icon            =   "frmPlanificacionTemporal.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8190
   ScaleWidth      =   17040
   WindowState     =   2  'Maximized
   Begin phGantXControl.phGantX phGantX1 
      Height          =   7935
      Left            =   5805
      TabIndex        =   1
      Top             =   0
      Width           =   11250
      StickyMode      =   0
      GantBackColor   =   14737632
      GantPyjamasColor=   16777215
      MoveInTimeWhenMoveRow=   0   'False
      MoveOver12Then24=   0   'False
      ColorTree       =   -16777211
      DragCursorTree  =   -12
      BeginProperty FontTree {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DrawShortLines  =   0   'False
      HideSelectionTree=   0   'False
      HotTrackTree    =   0   'False
      IndentTree      =   19
      ToolTipsTree    =   -1  'True
      AutoExpandLevelsTree=   0
      MaxLevelsTree   =   -1
      RightClickSelectTree=   0   'False
      MultiSelectTree =   -1  'True
      AutoExpandTree  =   0   'False
      RowSelectTree   =   0   'False
      ShowButtonsTree =   -1  'True
      ShowLinesTree   =   -1  'True
      CursorTree      =   0
      DateFormat      =   "DD/MM/YYYY"
      IndicatorOn     =   -1  'True
      PixelsForModeSwitch=   15
      Scale           =   3.23158914728682E-02
      Start           =   40276
      Stop            =   40296.84375
      TimeFormat      =   "HH:MM:SS"
      ZoomFactor      =   0.5
      ColorScaler     =   -16777201
      BeginProperty FontScaler {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CursorScaler    =   0
      FirstDayOfWeekScaler=   2
      ScalerHeight    =   45
      DefaultLinkStyle=   0
      DefaultLinkColor=   16711680
      StartSnapONOFF  =   -1  'True
      LengthSnapONOFF =   -1  'True
      StartSnap       =   3.47222222222222E-03
      LengthSnap      =   6.94444444444444E-04
      LengthSnapWhenMovingONOFF=   0   'False
      BeginProperty ScalerFont1 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty ScalerFont2 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty ScalerFont3 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ScalerFont1Color=   -16777208
      ScalerFont2Color=   -16777208
      ScalerFont3Color=   -16777208
      ScalerIndicatorStyle=   1
      ScalerWeekNumbers=   0   'False
      RescaleWithCtrl =   -1  'True
      TreeWidth       =   100
      GridHasInfoCol  =   0   'False
      GridHasHeader   =   -1  'True
      UseGrid         =   0   'False
      GridCellFocusedX=   -1
      GridCellFocusedY=   -1
      TimeLinks_Show  =   0
      BeginProperty TimeItemTextFont1 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TimeItemTextFont2 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TimeItemTextFont3 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TimeItemTextFont1Color=   -16777208
      TimeItemTextFont2Color=   -16777208
      TimeItemTextFont3Color=   -16777208
      CursorGantArea  =   -2
      FocusRectsGrid  =   -1  'True
      BorderStyleTreeGrid=   1
      DragModeTreeGrid=   1
      VerticalPyjamas =   0   'False
      TabStop         =   -1  'True
      AsyncRelease    =   0   'False
      GridStopEditOnFocusChange=   -1  'True
      ScaleLookAndFeel=   1
      ScaleScrollButtons=   1
      ScaleMinStart   =   0
      ScaleMaxStop    =   767011
      DrawLongLines   =   -1  'True
      PrintSettingsPrinterName=   "Send To OneNote 2007"
      TodayLineOnOff  =   -1  'True
      TodayLineColor  =   0
      TimeItemAutoScroll=   0
      StartDragWhenTimeMoveOutside=   0   'False
      PageSettingsIsLandscape=   0   'False
      GridVerticalScrollOutOfGrid=   -1  'True
      HdcToUseOnNewPage=   0
      AllowKeyNavigationInGrid=   -1  'True
      LongDatesToLeftEnd=   0   'False
      FavourMoveOverResizeOnSmallTimeItems=   -1  'True
      InplaceDateTimeClearStatesBetweenEdits=   -1  'True
      SupressOnUserDrawExceptions=   0   'False
   End
   Begin XtremeReportControl.ReportControl ReportControl 
      Height          =   6435
      Left            =   15
      TabIndex        =   0
      Top             =   1470
      Width           =   5700
      _Version        =   786432
      _ExtentX        =   10054
      _ExtentY        =   11351
      _StockProps     =   64
      BorderStyle     =   3
      PreviewMode     =   -1  'True
      AllowColumnRemove=   0   'False
      AllowColumnReorder=   0   'False
      AllowColumnSort =   0   'False
      ShowHeaderRows  =   -1  'True
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1365
      Left            =   60
      TabIndex        =   2
      Top             =   30
      Width           =   5670
      _Version        =   786432
      _ExtentX        =   10001
      _ExtentY        =   2408
      _StockProps     =   79
      Caption         =   "Datos OT"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.PushButton cmdBuscar 
         Default         =   -1  'True
         Height          =   345
         Left            =   2295
         TabIndex        =   5
         Top             =   255
         Width           =   900
         _Version        =   786432
         _ExtentX        =   1587
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Buscar"
         UseVisualStyle  =   -1  'True
      End
      Begin VB.TextBox txtOt 
         Height          =   285
         Left            =   870
         TabIndex        =   4
         Top             =   270
         Width           =   1215
      End
      Begin VB.Label lblDescripcion 
         BackColor       =   &H8000000B&
         Height          =   240
         Left            =   120
         TabIndex        =   7
         Top             =   930
         Width           =   5430
      End
      Begin VB.Label lblCliente 
         BackColor       =   &H8000000C&
         Height          =   240
         Left            =   120
         TabIndex        =   6
         Top             =   690
         Width           =   5430
      End
      Begin VB.Label Label1 
         Caption         =   "Nro. OT"
         Height          =   240
         Left            =   240
         TabIndex        =   3
         Top             =   315
         Width           =   1110
      End
   End
   Begin VB.Menu mnuPieza 
      Caption         =   "Pieza"
      Visible         =   0   'False
      Begin VB.Menu mnuTareaAPieza 
         Caption         =   "Agregar tarea a pieza"
      End
      Begin VB.Menu mnuAsignarRecursos 
         Caption         =   "Asignar Recursos"
      End
   End
End
Attribute VB_Name = "frmPlanificacionTemporal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private parentRow2Check As ReportRow
Dim detadto As DetalleOTConjuntoDTO
Private vpedido As OrdenTrabajo


Public Property Let Pedido(nvalue As OrdenTrabajo)
    Set vpedido = nvalue
    BuscarPedido
End Property
Private Sub AgregarTareas(row As ReportRow)
    Dim tn As IphDataEntity_Tree2
    Me.phGantX1.ClearTree
    Dim idDetalle As Long
    idDetalle = row.record.Tag

    Dim lista_ptp As Collection
    Dim ptp As PlaneamientoTiempoProceso
    If row.ParentRow Is Nothing Then    'es el detalle de la ot
        Dim detalle As DetalleOrdenTrabajo
        Set detalle = DAODetalleOrdenTrabajo.FindById(idDetalle)
        'Set lista_ptp = DAOTiemposProceso.FindAllByDetallePedidoIdAndPiezaId( idDetalle, detalle.pieza.Id, , True)
        Set lista_ptp = DAOTiemposProceso.FindAllByDetallePedidoId(idDetalle, , , True)
    Else
        Set detadto = DAODetalleOrdenTrabajo.FindConjuntoById(idDetalle)
        If detadto Is Nothing Then Exit Sub
        'Set lista_ptp = DAOTiemposProceso.FindAllByDetallePedidoIdAndPiezaId(idDetalle, detadto.pieza.Id, , True)
        Set lista_ptp = DAOTiemposProceso.FindAllByDetallePedidoId(row.ParentRow.record.Tag, idDetalle, , True)
    End If

    Dim c As Long
    c = 0
    For Each ptp In lista_ptp
        c = c + 1
        Dim newactivity As IphDataEntity_Tree
        If (phGantX1.CurrentDataEntityTree Is Nothing) Then
            Set newactivity = phGantX1.AddRootDataEntityTree

        Else
            Set newactivity = phGantX1.AddDataEntityTree(phGantX1.CurrentDataEntityTree)
        End If
        newactivity.CanEdit = True
        newactivity.text = ptp.Planificacion.Prioridad
        newactivity.UserVariantReference = ptp

        Dim time As IphDataEntity_GantTime
        If IsSomething(newactivity) Then
            Set time = phGantX1.AddGantTime(newactivity, 0)
            time.UserVariantReference = ptp

            If ptp.Planificacion.Id = 0 Then
                ptp.Planificacion.Inicio = Date
                ptp.Planificacion.Fin = DateAdd("d", 1, Date)
                ptp.Planificacion.idTiempoProceso = ptp.Id
                ptp.Planificacion.Color = ColorConstants.vbBlue
                ptp.Planificacion.Prioridad = c
            End If
            time.Start = ptp.Planificacion.Inicio
            time.Stop = ptp.Planificacion.Fin
            time.Color = ptp.Planificacion.Color
            Set tn = time.row.TreeNode

            phGantX1.GridCellValueSet 0, tn.GridRowIndex, ptp.Planificacion.Prioridad, -1
            phGantX1.GridCellValueSet 1, tn.GridRowIndex, ptp.Tarea.Description, -1
            phGantX1.GridCellValueSet 2, tn.GridRowIndex, IIf(ptp.Planificacion.Critica, "TRUE", "FALSE"), -1
            phGantX1.GridCellValueSet 3, tn.GridRowIndex, time.Start, -1
            phGantX1.GridCellValueSet 4, tn.GridRowIndex, time.Stop, -1
            phGantX1.GridCellValueSet 5, tn.GridRowIndex, DateDiff("D", time.Start, time.Stop), -1

        End If
    Next
End Sub
Private Sub cmdBuscar_Click()
    If LenB(Me.txtOt) > 0 And IsNumeric(Me.txtOt) Then
        Set vpedido = DAOOrdenTrabajo.FindById(Val(Me.txtOt))
        If IsSomething(vpedido) Then
            Set vpedido.Detalles = DAODetalleOrdenTrabajo.FindAllByOrdenTrabajo(vpedido.Id)
            MostrarGantt
        End If
    End If
End Sub
Private Sub limpiar()
    Me.lblCliente = vbNullString
    Me.lblDescripcion = vbNullString
End Sub

Private Sub Form_Load()
    Customize Me
    Me.phGantX1.GantBackColor = FormHelper.FondoCeleste
    ArmarColumnas
    ArmarColGantt
    If IsSomething(vpedido) Then
        MostrarGantt
    End If
End Sub
Private Sub MostrarGantt()
    Me.txtOt = vpedido.Id
    Me.lblCliente = "Cliente: " & vpedido.cliente.razon
    Me.lblDescripcion = "Descripcion: " & vpedido.descripcion & "  | Fecha entrega: " & vpedido.FechaEntrega
    ArmarGantt
    Me.phGantX1.Start = Date - 5
End Sub

Private Sub BuscarPedido()
    Set vpedido = DAOOrdenTrabajo.FindById(vpedido.Id)
    Set vpedido.Detalles = DAODetalleOrdenTrabajo.FindAllByOrdenTrabajo(vpedido.Id)
End Sub

Private Sub ArmarColGantt()
    phGantX1.UseGrid = True
    phGantX1.RescaleWithCtrl = True
    phGantX1.GridHasHeader = True
    phGantX1.AutoExpandLevelsTree = 3

    phGantX1.Start = Date - 15
    phGantX1.Stop = Date + 15


    phGantX1.GridColumnAdd
    phGantX1.GridColumnAdd
    phGantX1.GridColumnAdd
    phGantX1.GridColumnAdd
    phGantX1.GridColumnAdd

    phGantX1.GridColumnSet 0, "Prio", tekSimpleEdit, 40, True
    phGantX1.GridColumnSet 1, "Nombre", tekSimpleEdit, 170, True
    phGantX1.GridColumnSet 2, "Crit", tekBool, 20, True
    phGantX1.GridColumnSet 3, "Inicio ", tekDate, 70, True
    phGantX1.GridColumnSet 4, "Fin", tekDate, 70, True
    phGantX1.GridColumnSet 5, "Días", tekSimpleEdit, 50, True

    Me.phGantX1.TreeWidth = 430

    phGantX1.GridEditComboClear
    phGantX1.GridEditComboAdd
    phGantX1.GridEditComboAdd
    phGantX1.GridEditComboAdd

    phGantX1.GridLayoutPropAdd
    phGantX1.GridLayoutPropAdd
    phGantX1.GridLayoutPropAdd

    '--> ver lo que sige en theHelper
    phGantX1.GridLayoutPropSet 0, ColorConstants.vbYellow, ColorConstants.vbBlack, "tahoma", 8, 20, 20, False, True, False, True, False
    phGantX1.GridLayoutPropSet 1, ColorConstants.vbRed, ColorConstants.vbBlack, "tahoma", 8, 20, 20, False, True, False, True, False
    phGantX1.GridLayoutPropSet 2, ColorConstants.vbBlue, ColorConstants.vbWhite, "tahoma", 8, 20, 20, False, True, False, True, False


End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Me.phGantX1.Width = Me.ScaleWidth - Me.ReportControl.Width - 120
    Me.phGantX1.Height = Me.ScaleHeight
    Me.ReportControl.Height = Me.ScaleHeight - Me.ReportControl.Top

End Sub
Private Sub ArmarColumnas()
'--------------------
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

    Set Column = Me.ReportControl.Columns.Add(3, "F. Entrega", 17, True)
    Column.Icon = 0
    Column.Sortable = False
    Column.AllowDrag = False
    Column.AllowRemove = False


    Me.ReportControl.PaintManager.HorizontalGridStyle = xtpGridSmallDots
    Me.ReportControl.PaintManager.VerticalGridStyle = xtpGridSmallDots

End Sub

Private Sub ArmarGantt()
    Dim deta As DetalleOrdenTrabajo
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
    For Each deta In vpedido.Detalles
        Set record = Me.ReportControl.Records.Add
        record.Tag = deta.Id
        record.AddItem deta.item
        Set item = record.AddItem(deta.Pieza.nombre)

        record.PreviewText = deta.Nota
        record.AddItem deta.CantidadPedida
        record.AddItem deta.FechaEntrega
        If deta.Pieza.EsConjunto Then
            For Each tmpdeta2 In DAODetalleOrdenTrabajo.FindAllConjunto(deta.Id, deta.Pieza.Id)

                Set Record2 = record.Childs.Add()
                Record2.Tag = tmpdeta2.Id
                Record2.AddItem vbNullString
                Set item = Record2.AddItem(tmpdeta2.Pieza.nombre)
                Record2.AddItem tmpdeta2.Cantidad
                Record2.AddItem vbNullString
                If tmpdeta2.Pieza.EsConjunto Then
                    For Each tmpdeta3 In DAODetalleOrdenTrabajo.FindAllConjunto(deta.Id, tmpdeta2.Pieza.Id)
                        Set Record3 = Record2.Childs.Add
                        Record3.Tag = tmpdeta3.Id
                        Record3.AddItem vbNullString
                        Set item = Record3.AddItem(tmpdeta3.Pieza.nombre)
                        Record3.AddItem tmpdeta3.Cantidad
                        Record3.AddItem vbNullString
                        If tmpdeta3.Pieza.EsConjunto Then
                            For Each tmpdeta4 In DAODetalleOrdenTrabajo.FindAllConjunto(deta.Id, tmpdeta3.Pieza.Id)
                                Set Record4 = Record3.Childs.Add
                                Record4.Tag = tmpdeta4.Id
                                Record4.AddItem vbNullString
                                Set item = Record4.AddItem(tmpdeta4.Pieza.nombre)
                                Record4.AddItem tmpdeta4.Cantidad
                                Record4.AddItem vbNullString
                            Next tmpdeta4
                        End If
                    Next tmpdeta3
                End If
            Next tmpdeta2
        End If
    Next
    Me.ReportControl.Populate
End Sub

Private Sub Label2_Click()

End Sub



Private Sub mnuAsignarRecursos_Click()
    Dim row As ReportRow
    Dim col As New Collection
    Dim deta As DetalleOrdenTrabajo
    Dim detadto As DetalleOTConjuntoDTO

    For Each row In Me.ReportControl.SelectedRows
        If row.ParentRow Is Nothing Then    'es el detalle de la ot
            Set deta = DAODetalleOrdenTrabajo.FindById(row.record.Tag)
            col.Add deta
        Else    'es el detalle de algun detalle
            Set detadto = DAODetalleOrdenTrabajo.FindConjuntoById(row.record.Tag)
            col.Add detadto
        End If
    Next row

    Dim F As New frmAsigacionRecursos
    Load F
    F.llenar vpedido.Id, col
    F.Show

End Sub

Private Sub mnuTareaAPieza_Click()
    If Me.ReportControl.SelectedRows.count > 0 Then
        Dim row As ReportRow
        Dim col As Collection
        Dim deta As DetalleOrdenTrabajo
        Dim detadto As DetalleOTConjuntoDTO
        Set row = Me.ReportControl.SelectedRows(0)

        Dim F As New frmPlaneamientoAgregarTareaAProceso

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

        TareaAgregada = False
        F.pedido_id = vpedido.Id
        F.Show 1
        If TareaAgregada Then
            AgregarTareas Me.ReportControl.SelectedRows(0)
        End If
    End If


End Sub

Private Sub phGantX1_OnValueChangedGantTime(ByVal theGant As phGantXControl.IphGantX, ByVal theDataEntity As phGantXControl.IphDataEntity_GantTime)
    Dim ptpp As TiempoProcesoPlanificado
    Dim tn As IphDataEntity_Tree2
    Set ptpp = theDataEntity.UserVariantReference.Planificacion
    ptpp.Inicio = theDataEntity.Start
    ptpp.Fin = theDataEntity.Stop


    If ptpp.Critica Then
        ptpp.Color = VBA.Information.RGB(255, 0, 0)
    Else
        ptpp.Color = VBA.Information.RGB(0, 255, 0)
    End If


    Set tn = theDataEntity.row.TreeNode


    phGantX1.GridCellValueSet 3, tn.GridRowIndex, DateTime.DateValue(theDataEntity.Start), -1
    phGantX1.GridCellValueSet 4, tn.GridRowIndex, DateTime.DateValue(theDataEntity.Stop), -1
    phGantX1.GridCellValueSet 5, tn.GridRowIndex, DateDiff("D", theDataEntity.Start, theDataEntity.Stop), -1
    theDataEntity.Color = ptpp.Color


    If Not savePTPP(ptpp) Then
        MsgBox "Hubo un error al actualizar los valores.", vbCritical + vbOKOnly
    End If
End Sub


Private Function savePTPP(ptpp As TiempoProcesoPlanificado)
    savePTPP = DAOTiempoProcesoPlanificado.Save(ptpp)
End Function

Private Sub phGantX1_OnValueChangedGrid(ByVal theGant As phGantXControl.IphGantX3, ByVal theDataEntity As phGantXControl.IphDataEntity_Tree2, ByVal x As Long, ByVal y As Long, newValue As String)
    Dim ptpp As TiempoProcesoPlanificado
    Dim tn As IphDataEntity_Tree2

    Dim gt As IphDataEntity_GantTime2
    Set gt = theDataEntity.GantRow.DataLists.DataList(0).Items(0)

    Set ptpp = gt.UserVariantReference.Planificacion

    Set tn = gt.row.TreeNode


    ' phGantX1.GridCellValueSet 2, tn.GridRowIndex, DateTime.DateValue(gt.Start), -1
    'phGantX1.GridCellValueSet 3, tn.GridRowIndex, DateTime.DateValue(gt.Stop), -1



    If x = 2 And newValue <> vbNullString And newValue <> "" Then
        ptpp.Critica = IIf(newValue = "TRUE", True, False)
        If ptpp.Critica Then
            ptpp.Color = VBA.Information.RGB(255, 0, 0)
        Else
            ptpp.Color = VBA.Information.RGB(0, 255, 0)
        End If
        gt.Color = ptpp.Color

    End If

    If x = 0 And newValue <> "" And IsNumeric(newValue) Then
        ptpp.Prioridad = Val(newValue)
    End If


    If x = 3 Then
        If newValue <> "" Then
            gt.Start = newValue
            phGantX1.GridCellValueSet 5, tn.GridRowIndex, DateDiff("d", gt.Start, gt.Stop), -1
        End If
    End If

    If x = 4 Then
        If newValue <> "" Then
            gt.Stop = newValue
            phGantX1.GridCellValueSet 5, tn.GridRowIndex, DateDiff("d", gt.Start, gt.Stop), -1
        End If
    End If


    If x = 5 Then
        gt.Stop = DateAdd("d", Val(newValue), gt.Start)
        phGantX1.GridCellValueSet 4, tn.GridRowIndex, gt.Stop, -1
    End If


    ptpp.Inicio = gt.Start
    ptpp.Fin = gt.Stop

    If Not savePTPP(ptpp) Then
        MsgBox "Hubo un error al actualizar los valores.", vbCritical + vbOKOnly
    End If

End Sub

Private Sub ReportControl_MouseDown(Button As Integer, Shift As Integer, x As Long, y As Long)
    If Button = 2 Then
        Dim hitinfo As ReportHitTestInfo
        Set hitinfo = Me.ReportControl.HitTest(x, y)
        If Not hitinfo Is Nothing Then
            If Not hitinfo.row Is Nothing Then
                If Me.ReportControl.SelectedRows.count = 1 Then

                    Me.ReportControl.SelectedRows.DeleteAll
                    Me.ReportControl.SelectedRows.Add hitinfo.row
                    Me.ReportControl.FocusedRow = hitinfo.row
                    'If CLng(hitinfo.row.record.Tag) < 0 Then 'es tarea
                    '    Me.PopupMenu Me.mnuTarea
                    'Else

                    'End If
                ElseIf Me.ReportControl.SelectedRows.count > 1 Then

                End If
                Me.mnuTareaAPieza.Enabled = (Me.ReportControl.SelectedRows.count = 1)

                Me.PopupMenu Me.mnuPieza
            End If
        End If
    End If
End Sub

Private Sub ReportControl_RowDblClick(ByVal row As XtremeReportControl.IReportRow, ByVal item As XtremeReportControl.IReportRecordItem)
    row.Expanded = Not row.Expanded
End Sub
Private Sub ReportControl_SelectionChanged()
    If Me.ReportControl.SelectedRows.count > 0 Then
        Dim row As ReportRow
        Set row = Me.ReportControl.SelectedRows(0)
        AgregarTareas row
    End If
End Sub
