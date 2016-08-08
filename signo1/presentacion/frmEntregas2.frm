VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~3.OCX"
Begin VB.Form frmEntregas2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Entregas de OT"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12495
   Icon            =   "frmEntregas2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   12495
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   8280
      TabIndex        =   20
      Top             =   6960
      Visible         =   0   'False
      Width           =   1140
   End
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   3855
      Left            =   105
      TabIndex        =   13
      Top             =   4005
      Width           =   8040
      _Version        =   786432
      _ExtentX        =   14182
      _ExtentY        =   6800
      _StockProps     =   68
      AllowReorder    =   -1  'True
      Appearance      =   10
      Color           =   128
      ItemCount       =   2
      SelectedItem    =   1
      Item(0).Caption =   "Entregas"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "grpEntregasDetalle"
      Item(1).Caption =   "Facturas"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "GroupBox1"
      Begin XtremeSuiteControls.GroupBox grpEntregasDetalle 
         Height          =   3390
         Left            =   -69910
         TabIndex        =   14
         Top             =   375
         Visible         =   0   'False
         Width           =   7815
         _Version        =   786432
         _ExtentX        =   13785
         _ExtentY        =   5980
         _StockProps     =   79
         Caption         =   "Entregas del detalle seleccionado de la OT"
         UseVisualStyle  =   -1  'True
         Begin GridEX20.GridEX gridEntregas 
            Height          =   3060
            Left            =   90
            TabIndex        =   15
            Top             =   255
            Width           =   6135
            _ExtentX        =   10821
            _ExtentY        =   5398
            Version         =   "2.0"
            HoldSortSettings=   -1  'True
            BoundColumnIndex=   ""
            ReplaceColumnIndex=   ""
            ColumnAutoResize=   -1  'True
            MethodHoldFields=   -1  'True
            GroupByBoxVisible=   0   'False
            DataMode        =   99
            ColumnHeaderHeight=   285
            IntProp1        =   0
            IntProp2        =   0
            IntProp7        =   0
            ColumnsCount    =   4
            Column(1)       =   "frmEntregas2.frx":000C
            Column(2)       =   "frmEntregas2.frx":014C
            Column(3)       =   "frmEntregas2.frx":0284
            Column(4)       =   "frmEntregas2.frx":03B0
            SortKeysCount   =   1
            SortKey(1)      =   "frmEntregas2.frx":0470
            FormatStylesCount=   12
            FormatStyle(1)  =   "frmEntregas2.frx":04D8
            FormatStyle(2)  =   "frmEntregas2.frx":0610
            FormatStyle(3)  =   "frmEntregas2.frx":06C0
            FormatStyle(4)  =   "frmEntregas2.frx":0774
            FormatStyle(5)  =   "frmEntregas2.frx":084C
            FormatStyle(6)  =   "frmEntregas2.frx":0904
            FormatStyle(7)  =   "frmEntregas2.frx":09E4
            FormatStyle(8)  =   "frmEntregas2.frx":0B14
            FormatStyle(9)  =   "frmEntregas2.frx":0BA8
            FormatStyle(10) =   "frmEntregas2.frx":0C3C
            FormatStyle(11) =   "frmEntregas2.frx":0CD4
            FormatStyle(12) =   "frmEntregas2.frx":0D70
            ImageCount      =   0
            PrinterProperties=   "frmEntregas2.frx":0E08
         End
         Begin XtremeSuiteControls.PushButton btnExportarExcel 
            Height          =   465
            Left            =   6480
            TabIndex        =   16
            Top             =   2835
            Width           =   1140
            _Version        =   786432
            _ExtentX        =   2011
            _ExtentY        =   820
            _StockProps     =   79
            Caption         =   "Exportar"
            UseVisualStyle  =   -1  'True
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   3390
         Left            =   90
         TabIndex        =   17
         Top             =   375
         Width           =   7815
         _Version        =   786432
         _ExtentX        =   13785
         _ExtentY        =   5980
         _StockProps     =   79
         Caption         =   "Facturas del detalle seleccionado de la OT"
         UseVisualStyle  =   -1  'True
         Begin GridEX20.GridEX gridFacturas 
            Height          =   3060
            Left            =   90
            TabIndex        =   18
            Top             =   255
            Width           =   6105
            _ExtentX        =   10769
            _ExtentY        =   5398
            Version         =   "2.0"
            HoldSortSettings=   -1  'True
            BoundColumnIndex=   ""
            ReplaceColumnIndex=   ""
            ColumnAutoResize=   -1  'True
            MethodHoldFields=   -1  'True
            GroupByBoxVisible=   0   'False
            DataMode        =   99
            ColumnHeaderHeight=   285
            IntProp1        =   0
            IntProp2        =   0
            IntProp7        =   0
            ColumnsCount    =   3
            Column(1)       =   "frmEntregas2.frx":0FE0
            Column(2)       =   "frmEntregas2.frx":1120
            Column(3)       =   "frmEntregas2.frx":1258
            SortKeysCount   =   1
            SortKey(1)      =   "frmEntregas2.frx":1384
            FormatStylesCount=   12
            FormatStyle(1)  =   "frmEntregas2.frx":13EC
            FormatStyle(2)  =   "frmEntregas2.frx":1524
            FormatStyle(3)  =   "frmEntregas2.frx":15D4
            FormatStyle(4)  =   "frmEntregas2.frx":1688
            FormatStyle(5)  =   "frmEntregas2.frx":1760
            FormatStyle(6)  =   "frmEntregas2.frx":1818
            FormatStyle(7)  =   "frmEntregas2.frx":18F8
            FormatStyle(8)  =   "frmEntregas2.frx":1A28
            FormatStyle(9)  =   "frmEntregas2.frx":1ABC
            FormatStyle(10) =   "frmEntregas2.frx":1B50
            FormatStyle(11) =   "frmEntregas2.frx":1BE8
            FormatStyle(12) =   "frmEntregas2.frx":1C84
            ImageCount      =   0
            PrinterProperties=   "frmEntregas2.frx":1D1C
         End
         Begin XtremeSuiteControls.PushButton PushButton1 
            Height          =   465
            Left            =   6480
            TabIndex        =   19
            Top             =   2835
            Width           =   1140
            _Version        =   786432
            _ExtentX        =   2011
            _ExtentY        =   820
            _StockProps     =   79
            Caption         =   "Exportar"
            UseVisualStyle  =   -1  'True
         End
      End
   End
   Begin XtremeSuiteControls.PushButton mnuAtajo 
      Height          =   465
      Left            =   10560
      TabIndex        =   11
      Top             =   4665
      Width           =   1635
      _Version        =   786432
      _ExtentX        =   2884
      _ExtentY        =   820
      _StockProps     =   79
      Caption         =   "Remito rápido"
      UseVisualStyle  =   -1  'True
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7320
      Top             =   6720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin XtremeSuiteControls.GroupBox grpE 
      Height          =   1395
      Left            =   8280
      TabIndex        =   5
      Top             =   5880
      Width           =   4110
      _Version        =   786432
      _ExtentX        =   7250
      _ExtentY        =   2461
      _StockProps     =   79
      Caption         =   "Valores"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.Label lblPorcentajeEntregas 
         Height          =   195
         Left            =   210
         TabIndex        =   8
         Top             =   990
         Width           =   885
         _Version        =   786432
         _ExtentX        =   1561
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "% Entregas: "
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblPorcentajeFabricacion 
         Height          =   195
         Left            =   225
         TabIndex        =   7
         Top             =   645
         Width           =   1080
         _Version        =   786432
         _ExtentX        =   1905
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "% Fabricación: "
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblPorcentajeAvance 
         Height          =   195
         Left            =   225
         TabIndex        =   6
         Top             =   300
         Width           =   810
         _Version        =   786432
         _ExtentX        =   1429
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "% Avance: "
         AutoSize        =   -1  'True
      End
   End
   Begin XtremeSuiteControls.PushButton btnCerrar 
      Height          =   465
      Left            =   10560
      TabIndex        =   2
      Top             =   4140
      Width           =   1650
      _Version        =   786432
      _ExtentX        =   2910
      _ExtentY        =   820
      _StockProps     =   79
      Caption         =   "Cerrar OT"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.GroupBox grpDetalles 
      Height          =   3975
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   12375
      _Version        =   786432
      _ExtentX        =   21828
      _ExtentY        =   7011
      _StockProps     =   79
      Caption         =   "Detalles de OT"
      UseVisualStyle  =   -1  'True
      Begin GridEX20.GridEX gridDetalles 
         Height          =   3540
         Left            =   120
         TabIndex        =   1
         Top             =   255
         Width           =   12210
         _ExtentX        =   21537
         _ExtentY        =   6244
         Version         =   "2.0"
         PreviewRowIndent=   300
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         PreviewColumn   =   "pieza"
         PreviewRowLines =   1
         MultiSelect     =   -1  'True
         MethodHoldFields=   -1  'True
         GroupByBoxVisible=   0   'False
         DataMode        =   99
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   12
         Column(1)       =   "frmEntregas2.frx":1EF4
         Column(2)       =   "frmEntregas2.frx":2004
         Column(3)       =   "frmEntregas2.frx":20F0
         Column(4)       =   "frmEntregas2.frx":21F8
         Column(5)       =   "frmEntregas2.frx":2324
         Column(6)       =   "frmEntregas2.frx":2430
         Column(7)       =   "frmEntregas2.frx":256C
         Column(8)       =   "frmEntregas2.frx":26A8
         Column(9)       =   "frmEntregas2.frx":27E4
         Column(10)      =   "frmEntregas2.frx":2910
         Column(11)      =   "frmEntregas2.frx":2A2C
         Column(12)      =   "frmEntregas2.frx":2B2C
         FmtConditionsCount=   1
         FmtCondition(1) =   "frmEntregas2.frx":2C8C
         FormatStylesCount=   7
         FormatStyle(1)  =   "frmEntregas2.frx":2D50
         FormatStyle(2)  =   "frmEntregas2.frx":2E88
         FormatStyle(3)  =   "frmEntregas2.frx":2F38
         FormatStyle(4)  =   "frmEntregas2.frx":2FEC
         FormatStyle(5)  =   "frmEntregas2.frx":30C4
         FormatStyle(6)  =   "frmEntregas2.frx":317C
         FormatStyle(7)  =   "frmEntregas2.frx":325C
         ImageCount      =   0
         PrinterProperties=   "frmEntregas2.frx":32F0
      End
   End
   Begin XtremeSuiteControls.PushButton btnRemitar 
      Height          =   465
      Left            =   8400
      TabIndex        =   3
      Top             =   4200
      Width           =   1650
      _Version        =   786432
      _ExtentX        =   2910
      _ExtentY        =   820
      _StockProps     =   79
      Caption         =   "Remitar"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton btnAplicarRemito 
      Height          =   465
      Left            =   8400
      TabIndex        =   4
      Top             =   4680
      Width           =   1650
      _Version        =   786432
      _ExtentX        =   2910
      _ExtentY        =   820
      _StockProps     =   79
      Caption         =   "Aplicar Remito"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton cmdPreconteo 
      Height          =   465
      Left            =   10560
      TabIndex        =   12
      Top             =   5160
      Width           =   1635
      _Version        =   786432
      _ExtentX        =   2884
      _ExtentY        =   820
      _StockProps     =   79
      Caption         =   "Planilla de Preconteo"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton btnTomarDeStock 
      Height          =   465
      Left            =   8400
      TabIndex        =   21
      Top             =   5160
      Width           =   1650
      _Version        =   786432
      _ExtentX        =   2910
      _ExtentY        =   820
      _StockProps     =   79
      Caption         =   "Tomar de Stock"
      UseVisualStyle  =   -1  'True
   End
   Begin VB.Label lblCentroCostos 
      Height          =   240
      Left            =   8160
      TabIndex        =   10
      Top             =   7680
      Width           =   4140
   End
   Begin VB.Label lblCliente 
      Height          =   240
      Left            =   8280
      TabIndex        =   9
      Top             =   7440
      Width           =   4140
   End
   Begin VB.Menu mnuEmergente 
      Caption         =   "menu"
      Visible         =   0   'False
      Begin VB.Menu mnuVerDesarrollo 
         Caption         =   "Ver Desarrollo..."
      End
      Begin VB.Menu mnuARchivos 
         Caption         =   "Archivos Asociados"
         Begin VB.Menu mnuArchivoPieza 
            Caption         =   "De la Pieza..."
         End
         Begin VB.Menu mnuDelDetalle 
            Caption         =   "Del Detalle..."
         End
      End
   End
End
Attribute VB_Name = "frmEntregas2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Implements ISuscriber


Private valid As Boolean
Option Explicit
Private m_ot As OrdenTrabajo
Private detalle As DetalleOrdenTrabajo
Private Entregas As Collection
Private facturas As Collection
Private detalleRemito As remitoDetalle
Private detalleFactura As FacturaDetalle
Dim claseP As New classPlaneamiento
Dim rtoRapido As Boolean
Private m_id_suscriber As String
Private CantArchivos As New Dictionary
Private CantArchivosDetalle As New Dictionary





Public Sub SetOrdenTrabajo(Ot As OrdenTrabajo)
    Set m_ot = Ot
    Me.caption = "Entregas de la OT Nº " & m_ot.IdFormateado
    CargaDetalles
End Sub

Private Sub CargaDetalles()
    Me.gridDetalles.ItemCount = 0
    Me.gridEntregas.ItemCount = 0

    Set m_ot.Detalles = DAODetalleOrdenTrabajo.FindAllByOrdenTrabajo(m_ot.id, True, True, True)
    Me.gridDetalles.ItemCount = m_ot.Detalles.count

    Me.lblPorcentajeAvance.caption = "% Avance: " & m_ot.PorcentajeAvance
    Me.lblPorcentajeEntregas.caption = "% Entregas: " & m_ot.PorcentajeEntregas
    Me.lblPorcentajeFabricacion.caption = "% Fabricación: " & m_ot.PorcentajeFabricacion

    Me.lblCentroCostos = "Centro de costos: " & m_ot.cliente.razon
    If IsSomething(m_ot.ClienteFacturar) Then
        Me.lblCliente = "Cliente: " & m_ot.ClienteFacturar.razon
    End If

End Sub


Private Sub btnAplicarRemito_Click()
    Dim detaOT As DetalleOrdenTrabajo

    If Me.gridDetalles.SelectedItems.count = 1 Then
        Set detaOT = m_ot.Detalles.item(Me.gridDetalles.SelectedItems(1).RowIndex)

        If detaOT.CantidadFabricados + detaOT.ReservaStock - detaOT.CantidadEntregada > 0 Then
            'si hay elementos disponibles, procedo con elegir el item del remito que voy a aplicar
            'a esta OT
            Dim fListaRemito As New frmPlaneamientoRemitosListaProceso

            fListaRemito.mostrar = -1
            Set Selecciones.RemitoElegido = Nothing
            fListaRemito.Show 1


            If IsSomething(Selecciones.RemitoElegido) Then
                Dim fRemitoDet As New frmPlaneamientoRemitoVer
                Set fRemitoDet.Remito = Selecciones.RemitoElegido
                fRemitoDet.Usable = True
                fRemitoDet.Show
            Else
                Exit Sub
            End If


        Else
            MsgBox "No hay elementos disponibles para entregar.", vbInformation
        End If

    Else
        MsgBox "Debe seleccionar un item.", vbExclamation
    End If





End Sub

Private Sub btnCerrar_Click()
    If m_ot.TodoEntregado Then
        'se cierra
        If DAOOrdenTrabajo.Cerrar(m_ot) Then
            Set m_ot = DAOOrdenTrabajo.FindById(m_ot.id)
            CargaDetalles

            MsgBox "La OT " & m_ot.id & " se cerró correctamente.", vbInformation

            Dim v As New clsEventoObserver
            Set v.Elemento = m_ot
            v.EVENTO = modificar_
            Set v.Originador = Me
            v.Tipo = ordenesTrabajo
            Channel.Notificar v, TipoSuscripcion.ordenesTrabajo


            Unload Me
        End If
    Else
        If m_ot.estado = EstadoOT_Finalizado Then    'cerrada
            MsgBox "La OT ya se encuentra cerrada.", vbExclamation
        Else
            If m_ot.PuedeCerrarse Then
                Dim fEntTotal As New frmEntregaTotal
                fEntTotal.Pedido = m_ot.id
                fEntTotal.Show 1
                Unload Me
            Else
                MsgBox "Para cerrar la OT debe tener todo fabricado o proveniente de stock.", vbExclamation
            End If
        End If
    End If
End Sub


Private Sub btnImprimir_Click()
    On Error GoTo err4
    Me.CommonDialog1.ShowPrinter
    Dim x As Long
    For x = 1 To Me.CommonDialog1.Copies
        ImprimirEntregas
    Next
    Exit Sub
err4:
    MsgBox Err.Description
End Sub

Private Sub btnExportarExcel_Click()
    Dim detalle As DetalleOrdenTrabajo
    Dim Entregas As Collection
    Dim remitoDetalle As remitoDetalle

    Dim xlWorkbook As New Excel.Workbook
    Dim xlWorksheet As New Excel.Worksheet
    Dim xlApplication As New Excel.Application

    Set xlWorkbook = xlApplication.Workbooks.Add
    Set xlWorksheet = xlWorkbook.Worksheets.item(1)

    xlWorksheet.Activate

    'fila, columna

    xlWorksheet.Cells(1, 1).value = "Detalles de Entregas"
    xlWorksheet.Cells(2, 1).value = "C.Costos:"
    xlWorksheet.Cells(3, 1).value = "Referencia:"
    xlWorksheet.Cells(4, 1).value = "Fecha Entrega:"

    xlWorksheet.Cells(1, 2).value = "OT Nº " & m_ot.IdFormateado & " al dia " & Date
    xlWorksheet.Cells(2, 2).value = m_ot.cliente.razon & " Cliente: " & m_ot.ClienteFacturar.razon
    xlWorksheet.Cells(3, 2).value = m_ot.descripcion
    xlWorksheet.Range(xlWorksheet.Cells(3, 2), xlWorksheet.Cells(3, 2)).HorizontalAlignment = xlLeft
    xlWorksheet.Cells(4, 2).value = m_ot.FechaEntrega
    xlWorksheet.Range(xlWorksheet.Cells(4, 2), xlWorksheet.Cells(4, 2)).HorizontalAlignment = xlLeft

    xlWorksheet.Cells(6, 1) = "Item"
    xlWorksheet.Cells(6, 2) = "Detalle"
    xlWorksheet.Cells(6, 3) = "Cant Ped"
    xlWorksheet.Cells(6, 4) = "Cant Entreg"

    Dim row As Long: row = 7

    For Each detalle In m_ot.Detalles
        xlWorksheet.Cells(row, 1) = Format(detalle.item, "'000")
        xlWorksheet.Range(xlWorksheet.Cells(row, 1), xlWorksheet.Cells(row, 1)).HorizontalAlignment = xlLeft
        xlWorksheet.Cells(row, 2) = detalle.Pieza.nombre
        xlWorksheet.Cells(row, 3) = detalle.CantidadPedida
        xlWorksheet.Cells(row, 4) = detalle.CantidadEntregada

        Set Entregas = DAORemitoSDetalle.FindAllByDetallePedido(detalle.id)
        If Entregas.count > 0 Then
            row = row + 1
            xlWorksheet.Cells(row, 2) = "Cant Entreg"
            xlWorksheet.Range(xlWorksheet.Cells(row, 2), xlWorksheet.Cells(row, 2)).HorizontalAlignment = xlRight
            xlWorksheet.Cells(row, 3) = "Remito"
            xlWorksheet.Range(xlWorksheet.Cells(row, 3), xlWorksheet.Cells(row, 3)).HorizontalAlignment = xlRight
            xlWorksheet.Cells(row, 4) = "Fecha"
            xlWorksheet.Range(xlWorksheet.Cells(row, 4), xlWorksheet.Cells(row, 4)).HorizontalAlignment = xlRight
            Dim rto As Remito

            For Each remitoDetalle In Entregas


                Set rto = DAORemitoS.FindById(remitoDetalle.Remito)
                If rto.estado = RemitoAprobado Then
                    row = row + 1
                    xlWorksheet.Cells(row, 2) = remitoDetalle.Cantidad
                    xlWorksheet.Cells(row, 3) = remitoDetalle.RemitoAlQuePertenece.numero
                    xlWorksheet.Cells(row, 4) = CDate(CLng(remitoDetalle.FEcha))
                End If
            Next

        End If

        row = row + 2
    Next detalle


    'autosize
    xlApplication.ScreenUpdating = False
    Dim wkSt As String
    wkSt = xlWorksheet.Name
    xlWorksheet.Cells.EntireColumn.AutoFit
    xlWorkbook.Sheets(wkSt).Select
    xlApplication.ScreenUpdating = True
    ''

    Dim ruta As String
    ruta = Environ$("TEMP")
    If LenB(ruta) = 0 Then ruta = Environ$("TMP")
    If LenB(ruta) = 0 Then ruta = App.path
    ruta = ruta & "\" & funciones.CreateGUID() & ".xls"

    xlWorkbook.SaveAs ruta

    xlWorkbook.Saved = True
    xlWorkbook.Close
    xlApplication.Quit

    ShellExecute -1, "open", ruta, "", "", 4

    Set xlWorksheet = Nothing
    Set xlWorkbook = Nothing
    Set xlApplication = Nothing

End Sub

Private Sub btnRemitar_Click()
    Dim item As JSSelectedItem
    Dim detaOT As DetalleOrdenTrabajo
    Dim detasId() As Long
    ReDim detasId(0) As Long

    For Each item In Me.gridDetalles.SelectedItems
        Set detaOT = m_ot.Detalles.item(item.RowIndex)
        If detasId(0) <> 0 Then
            ReDim Preserve detasId(UBound(detasId, 1) + 1)
        End If
        detasId(UBound(detasId, 1)) = detaOT.id


        If (detaOT.CantidadFabricados + detaOT.ReservaStock) = 0 Then
            MsgBox "El item " & detaOT.item & " debe estar totalmente fabricado.", vbExclamation
            Exit Sub
        End If

    Next

    If Me.gridDetalles.SelectedItems.count = 1 Then    'entrega 1
        Dim fEntrega As New frmPlaneamientoRealizarEntrega

        Set detaOT = m_ot.Detalles.item(Me.gridDetalles.SelectedItems(1).RowIndex)
        Set fEntrega.deta = detaOT
        fEntrega.lblIdPieza = detaOT.id    ' detaOT.pieza.Id
        fEntrega.lblPieza = detaOT.Pieza.nombre
        fEntrega.lblPedidos = detaOT.CantidadPedida
        fEntrega.Text1 = detaOT.CantidadPedida - detaOT.CantidadFabricados
        fEntrega.lblFabricados = detaOT.CantidadFabricados
        fEntrega.lblEntregados = detaOT.CantidadEntregada
        fEntrega.lblDeStock = detaOT.ReservaStock
        fEntrega.lblOT = m_ot.id
        fEntrega.lblItem = detaOT.item
        fEntrega.Show 1

    ElseIf Me.gridDetalles.SelectedItems.count > 1 Then
        Dim f22 As New frmPlaneamientoRealizarEntregaMultiple
        f22.idP = m_ot.id
        f22.vector detasId
        f22.Show 1
    End If

    CargaDetalles
End Sub

Private Sub btnTomarDeStock_Click()
On Error GoTo err1
Dim res As String
res = InputBox("Ingrese la cantidad a tomar de stock (Máximo " & detalle.Pieza.CantidadStock & ")", "Reserva de Stock", "0")
If IsNumeric(res) Then

    Dim reserva As Double: reserva = Val(res)
    
    If DAOOrdenTrabajo.DescontarReservaDetalle(detalle, reserva) Then
            MsgBox "Reserva de Stock realizada!", vbInformation
    End If

End If
Exit Sub

err1:
MsgBox Err.Description

End Sub

Private Sub cmdPreconteo_Click()
    If Not Permisos.PlanRemitosControl Then
        MsgBox "No tiene permisos para realizar esta acción!", vbCritical, "Error"
        Exit Sub
    End If


    If MsgBox("¿Seguro de crear una planilla de preconteo para esta OT?", vbYesNo, "Consulta") = vbYes Then

        DAOOrdenTrabajo.ImprimirPreconteo m_ot
    End If
End Sub

Private Sub Command1_Click()
    conectar.BeginTransaction
    On Error GoTo error1
    Dim c As Long
    Dim q As String
    q = "SELECT  e.idDetallePedido AS id ,a.fechaEmision, f.cantidad, f.Valor, a.NroFactura FROM AdminFacturasDetalleNueva  f LEFT JOIN AdminFacturas a ON f.idFactura=a.id INNER JOIN entregas e ON f.idEntrega=e.id where a.estado<>3 ORDER BY a.NroFactura"


    Dim deta As FacturaDetalle
    Dim rs As Recordset
    Dim rs2 As Recordset
    Set rs = conectar.RSFactory(q)
    Dim cont As Long
    c = 0
    While Not rs.EOF And Not rs.BOF

        q = "select * from detalles_pedidos_cantidad where id_detalle_pedido=" & rs!id    '& " and tipo_cantidad=2"
        Set rs2 = conectar.RSFactory(q)
        cont = 0
        While Not rs2.EOF And Not rs2.BOF
            cont = cont + 1
            rs2.MoveNext
        Wend



        If Not conectar.execute("INSERT INTO sp.detalles_pedidos_cantidad (id_detalle_pedido, cantidad, fecha, tipo_cantidad, monto) VALUES  (' " & rs!id & " ',   '" & rs!Cantidad & " ',  '" & funciones.datetimeFormateada(Now) & "',    '2',    ' " & rs!Valor & " '     )") Then GoTo error1
        c = c + 1
        Debug.Print c


        rs.MoveNext

    Wend
    conectar.CommitTransaction

    Exit Sub
error1:
    conectar.RollBackTransaction

End Sub

Private Sub Form_Load()
    Customize Me
    GridEXHelper.CustomizeGrid Me.gridDetalles
    GridEXHelper.CustomizeGrid Me.gridEntregas
    GridEXHelper.CustomizeGrid Me.gridFacturas
    Me.gridDetalles.ItemCount = 0
    Me.gridEntregas.ItemCount = 0
    Channel.AgregarSuscriptor Me, FacturarRemitosDetalle_, True

    Set CantArchivos = DAOArchivo.GetCantidadArchivosPorReferencia(OA_Piezas)
    Set CantArchivosDetalle = DAOArchivo.GetCantidadArchivosPorReferencia(OA_OrdenesTrabajoDetalle)

valid = (m_ot.TipoOrden = OT_TRADICIONAL Or m_ot.TipoOrden = OT_TRADICIONAL)

Me.btnAplicarRemito.Enabled = valid
Me.btnCerrar.Enabled = valid
Me.mnuAtajo.Enabled = valid
Me.btnRemitar.Enabled = valid
Me.btnTomarDeStock.Enabled = valid
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Channel.RemoverSuscripcionTotal Me
End Sub

Private Sub gridDetalles_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim idx As Long
    idx = Me.gridDetalles.RowIndex(Me.gridDetalles.row)

    If Button = 2 And idx > 0 Then

        Me.mnuArchivoPieza.caption = "De la Pieza... " & CantArchivos.item(detalle.Pieza.id)

        Me.mnuDelDetalle.caption = "Del Detalle... " & CantArchivosDetalle.item(detalle.id)



        Me.PopupMenu Me.mnuEmergente
    End If

End Sub

Private Sub gridDetalles_RowFormat(RowBuffer As GridEX20.JSRowData)
    On Error Resume Next

    Set detalle = m_ot.Detalles.item(RowBuffer.RowIndex)

    If detalle.FechaEntrega < Date Then
        RowBuffer.CellStyle(5) = GridEXHelper.GRID_FORMATSTYLE_ROJO
    ElseIf detalle.FechaEntrega > Date Then
        RowBuffer.CellStyle(5) = GridEXHelper.GRID_FORMATSTYLE_VERDE


    End If

End Sub

Private Sub gridDetalles_SelectionChange()
    Me.gridEntregas.ItemCount = 0

    If Me.gridDetalles.RowIndex(Me.gridDetalles.row) > 0 Then
        Set detalle = m_ot.Detalles.item(Me.gridDetalles.RowIndex(Me.gridDetalles.row))
        Set Entregas = DAORemitoSDetalle.FindAllByDetallePedido(detalle.id)

        If Permisos.AdminFacturaConsultas Then
            Me.gridFacturas.ItemCount = 0
            Set facturas = DAOFacturaDetalles.FindAll("entregas.idDetallePedido =" & detalle.id, True)

            Me.gridFacturas.ItemCount = facturas.count
        End If
        Me.gridEntregas.ItemCount = Entregas.count


    End If


End Sub

Private Sub gridDetalles_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex > 0 And m_ot.Detalles.count > 0 Then
        Set detalle = m_ot.Detalles.item(RowIndex)
        Values(1) = detalle.item
        Values(2) = detalle.Nota
        Values(3) = detalle.Pieza.nombre
        
        Values(4) = detalle.CantidadPedida
        
        If m_ot.TipoOrden = OT_STOCK Then
            Values(4) = Values(4) & " (" & DAODetalleOrdenTrabajo.PendientesEntregaPorPieza(detalle.Pieza.id) & ")"
        End If
        
        Values(5) = detalle.FechaEntrega
        Values(6) = detalle.CantidadFabricados
        Values(7) = detalle.CantidadEntregada
        Values(8) = detalle.CantidadPedida - detalle.CantidadEntregada
        Values(9) = detalle.CantidadFacturada
        Values(10) = detalle.ReservaStock
        Values(11) = detalle.Pieza.UnidadMedida
        Values(12) = detalle.Pieza.CantidadStock
    End If
End Sub

Private Sub gridEntregas_RowFormat(RowBuffer As GridEX20.JSRowData)
    On Error Resume Next

    Set detalleRemito = Entregas.item(RowBuffer.RowIndex)

    If detalleRemito.RemitoAlQuePertenece.estado = RemitoPendiente Then
        RowBuffer.RowStyle = "Pendiente"
    Else

        If detalleRemito.RemitoAlQuePertenece.estado = RemitoAnulado Then
            RowBuffer.RowStyle = GridEXHelper.GRID_FORMATSTYLE_ANULADA

        End If

        If detalleRemito.RemitoAlQuePertenece.EstadoFacturado = RemitoNoFacturado Then
            RowBuffer.CellStyle(4) = "fondo_rojo"
        ElseIf detalleRemito.RemitoAlQuePertenece.EstadoFacturado = RemitoFacturadoTotal Then
            RowBuffer.CellStyle(4) = "fondo_verde"
        ElseIf detalleRemito.RemitoAlQuePertenece.EstadoFacturado = RemitoFacturadoParcial Then
            RowBuffer.CellStyle(4) = "fondo_naranja"
        ElseIf detalleRemito.RemitoAlQuePertenece.EstadoFacturado = RemitoNoFacturable Then
            RowBuffer.CellStyle(4) = "fondo_negro"
        End If
    End If
End Sub

Private Sub gridEntregas_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    On Error Resume Next
    If RowIndex > 0 And Entregas.count > 0 Then
        Set detalleRemito = Entregas.item(RowIndex)
        Values(1) = detalleRemito.RemitoAlQuePertenece.numero
        Values(2) = detalleRemito.Cantidad
        Values(3) = detalleRemito.FEcha
    End If
End Sub

Private Sub ImprimirEntregas()
    Dim rs As Recordset
    Dim rs2 As Recordset
    Printer.Font.Size = 10
    Printer.Font.Bold = True
    Printer.Orientation = 1
    Printer.Print "DETALLE DE ENTREGAS O/T Nro " & m_ot.IdFormateado & " al día " & Date

    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)
    Printer.Print

    '    Set rs = conectar.RSFactory("select p.*,c.razon from pedidos p inner join clientes c on p.idcliente=c.id where p.id=" & idOt)
    '    If Not rs.EOF And Not rs.BOF Then
    '        cli = rs!idCliente
    '        clie = rs!Razon
    '        referencia = rs!Descripcion
    '        entrega = rs!FechaEntrega
    '    Else
    '        Exit Sub
    '    End If

    Printer.Print "C.Costos: " & m_ot.cliente.id & " - " & m_ot.cliente.razon,
    Printer.Print "Referencia: " & UCase(m_ot.descripcion)
    Printer.Print "Entrega: " & Format(m_ot.FechaEntrega, "dd-mm-yyyy")
    Printer.Print
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)


    'Acá se imprimen los encabezados de la Lista
    Printer.Print Tab(1);
    Printer.Print "Item";
    Printer.Print Tab(10);
    Printer.Print "Detalle";
    Printer.Print Tab(80);
    Printer.Print "Cant";
    Printer.Print Tab(90);
    Printer.Print "Entregados"
    Printer.Font.Bold = False
    Printer.Print
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)


    'aca se imprime la lista de elementos con sus entregas
    Dim detalle As DetalleOrdenTrabajo
    Dim Entregas As Collection
    Dim detalleRemito As remitoDetalle
    Dim remi As Remito

    For Each detalle In m_ot.Detalles

        Printer.Print Tab(1);
        Printer.Print Format(detalle.item, "000");
        Printer.Print Tab(12);
        Printer.Print UCase(detalle.Pieza.nombre);
        Printer.Print Tab(90);
        Printer.Print detalle.CantidadPedida;
        Printer.Print Tab(100);
        Printer.Print detalle.CantidadEntregada

        Set Entregas = DAORemitoSDetalle.FindAllByDetallePedido(detalle.id)

        If Entregas.count > 0 Then
            Printer.Print
            Printer.FontBold = True
            Printer.Print Tab(65);
            Printer.Print "Cant";
            Printer.Print Tab(75);
            Printer.Print "Remito";
            Printer.Print Tab(85);
            Printer.Print "Fecha";
            Printer.FontBold = False
        End If

        For Each detalleRemito In Entregas
            Printer.Print Tab(72);
            Printer.Print detalleRemito.Cantidad;
            Printer.Print Tab(82);
            Set remi = DAORemitoS.FindById(detalleRemito.Remito)
            If IsSomething(remi) Then
                Printer.Print remi.numero
            Else
                Printer.Print
            End If
            Printer.Print Tab(92);
            Printer.Print CDate(CLng(detalleRemito.FEcha))
        Next

        Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)
        Printer.Print
    Next detalle


    Printer.Print

    Printer.Print "Fecha emisión " & Date


    Printer.EndDoc

End Sub


Private Sub gridFacturas_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    On Error Resume Next
    If RowIndex > 0 And facturas.count > 0 Then
        Set detalleFactura = facturas.item(RowIndex)
        Values(1) = detalleFactura.Factura.numero
        Values(2) = detalleFactura.Cantidad
        Values(3) = detalleFactura.Factura.FechaEmision
    End If
End Sub

Private Property Get ISuscriber_id() As String
    If LenB(m_id_suscriber) = 0 Then m_id_suscriber = funciones.CreateGUID()
    ISuscriber_id = m_id_suscriber
End Property

Private Function ISuscriber_Notificarse(EVENTO As clsEventoObserver) As Variant
    On Error GoTo err1
    If rtoRapido Then Exit Function



    If EVENTO.Tipo = FacturarRemitosDetalle_ Then
        'If ReadOnly Then

        'aplicacion de detalle remito post facturacion
        If Not IsSomething(EVENTO.Elemento) Then Exit Function

        Set detalle = Nothing
        If Me.gridDetalles.row > 0 Then
            Set detalle = m_ot.Detalles.item(Me.gridDetalles.RowIndex(Me.gridDetalles.row))
        End If



        If EVENTO.Elemento.count > 0 And IsSomething(detalle) Then
            Dim redeta As remitoDetalle
            Set redeta = EVENTO.Elemento.item(1)



            If redeta.Cantidad <= (detalle.CantidadPedida - detalle.CantidadEntregada) Then

                If redeta.Origen = OrigenRemitoConcepto Then
                    If MsgBox("¿Está seguro de aplicar este remito a este item de la OT?", vbYesNo + vbQuestion) = vbYes Then
                        If claseP.aplicarRemitoAOT(m_ot.id, redeta.id, detalle.id, redeta.Cantidad) Then
                            MsgBox "Remito aplicado correctamente.", vbInformation
                            CargaDetalles
                        Else
                            MsgBox "Se produjo algún error. No se graban los cambios.", vbCritical
                        End If
                    End If
                Else
                    Err.Raise 9991, , "El item del remito debe ser concepto para poder aplicar"
                End If



            Else

                If MsgBox("La cantidad de la entrega [" & redeta.Cantidad & "] debe ser menor o igual a la cantidad pedida menos la cantidad entregada del detalle de la OT. " + vbNewLine + "Aplicar de todas formas?", vbYesNo, "Consulta") = vbYes Then
                    Err.Raise 9992, , "Debe ser iguales en cantidad para poder aplicar. Modulo:FrmEntregas2."
                    If MsgBox("¿Está seguro de aplicar este remito a este item de la OT?", vbYesNo + vbQuestion) = vbYes Then



                        If claseP.aplicarRemitoAOT(m_ot.id, redeta.id, detalle.id, detalle.CantidadPedida - detalle.CantidadEntregada) Then
                            MsgBox "Remito aplicado correctamente.", vbInformation
                            CargaDetalles
                        Else
                            MsgBox "Se produjo algún error. No se graban los cambios.", vbCritical
                        End If


                    End If
                Else
                    Err.Raise 9992, , "Debe ser iguales en cantidad para poder aplicar"
                End If

            End If



        End If
    End If
    Exit Function
err1:
    MsgBox Err.Description, vbExclamation + vbOKOnly
End Function


Private Function QuickRemito(itemsDisponibles As Long) As Boolean
    itemsDisponibles = funciones.itemsPorRemito
    On Error GoTo err1
    Dim SalioDelFor As Boolean
    Set m_ot.Detalles = DAODetalleOrdenTrabajo.FindAllByOrdenTrabajo(m_ot.id, True, True, True)
    Dim deta As DetalleOrdenTrabajo
    Dim Cant As Double
    Dim cantEntrega As Double
    Dim Remito As Remito
    Dim redeta As remitoDetalle

    Set Remito = New Remito
    Remito.FEcha = Now
    Remito.detalle = m_ot.descripcion
    Set Remito.cliente = m_ot.ClienteFacturar

    Remito.estado = RemitoPendiente
    Remito.EstadoFacturado = RemitoNoFacturado

    Set Remito.usuarioCreador = funciones.GetUserObj
    Remito.numero = DAORemitoS.ProximoRemito
    Set Remito.Detalles = New Collection
    If Not DAORemitoS.Guardar(Remito, False) Then GoTo err1

    Dim muestra_observaciones As Boolean


    muestra_observaciones = MsgBox("Desea incluir las observaciones de los items de la orden?", vbYesNo, "Consulta") = vbYes
    For Each deta In m_ot.Detalles


        'realizo el seguimiento
        Cant = deta.CantidadPedida - deta.CantidadFabricados
        If Cant > 0 Then

            deta.CantidadFabricados = deta.CantidadFabricados + Cant
            deta.CantidadFabricadosStatic = deta.CantidadFabricadosStatic + Cant
            If Not DAODetalleOrdenTrabajo.SaveCantidad(deta.id, Cant, CantidadFabricada_, 0, Remito.id, 0, 0, 0) Then GoTo err1
        End If

        'realizo la entrega
        cantEntrega = deta.CantidadPedida - deta.CantidadEntregada
        If cantEntrega > 0 Then
            Set redeta = New remitoDetalle
            redeta.Cantidad = cantEntrega
            Set redeta.DetallePedido = deta
            redeta.EstadoRemito = Remito.estado
            redeta.facturable = True
            redeta.Facturado = False
            redeta.FEcha = Now


            If muestra_observaciones Then
                redeta.observaciones = deta.Nota
            End If

            redeta.idDetallePedido = deta.id
            redeta.idpedido = m_ot.id
            redeta.Origen = OrigenRemitoOt
            redeta.Remito = Remito.id
            redeta.Valor = deta.Precio
            redeta.ValorModificado = False
            Set redeta.RemitoAlQuePertenece = Remito

            If itemsDisponibles = 0 Then
                SalioDelFor = True
                Exit For

            End If

            Remito.Detalles.Add redeta
            deta.CantidadEntregadaStatic = redeta.Cantidad + deta.CantidadEntregadaStatic
            deta.CantidadEntregada = redeta.Cantidad + deta.CantidadEntregada

            'If Not DAODetalleOrdenTrabajo.SaveCantidad(deta.id, redeta.Cantidad, CantidadEntregada_, redeta.valor) Then GoTo ERR1

            If LenB(redeta.observaciones) > 0 Then
                itemsDisponibles = itemsDisponibles - 2
            Else
                itemsDisponibles = itemsDisponibles - 1
            End If

            If Not DAODetalleOrdenTrabajo.Save(deta) Then GoTo err1

        End If
    Next deta

    If Not DAORemitoS.Guardar(Remito, True) Then GoTo err1
    MsgBox "Orden remitada correctamente!", vbExclamation
    If SalioDelFor Then
        If MsgBox("Quedan items para procesar, ¿hacerlo con un nuevo remito?", vbYesNo, "Consulta") = vbYes Then
            If Not QuickRemito(itemsDisponibles) Then GoTo err1
        End If
    End If
    QuickRemito = True
    CargaDetalles
    Exit Function
err1:
    QuickRemito = False


End Function

Private Sub mnuArchivoPieza_Click()
    gridDetalles_SelectionChange
    Dim F As New frmArchivos2
    F.Origen = OrigenArchivos.OA_Piezas
    F.ObjetoId = detalle.Pieza.id
    F.caption = "OT Nº " & m_ot.IdFormateado & " - Item " & detalle.item
    F.Show
End Sub

Private Sub mnuAtajo_Click()
    If Not Permisos.PlanRemitosControl Then
        MsgBox "No tiene permisos para realizar esta acción!", vbCritical, "Error"
        Exit Sub
    End If


    If MsgBox("¿Seguro de crear un remito rápido para esta OT?", vbYesNo, "Consulta") = vbYes Then
        rtoRapido = True
        On Error GoTo err1
        conectar.BeginTransaction
        Dim itemsDisponibles As Long
        itemsDisponibles = funciones.itemsPorRemito


        If Not QuickRemito(itemsDisponibles) Then GoTo err1
        conectar.CommitTransaction
    End If
    rtoRapido = False
    Exit Sub

err1:
    conectar.RollBackTransaction
End Sub


Private Sub mnuDelDetalle_Click()
    gridDetalles_SelectionChange
    Dim F As New frmArchivos2
    F.Origen = OrigenArchivos.OA_OrdenesTrabajoDetalle
    F.ObjetoId = detalle.id
    F.caption = "OT Nº " & m_ot.IdFormateado & " - Item " & detalle.item
    F.Show
End Sub

Private Sub mnuVerDesarrollo_Click()
    gridDetalles_SelectionChange
    'Dim idx As Long
    'idx = Me.grid.RowIndex(Me.grid.row)
    'If idx > 0 Then
    Dim F As New frmDesarrollo
    Load F
    F.CargarPieza detalle.Pieza.id     'm_ot.Detalles(idx).Pieza.Id
    F.Show

    '    End If
End Sub
