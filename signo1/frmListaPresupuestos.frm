VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmVentasPresupuestoLista 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Lista de cotizaciones"
   ClientHeight    =   7065
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14970
   Icon            =   "frmListaPresupuestos.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7065
   ScaleWidth      =   14970
   Begin XtremeSuiteControls.PushButton cmdImprimir 
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   6585
      Width           =   1095
      _Version        =   786432
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Imprimir"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   2205
      Left            =   120
      TabIndex        =   10
      Top             =   60
      Width           =   10035
      _Version        =   786432
      _ExtentX        =   17701
      _ExtentY        =   3889
      _StockProps     =   79
      Caption         =   "Parámetros de búsqueda"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.PushButton Command5 
         Default         =   -1  'True
         Height          =   375
         Left            =   8505
         TabIndex        =   4
         Top             =   1440
         Width           =   1335
         _Version        =   786432
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Buscar"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cboClientes 
         Height          =   315
         Left            =   4860
         TabIndex        =   0
         Top             =   240
         Width           =   4440
         _Version        =   786432
         _ExtentX        =   7832
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Text            =   "ComboBox1"
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Left            =   4860
         TabIndex        =   2
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CheckBox chkVerDetalle 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mostrar Detalles"
         Height          =   255
         Left            =   8160
         TabIndex        =   5
         Top             =   1860
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.TextBox txtFiltro 
         Height          =   285
         Left            =   4860
         TabIndex        =   3
         Top             =   675
         Width           =   4935
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1815
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   3201
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Estado"
            Object.Width           =   4939
         EndProperty
      End
      Begin XtremeSuiteControls.PushButton CMDsINCliente 
         Height          =   255
         Left            =   9420
         TabIndex        =   1
         Top             =   270
         Width           =   390
         _Version        =   786432
         _ExtentX        =   688
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "X"
         UseVisualStyle  =   -1  'True
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Número"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3660
         TabIndex        =   13
         Top             =   1110
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3780
         TabIndex        =   12
         Top             =   300
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3660
         TabIndex        =   11
         Top             =   690
         Width           =   1095
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5520
      Top             =   6600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   13
      ImageHeight     =   13
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaPresupuestos.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaPresupuestos.frx":035E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaPresupuestos.frx":07B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaPresupuestos.frx":0C06
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaPresupuestos.frx":105A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaPresupuestos.frx":14AE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin GridEX20.GridEX grilla 
      Height          =   4215
      Left            =   -15
      TabIndex        =   7
      Top             =   2340
      Width           =   16095
      _ExtentX        =   28390
      _ExtentY        =   7435
      Version         =   "2.0"
      PreviewRowIndent=   200
      AutomaticSort   =   -1  'True
      HoldSortSettings=   -1  'True
      DefaultGroupMode=   1
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      PreviewColumn   =   3
      PreviewRowLines =   1
      ColumnAutoResize=   -1  'True
      MultiSelect     =   -1  'True
      MethodHoldFields=   -1  'True
      DatabaseName    =   "3"
      ForeColorInfoText=   -2147483639
      BackColorInfoText=   8421504
      GroupByBoxInfoText=   "Arrastre una columna para agrupar"
      AllowEdit       =   0   'False
      BackColorGBBox  =   8421504
      BackColorHeader =   16761024
      ImageWidth      =   14
      ImageHeight     =   14
      ImageCount      =   6
      ImagePicture1   =   "frmListaPresupuestos.frx":1902
      ImagePicture2   =   "frmListaPresupuestos.frx":1C54
      ImagePicture3   =   "frmListaPresupuestos.frx":1FA6
      ImagePicture4   =   "frmListaPresupuestos.frx":22F8
      ImagePicture5   =   "frmListaPresupuestos.frx":264A
      ImagePicture6   =   "frmListaPresupuestos.frx":299C
      RowHeaders      =   -1  'True
      DataMode        =   99
      HeaderFontName  =   "Tahoma"
      FontName        =   "Tahoma"
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp7        =   0
      ColumnsCount    =   10
      Column(1)       =   "frmListaPresupuestos.frx":2CEE
      Column(2)       =   "frmListaPresupuestos.frx":2E62
      Column(3)       =   "frmListaPresupuestos.frx":2F32
      Column(4)       =   "frmListaPresupuestos.frx":3056
      Column(5)       =   "frmListaPresupuestos.frx":3146
      Column(6)       =   "frmListaPresupuestos.frx":3216
      Column(7)       =   "frmListaPresupuestos.frx":3372
      Column(8)       =   "frmListaPresupuestos.frx":3462
      Column(9)       =   "frmListaPresupuestos.frx":355A
      Column(10)      =   "frmListaPresupuestos.frx":362E
      SortKeysCount   =   1
      SortKey(1)      =   "frmListaPresupuestos.frx":3706
      FmtConditionsCount=   1
      FmtCondition(1) =   "frmListaPresupuestos.frx":376E
      FormatStylesCount=   12
      FormatStyle(1)  =   "frmListaPresupuestos.frx":3832
      FormatStyle(2)  =   "frmListaPresupuestos.frx":395A
      FormatStyle(3)  =   "frmListaPresupuestos.frx":3A0A
      FormatStyle(4)  =   "frmListaPresupuestos.frx":3ABE
      FormatStyle(5)  =   "frmListaPresupuestos.frx":3B96
      FormatStyle(6)  =   "frmListaPresupuestos.frx":3C6E
      FormatStyle(7)  =   "frmListaPresupuestos.frx":3D4E
      FormatStyle(8)  =   "frmListaPresupuestos.frx":3E46
      FormatStyle(9)  =   "frmListaPresupuestos.frx":3EFA
      FormatStyle(10) =   "frmListaPresupuestos.frx":3FAE
      FormatStyle(11) =   "frmListaPresupuestos.frx":3FFA
      FormatStyle(12) =   "frmListaPresupuestos.frx":4082
      ImageCount      =   6
      ImagePicture(1) =   "frmListaPresupuestos.frx":4146
      ImagePicture(2) =   "frmListaPresupuestos.frx":4498
      ImagePicture(3) =   "frmListaPresupuestos.frx":47EA
      ImagePicture(4) =   "frmListaPresupuestos.frx":4B3C
      ImagePicture(5) =   "frmListaPresupuestos.frx":4E8E
      ImagePicture(6) =   "frmListaPresupuestos.frx":51E0
      PrinterProperties=   "frmListaPresupuestos.frx":5532
   End
   Begin XtremeSuiteControls.PushButton cmdEstad 
      Height          =   375
      Left            =   1320
      TabIndex        =   9
      Top             =   6600
      Width           =   1095
      _Version        =   786432
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Estadísticas"
      UseVisualStyle  =   -1  'True
   End
   Begin VB.Menu m4 
      Caption         =   "m5"
      Visible         =   0   'False
      Begin VB.Menu numero 
         Caption         =   "numero"
         Enabled         =   0   'False
      End
      Begin VB.Menu EditPres 
         Caption         =   "Editar..."
      End
      Begin VB.Menu AprobarPresu 
         Caption         =   "Aprobar"
      End
      Begin VB.Menu mnuCrearOT 
         Caption         =   "Crear OT..."
      End
      Begin VB.Menu detalle 
         Caption         =   "Ver Detalles..."
      End
      Begin VB.Menu historic 
         Caption         =   "Ver Historial..."
      End
      Begin VB.Menu verIncidencias 
         Caption         =   "Ver Incidencias..."
      End
      Begin VB.Menu Archivos 
         Caption         =   "Archivos Asociados..."
      End
      Begin VB.Menu scanear 
         Caption         =   "Adquirir..."
      End
      Begin VB.Menu n6 
         Caption         =   "-"
      End
      Begin VB.Menu desactiva 
         Caption         =   "Desactivar"
      End
      Begin VB.Menu ncotizar 
         Caption         =   "No Cotizar"
      End
      Begin VB.Menu enviar 
         Caption         =   "Enviar..."
      End
      Begin VB.Menu recotiza 
         Caption         =   "Re-Cotizar"
      End
   End
End
Attribute VB_Name = "frmVentasPresupuestoLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Implements ISuscriber

Dim srow As Long
Dim suscriber_id As String
Private tmpIncidencias As New Dictionary
Private tmpArchivos As New Dictionary
Dim rec2 As clsPresupuesto
Dim estados As New Collection
Dim filtro As String
Dim presupuestos As Collection
Dim rectmp As clsPresupuesto
Dim marca As Integer
Dim vAccion As Integer

Private Sub AprobarPresu_Click()
    If MsgBox("¿Está seguro de aprobar el presupuesto?", vbYesNo, "Confirmación") = vbYes Then
        If DAOPresupuestos.aprobar(rectmp) Then
            ROWS_ = grilla.RowIndex(grilla.row)
            grilla.RefreshRowIndex ROWS_
        Else
            llenar_Grilla
        End If
    End If
End Sub
Private Sub archivos_Click()
    Dim frmarchi1 As New frmArchivos2
    frmarchi1.Origen = 2
    frmarchi1.ObjetoId = rectmp.Id
    frmarchi1.caption = "Presupuesto Nº " & rectmp.Id
    frmarchi1.Show
End Sub
Private Sub chkVerDetalle_Click()
    mostrarDetalles
End Sub

Private Sub cmdImprimir_Click()
    Dim elegidos As Boolean
    If grilla.SelectedItems.count > 1 Then
        elegidos = True
    Else
        elegidos = False
    End If

    With Me.grilla.PrinterProperties
        .FitColumns = True
        .RepeatHeaders = True
        .Orientation = jgexPPLandscape
        .HeaderString(jgexHFCenter) = "Presupuestos"
        .FooterString(jgexHFCenter) = Now
    End With
    Load frmPrintPreview
    frmPrintPreview.Move Me.Left, Me.Top, Me.Width, Me.Height
    grilla.PrintPreview frmPrintPreview.GEXPreview1, elegidos
    frmPrintPreview.Show 1
End Sub
Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub Command3_Click()
    Unload Me
End Sub
Private Sub Command4_Click()
    frmVentasPresupuestosLista.Show
End Sub
Public Sub llenarEstados()
    Set estados = Nothing
    Dim i As ListItem
    For P = 1 To Me.ListView1.ListItems.count
        Set i = Me.ListView1.ListItems(P)
        If i.Checked Then
            estados.Add i.Tag
        End If
    Next
End Sub



Private Sub CMDsINCliente_Click()
    Me.cboClientes.ListIndex = -1
End Sub


Private Sub cmdEstad_Click()
    Dim P As JSSelectedItem
    Dim pto As clsPresupuesto
    Dim dp As clsPresupuestoDetalle
    Dim listadtopiezacantidad As New Collection

    Dim colHistoricos As Collection
    Dim dhp As clsPresupuestoDetalleHistorico


    Dim sectoresTiempo As New Collection


    For Each P In Me.grilla.SelectedItems
        Set pto = presupuestos.item(P.RowIndex)
        Set pto.DetallePresupuesto = DAOPresupuestosDetalle.GetAllByPresupuesto(pto)

        If pto.EstadoPresupuesto = EstadoPresupuesto.Desactivado Then
            'no hace nada
        Else
            If pto.EstadoPresupuesto = EstadoPresupuesto.ACotizar_ Then
                'lo hace por pieza
                For Each dp In pto.DetallePresupuesto
                    Set dto = New DTOPiezaCantidad
                    Set dto.Pieza = dp.Pieza
                    dto.Cantidad = dp.Cantidad
                    listadtopiezacantidad.Add dto
                Next dp
            Else
                'lo hace por historico
                For Each dp In pto.DetallePresupuesto
                    Set colHistoricos = DAODetallePresupuestoHistorico.FindAllByDetallePresupuestoId(dp.Id)
                    For Each dhp In colHistoricos
                        ProcessChild sectoresTiempo, dhp, dp.Cantidad
                    Next dhp
                Next dp
            End If
        End If
    Next P

    Dim frmEstad As New frmEstadistiacasEnCurso
    frmEstad.caption = "Estadisticas de presupuestos seleccionados"
    frmEstad.conjGrabado = True
    Set frmEstad.col = MergeEstadisticas(sectoresTiempo, DAOPieza.ListaDTOTiempoPorSector(listadtopiezacantidad))
    frmEstad.LlenarGridDesdeOT
    frmEstad.Show

End Sub

Private Sub ProcessChild(sectoresTiempo As Collection, dH As clsPresupuestoDetalleHistorico, ByVal Cant As Double)
    Dim mdoHist As PresupuestoDetalleHistoricoMDO
    Dim tmp As PresupuestoDetalleHistoricoMDO

    For Each mdoHist In dH.historicoMDO
        AddTarea sectoresTiempo, mdoHist.Tarea, Cant, mdoHist.CantOperarios, mdoHist.Tiempo
    Next mdoHist

    For Each tmp In dH.HistoricoHijos
        ProcessChild sectoresTiempo, tmp, Cant
    Next tmp
End Sub

Private Function MergeEstadisticas(colOT As Collection, colPiezas As Collection) As Collection
    Dim tmpSectorTiempo As DTOSectoresTiempo
    Dim sectorTiempo As DTOSectoresTiempo
    Dim tareaTiempo As DTOTareaTiempo
    Dim tmpTareaTiempo As DTOTareaTiempo

    For Each sectorTiempo In colPiezas
        If funciones.BuscarEnColeccion(colOT, CStr(sectorTiempo.Sector.Id)) Then
            Set tmpSectorTiempo = colOT.item(CStr(sectorTiempo.Sector.Id))
        Else
            Set tmpSectorTiempo = New DTOSectoresTiempo
            Set tmpSectorTiempo.Sector = sectorTiempo.Sector
            colOT.Add tmpSectorTiempo, CStr(tmpSectorTiempo.Sector.Id)
        End If

        For Each tareaTiempo In sectorTiempo.ListaDtoTareaTiempo
            If BuscarEnColeccion(tmpSectorTiempo.ListaDtoTareaTiempo, CStr(tareaTiempo.Tarea.Id)) Then
                Set tmpTareaTiempo = tmpSectorTiempo.ListaDtoTareaTiempo.item(CStr(tareaTiempo.Tarea.Id))
            Else
                Set tmpTareaTiempo = New DTOTareaTiempo
                Set tmpTareaTiempo.Tarea = tareaTiempo.Tarea
                tmpSectorTiempo.ListaDtoTareaTiempo.Add tmpTareaTiempo, CStr(tmpTareaTiempo.Tarea.Id)
            End If

            tmpTareaTiempo.Tiempo = tmpTareaTiempo.Tiempo + tareaTiempo.Tiempo

        Next tareaTiempo

    Next sectorTiempo


    Set MergeEstadisticas = colOT
End Function

Private Sub AddTarea(sectoresTiempo As Collection, Tarea As clsTarea, CantidadPedida As Double, OperariosCotizado As Long, TiempoCotizado As Double)
    Dim tiempoSectorDTO As DTOSectoresTiempo
    Dim tareaTiempoDTO As DTOTareaTiempo
    Dim Tiempo As Double

    If Tarea.CantPorProc = 1 Then
        Tiempo = (OperariosCotizado * TiempoCotizado * CantidadPedida) / 60
    Else
        Tiempo = (OperariosCotizado * TiempoCotizado) / 60
    End If

    If BuscarEnColeccion(sectoresTiempo, CStr(Tarea.SectorID)) Then
        Set tiempoSectorDTO = sectoresTiempo.item(CStr(Tarea.SectorID))
    Else
        Set tiempoSectorDTO = New DTOSectoresTiempo
        Set tiempoSectorDTO.Sector = DAOSectores.GetById(Tarea.SectorID)
        sectoresTiempo.Add tiempoSectorDTO, CStr(Tarea.Sector.Id)
    End If


    If BuscarEnColeccion(tiempoSectorDTO.ListaDtoTareaTiempo, CStr(Tarea.Id)) Then
        Set tareaTiempoDTO = tiempoSectorDTO.ListaDtoTareaTiempo.item(CStr(Tarea.Id))
    Else
        Set tareaTiempoDTO = New DTOTareaTiempo
        Set tareaTiempoDTO.Tarea = Tarea
        tiempoSectorDTO.ListaDtoTareaTiempo.Add tareaTiempoDTO, CStr(Tarea.Id)
    End If
    tareaTiempoDTO.Tiempo = tareaTiempoDTO.Tiempo + Tiempo

End Sub

Private Sub Command5_Click()
    llenar_Grilla
End Sub
Private Sub desactiva_Click()
    If MsgBox("¿Está seguro de desactivar este presupuesto?", vbYesNo, "Confirmación") = vbYes Then    '
        A = grilla.RowIndex(grilla.row)
        If Not DAOPresupuestos.desactivar(rectmp) Then
            MsgBox "Se produjo algún error al intentar desactivar el presupuesto!", vbCritical, "Error"
        Else
            grilla.RefreshRowIndex A
        End If
    End If
End Sub
Private Sub detalle_Click()
    Dim frmver As frmVentasPresupuestoDetalle
    If Permisos.VentasCotizConsultas Then
        Set frmver = New frmVentasPresupuestoDetalle
        frmver.presupuesto = rectmp
        frmver.Show
    Else
        sinAcceso
    End If
End Sub
Private Sub EditPres_Click()
    Dim frmNuevo As New frmVentasPresupuestoEditar
    If Permisos.VentasCotizControl Then
        Set frmNuevo = New frmVentasPresupuestoEditar
        frmNuevo.nroPresu = rectmp.Id
        frmNuevo.Show
    Else
        sinAcceso
    End If
End Sub
Private Sub enviar_Click()

    On Error GoTo A
    Dim T As Long
    T = grilla.RowIndex(grilla.row)
    If DAOPresupuestos.enviar(rectmp) Then
        MsgBox "Presupuesto Exportado Correctamente!", vbInformation, "Información"
        grilla.RefreshRowIndex T
    Else
        MsgBox "No se envio el presupuesto", vbCritical, "Error"
    End If

    Exit Sub
    Set baseV = Nothing
A:
    MsgBox Err.Description
End Sub
Private Sub mostrarDetalles()
    If Me.chkVerDetalle.value = 1 Then
        Me.grilla.Gridlines = jgexGLHorizontal
        grilla.PreviewRowLines = 3
    Else
        grilla.Gridlines = jgexGLBoth
        grilla.PreviewRowLines = 0

    End If
End Sub

Private Sub Form_Activate()
    Me.grilla.Refresh
End Sub

Private Sub Form_Load()
    FormHelper.Customize Me
    mostrarDetalles
    GridEXHelper.CustomizeGrid Me.grilla, True
    suscriber_id = funciones.CreateGUID
    DAOCliente.llenarComboXtremeSuite Me.cboClientes, True, True, False
    Me.cboClientes.ListIndex = -1
    llenar_Grilla

    llenar_lista_estados
    Channel.AgregarSuscriptor Me, Presupuestos_
    
        Me.caption = caption & " (" & Name & ")"
        
        
End Sub
Private Sub Form_Terminate()
    Channel.RemoverSuscripcionTotal Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Channel.RemoverSuscripcionTotal Me
End Sub

Private Sub grilla_BeforePrintPage(ByVal PageNumber As Long, ByVal nPages As Long)
    grilla.PrinterProperties.FooterString(jgexHFRight) = "Página" & PageNumber & " de " & nPages
End Sub
Private Sub grilla_ColumnHeaderClick(ByVal Column As GridEX20.JSColumn)
    GridEXHelper.ColumnHeaderClick grilla, Column
End Sub
Private Sub grilla_DblClick()
    If grilla.RowIndex(grilla.row) = 0 Then Exit Sub
    Dim frmver As frmVentasPresupuestoDetalle
    Dim FRMEDIT As frmVentasPresupuestoEditar
    If rectmp.EstadoPresupuesto = ACotizar_ Then
        If Permisos.VentasCotizControl Then
            Set FRMEDIT = New frmVentasPresupuestoEditar
            FRMEDIT.nroPresu = rectmp.Id
            FRMEDIT.Show
        Else
            sinAcceso
        End If
    Else
        If Permisos.VentasCotizConsultas Then
            Set frmver = New frmVentasPresupuestoDetalle
            frmver.presupuesto = rectmp
            frmver.Show
        Else
            sinAcceso
        End If

    End If
End Sub

Private Sub grilla_FetchIcon(ByVal RowIndex As Long, ByVal ColIndex As Integer, ByVal RowBookmark As Variant, ByVal IconIndex As GridEX20.JSRetInteger)
    Set rec2 = presupuestos(RowIndex)
    If rec2.EstadoPresupuesto = Enviado_ Then
        IconIndex = 1
    ElseIf rec2.EstadoPresupuesto = ACotizar_ Then
        IconIndex = 2
    ElseIf rec2.EstadoPresupuesto = EstadoPresupuesto.Pendiente_ Then
        IconIndex = 3
    ElseIf rec2.EstadoPresupuesto = EstadoPresupuesto.Procesado_ Then
        IconIndex = 4
    ElseIf rec2.EstadoPresupuesto = NoCotizado Then
        IconIndex = 5
    ElseIf rec2.EstadoPresupuesto = Desactivado Then
        IconIndex = 6
    End If
End Sub

Private Sub grilla_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim est As EstadoPresupuesto
    Dim gr
    gr = grilla.RowIndex(grilla.row)
    If gr = 0 Or presupuestos.count = 0 Then Exit Sub
    Set rectmp = presupuestos.item(gr)
    est = rectmp.EstadoPresupuesto
    If Button = 2 Then
        Me.numero.caption = "[ Nro. " & rectmp.Id & " ]"
        If est = 6 Then
            Me.ncotizar.Enabled = True
            Me.EditPres.Enabled = True
            If Permisos.ventasCotizAprobaciones = False Then
                Me.AprobarPresu.Enabled = False
            Else
                Me.AprobarPresu.Enabled = True
            End If
            Me.recotiza.Enabled = True

        Else
            Me.ncotizar.Enabled = False
            Me.AprobarPresu.Enabled = False
            Me.EditPres.Enabled = False
            Me.recotiza.Enabled = True
        End If
        If est = 1 Then
            Me.enviar.Enabled = True
        Else
            Me.enviar.Enabled = False
        End If

        Me.mnuCrearOT.Enabled = (est = Enviado_)

        If Not Permisos.SistemaArchivosVer Then
            Me.archivos = False
        End If

        Me.mnuCrearOT.Enabled = (Permisos.PlanOTcontrol)

        Me.PopupMenu Me.m4
    End If
End Sub

Private Sub grilla_RowFormat(RowBuffer As GridEX20.JSRowData)
    If RowBuffer.RowIndex = 0 Then Exit Sub
    Set rectmp = presupuestos.item(RowBuffer.RowIndex)
    If RowBuffer.value(8) < Now Then
        RowBuffer.CellStyle(8) = "Vencidos"
    ElseIf RowBuffer.value(8) = Format(Now, "dd/mm/yyyy") Then
        RowBuffer.CellStyle(8) = "VenceHoy"
    End If
    If RowBuffer.value(9) > 0 Then
        RowBuffer.CellStyle(9) = "HayArchivosIncidencias"
    End If
    If RowBuffer.value(10) > 0 Then
        RowBuffer.CellStyle(10) = "HayArchivosIncidencias"
    End If
    RowBuffer.CellStyle(1) = "codigo"
    If rectmp.EstadoPresupuesto = Desactivado Then
        RowBuffer.RowStyle = "Desactivado"
    End If
End Sub
Private Sub grilla_SelectionChange()
    srow = Me.grilla.RowIndex(Me.grilla.row)
    If srow > 0 Then
        Set rectmp = presupuestos.item(srow)
    End If
End Sub
Private Sub historic_Click()
    DAOPresupuestoHistorial.getAllByPresu rectmp.Id, True
End Sub
Private Property Get ISuscriber_id() As String
    ISuscriber_id = suscriber_id
End Property

Public Function ISuscriber_Notificarse(EVENTO As clsEventoObserver) As Variant
    Dim i As Long
    Dim tmp As clsPresupuesto
    If EVENTO.EVENTO = agregar_ Then
        presupuestos.Add EVENTO.Elemento
        llenar_Grilla
    ElseIf EVENTO.EVENTO = modificar_ Then
        Set tmp = EVENTO.Elemento

        For i = presupuestos.count To 1 Step -1
            If presupuestos(i).Id = tmp.Id Then
                '                Set rectmp = presupuestos(i)
                '                rectmp.Id = tmp.Id
                '                rectmp.detalle = tmp.detalle
                '                rectmp.EstadoPresupuesto = tmp.EstadoPresupuesto
                '                Set rectmp.Cliente = tmp.Cliente
                '                rectmp.FechaEntrega = tmp.FechaEntrega
                '                Set rectmp.Moneda = tmp.Moneda
                '                Set rectmp.UsuarioModificado = tmp.UsuarioModificado
                '                rectmp.VencimientoPresupuesto = tmp.VencimientoPresupuesto
                '                rectmp.FechaModificado = tmp.FechaModificado

                presupuestos.remove i

                If presupuestos.count > 0 Then
                    If i = 1 Then
                        presupuestos.Add tmp, CStr(tmp.Id), 1
                    ElseIf (i - 1) = presupuestos.count Then
                        presupuestos.Add tmp, CStr(tmp.Id), , i - 1
                    Else
                        presupuestos.Add tmp, CStr(tmp.Id), i
                    End If
                Else
                    presupuestos.Add tmp, CStr(tmp.Id)
                End If

                grilla.RefreshRowIndex i
                Exit For
            End If
        Next

    End If
End Function

Private Sub mnuCrearOT_Click()
    If MsgBox("¿Está seguro de crear una OT desde este PPTO?", vbYesNo) = vbNo Then Exit Sub


    If IsSomething(rectmp) Then
        If rectmp.EstadoPresupuesto = Enviado_ Then
            Dim Ot As OrdenTrabajo
            Set rectmp.DetallePresupuesto = DAOPresupuestosDetalle.GetAllByPresupuesto(rectmp)
            Set Ot = DAOPresupuestos.CrearOT(rectmp, -1, rectmp.detalle)
            If IsSomething(Ot) Then
                Dim v As New clsEventoObserver
                Set v.Elemento = Ot
                v.EVENTO = agregar_
                Set v.Originador = Me
                v.Tipo = ordenesTrabajo

                Channel.Notificar v, TipoSuscripcion.ordenesTrabajo
                MsgBox "OT creada con exito. Número " & Ot.IdFormateado, vbInformation
            Else
                MsgBox "Ocurrió un error al crear la OT.", vbCritical
            End If
        End If
    End If
End Sub

Private Sub ncotizar_Click()

    If MsgBox("¿Está seguro de no cotizar este presupuesto?", vbYesNo, "Confirmación") = vbYes Then
        A = grilla.RowIndex(grilla.row)
        DAOPresupuestos.NoCotizar rectmp
        grilla.RefreshRowIndex A
    End If

    Set baseV = Nothing
End Sub
Private Sub llenar_lista_estados()
    Dim x As ListItem
    enums.EnumEstadoPresupuesto 0, matriz
    cont = 0
    Dim F As Variant
    For Each F In matriz
        If F <> Empty Then
            Select Case F
                Case "Enviado": icono = 1
                Case "A Cotizar": icono = 2
                Case "Pendiente": icono = 3
                Case "Procesado": icono = 4
                Case "No Cotizado": icono = 5
                Case "Desactivado": icono = 6
                Case Else: icono = 0
            End Select
            Set x = Me.ListView1.ListItems.Add(, , " " & F, 1, icono)
            x.Tag = cont
        End If
        cont = cont + 1
    Next
End Sub

Private Sub recotiza_Click()

    If Permisos.VentasCotizControl Then
        DAOPresupuestos.ReCotizar rectmp
        llenar_Grilla
    Else
        sinAcceso
    End If

    Set baseV = Nothing
End Sub

Private Sub scanear_Click()
    Dim archivos As New classArchivos
    archivos.escanearDocumento 2, rectmp.Id
End Sub
Private Sub txtCodigo_GotFocus()
    foco Me.txtCodigo
End Sub
Private Sub txtFiltro_GotFocus()
    foco Me.txtFiltro
End Sub
Private Sub verIncidencias_Click()
    Dim inci As New frmVerIncidencias
    inci.referencia = rectmp.Id
    inci.Origen = 1
    inci.Show
End Sub
Private Sub llenar_Grilla()
    Set tmpIncidencias = DAOIncidencias.GetCantidadIncidenciasPorReferencia(OI_Presupuestos)
    Set tmpArchivos = DAOArchivo.GetCantidadArchivosPorReferencia(OA_Presupuestos)
    Dim NRO As Long
    Dim cliente As Long
    grilla.ItemCount = 0
    llenarEstados
    If Me.cboClientes.ListIndex >= 0 Then cliente = Me.cboClientes.ItemData(Me.cboClientes.ListIndex) Else cliente = -1
    filtro = Trim(Me.txtFiltro)
    If Trim(txtCodigo) = vbNullString Or Not IsNumeric(txtCodigo) Then
        NRO = 0
    Else
        NRO = CLng(Me.txtCodigo)
    End If
    Set presupuestos = DAOPresupuestos.GetAll(filtro, estados, cliente, NRO)
    grilla.ItemCount = presupuestos.count
    Me.caption = "Presupuestos [ Cantidad: " & presupuestos.count & " ]"
    GridEXHelper.AutoSizeColumns Me.grilla, True
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Me.grilla.Width = Me.ScaleWidth
    Me.grilla.Height = Me.ScaleHeight - 3000
    Me.GroupBox1.Width = Me.ScaleWidth - 200
    PosY = Me.grilla.Top + Me.grilla.Height
    dify = Me.Height - PosY
    posyy = (dify - (Me.cmdImprimir.Height * 2))
    Me.cmdImprimir.Top = PosY + (posyy / 2) - 50
    Me.cmdEstad.Top = PosY + (posyy / 2) - 50


End Sub

Private Sub grilla_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Set rectmp = presupuestos.item(RowIndex)
    With rectmp
        Values(1) = Format(.Id, "0000")
        Values(2) = .cliente.razon
        Values(3) = .detalle
        Values(4) = .FechaEntrega
        Values(5) = .UsuarioCreado.usuario
        Values(6) = enums.EnumEstadoPresupuesto(.EstadoPresupuesto)
        Values(7) = .fechaCreado
        Values(8) = .VencimientoPresupuesto
        Values(9) = IIf(IsEmpty(tmpArchivos(.Id)), 0, tmpArchivos(.Id))
        Values(10) = IIf(IsEmpty(tmpIncidencias(.Id)), 0, tmpIncidencias(.Id))
    End With
End Sub




