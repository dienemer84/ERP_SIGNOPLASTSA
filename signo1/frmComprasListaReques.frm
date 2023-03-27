VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmComprasRequeLista 
   BackColor       =   &H00FF8080&
   Caption         =   "Lista de Requerimientos"
   ClientHeight    =   6480
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   13890
   Icon            =   "frmComprasListaReques.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6480
   ScaleWidth      =   13890
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1650
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   13740
      _Version        =   786432
      _ExtentX        =   24236
      _ExtentY        =   2910
      _StockProps     =   79
      Caption         =   "Filtros"
      UseVisualStyle  =   -1  'True
      Begin VB.ListBox lstEstados 
         Height          =   1185
         Left            =   9585
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   17
         Top             =   285
         Width           =   2520
      End
      Begin VB.TextBox txtOTDestino 
         Height          =   285
         Left            =   1140
         TabIndex        =   2
         Top             =   1110
         Width           =   975
      End
      Begin VB.TextBox txtNumero 
         Height          =   285
         Left            =   1140
         TabIndex        =   1
         Top             =   285
         Width           =   975
      End
      Begin XtremeSuiteControls.PushButton cmdBorrarSector 
         Height          =   285
         Left            =   4305
         TabIndex        =   3
         Top             =   690
         Width           =   330
         _Version        =   786432
         _ExtentX        =   582
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "X"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cboSectores 
         Height          =   315
         Left            =   1140
         TabIndex        =   4
         Top             =   675
         Width           =   3135
         _Version        =   786432
         _ExtentX        =   5530
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.GroupBox GroupBox2 
         Height          =   1050
         Left            =   4725
         TabIndex        =   5
         Top             =   345
         Width           =   4695
         _Version        =   786432
         _ExtentX        =   8281
         _ExtentY        =   1852
         _StockProps     =   79
         Caption         =   "Fecha Entrega"
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.DateTimePicker dtpDesde 
            Height          =   315
            Left            =   825
            TabIndex        =   6
            Top             =   615
            Width           =   1470
            _Version        =   786432
            _ExtentX        =   2593
            _ExtentY        =   556
            _StockProps     =   68
            CheckBox        =   -1  'True
            Format          =   1
         End
         Begin XtremeSuiteControls.DateTimePicker dtpHasta 
            Height          =   315
            Left            =   3000
            TabIndex        =   7
            Top             =   615
            Width           =   1470
            _Version        =   786432
            _ExtentX        =   2593
            _ExtentY        =   556
            _StockProps     =   68
            CheckBox        =   -1  'True
            Format          =   1
         End
         Begin XtremeSuiteControls.ComboBox cboRangos 
            Height          =   315
            Left            =   825
            TabIndex        =   8
            Top             =   225
            Width           =   3645
            _Version        =   786432
            _ExtentX        =   6429
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Style           =   2
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.Label Label7 
            Height          =   195
            Left            =   240
            TabIndex        =   11
            Top             =   285
            Width           =   480
            _Version        =   786432
            _ExtentX        =   847
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Rango"
            BackColor       =   12632256
            AutoSize        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label5 
            Height          =   195
            Left            =   255
            TabIndex        =   10
            Top             =   660
            Width           =   465
            _Version        =   786432
            _ExtentX        =   820
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Desde"
            BackColor       =   12632256
            AutoSize        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label6 
            Height          =   195
            Left            =   2430
            TabIndex        =   9
            Top             =   675
            Width           =   420
            _Version        =   786432
            _ExtentX        =   741
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Hasta"
            BackColor       =   12632256
            AutoSize        =   -1  'True
         End
      End
      Begin XtremeSuiteControls.PushButton cmdBuscar 
         Height          =   465
         Left            =   12270
         TabIndex        =   16
         Top             =   990
         Width           =   1260
         _Version        =   786432
         _ExtentX        =   2222
         _ExtentY        =   820
         _StockProps     =   79
         Caption         =   "Buscar"
         BackColor       =   15786449
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   195
         Left            =   210
         TabIndex        =   14
         Top             =   1140
         Width           =   810
         _Version        =   786432
         _ExtentX        =   1429
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "OT Destino"
         Alignment       =   1
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblNumero 
         Height          =   195
         Left            =   465
         TabIndex        =   13
         Top             =   330
         Width           =   555
         _Version        =   786432
         _ExtentX        =   979
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Número"
         Alignment       =   1
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblSector 
         Height          =   195
         Left            =   540
         TabIndex        =   12
         Top             =   720
         Width           =   465
         _Version        =   786432
         _ExtentX        =   820
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Sector"
         Alignment       =   1
         AutoSize        =   -1  'True
      End
   End
   Begin GridEX20.GridEX grilla 
      Height          =   4590
      Left            =   60
      TabIndex        =   15
      Top             =   1830
      Width           =   13740
      _ExtentX        =   24236
      _ExtentY        =   8096
      Version         =   "2.0"
      HoldSortSettings=   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      MultiSelect     =   -1  'True
      MethodHoldFields=   -1  'True
      Options         =   -1
      RecordsetType   =   1
      ForeColorInfoText=   0
      BackColorInfoText=   16777215
      GroupByBoxInfoText=   "Arrastrar el encabezado de una columna para agrupar"
      AllowEdit       =   0   'False
      BorderStyle     =   3
      BackColorGBBox  =   16744576
      BackColorHeader =   16761024
      DataMode        =   99
      HeaderFontBold  =   -1  'True
      HeaderFontWeight=   700
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   9
      Column(1)       =   "frmComprasListaReques.frx":000C
      Column(2)       =   "frmComprasListaReques.frx":0180
      Column(3)       =   "frmComprasListaReques.frx":0268
      Column(4)       =   "frmComprasListaReques.frx":0390
      Column(5)       =   "frmComprasListaReques.frx":0460
      Column(6)       =   "frmComprasListaReques.frx":0540
      Column(7)       =   "frmComprasListaReques.frx":0610
      Column(8)       =   "frmComprasListaReques.frx":070C
      Column(9)       =   "frmComprasListaReques.frx":080C
      GroupCount      =   1
      Group(1)        =   "frmComprasListaReques.frx":0904
      SortKeysCount   =   1
      SortKey(1)      =   "frmComprasListaReques.frx":096C
      FormatStylesCount=   9
      FormatStyle(1)  =   "frmComprasListaReques.frx":09D4
      FormatStyle(2)  =   "frmComprasListaReques.frx":0B0C
      FormatStyle(3)  =   "frmComprasListaReques.frx":0BBC
      FormatStyle(4)  =   "frmComprasListaReques.frx":0C70
      FormatStyle(5)  =   "frmComprasListaReques.frx":0D24
      FormatStyle(6)  =   "frmComprasListaReques.frx":0DDC
      FormatStyle(7)  =   "frmComprasListaReques.frx":0EBC
      FormatStyle(8)  =   "frmComprasListaReques.frx":0F4C
      FormatStyle(9)  =   "frmComprasListaReques.frx":0FE0
      ImageCount      =   0
      PrinterProperties=   "frmComprasListaReques.frx":1070
   End
   Begin VB.Menu menu_1 
      Caption         =   "menu"
      Visible         =   0   'False
      Begin VB.Menu editReq 
         Caption         =   "Editar"
      End
      Begin VB.Menu finalizar 
         Caption         =   "Finalizar"
      End
      Begin VB.Menu aprobar 
         Caption         =   "Aprobar"
      End
      Begin VB.Menu procesar 
         Caption         =   "Procesar Proveedores"
      End
      Begin VB.Menu fin_proceso 
         Caption         =   "Fin Proceso"
      End
      Begin VB.Menu crearPO 
         Caption         =   "Crear PO"
      End
      Begin VB.Menu mnuAnular 
         Caption         =   "Anular"
      End
      Begin VB.Menu mnuExportarExcel 
         Caption         =   "Exportar a Excel"
      End
      Begin VB.Menu df 
         Caption         =   "-"
      End
      Begin VB.Menu verDetalle 
         Caption         =   "Ver Requerimiento"
      End
      Begin VB.Menu mnuProveedoresReq 
         Caption         =   "Ver Proveedores del Req"
      End
      Begin VB.Menu verHistorial 
         Caption         =   "Historial"
      End
   End
End
Attribute VB_Name = "frmComprasRequeLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements ISuscriber
Dim suscriber_id As String
Dim frmNuevo As frmComprasRequesNuevo
Dim rectmp As clsRequerimiento
Dim rows As Long
Dim requerimientos As Collection

Private grillaActual As GridEX

Private loaded As Boolean

Private vencidos As Dictionary
Private vencenhoy As Dictionary
Private avencer As Dictionary

Private Sub llenar_Grilla()


    Set vencidos = DAORequeMateriales.FindAllRequesVencidos
    Set vencenhoy = DAORequeMateriales.FindAllRequesVencenHoy
    Set avencer = DAORequeMateriales.FindAllRequesIdProximosAVencer

    Set requerimientos = DAORequerimiento.FindAll(GetFilter(), True, False, True, False)
    grillaActual.ItemCount = 0
    grillaActual.ItemCount = requerimientos.count
    grillaActual.ReBind
    grillaActual.Refresh
    grillaActual.RefreshGroups False
    GridEXHelper.AutoSizeColumns Me.grilla
End Sub
Private Sub aprobar_Click()
    Dim A As Long
    A = grillaActual.RowIndex(grillaActual.row)
    If A > 0 And grillaActual.rowcount > 0 Then
        Set rectmp = requerimientos.item(A)
        Set rectmp = DAORequerimiento.FindById(rectmp.Id, True, True, True, True)   'lo traigo full porque necesito los materiales
        If MsgBox("¿Está seguro de aprobar el requerimiento?", vbYesNo + vbQuestion) = vbYes Then
            If DAORequerimiento.aprobar(rectmp) Then
                llenar_Grilla
                MsgBox "Aprobación exitosa.", vbInformation
            Else
                MsgBox "Se produjo algún error." & vbNewLine & "Revise que todos los detalles del requerimiento esten en estado finalizado.", vbCritical, "Error"
            End If
        End If
    End If
End Sub
Private Sub Command1_Click()
    Unload Me
End Sub
Private Sub Command2_Click()
    Dim elegidos As Boolean
    If grillaActual.SelectedItems.count > 1 Then
        elegidos = True
    Else
        elegidos = False
    End If
    With grillaActual.PrinterProperties
        .FitColumns = True
        .RepeatHeaders = True
        .Orientation = jgexPPLandscape
        .HeaderString(jgexHFCenter) = "Listado de requerimientos"
        .FooterString(jgexHFCenter) = Now
    End With
    Load frmPrintPreview
    frmPrintPreview.Move Me.Left, Me.Top, Me.Width, Me.Height
    grillaActual.PrintPreview frmPrintPreview.GEXPreview1, elegidos
    frmPrintPreview.Show 1
End Sub


Private Sub cboRangos_Click()
    funciones.CalculateDateRange Me.cboRangos, Me.dtpDesde, Me.dtpHasta

End Sub





Private Sub cmdBorrarSector_Click()
    Me.cboSectores.ListIndex = -1
End Sub

Private Sub cmdBuscar_Click()
    llenar_Grilla
End Sub

Private Sub crearPO_Click()
    Dim A As Long
    A = grillaActual.RowIndex(grillaActual.row)
    If A > 0 And grillaActual.rowcount > 0 Then
        Set rectmp = requerimientos.item(A)
        Set rectmp = DAORequerimiento.FindById(rectmp.Id, True, True, True, True)     'lo traigo full porque necesito los materiales
        If MsgBox("¿Está seguro de crear peticiones de oferta para este Requerimiento?", vbYesNo + vbQuestion) = vbYes Then
            If DAOPeticionOferta.Nueva(rectmp) Then
                llenar_Grilla
                MsgBox "Peticiones de Oferta creadas con éxito!", vbInformation, "Información"
            End If
        End If
    End If


End Sub







Private Sub editReq_Click()
    Dim A As Long
    A = grillaActual.RowIndex(grillaActual.row)
    If A > 0 And grillaActual.rowcount > 0 Then
        Set frmNuevo = New frmComprasRequesNuevo
        Set rectmp = requerimientos.item(A)

        frmNuevo.Requerimiento = rectmp.Id
        frmNuevo.Show
    End If
End Sub

Private Sub fin_proceso_Click()
    Dim A As Long
    A = grillaActual.RowIndex(grillaActual.row)
    If A > 0 And grillaActual.rowcount > 0 Then
        Set rectmp = requerimientos.item(A)
        Set rectmp = DAORequerimiento.FindById(rectmp.Id, True, True, True, True)    'lo traigo full porque necesito los materiales
        If MsgBox("¿Está seguro terminar el proceso?", vbYesNo + vbQuestion) = vbYes Then
            If DAORequerimiento.FinProceso(rectmp) Then
                llenar_Grilla
                MsgBox "Finalización exitosa!", vbInformation, "Información"
            Else
                MsgBox "Se produjo algún error, no se realizarán cambios!", vbCritical, "Error"
            End If
        End If
    End If
End Sub

Private Sub finalizar_Click()
    Dim A As Long
    A = grillaActual.RowIndex(grillaActual.row)
    If A > 0 And grillaActual.rowcount > 0 Then
        Set rectmp = requerimientos.item(A)
        Set rectmp = DAORequerimiento.FindById(rectmp.Id, True, True, True, True)    'lo traigo full porque necesito los materiales

        If rectmp.ValidarEntregas Then
            If MsgBox("¿Está seguro de finalizar el requerimiento?", vbYesNo + vbQuestion) = vbYes Then
                If DAORequerimiento.finalizar(rectmp) Then
                    llenar_Grilla
                    MsgBox "Finalización exitosa.", vbInformation + vbOKOnly
                Else
                    MsgBox "Se produjo algún error, no se realizarán cambios!", vbCritical, "Error"
                End If
            End If
        Else
            MsgBox "Debe reveer las entregas, no puede finalizar el requerimiento en estas condiciones.!", vbInformation, "Información"
        End If
    End If
End Sub



Private Sub Form_Load()
    FormHelper.Customize Me

    GridEXHelper.CustomizeGrid Me.grilla, True, False

    suscriber_id = funciones.CreateGUID
    Channel.AgregarSuscriptor Me, RequerimientosCompra_

    Dim i As Long

    For i = 0 To UBound(enums.estados_Reques) - 1
        Me.lstEstados.AddItem enums.enumEstadoRequeCompra(i)
        Me.lstEstados.ItemData(Me.lstEstados.NewIndex) = i
    Next i

    For i = 0 To Me.lstEstados.ListCount - 1
        Me.lstEstados.Selected(i) = True
    Next i
    Me.lstEstados.ListIndex = 0

    DAOSectores.LlenarComboXtreme Me.cboSectores
    'Me.cboSectores.ListIndex = -1


    funciones.FillComboBoxDateRanges Me.cboRangos
    For i = 0 To Me.cboRangos.ListCount - 1
        If Me.cboRangos.ItemData(i) = DateRangeValue.DRV_YearCurrent Then Exit For
    Next i
    Me.cboRangos.ListIndex = i

    Set grillaActual = Me.grilla
    llenar_Grilla


    loaded = True
End Sub
Private Sub Form_Resize()
    On Error Resume Next
    Dim widthOffset As Long: widthOffset = 150
    Dim heightOffset As Long: heightOffset = 1900


    Me.grilla.Height = Me.ScaleHeight - heightOffset
    Me.grilla.Width = Me.ScaleWidth - widthOffset
    Me.grilla.ColumnAutoResize = True


End Sub
Private Sub Form_Unload(Cancel As Integer)
    Channel.RemoverSuscripcionTotal Me
End Sub




Private Sub grilla_ColumnHeaderClick(ByVal Column As GridEX20.JSColumn)
    funciones.ordenar_grilla Column, Me.grilla
End Sub

Private Sub grilla_DblClick()
    verDetalle_Click
End Sub

Private Sub grilla_GroupByBoxHeaderClick(ByVal Group As GridEX20.JSGroup)
    GridEXHelper.GroupByBoxHeaderClick Group
End Sub

Private Sub grilla_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then MouseUp Me.grilla
End Sub

Private Sub MouseUp(gri As GridEX)

    Dim r As Recordset
    Dim est As EstadoRequeCompra
    Dim gr As Long
    Dim idr As Long
    gr = gri.RowIndex(gri.row)
    If gr = 0 Then Exit Sub
    Set rectmp = requerimientos.item(gr)
    idr = rectmp.Id
    est = rectmp.estado

    Me.editReq.Enabled = (est = EnEdición_ And Permisos.ComprasRequesControl) Or (est = EstadoRequeCompra.Finalizado_ And Permisos.ComprasRequesAprobaciones)
    Me.aprobar = (est = EstadoRequeCompra.Finalizado_) And Permisos.ComprasRequesAprobaciones
    Me.procesar.Enabled = (est = Aprobado_ Or est = EnProceso_) And Permisos.ComprasRequesProcesar
    Me.crearPO.Enabled = (est = EstadoRequeCompra.Procesado_) Or (est = ProcesadoParcial_) And Permisos.ComprasPOCrear
    Me.finalizar.Enabled = (est = EnEdición_) And Permisos.ComprasRequesControl
    Me.fin_proceso.Enabled = (est = EnProceso_) And Permisos.ComprasRequesProcesar
    Me.mnuAnular.Enabled = (est <> Anulado) And Permisos.ComprasRequesAnular

    Me.PopupMenu menu_1
End Sub


Private Sub grilla_RowFormat(RowBuffer As GridEX20.JSRowData)
    On Error GoTo E
    Set rectmp = requerimientos(RowBuffer.RowIndex)

    If vencidos.Exists(CStr(rectmp.Id)) Then RowBuffer.CellStyle(7) = "vencidos"
    If vencenhoy.Exists(CStr(rectmp.Id)) Then RowBuffer.CellStyle(8) = "vencenhoy"
    If avencer.Exists(CStr(rectmp.Id)) Then RowBuffer.CellStyle(9) = "avencer"

E:
End Sub

Private Sub grilla_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    hydrate RowIndex, Values
End Sub

Private Sub hydrate(RowIndex As Long, Values As GridEX20.JSRowData)
    Set rectmp = requerimientos.item(RowIndex)
    Dim ote As String
    With rectmp
        'If .tipo = stock_ Then ote = vbNullString Else ote = " " & .DestinoOT
        Values(1) = .Id
        Values(2) = .Sector.Sector
        Values(3) = .fechaCreado
        Values(4) = .StringDestino    ' enums.enumDestino(.tipo) & ote
        Values(5) = .Usuario_creador.usuario
        Values(6) = enums.enumEstadoRequeCompra(.estado)
    End With
End Sub



Private Property Get ISuscriber_id() As String
    ISuscriber_id = suscriber_id
End Property

Private Function ISuscriber_Notificarse(EVENTO As clsEventoObserver) As Variant
    Dim tmp As clsRequerimiento
    If EVENTO.EVENTO = agregar_ Then
        requerimientos.Add EVENTO.Elemento, CStr(EVENTO.Elemento.Id)
        grillaActual.ItemCount = requerimientos.count
    ElseIf EVENTO.EVENTO = modificar_ Then
        '        Set tmp = evento.Elemento
        '        'requerimientos.Remove CStr(tmp.Id)
        '        'requerimientos.Add tmp, CStr(tmp.Id)
        '
        '        Dim i As Long
        '        For i = 1 To requerimientos.count
        '            If requerimientos.Item(i).Id = tmp.Id Then
        '                'Set requerimientos.Item(i) = tmp
        '                grilla.RefreshRowIndex i
        '                Exit Function
        '            End If
        '        Next i


        llenar_Grilla
    End If

End Function



Private Function GetFilter() As String

    Dim filtro As String: filtro = " 1 = 1 "

    Dim i As Long
    Dim estados() As Variant
    ReDim estados(0)

    For i = 0 To Me.lstEstados.ListCount - 1
        If Me.lstEstados.Selected(i) Then
            ReDim Preserve estados(UBound(estados, 1) + 1)
            estados(UBound(estados, 1)) = Me.lstEstados.ItemData(i)
        End If
    Next i

    If UBound(estados, 1) > 0 Then
        filtro = filtro & " and estado in ("
        estados(0) = estados(1)    'le paso el estado del 1ro para que no quede vacio
        filtro = filtro & Join(estados, ", ")
        filtro = filtro & ")"
    End If

    If LenB(Me.txtNumero.text) > 0 Then
        filtro = filtro & " AND id = " & Val(Me.txtNumero.text)
    End If

    If Me.cboSectores.ListIndex <> -1 Then
        filtro = filtro & " AND idSector = " & Me.cboSectores.ItemData(Me.cboSectores.ListIndex)
    End If

    If LenB(Me.txtOTDestino.text) > 0 Then
        filtro = filtro & " AND idPedido = " & Val(Me.txtOTDestino.text)
    End If

    If Not IsNull(Me.dtpDesde.value) Then
        filtro = filtro & " AND FechaCreado >= " & conectar.Escape(Me.dtpDesde.value)
    End If

    If Not IsNull(Me.dtpHasta.value) Then
        filtro = filtro & " AND FechaCreado <= " & conectar.Escape(Me.dtpHasta.value)
    End If


    GetFilter = filtro
End Function

Private Sub mnuAnular_Click()
    Dim A As Long
    A = grillaActual.RowIndex(grillaActual.row)
    If A > 0 And grillaActual.rowcount > 0 Then

        If MsgBox("¿Está seguro de anular el requerimiento?", vbYesNo + vbQuestion) = vbYes Then
            Set rectmp = requerimientos.item(A)
            Set rectmp = DAORequerimiento.FindById(rectmp.Id, True, True, True, True)

            If DAORequerimiento.Anular(rectmp) Then
                llenar_Grilla
                MsgBox "El requerimiento ha sido anulado.", vbOKOnly + vbInformation
            Else
                MsgBox "No se pudo anular el requerimiento.", vbCritical + vbOKOnly
            End If

        End If
    End If
End Sub

Private Sub mnuExportarExcel_Click()
    Dim A As Long
    A = grillaActual.RowIndex(grillaActual.row)
    If A > 0 And grillaActual.rowcount > 0 Then
        Set rectmp = requerimientos.item(A)
        DAORequerimiento.ExportExcel rectmp.Id, (MsgBox("¿Desea exportar con información acerca de la OC y las PO?", vbYesNo + vbQuestion) = vbYes)
    End If

End Sub

Private Sub mnuProveedoresReq_Click()
    Dim A As Long
    A = grillaActual.RowIndex(grillaActual.row)
    If A > 0 And grillaActual.rowcount > 0 Then
        Set rectmp = requerimientos.item(A)
        Set rectmp = DAORequerimiento.FindById(rectmp.Id, True, True, True, True)    'lo traigo full porque necesito los materiales

        Dim F As New frmComprasRequesProcesar
        F.ReadOnly = True
        F.reque = rectmp
        F.Show
    End If
End Sub

Private Sub procesar_Click()
    Dim A As Long
    A = grillaActual.RowIndex(grillaActual.row)
    Dim idReque As Long
    If A > 0 And grillaActual.rowcount > 0 Then
        Set rectmp = requerimientos.item(A)
        idReque = rectmp.Id
        Set rectmp = DAORequerimiento.FindById(rectmp.Id, True, True, True, True)  'lo traigo full porque necesito los materiales
        If rectmp.estado = Aprobado_ Then
            If MsgBox("¿Está seguro de procesar el requerimiento seleccionado?", vbYesNo + vbQuestion) = vbYes Then
                If DAORequerimiento.procesar(rectmp) Then
                    llenar_Grilla
                    MsgBox "El requerimiento esta listo para procesar sus proveedores.", vbInformation + vbOKOnly
                    Dim F As New frmComprasRequesProcesar
                    F.reque = DAORequerimiento.FindById(idReque, True, True, True, True)  'lo traigo full porque necesito los materiales
                    F.Show
                Else
                    MsgBox "Se produjo algún error al procesar el requerimiento!. Revise que todos los items esten en estado aprobado.", vbCritical, "Error"
                End If
            End If
        ElseIf rectmp.estado = EnProceso_ Then
            frmComprasRequesProcesar.reque = rectmp
            frmComprasRequesProcesar.Show
        End If
    End If
End Sub







Private Sub verDetalle_Click()
    Dim A As Long
    A = grillaActual.RowIndex(grillaActual.row)
    If A > 0 And grillaActual.rowcount > 0 Then
        Set frmNuevo = New frmComprasRequesNuevo
        Set rectmp = requerimientos.item(A)

        frmNuevo.SoloVer = True
        frmNuevo.Requerimiento = rectmp.Id
        frmNuevo.Show
    End If

End Sub

Private Sub verHistorial_Click()
    Dim req As clsRequerimiento
    Dim A As Long
    A = grillaActual.RowIndex(grillaActual.row)
    If A > 0 And grillaActual.rowcount > 0 Then
        Set req = requerimientos.item(A)
        frmHistoriales.lista = DAORequeHistorial.getAllByIdReque(req.Id)
        frmHistoriales.Show
    End If

End Sub
