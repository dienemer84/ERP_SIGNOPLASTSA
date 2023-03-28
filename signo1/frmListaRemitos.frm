VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmPlaneamientoRemitosLista 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Remitos"
   ClientHeight    =   8805
   ClientLeft      =   1455
   ClientTop       =   2130
   ClientWidth     =   13500
   Icon            =   "frmListaRemitos.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8805
   ScaleWidth      =   13500
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1545
      Left            =   120
      TabIndex        =   1
      Top             =   90
      Width           =   12525
      _Version        =   786432
      _ExtentX        =   22093
      _ExtentY        =   2725
      _StockProps     =   79
      Caption         =   "Parámetros de búsqueda"
      UseVisualStyle  =   -1  'True
      Begin VB.TextBox txtDescripcion 
         Height          =   285
         Left            =   1425
         TabIndex        =   3
         Top             =   645
         Width           =   4215
      End
      Begin VB.TextBox txtNroRemito 
         Height          =   285
         Left            =   1425
         TabIndex        =   2
         Top             =   240
         Width           =   1695
      End
      Begin XtremeSuiteControls.ComboBox cboClientes 
         Height          =   315
         Left            =   4530
         TabIndex        =   7
         Top             =   240
         Width           =   6015
         _Version        =   786432
         _ExtentX        =   10610
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   285
         Left            =   10620
         TabIndex        =   8
         Top             =   270
         Width           =   495
         _Version        =   786432
         _ExtentX        =   873
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "X"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton cmdBuscar 
         Default         =   -1  'True
         Height          =   375
         Left            =   11280
         TabIndex        =   9
         Top             =   720
         Width           =   1095
         _Version        =   786432
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Buscar"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.DateTimePicker dtpDesde 
         Height          =   315
         Left            =   5730
         TabIndex        =   10
         Top             =   1065
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
         Left            =   7905
         TabIndex        =   11
         Top             =   1065
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
         Left            =   1380
         TabIndex        =   12
         Top             =   1050
         Width           =   3645
         _Version        =   786432
         _ExtentX        =   6429
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.PushButton cmdImprimir 
         Height          =   375
         Left            =   11280
         TabIndex        =   16
         Top             =   240
         Width           =   1080
         _Version        =   786432
         _ExtentX        =   1905
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Imprimir"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cboEstadoFacturado 
         Height          =   315
         Left            =   6600
         TabIndex        =   17
         Top             =   645
         Width           =   1590
         _Version        =   786432
         _ExtentX        =   2805
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.PushButton cmdLimpiaEstado 
         Height          =   255
         Left            =   8280
         TabIndex        =   19
         Top             =   690
         Width           =   300
         _Version        =   786432
         _ExtentX        =   529
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "X"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label4 
         Height          =   195
         Left            =   5760
         TabIndex        =   18
         Top             =   710
         Width           =   720
         _Version        =   786432
         _ExtentX        =   1270
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Facturado"
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label7 
         Height          =   195
         Left            =   795
         TabIndex        =   15
         Top             =   1110
         Width           =   480
         _Version        =   786432
         _ExtentX        =   847
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Rango"
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label5 
         Height          =   195
         Left            =   5160
         TabIndex        =   14
         Top             =   1110
         Width           =   465
         _Version        =   786432
         _ExtentX        =   820
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Desde"
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label6 
         Height          =   195
         Left            =   7335
         TabIndex        =   13
         Top             =   1125
         Width           =   420
         _Version        =   786432
         _ExtentX        =   741
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Hasta"
         AutoSize        =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         Height          =   255
         Left            =   3315
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   710
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Número"
         Height          =   255
         Left            =   270
         TabIndex        =   4
         Top             =   270
         Width           =   1095
      End
   End
   Begin GridEX20.GridEX grilla 
      Height          =   6615
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   11668
      Version         =   "2.0"
      PreviewRowIndent=   200
      DefaultGroupMode=   1
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      PreviewColumn   =   6
      PreviewRowLines =   1
      ColumnAutoResize=   -1  'True
      MethodHoldFields=   -1  'True
      DatabaseName    =   " "
      ForeColorInfoText=   0
      AllowEdit       =   0   'False
      ImageCount      =   1
      ImagePicture1   =   "frmListaRemitos.frx":000C
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp7        =   0
      ColumnsCount    =   10
      Column(1)       =   "frmListaRemitos.frx":0326
      Column(2)       =   "frmListaRemitos.frx":044E
      Column(3)       =   "frmListaRemitos.frx":051E
      Column(4)       =   "frmListaRemitos.frx":0636
      Column(5)       =   "frmListaRemitos.frx":0706
      Column(6)       =   "frmListaRemitos.frx":07D6
      Column(7)       =   "frmListaRemitos.frx":08C2
      Column(8)       =   "frmListaRemitos.frx":0996
      Column(9)       =   "frmListaRemitos.frx":0A8A
      Column(10)      =   "frmListaRemitos.frx":0B7E
      FormatStylesCount=   14
      FormatStyle(1)  =   "frmListaRemitos.frx":0C6A
      FormatStyle(2)  =   "frmListaRemitos.frx":0DA2
      FormatStyle(3)  =   "frmListaRemitos.frx":0E52
      FormatStyle(4)  =   "frmListaRemitos.frx":0F06
      FormatStyle(5)  =   "frmListaRemitos.frx":0FDE
      FormatStyle(6)  =   "frmListaRemitos.frx":1096
      FormatStyle(7)  =   "frmListaRemitos.frx":1176
      FormatStyle(8)  =   "frmListaRemitos.frx":1236
      FormatStyle(9)  =   "frmListaRemitos.frx":1316
      FormatStyle(10) =   "frmListaRemitos.frx":1416
      FormatStyle(11) =   "frmListaRemitos.frx":14F2
      FormatStyle(12) =   "frmListaRemitos.frx":15C2
      FormatStyle(13) =   "frmListaRemitos.frx":1696
      FormatStyle(14) =   "frmListaRemitos.frx":174E
      ImageCount      =   1
      ImagePicture(1) =   "frmListaRemitos.frx":17DA
      PrinterProperties=   "frmListaRemitos.frx":1AF4
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   360
      Top             =   7680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Menu mnuRtos 
      Caption         =   "mnuRtos"
      Visible         =   0   'False
      Begin VB.Menu numero 
         Caption         =   "Numero"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuEditar 
         Caption         =   "Editar..."
      End
      Begin VB.Menu endRto 
         Caption         =   "Aprobar..."
      End
      Begin VB.Menu AnularRto 
         Caption         =   "Anular..."
      End
      Begin VB.Menu mnuNoFacturable 
         Caption         =   "No Facturable..."
      End
      Begin VB.Menu mnuValorizar 
         Caption         =   "Valorizar..."
      End
      Begin VB.Menu mnuFacturarRemito 
         Caption         =   "Facturar..."
      End
      Begin VB.Menu adan 
         Caption         =   "-"
      End
      Begin VB.Menu printRto 
         Caption         =   "Imprimir..."
      End
      Begin VB.Menu verRto 
         Caption         =   "Ver detalle..."
      End
      Begin VB.Menu archivos 
         Caption         =   "Archivos Asociados..."
      End
      Begin VB.Menu scanear 
         Caption         =   "Adquirir..."
      End
      Begin VB.Menu mnuHistorico 
         Caption         =   "Historial..."
      End
      Begin VB.Menu nada22 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDefinirBultos 
         Caption         =   "Definir Bultos..."
      End
      Begin VB.Menu mnuPrintBultos 
         Caption         =   "Imprimir Lista Bultos..."
      End
      Begin VB.Menu mnuControlCarga 
         Caption         =   "Imprimir Control Carga..."
      End
   End
End
Attribute VB_Name = "frmPlaneamientoRemitosLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements ISuscriber
Dim remitos As Collection
Dim tmpRto As Remito
Dim filtro As String
Dim id_suscriber As String
Public VerInfoAdministracion As Boolean
Public VerInfoPlaneamiento As Boolean
Dim claseP As New classPlaneamiento
Dim facturasRemitos As Dictionary
Dim m_Archivos As Dictionary
Private Sub AnularRto_Click()
    Dim IdRemito As Long
    Dim A As Long
    A = Me.grilla.RowIndex(Me.grilla.row)
    If MsgBox("¿Está seguro de anular el remito?", vbYesNo, "Confirmación") = vbYes Then
        If DAORemitoS.Anular(tmpRto) Then
            MsgBox "Remito anulado con éxito!", vbExclamation, "Información"
            Me.grilla.RefreshRowIndex A
        Else
            MsgBox "Se produjo algún error al anular el remito!", vbExclamation, "Información"
        End If
    End If

End Sub

Private Sub archivos_Click()

    Dim frmarchi1 As New frmArchivos2
    frmarchi1.Origen = OA_Remitos
    frmarchi1.ObjetoId = tmpRto.Id
    frmarchi1.caption = "Remito " & tmpRto.numero
    frmarchi1.Show


    frmarchi1.Show



End Sub

Private Sub cboRangos_Click()
    funciones.CalculateDateRange Me.cboRangos, Me.dtpDesde, Me.dtpHasta
End Sub

Private Sub cmdBuscar_Click()
    listaRemitos
End Sub
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub cmdImprimir_Click()
    Dim pro As String
    Dim q As String
    If Me.cboClientes.ListIndex > -1 Then
        pro = " Cliente: " & Me.cboClientes.text
    End If

    With Me.grilla.PrinterProperties

        .FitColumns = True
        .RepeatHeaders = True
        .Orientation = jgexPPLandscape
        .HeaderString(jgexHFCenter) = "Listado de Remitos"
        If LenB(pro) > 1 Then
            .HeaderString(jgexHFLeft) = pro
        End If

        If Not IsNull(Me.dtpDesde) Then
            q = "Desde " & Format(Me.dtpDesde, "dd-mm-yyyy") & Chr(10)
        End If
        If Not IsNull(Me.dtpHasta) Then
            q = q & "Hasta " & Format(Me.dtpHasta, "dd-mm-yyyy") & Chr(10)
        End If

        If IsNull(Me.dtpHasta) And IsNull(Me.dtpDesde) Then
            q = "PERIODO SIN ESPECIFICAR" & Chr(10)
        End If


        If LenB(q) > 1 Then
            .HeaderString(jgexHFRight) = q
        End If

        .FooterString(jgexHFCenter) = Now

    End With
    Load frmPrintPreview
    frmPrintPreview.Move Me.Left, Me.Top, Me.Width, Me.Height
    Me.grilla.PrintPreview frmPrintPreview.GEXPreview1
    frmPrintPreview.Show 1
End Sub

Private Sub cmdLimpiaEstado_Click()
    Me.cboEstadoFacturado.ListIndex = -1
End Sub

Private Sub endRto_Click()
    On Error GoTo err454
    Dim A As Long
    Dim rtoNro As Long
    rtoNro = tmpRto.Id
    A = Me.grilla.RowIndex(Me.grilla.row)
    If MsgBox("¿Desea aprobar el remito seleccionado?", vbYesNo, "Confirmación") = vbYes Then
        If DAORemitoS.aprobar(tmpRto) Then
            If MsgBox("El remito se aprobó correctamente." & Chr(10) & "¿Desea imprimirlo ahora?", vbYesNo, "Confirmación") = vbYes Then
                CD.Flags = cdlPDUseDevModeCopies
                CD.Copies = 5
                CD.ShowPrinter
                Dim i As Long
                For i = 1 To CD.Copies
                    DAORemitoS.ImprimirRemito rtoNro
                Next i
            End If
            Me.grilla.RefreshRowIndex A
        Else
            MsgBox "Se produjo un error al aprobar el remito.", vbCritical, "Error"
        End If
    End If
    Exit Sub
err454:

End Sub
Private Sub listaRemitos()
    Set m_Archivos = DAOArchivo.GetCantidadArchivosPorReferencia(OA_Remitos)

    filtro = "1 = 1"
    If LenB(Me.txtDescripcion.text) > 0 Then
        filtro = filtro & " and " & DAORemitoS.TABLA_REMITO & "." & DAORemitoS.CAMPO_DETALLE & " like '%" & Trim(Me.txtDescripcion) & "%'"
    End If
    If Me.cboClientes.ListIndex > -1 Then
        filtro = filtro & " and " & DAORemitoS.TABLA_CLIENTE & ".id=" & (Me.cboClientes.ItemData(Me.cboClientes.ListIndex))
    End If

    If Not IsEmpty(Me.txtNroRemito) And IsNumeric(txtNroRemito) Then
        filtro = filtro & "  and " & DAORemitoS.TABLA_REMITO & "." & DAORemitoS.CAMPO_NUMERO & "=" & Me.txtNroRemito
    End If

    If Not IsNull(Me.dtpDesde.value) Then
        filtro = filtro & " and  " & DAORemitoS.TABLA_REMITO & "." & DAORemitoS.CAMPO_FECHA & " >= " & conectar.Escape(Format(Me.dtpDesde.value, "yyyy-mm-dd 00:00:00"))
    End If

    If Not IsNull(Me.dtpHasta.value) Then
        filtro = filtro & " and  " & DAORemitoS.TABLA_REMITO & "." & DAORemitoS.CAMPO_FECHA & " <= " & conectar.Escape(Format(Me.dtpHasta.value, "yyyy-mm-dd 23:59:59"))
    End If


    If Me.cboEstadoFacturado.ListIndex > -1 Then
        filtro = filtro & " and " & DAORemitoS.CAMPO_ESTADO_FACTURADO & "=" & (Me.cboEstadoFacturado.ItemData(Me.cboEstadoFacturado.ListIndex))
    End If




    Set remitos = DAORemitoS.FindAll("and " & filtro)

    Dim remi As Remito
    Dim remitosId As New Collection
    For Each remi In remitos
        remitosId.Add remi.Id
    Next
    Set facturasRemitos = New Dictionary

    If remitosId.count > 0 Then
        Set facturasRemitos = DAOFactura.FindAllByRemitos(remitosId)
    End If


    Me.grilla.ItemCount = 0
    Me.grilla.ItemCount = remitos.count
    Me.grilla.Update

End Sub


Private Sub Form_Load()
    FormHelper.Customize Me

    DAOCliente.llenarComboXtremeSuite Me.cboClientes, False, True, False
    Me.cboClientes.ListIndex = -1
    GridEXHelper.CustomizeGrid Me.grilla, True
    id_suscriber = funciones.CreateGUID
    Me.grilla.Columns(8).Visible = VerInfoAdministracion
    Me.cboEstadoFacturado.Visible = VerInfoAdministracion
    Channel.AgregarSuscriptor Me, Remitos_
    Me.grilla.ItemCount = 0
    Me.grilla.Update
    Dim i As Integer
    funciones.FillComboBoxDateRanges Me.cboRangos
    For i = 0 To Me.cboRangos.ListCount - 1
        If Me.cboRangos.ItemData(i) = DateRangeValue.DRV_YearCurrent Then Exit For
    Next i
    Me.cboRangos.ListIndex = i


    Me.cboEstadoFacturado.Clear
    Me.cboEstadoFacturado.AddItem enums.EnumEstadoRemitoFacturado(EstadoRemitoFacturado.RemitoNoFacturado)
    Me.cboEstadoFacturado.ItemData(Me.cboEstadoFacturado.NewIndex) = EstadoRemitoFacturado.RemitoNoFacturado
    Me.cboEstadoFacturado.AddItem enums.EnumEstadoRemitoFacturado(EstadoRemitoFacturado.RemitoNoFacturable)

    Me.cboEstadoFacturado.ItemData(Me.cboEstadoFacturado.NewIndex) = EstadoRemitoFacturado.RemitoNoFacturable
    Me.cboEstadoFacturado.AddItem enums.EnumEstadoRemitoFacturado(EstadoRemitoFacturado.RemitoFacturadoParcial)
    Me.cboEstadoFacturado.ItemData(Me.cboEstadoFacturado.NewIndex) = EstadoRemitoFacturado.RemitoFacturadoParcial

    Me.cboEstadoFacturado.AddItem enums.EnumEstadoRemitoFacturado(EstadoRemitoFacturado.RemitoFacturadoTotal)

    Me.cboEstadoFacturado.ItemData(Me.cboEstadoFacturado.NewIndex) = EstadoRemitoFacturado.RemitoFacturadoTotal


    listaRemitos

End Sub
Private Sub Form_Resize()
    On Error Resume Next
    Me.grilla.Width = Me.ScaleWidth - 180
    Me.grilla.Height = Me.ScaleHeight - 1800
    Me.GroupBox1.Width = Me.ScaleWidth
End Sub


Private Sub Form_Terminate()
    Channel.RemoverSuscripcionTotal Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Channel.RemoverSuscripcionTotal Me
End Sub

Private Sub grilla_ColumnHeaderClick(ByVal Column As GridEX20.JSColumn)
    GridEXHelper.ColumnHeaderClick grilla, Column
End Sub

Private Sub grilla_FetchIcon(ByVal RowIndex As Long, ByVal ColIndex As Integer, ByVal RowBookmark As Variant, ByVal IconIndex As GridEX20.JSRetInteger)
    On Error Resume Next
    If ColIndex = 10 And m_Archivos.item(tmpRto.Id) > 0 Then
        IconIndex = 1
    End If
End Sub
Private Sub grilla_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim row As Long
    row = grilla.RowIndex(grilla.row)
    If row > 0 Then
        If Button = 2 Then
            grilla_SelectionChange
            Me.numero.caption = "[ Nro. " & tmpRto.numero & " ]"




            If VerInfoAdministracion And tmpRto.estado <> RemitoAnulado Then
                Me.mnuValorizar.Enabled = (tmpRto.EstadoFacturado = RemitoNoFacturado Or tmpRto.EstadoFacturado = RemitoFacturadoParcial)

                If tmpRto.EstadoFacturado = RemitoNoFacturable Or tmpRto.EstadoFacturado = RemitoNoFacturado Then
                    Me.mnuNoFacturable.Enabled = True
                    If tmpRto.EstadoFacturado = RemitoNoFacturable Then
                        Me.mnuNoFacturable.caption = "Hacer Facturable..."
                    ElseIf tmpRto.EstadoFacturado = RemitoNoFacturado Then
                        Me.mnuNoFacturable.caption = "Hacer No Facturable..."
                    End If
                Else
                    Me.mnuNoFacturable.Enabled = False
                End If


            Else
                Me.mnuValorizar = False
                Me.mnuNoFacturable.Enabled = False
            End If



            Me.archivos = Permisos.SistemaArchivosVer
            Me.AnularRto = (tmpRto.estado = EstadoRemito.RemitoAprobado)
            Me.printRto.Enabled = (tmpRto.estado = EstadoRemito.RemitoAprobado Or tmpRto.estado = RemitoAnulado)
            Me.endRto.Enabled = (Permisos.planRemitosAprobaciones And tmpRto.estado = RemitoPendiente)
            Me.mnuEditar.Enabled = (tmpRto.estado = RemitoPendiente)
            Me.mnuFacturarRemito.Enabled = (tmpRto.estado = EstadoRemito.RemitoAprobado)

            Me.PopupMenu Me.mnuRtos
        End If
    End If
End Sub
Private Sub grilla_RowFormat(RowBuffer As GridEX20.JSRowData)
    If RowBuffer.RowIndex > 0 Then
        Set tmpRto = remitos(RowBuffer.RowIndex)

        If tmpRto.estado = RemitoAnulado Then
            RowBuffer.RowStyle = "Anulado"
        Else
            If tmpRto.estado = EstadoRemito.RemitoAprobado Then
                RowBuffer.CellStyle(4) = "EstadoAprobado"
            ElseIf tmpRto.estado = RemitoPendiente Then
                RowBuffer.CellStyle(4) = "EstadoPendiente"
            End If
        End If
        If VerInfoAdministracion Then
            'solo pierdo tiempo en formatear si me piden ver los datos
            If tmpRto.EstadoFacturado = RemitoNoFacturable Then
                RowBuffer.CellStyle(8) = "NoFacturable"
            Else
                If tmpRto.EstadoFacturado = RemitoFacturadoTotal Then
                    RowBuffer.CellStyle(8) = "Total"
                ElseIf tmpRto.EstadoFacturado = RemitoFacturadoParcial Then
                    RowBuffer.CellStyle(8) = "Parcial"
                ElseIf tmpRto.EstadoFacturado = RemitoNoFacturado Then
                    RowBuffer.CellStyle(8) = "NoFacturado"
                End If

            End If
        End If
    End If
End Sub

Private Sub grilla_SelectionChange()
    Dim it As Long: it = grilla.RowIndex(grilla.row)
    If it > 0 Then
        Set tmpRto = remitos.item(it)
    Else
        Set tmpRto = Nothing
    End If
End Sub

Private Sub grilla_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    On Error Resume Next
    Set tmpRto = remitos.item(RowIndex)
    With Values
        .value(1) = tmpRto.numero
        .value(6) = tmpRto.detalle
        .value(3) = tmpRto.FEcha
        .value(4) = enums.EnumEstadoRemito(tmpRto.estado)
        .value(5) = tmpRto.usuarioCreador.usuario
        If IsSomething(tmpRto.cliente) Then .value(2) = tmpRto.cliente.razon
        If IsSomething(tmpRto.usuarioAprobador) Then
            .value(7) = tmpRto.usuarioAprobador.usuario
        Else
            .value(7) = vbNullString
        End If
        .value(8) = tmpRto.VerEstadoFacturado

        If facturasRemitos.Exists(CStr(tmpRto.numero)) Then
            If LenB(facturasRemitos.item(CStr(tmpRto.numero))) >= 0 Then
                .value(9) = Left(facturasRemitos.item(CStr(tmpRto.numero)), Len(facturasRemitos.item(CStr(tmpRto.numero))) - 2)
            End If
        End If
        Dim Cant As Long


        .value(10) = "(" & Val(m_Archivos.item(tmpRto.Id)) & ")"
    End With

End Sub


Private Property Get ISuscriber_id() As String
    ISuscriber_id = id_suscriber
End Property
Private Function ISuscriber_Notificarse(EVENTO As clsEventoObserver) As Variant
    Dim i As Long
    Dim tmp As Remito
    If EVENTO.EVENTO = agregar_ Then
        remitos.Add EVENTO.Elemento
        listaRemitos
    ElseIf EVENTO.EVENTO = modificar_ Then
        Set tmp = EVENTO.Elemento
        For i = remitos.count To 1 Step -1
            If remitos(i).Id = tmp.Id Then
                Set tmpRto = remitos(i)
                tmpRto.Id = tmp.Id
                tmpRto.detalle = tmp.detalle
                tmpRto.estado = tmp.estado
                tmpRto.EstadoFacturado = tmp.EstadoFacturado
                Set tmpRto.cliente = tmp.cliente
                Set tmpRto.usuarioCreador = tmp.usuarioCreador
                grilla.RefreshRowIndex i
                Exit For
            End If
        Next
    End If
End Function

Private Sub mnuControlCarga_Click()
    On Error GoTo err1
    If MsgBox("¿Seguro de imprimir el control de cargas de este remito?", vbYesNo, "Confirmación") = vbYes Then
        DAORemitoS.ImprimirControlCarga tmpRto
        Exit Sub
err1:
        MsgBox Err.Description
    End If
End Sub

Private Sub mnuDefinirBultos_Click()
    On Error GoTo err1
    Dim Cant As Integer





    Cant = InputBox("Ingrese cantidad de bultos", "Cantidad de bultos", tmpRto.CantidadBultos)

    If Cant <> tmpRto.CantidadBultos And Cant > 0 Then
        tmpRto.CantidadBultos = Cant
        If DAORemitoS.Save(tmpRto, False, True) Then

            MsgBox "Actualización exitosa!", vbOKOnly, "Información"
        Else
            MsgBox "Se produjo algún error! " & vbNewLine & "No se guardarán los cambios", vbOKOnly, "Error"
        End If

    End If


err1:

End Sub

Private Sub mnuEditar_Click()
    Dim frm3 As frmPlaneamientoRemitoVer
    Set frm3 = New frmPlaneamientoRemitoVer
    frm3.editar = True
    Set frm3.Remito = tmpRto
    frm3.valorizable = Me.VerInfoAdministracion
    frm3.conceptuable = (tmpRto.estado = RemitoPendiente)
    frm3.MostrarInfoAdministracion = VerInfoAdministracion
    frm3.Show
End Sub
Private Sub mnuHistorico_Click()
    Dim frm As New frmHistoriales
    frm.lista = DAORemitoHistorico.getAllByIdRemito(tmpRto.Id)
    frm.Show

End Sub

Private Sub mnuNoFacturable_Click()
    Dim A As Long
    A = grilla.RowIndex(grilla.row)
    DAORemitoS.CambiarEstadoFacturable tmpRto

    grilla.RefreshRowIndex A
End Sub

Private Sub mnuPrintBultos_Click()

    If tmpRto.CantidadBultos > 0 Then

        If MsgBox("¿Seguro de imprimir " & tmpRto.CantidadBultos & " rótulos?", vbYesNo, "Consulta") = vbYes Then
            DAORemitoS.ImprimirBultos tmpRto
        End If
    End If
End Sub

Private Sub mnuValorizar_Click()
    Dim frm2 As frmPlaneamientoRemitoVer
    Set frm2 = New frmPlaneamientoRemitoVer
    Set frm2.Remito = tmpRto
    ' frm2.editar = True
    frm2.valorizable = True And VerInfoAdministracion
    frm2.MostrarInfoAdministracion = True And VerInfoAdministracion
    frm2.Show
End Sub
Private Sub printRto_Click()
    On Error GoTo err444:
    Dim rs As Recordset
    Dim rto As Long
    rto = tmpRto.Id
    Set rs = conectar.RSFactory("select impreso from remitos where id=" & rto)
    Dim est As Long
    Dim i As Long
    If Not rs.BOF And Not rs.EOF Then est = rs!impreso Else Exit Sub
    If est > 0 Then
        If MsgBox("Este remito ya fué impreso," & Chr(10) & "¿Desea volver a imprimir?", vbYesNo, "Confirmación") = vbYes Then
            CD.Flags = cdlPDUseDevModeCopies
            CD.Copies = 5
            CD.ShowPrinter
            For i = 1 To CD.Copies
                DAORemitoS.ImprimirRemito rto
            Next i
        End If
    Else
        If MsgBox("Este remito no fue impreso." & Chr(10) & "¿Desea imprimirlo?", vbYesNo) = vbYes Then
            CD.Copies = 5
            CD.ShowPrinter
            For i = 1 To CD.Copies
                DAORemitoS.ImprimirRemito rto
            Next
        End If
    End If
    Exit Sub
err444:
End Sub
Private Sub PushButton1_Click()
    Me.cboClientes.ListIndex = -1
End Sub
Private Sub scanear_Click()
    On Error Resume Next
    grilla_SelectionChange
    Dim archivos As New classArchivos
    If archivos.escanearDocumento(OA_Remitos, tmpRto.Id) Then
        Set m_Archivos = DAOArchivo.GetCantidadArchivosPorReferencia(OA_Remitos)
        Me.grilla.RefreshRowIndex (tmpRto.Id)
    End If
End Sub
Private Sub txtDescripcion_GotFocus()
    foco Me.txtDescripcion
End Sub
Private Sub txtNroRemito_GotFocus()
    foco Me.txtNroRemito
End Sub
Private Sub txtNroRemito_Validate(Cancel As Boolean)
'    funciones.ValidarTextBox Me.txtNroRemito, Cancel
End Sub
Private Sub verRto_Click()
    Dim frm As frmPlaneamientoRemitoVer
    Set frm = New frmPlaneamientoRemitoVer
    Set frm.Remito = tmpRto
    frm.MostrarInfoAdministracion = VerInfoAdministracion
    frm.Show
End Sub
