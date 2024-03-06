VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmAdminCobranzasLista 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Recibos"
   ClientHeight    =   8445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   17550
   Icon            =   "frmAdminCobranzasLista.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8445
   ScaleWidth      =   17550
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.GroupBox GroupBox 
      Height          =   1095
      Left            =   12000
      TabIndex        =   19
      Top             =   120
      Width           =   5415
      _Version        =   786432
      _ExtentX        =   9551
      _ExtentY        =   1931
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.ProgressBar progreso 
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   600
         Width           =   5175
         _Version        =   786432
         _ExtentX        =   9128
         _ExtentY        =   661
         _StockProps     =   93
         Appearance      =   6
      End
      Begin XtremeSuiteControls.Label lblTotalRecibido 
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   22
         Top             =   240
         Width           =   1575
         _Version        =   786432
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "$ 00,00"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label Label 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   1095
         _Version        =   786432
         _ExtentX        =   1931
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Total Recibido:"
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Parámetros de Búsqueda"
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11775
      Begin XtremeSuiteControls.PushButton btnLimpiarCliente 
         Height          =   255
         Left            =   6240
         TabIndex        =   18
         Top             =   385
         Width           =   420
         _Version        =   786432
         _ExtentX        =   741
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "X"
         UseVisualStyle  =   -1  'True
      End
      Begin VB.TextBox txtNroRecibo 
         Height          =   285
         Left            =   1320
         TabIndex        =   9
         Top             =   795
         Width           =   1500
      End
      Begin XtremeSuiteControls.ComboBox cboCliente 
         Height          =   315
         Left            =   1320
         TabIndex        =   2
         Top             =   360
         Width           =   4785
         _Version        =   786432
         _ExtentX        =   8440
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.GroupBox GroupBox2 
         Height          =   1650
         Left            =   6960
         TabIndex        =   11
         Top             =   240
         Width           =   4695
         _Version        =   786432
         _ExtentX        =   8281
         _ExtentY        =   2910
         _StockProps     =   79
         Caption         =   "Fecha Emision"
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.DateTimePicker dtpDesde 
            Height          =   315
            Left            =   825
            TabIndex        =   12
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
            TabIndex        =   13
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
            TabIndex        =   14
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
         Begin XtremeSuiteControls.Label Label6 
            Height          =   195
            Left            =   2430
            TabIndex        =   17
            Top             =   675
            Width           =   420
            _Version        =   786432
            _ExtentX        =   741
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Hasta"
            AutoSize        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label5 
            Height          =   195
            Left            =   255
            TabIndex        =   16
            Top             =   660
            Width           =   465
            _Version        =   786432
            _ExtentX        =   820
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Desde"
            AutoSize        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label7 
            Height          =   195
            Left            =   240
            TabIndex        =   15
            Top             =   285
            Width           =   480
            _Version        =   786432
            _ExtentX        =   847
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Rango"
            AutoSize        =   -1  'True
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "Nro Recibo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   825
         Width           =   975
      End
      Begin VB.Label P 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
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
         Height          =   195
         Left            =   615
         TabIndex        =   1
         Top             =   420
         Width           =   600
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   12000
      TabIndex        =   5
      Top             =   1320
      Width           =   5415
      Begin XtremeSuiteControls.PushButton btnExportar 
         Height          =   465
         Left            =   1920
         TabIndex        =   6
         Top             =   240
         Width           =   1515
         _Version        =   786432
         _ExtentX        =   2672
         _ExtentY        =   820
         _StockProps     =   79
         Caption         =   "Exportar"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton cmdBuscar 
         Height          =   465
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1515
         _Version        =   786432
         _ExtentX        =   2672
         _ExtentY        =   820
         _StockProps     =   79
         Caption         =   "Buscar"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton cmdImprimir 
         Height          =   465
         Left            =   3720
         TabIndex        =   8
         Top             =   240
         Width           =   1515
         _Version        =   786432
         _ExtentX        =   2672
         _ExtentY        =   820
         _StockProps     =   79
         Caption         =   "Imprimir"
         UseVisualStyle  =   -1  'True
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   19560
      TabIndex        =   4
      Top             =   6120
      Visible         =   0   'False
      Width           =   1140
   End
   Begin GridEX20.GridEX grilla_recibos 
      Height          =   6000
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   17295
      _ExtentX        =   30506
      _ExtentY        =   10583
      Version         =   "2.0"
      HoldSortSettings=   -1  'True
      DefaultGroupMode=   1
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      ColumnAutoResize=   -1  'True
      MethodHoldFields=   -1  'True
      ImageCount      =   2
      ImagePicture1   =   "frmAdminCobranzasLista.frx":000C
      ImagePicture2   =   "frmAdminCobranzasLista.frx":0326
      RowHeaders      =   -1  'True
      DataMode        =   99
      ColumnHeaderHeight=   285
      FrozenColumns   =   1
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   12
      Column(1)       =   "frmAdminCobranzasLista.frx":0640
      Column(2)       =   "frmAdminCobranzasLista.frx":07CC
      Column(3)       =   "frmAdminCobranzasLista.frx":0920
      Column(4)       =   "frmAdminCobranzasLista.frx":0A40
      Column(5)       =   "frmAdminCobranzasLista.frx":0B8C
      Column(6)       =   "frmAdminCobranzasLista.frx":0CD4
      Column(7)       =   "frmAdminCobranzasLista.frx":0E44
      Column(8)       =   "frmAdminCobranzasLista.frx":0FAC
      Column(9)       =   "frmAdminCobranzasLista.frx":111C
      Column(10)      =   "frmAdminCobranzasLista.frx":1278
      Column(11)      =   "frmAdminCobranzasLista.frx":13D0
      Column(12)      =   "frmAdminCobranzasLista.frx":1518
      FormatStylesCount=   10
      FormatStyle(1)  =   "frmAdminCobranzasLista.frx":1698
      FormatStyle(2)  =   "frmAdminCobranzasLista.frx":17D0
      FormatStyle(3)  =   "frmAdminCobranzasLista.frx":1880
      FormatStyle(4)  =   "frmAdminCobranzasLista.frx":1934
      FormatStyle(5)  =   "frmAdminCobranzasLista.frx":1A0C
      FormatStyle(6)  =   "frmAdminCobranzasLista.frx":1AC4
      FormatStyle(7)  =   "frmAdminCobranzasLista.frx":1BA4
      FormatStyle(8)  =   "frmAdminCobranzasLista.frx":1C50
      FormatStyle(9)  =   "frmAdminCobranzasLista.frx":1D00
      FormatStyle(10) =   "frmAdminCobranzasLista.frx":1DE0
      ImageCount      =   2
      ImagePicture(1) =   "frmAdminCobranzasLista.frx":1E8C
      ImagePicture(2) =   "frmAdminCobranzasLista.frx":21A6
      PrinterProperties=   "frmAdminCobranzasLista.frx":24C0
   End
   Begin VB.Menu mnu 
      Caption         =   "mnu"
      Visible         =   0   'False
      Begin VB.Menu nro 
         Caption         =   "nro"
         Enabled         =   0   'False
      End
      Begin VB.Menu verRecibo 
         Caption         =   "Ver..."
      End
      Begin VB.Menu editarRecibo 
         Caption         =   "Editar..."
      End
      Begin VB.Menu aprobarRecibo 
         Caption         =   "Aprobar..."
      End
      Begin VB.Menu mnuAplicarComprobante 
         Caption         =   "Aplicar Cbtes..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuAnular 
         Caption         =   "Anular..."
      End
      Begin VB.Menu nn 
         Caption         =   "-"
      End
      Begin VB.Menu imprimirRecibo 
         Caption         =   "Imprimir..."
      End
      Begin VB.Menu mnuAdquirir 
         Caption         =   "Adquirir"
      End
      Begin VB.Menu mnuArchivos 
         Caption         =   "Archivos Asociados..."
      End
   End
End
Attribute VB_Name = "frmAdminCobranzasLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Implements ISuscriber
Dim vId As String
'Dim strsql As String
Dim rs As New Recordset
'Dim clasea As New classAdministracion
Dim recibos As Collection    'of recibo
Dim Recibo As Recibo
Private tmpIncidencias As New Dictionary
Private tmpArchivos As New Dictionary

Private Sub aprobarRecibo_Click()
    If MsgBox("¿Está seguro de aprobar este recibo?", vbYesNo, "Confirmación") = vbYes Then

        Set Recibo = DAORecibo.FindById(Recibo.Id, True, True, True, True, True)

        If DAORecibo.aprobar(Recibo) Then
            MsgBox "Aprobación exitosa!", vbInformation, "Información"
            llenarLista
        Else
            MsgBox "Error, no se aprobó el recibo!", vbCritical, "Error"
        End If
    End If

End Sub

Private Sub btnExportar_Click()
    Me.progreso.Visible = True
    'Me.lblExportando.Visible = True

    If IsSomething(recibos) Then
        '        If Not DAOFacturaProveedor.ExportarColeccion(facturas, Me.progreso) Then GoTo err1
        If Not DAORecibo.ExportarColeccion(recibos, Me.progreso) Then GoTo err1
    End If

    Me.progreso.Visible = False
    'Me.lblExportando.Visible = False

    Exit Sub
err1:
    MsgBox "Se produjo un error al exportar!", vbCritical, "Error"

End Sub

Private Sub btnLimpiarCliente_Click()
    Me.cboCliente.ListIndex = -1

End Sub

Private Sub cboRangos_Click()
    funciones.CalculateDateRange Me.cboRangos, Me.dtpDesde, Me.dtpHasta

End Sub

Private Sub cmdImprimir_Click()
    With Me.grilla_recibos.PrinterProperties
        .FitColumns = True
        .RepeatHeaders = True
        .Orientation = jgexPPLandscape
        .HeaderString(jgexHFCenter) = "Recibos emitidos"
        .FooterString(jgexHFCenter) = Now
    End With
    Load frmPrintPreview
    frmPrintPreview.Move Me.Left, Me.Top, Me.Width, Me.Height
    grilla_recibos.PrintPreview frmPrintPreview.GEXPreview1
    frmPrintPreview.Show 1
    
End Sub

Private Sub Command1_Click()
    Dim T As String
    T = "SELECT * FROM AdminRecibos WHERE tot_estatico_recibo = 0"

    Dim rs As Collection
    Set rs = DAORecibo.FindAll("tot_estatico_recibo = 0", , , , , True)
    Dim rc As Recibo
    
    For Each rc In rs
        If rc.estado = EstadoRecibo.Aprobado Then
            rc.TotalEstatico.TotalReciboEstatico = rc.total
        End If
    Next

End Sub

Public Sub cmdBuscar_Click()
    llenarLista
    
End Sub

Private Sub editarRecibo_Click()
    Dim F As New frmAdminCobranzasNuevoRecibo
    F.editar = True
    F.cmdActualizar.Enabled = False
    F.cmdGuardar.Enabled = True
    F.reciboId = Recibo.Id
    F.Show

End Sub

Private Sub Form_Load()
    FormHelper.Customize Me
    GridEXHelper.CustomizeGrid Me.grilla_recibos, True, False
    vId = funciones.CreateGUID
    Channel.AgregarSuscriptor Me, Recibos_
    DAOCliente.llenarComboXtremeSuite Me.cboCliente, True, True
    Dim i As Integer
    funciones.FillComboBoxDateRanges Me.cboRangos
    For i = 0 To Me.cboRangos.ListCount - 1
        If Me.cboRangos.ItemData(i) = DateRangeValue.DRV_YearCurrent Then Exit For
    Next i
    Me.cboRangos.ListIndex = i
    llenarLista

    ''Me.caption = caption & "(" & Name & ")"

End Sub

Private Sub llenarLista()
    Set tmpIncidencias = DAOIncidencias.GetCantidadIncidenciasPorReferencia(OI_Recibos)
    Set tmpArchivos = DAOArchivo.GetCantidadArchivosPorReferencia(OA_Recibos)
    Dim F As String
    Dim TotalRecibido As Double
    
    F = "1 = 1"

    If Me.cboCliente.ListIndex > 0 Then F = F & " and rec.idCliente = " & Me.cboCliente.ItemData(Me.cboCliente.ListIndex)

    If LenB(Me.txtNroRecibo.text) > 0 Then F = F & " AND rec.id = " & Me.txtNroRecibo.text



    If Not IsNull(Me.dtpDesde.value) Then
        F = F & " AND rec.Fecha >= " & conectar.Escape(Me.dtpDesde.value)
    End If

    If Not IsNull(Me.dtpHasta.value) Then
        F = F & " AND rec.Fecha <= " & conectar.Escape(Me.dtpHasta.value)
    End If

    Set recibos = DAORecibo.FindAll(F)
    
    Me.grilla_recibos.ItemCount = 0
    Me.grilla_recibos.ItemCount = recibos.count
   
' TOTALIZADOR DE VALORES RECIBIDOS POR RECIBO
    For Each Recibo In recibos
        TotalRecibido = TotalRecibido + Recibo.TotalEstatico.TotalRecibidoEstatico
    If Recibo.estado = Reciboanulado Then
    TotalRecibido = TotalRecibido - Recibo.TotalEstatico.TotalRecibidoEstatico
    End If
    
    Next
    
    Me.lblTotalRecibido(1).caption = FormatCurrency(funciones.FormatearDecimales(TotalRecibido))

    Me.caption = "Recibos (" & recibos.count & " Recibos encontrados)"

End Sub

Private Sub Form_Resize()
    If Me.ScaleHeight > 0 Then
        Me.grilla_recibos.Height = Me.ScaleHeight - 1800
    End If

End Sub

Private Sub grilla_recibos_ColumnHeaderClick(ByVal Column As GridEX20.JSColumn)
    GridEXHelper.ColumnHeaderClick Me.grilla_recibos, Column
    
End Sub

Private Sub grilla_recibos_DblClick()
    verRecibo_Click
    
End Sub

Private Sub grilla_recibos_FetchIcon(ByVal rowIndex As Long, ByVal ColIndex As Integer, ByVal RowBookmark As Variant, ByVal IconIndex As GridEX20.JSRetInteger)
    On Error Resume Next

    Recibo = recibos.item(rowIndex)

    If ColIndex = 6 And tmpArchivos.item(Recibo.Id) > 0 Then
        IconIndex = 1
    End If
    
End Sub

Private Sub grilla_recibos_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        SeleccionarRecibo
        Me.nro.caption = "[ Nro. " & Format(Recibo.Id, "0000") & " ]"

        If Recibo.estado = EstadoRecibo.Pendiente Then   'pendiente
            Me.editarRecibo.Enabled = True
            Me.imprimirRecibo.Enabled = False
            Me.mnuAnular.Enabled = False
            Me.aprobarRecibo.Enabled = True
            Me.mnuAplicarComprobante.Enabled = False
            If Permisos.AdminCobrosAprobaciones = False Then
                Me.aprobarRecibo.Enabled = False
            End If

        ElseIf Recibo.estado = EstadoRecibo.Aprobado Then      'aprobado
            Me.imprimirRecibo.Enabled = True
            Me.editarRecibo.Enabled = False
            Me.aprobarRecibo.Enabled = False
            Me.mnuAnular.Enabled = True

            If Permisos.AdminCobrosAprobaciones = False Then
                Me.mnuAnular.Enabled = False
            End If

            If Recibo.ACuentaDisponible > 0 Then
                Me.mnuAplicarComprobante.Enabled = True
            End If

            If Recibo.ACuentaDisponible <= 0 Then
                Me.mnuAplicarComprobante.Enabled = False
            End If

        ElseIf Recibo.estado = EstadoRecibo.Pendiente Then    'pendiente
            Me.editarRecibo.Enabled = False
            Me.imprimirRecibo.Enabled = True
            Me.aprobarRecibo.Enabled = False
            Me.mnuAnular.Enabled = False

        ElseIf Recibo.estado = EstadoRecibo.Reciboanulado Then    'anulado
            Me.editarRecibo.Enabled = False
            Me.imprimirRecibo.Enabled = True
            Me.aprobarRecibo.Enabled = False
            Me.mnuAnular.Enabled = False
        
        End If

        Me.PopupMenu Me.mnu

    End If

End Sub

Private Sub grilla_recibos_RowFormat(RowBuffer As GridEX20.JSRowData)
    If RowBuffer.rowIndex = 0 Then Exit Sub
    If RowBuffer.value(11) = "Aprobado" Then
        RowBuffer.CellStyle(11) = "Verde"
    ElseIf RowBuffer.value(11) = "Anulado" Then

        RowBuffer.CellStyle(11) = "Anulado"
    Else
        RowBuffer.CellStyle(11) = "Rojo"
    End If

    If RowBuffer.value(10) > 0 Then
        RowBuffer.CellStyle(10) = "HayArchivosIncidencias"
    End If

End Sub

Private Sub grilla_recibos_SelectionChange()
   SeleccionarRecibo

End Sub

Private Sub SeleccionarRecibo()
    On Error Resume Next
    Set Recibo = recibos.item(Me.grilla_recibos.rowIndex(Me.grilla_recibos.row))

End Sub

Private Sub grilla_recibos_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Set Recibo = recibos.item(rowIndex)
    
    Values(1) = Recibo.Id
    Values(2) = Format(Recibo.FEcha, "yyyy/mm/dd", vbSunday)

    Values(3) = Recibo.Cliente.razon
    Values(4) = Recibo.FechaCreacion
    Values(5) = Recibo.moneda.NombreCorto
    
    If Recibo.estado = 1 Or Recibo.estado = Reciboanulado Then
    Values(6) = "-"
    Values(7) = "-"
    Values(8) = "-"
        
    Values(9) = "-"
    Values(10) = "-"
    
    Else
        Values(7) = Replace(FormatCurrency(funciones.FormatearDecimales(Recibo.TotalEstatico.TotalReciboEstatico + Recibo.TotalRetenciones)), "$", "")
        Values(8) = Replace(FormatCurrency(funciones.FormatearDecimales(Recibo.TotalRetenciones)), "$", "")
        'VALORES POR PAGAR
        Values(6) = Replace(FormatCurrency(funciones.FormatearDecimales(Recibo.TotalEstatico.TotalReciboEstatico)), "$", "")

        'IMPORTE TOTAL PAGADO
        Values(9) = Replace(FormatCurrency(funciones.FormatearDecimales(Recibo.TotalEstatico.TotalRecibidoEstatico)), "$", "")

        'IMPORTE SALDO TOTAL A CUENTA
        Values(10) = Replace(FormatCurrency(funciones.FormatearDecimales(Recibo.ACuentaDisponible)), "$", "")

End If
    
    Values(11) = enums.EnumEstadoRecibo(Recibo.estado)
    
    Values(12) = IIf(IsEmpty(tmpArchivos(Recibo.Id)), 0, tmpArchivos(Recibo.Id))
    
End Sub


Private Sub imprimirRecibo_Click()
    If MsgBox("¿Desea imprimir el recibo?", vbYesNo, "Confirmación") = vbYes Then
        DAORecibo.Imprimir Recibo.Id

    End If

End Sub

Private Property Get ISuscriber_id() As String
    ISuscriber_id = vId
    
End Property

Private Function ISuscriber_Notificarse(EVENTO As clsEventoObserver) As Variant
    Dim tmp As Recibo
    If EVENTO.EVENTO = agregar_ Then
        llenarLista
        Me.grilla_recibos.Refresh

    ElseIf EVENTO.EVENTO = modificar_ Then

        Set tmp = EVENTO.Elemento

        Dim i As Long
        For i = recibos.count To 1 Step -1

            If recibos(i).Id = tmp.Id Then

                recibos.remove i
                If recibos.count > 0 Then
                    If i = 1 Then    'ver esto cuand oes un solo item
                        recibos.Add tmp, CStr(tmp.Id), 1
                    ElseIf (i - 1) = recibos.count Then
                        recibos.Add tmp, CStr(tmp.Id), , i - 1
                    Else
                        recibos.Add tmp, CStr(tmp.Id), i
                    End If
                Else
                    recibos.Add tmp, CStr(tmp.Id)
                End If

                Me.grilla_recibos.RefreshRowIndex i

                Exit For
            End If
        Next
    End If

End Function

Private Sub mnuAdquirir_Click()
    On Error Resume Next
    Dim archivos As New classArchivos
    If archivos.escanearDocumento(OrigenArchivos.OA_Recibos, Recibo.Id) Then
        Set m_Archivos = DAOArchivo.GetCantidadArchivosPorReferencia(OA_Recibos)
        Me.grilla_recibos.RefreshRowIndex (Recibo.Id)
    End If
    
End Sub

Private Sub mnuAnular_Click()
    On Error GoTo err1
    
    If MsgBox("¿Desera anular el recibo número " & Recibo.Id & " ?" & Chr(10) & "Esta acción no tiene rollback", vbYesNo, "Confirmación") = vbYes Then
        DAORecibo.Anular Recibo
        MsgBox "Recibo anulado con éxito!", vbInformation, "Información"
    End If

    Exit Sub
err1:
    MsgBox Err.Description
    
End Sub

Private Sub mnuAplicarComprobante_Click()
    Dim F As New frmAdminCobranzasNuevoRecibo
    F.editar = True
    F.cmdActualizar.Enabled = True
    F.cmdGuardar.Enabled = False
    F.reciboId = Recibo.Id
    F.Show

End Sub

Private Sub mnuArchivos_Click()
    Dim frmarchi1 As New frmArchivos2
    frmarchi1.Origen = OA_Recibos
    frmarchi1.ObjetoId = Recibo.Id
    frmarchi1.caption = "Recibo Nº " & Recibo.Id
    frmarchi1.Show
    
End Sub

Private Sub txtNroRecibo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        llenarLista
    End If
    
End Sub

Private Sub verRecibo_Click()
'31.10.2022- SE AGREGA ESTA LINEA PARA QUE PRIMERO EJECUTE LA FUNCION SELECCIONARRECIBO
    SeleccionarRecibo
    On Error GoTo err1
    Dim F As New frmAdminCobranzasNuevoRecibo
    F.editar = False
    F.reciboId = Recibo.Id
    F.Show
    Exit Sub
err1:

End Sub
