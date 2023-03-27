VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmAdminCobranzasListaAnticipo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Recibos de Anticipo"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   13740
   Begin VB.TextBox txtNroRecibo 
      Height          =   285
      Left            =   1395
      TabIndex        =   3
      Top             =   600
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   12330
      TabIndex        =   0
      Top             =   165
      Visible         =   0   'False
      Width           =   1140
   End
   Begin XtremeSuiteControls.ComboBox cboCliente 
      Height          =   315
      Left            =   1395
      TabIndex        =   1
      Top             =   165
      Width           =   5865
      _Version        =   786432
      _ExtentX        =   10345
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.PushButton cmdBuscar 
      Height          =   345
      Left            =   4155
      TabIndex        =   2
      Top             =   585
      Width           =   1515
      _Version        =   786432
      _ExtentX        =   2672
      _ExtentY        =   609
      _StockProps     =   79
      Caption         =   "Buscar"
      BackColor       =   12632256
      UseVisualStyle  =   -1  'True
   End
   Begin GridEX20.GridEX grilla_recibos 
      Height          =   5640
      Left            =   0
      TabIndex        =   4
      Top             =   1110
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   9948
      Version         =   "2.0"
      HoldSortSettings=   -1  'True
      DefaultGroupMode=   1
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      ColumnAutoResize=   -1  'True
      MethodHoldFields=   -1  'True
      ImageCount      =   2
      ImagePicture1   =   "frmAdminCobranzasListaAnticipo.frx":0000
      ImagePicture2   =   "frmAdminCobranzasListaAnticipo.frx":031A
      RowHeaders      =   -1  'True
      DataMode        =   99
      ColumnHeaderHeight=   285
      FrozenColumns   =   1
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   8
      Column(1)       =   "frmAdminCobranzasListaAnticipo.frx":0634
      Column(2)       =   "frmAdminCobranzasListaAnticipo.frx":07C0
      Column(3)       =   "frmAdminCobranzasListaAnticipo.frx":0914
      Column(4)       =   "frmAdminCobranzasListaAnticipo.frx":0A34
      Column(5)       =   "frmAdminCobranzasListaAnticipo.frx":0BA8
      Column(6)       =   "frmAdminCobranzasListaAnticipo.frx":0CF0
      Column(7)       =   "frmAdminCobranzasListaAnticipo.frx":0E50
      Column(8)       =   "frmAdminCobranzasListaAnticipo.frx":0F98
      FormatStylesCount=   10
      FormatStyle(1)  =   "frmAdminCobranzasListaAnticipo.frx":1118
      FormatStyle(2)  =   "frmAdminCobranzasListaAnticipo.frx":1250
      FormatStyle(3)  =   "frmAdminCobranzasListaAnticipo.frx":1300
      FormatStyle(4)  =   "frmAdminCobranzasListaAnticipo.frx":13B4
      FormatStyle(5)  =   "frmAdminCobranzasListaAnticipo.frx":148C
      FormatStyle(6)  =   "frmAdminCobranzasListaAnticipo.frx":1544
      FormatStyle(7)  =   "frmAdminCobranzasListaAnticipo.frx":1624
      FormatStyle(8)  =   "frmAdminCobranzasListaAnticipo.frx":16D0
      FormatStyle(9)  =   "frmAdminCobranzasListaAnticipo.frx":1780
      FormatStyle(10) =   "frmAdminCobranzasListaAnticipo.frx":1860
      ImageCount      =   2
      ImagePicture(1) =   "frmAdminCobranzasListaAnticipo.frx":190C
      ImagePicture(2) =   "frmAdminCobranzasListaAnticipo.frx":1C26
      PrinterProperties=   "frmAdminCobranzasListaAnticipo.frx":1F40
   End
   Begin XtremeSuiteControls.PushButton cmdImprimir 
      Height          =   345
      Left            =   5760
      TabIndex        =   5
      Top             =   585
      Width           =   1515
      _Version        =   786432
      _ExtentX        =   2672
      _ExtentY        =   609
      _StockProps     =   79
      Caption         =   "Imprimir"
      BackColor       =   12632256
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.GroupBox GroupBox2 
      Height          =   1050
      Left            =   7395
      TabIndex        =   6
      Top             =   0
      Width           =   4695
      _Version        =   786432
      _ExtentX        =   8281
      _ExtentY        =   1852
      _StockProps     =   79
      Caption         =   "Fecha Emision"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.DateTimePicker dtpDesde 
         Height          =   315
         Left            =   825
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
      Begin XtremeSuiteControls.DateTimePicker dtpHasta 
         Height          =   315
         Left            =   3000
         TabIndex        =   8
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
         TabIndex        =   9
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
         TabIndex        =   12
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
      Begin XtremeSuiteControls.Label Label5 
         Height          =   195
         Left            =   255
         TabIndex        =   11
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
      Begin XtremeSuiteControls.Label Label7 
         Height          =   195
         Left            =   240
         TabIndex        =   10
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
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Nro Recibo:"
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
      Left            =   315
      TabIndex        =   14
      Top             =   630
      Width           =   1035
   End
   Begin VB.Label P 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente:"
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
      Left            =   690
      TabIndex        =   13
      Top             =   225
      Width           =   660
   End
   Begin VB.Menu mnu 
      Caption         =   "mnu"
      Visible         =   0   'False
      Begin VB.Menu nro 
         Caption         =   "nro"
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
      Begin VB.Menu mnuAnular 
         Caption         =   "Anular..."
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu nn 
         Caption         =   "-"
      End
      Begin VB.Menu imprimirRecibo 
         Caption         =   "Imprimir..."
      End
      Begin VB.Menu mnuAdquirir 
         Caption         =   "Adquirir"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuArchivos 
         Caption         =   "Archivos Asociados..."
      End
   End
End
Attribute VB_Name = "frmAdminCobranzasListaAnticipo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Implements ISuscriber
Dim vId As String
Dim strsql As String
Dim rs As New Recordset
Dim clasea As New classAdministracion
Dim recibos As Collection    'of recibo
Dim recibo As recibo
Private tmpIncidencias As New Dictionary
Private tmpArchivos As New Dictionary

Private Sub aprobarRecibo_Click()
    Dim idRecibo As Long
    If MsgBox("¿Está seguro de aprobar este recibo?", vbYesNo, "Confirmación") = vbYes Then

        Set recibo = DAOReciboAnticipo.FindById(recibo.id, True, True, True, True, True)

        If DAOReciboAnticipo.aprobar(recibo) Then
            MsgBox "Aprobación exitosa!", vbInformation, "Información"
            llenarLista
        Else
            MsgBox "Error, no se aprobó el recibo!", vbCritical, "Error"
        End If
    End If
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
    T = "SELECT * FROM AdminRecibosAnticipo WHERE tot_estatico_recibo = 0"

    Dim rs As Collection

    Set rs = DAOReciboAnticipo.FindAll("tot_estatico_recibo = 0", , , , , True)

    Dim rc As recibo


    For Each rc In rs
        If rc.estado = EstadoRecibo.Aprobado Then
            rc.totalEstatico.TotalReciboEstatico = rc.Total
        End If
    Next


End Sub

Private Sub cmdBuscar_Click()
    llenarLista
End Sub

Private Sub editarRecibo_Click()

    Dim F As New frmAdminCobranzasNuevoReciboAnticipo
    F.editar = True
    F.reciboId = recibo.id
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
    
    'Me.caption = caption & "(" & Name & ")"
        
End Sub

Private Sub llenarLista()

    Set tmpIncidencias = DAOIncidencias.GetCantidadIncidenciasPorReferencia(OI_Recibos)
    Set tmpArchivos = DAOArchivo.GetCantidadArchivosPorReferencia(OA_Recibos)
    Dim F As String
    F = "1 = 1"

    If Me.cboCliente.ListIndex > 0 Then F = F & " and rec.idCliente = " & Me.cboCliente.ItemData(Me.cboCliente.ListIndex)

    If LenB(Me.txtNroRecibo.text) > 0 Then F = F & " AND rec.id = " & Me.txtNroRecibo.text



    If Not IsNull(Me.dtpDesde.value) Then
        F = F & " AND rec.Fecha >= " & conectar.Escape(Me.dtpDesde.value)
    End If

    If Not IsNull(Me.dtpHasta.value) Then
        F = F & " AND rec.Fecha <= " & conectar.Escape(Me.dtpHasta.value)
    End If


    Set recibos = DAOReciboAnticipo.FindAll(F)
    Me.grilla_recibos.ItemCount = 0
    Me.grilla_recibos.ItemCount = recibos.count

End Sub

Private Sub Form_Resize()
    If Me.ScaleHeight > 0 Then
        Me.grilla_recibos.Height = Me.ScaleHeight - 1200
    End If
    Me.grilla_recibos.Width = Me.ScaleWidth
End Sub

Private Sub grilla_recibos_ColumnHeaderClick(ByVal Column As GridEX20.JSColumn)
    GridEXHelper.ColumnHeaderClick Me.grilla_recibos, Column
End Sub

Private Sub grilla_recibos_DblClick()
    verRecibo_Click
End Sub

Private Sub grilla_recibos_FetchIcon(ByVal RowIndex As Long, ByVal ColIndex As Integer, ByVal RowBookmark As Variant, ByVal IconIndex As GridEX20.JSRetInteger)
    On Error Resume Next
SeleccionarRecibo
    recibo = recibos.item(RowIndex)

    If ColIndex = 6 And tmpArchivos.item(recibo.id) > 0 Then
        IconIndex = 1
    End If
End Sub

Private Sub grilla_recibos_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.nro.caption = "[ Nro. " & Format(recibo.id, "0000") & " ]"
        If recibo.estado = EstadoRecibo.Pendiente Then   'pendiente
            Me.editarRecibo.Enabled = True
            Me.imprimirRecibo.Enabled = False
            Me.mnuAnular.Enabled = False
            Me.aprobarRecibo.Enabled = True
            If Permisos.AdminCobrosAprobaciones = False Then
                Me.aprobarRecibo.Enabled = False
            End If
        ElseIf recibo.estado = EstadoRecibo.Aprobado Then      'aprobado
            Me.imprimirRecibo.Enabled = True
            Me.editarRecibo.Enabled = False
            Me.aprobarRecibo.Enabled = False
            Me.mnuAnular.Enabled = True
            If Permisos.AdminCobrosAprobaciones = False Then
                Me.mnuAnular.Enabled = False
            End If

        ElseIf recibo.estado = EstadoRecibo.Pendiente Then    'anulado
            Me.editarRecibo.Enabled = False
            Me.imprimirRecibo.Enabled = True
            Me.aprobarRecibo.Enabled = False
            Me.mnuAnular.Enabled = False

        ElseIf recibo.estado = EstadoRecibo.Reciboanulado Then    'anulado
            Me.editarRecibo.Enabled = False
            Me.imprimirRecibo.Enabled = True
            Me.aprobarRecibo.Enabled = False
            Me.mnuAnular.Enabled = False
        End If
        Me.PopupMenu Me.mnu
    End If

End Sub

Private Sub grilla_recibos_RowFormat(RowBuffer As GridEX20.JSRowData)
    If RowBuffer.RowIndex = 0 Then Exit Sub
    If RowBuffer.value(7) = "Aprobado" Then
        RowBuffer.CellStyle(7) = "Verde"
    ElseIf RowBuffer.value(7) = "Anulado" Then

        RowBuffer.CellStyle(7) = "Anulado"
    Else
        RowBuffer.CellStyle(7) = "Rojo"
    End If

    If RowBuffer.value(8) > 0 Then
        RowBuffer.CellStyle(8) = "HayArchivosIncidencias"
    End If



End Sub

Private Sub grilla_recibos_SelectionChange()
   
    SeleccionarRecibo
    
End Sub

Private Sub SeleccionarRecibo()
    On Error Resume Next
    Set recibo = recibos.item(Me.grilla_recibos.RowIndex(Me.grilla_recibos.row))

End Sub

Private Sub grilla_recibos_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Set recibo = recibos.item(RowIndex)
    Values(1) = recibo.id
    Values(2) = Format(recibo.FEcha, "yyyy/mm/dd", vbSunday)
    
    
    
    Values(3) = recibo.cliente.razon
    Values(4) = recibo.fechaCreacion
    Values(5) = recibo.moneda.NombreCorto
    'Values(6) = funciones.FormatearDecimales(recibo.totalEstatico.TotalReciboEstatico + recibo.TotalRetenciones)
    'Values(6) = funciones.FormatearDecimales(recibo.ACuentaDisponible)
    Values(6) = Replace(FormatCurrency(funciones.FormatearDecimales(recibo.ACuentaDisponible)), "$", "")
    Values(7) = enums.EnumEstadoRecibo(recibo.estado)
    Values(8) = IIf(IsEmpty(tmpArchivos(recibo.id)), 0, tmpArchivos(recibo.id))
End Sub


Private Sub imprimirRecibo_Click()
    If MsgBox("¿Desea imprimir el recibo?", vbYesNo, "Confirmación") = vbYes Then
        DAOReciboAnticipo.Imprimir recibo.id

    End If

End Sub

Private Property Get ISuscriber_id() As String
    ISuscriber_id = vId
End Property

Private Function ISuscriber_Notificarse(EVENTO As clsEventoObserver) As Variant
    Dim tmp As recibo
    If EVENTO.EVENTO = agregar_ Then
        llenarLista
        Me.grilla_recibos.Refresh
    ElseIf EVENTO.EVENTO = modificar_ Then
        Set tmp = EVENTO.Elemento

        Dim i As Long
        For i = recibos.count To 1 Step -1

            If recibos(i).id = tmp.id Then

                recibos.remove i
                If recibos.count > 0 Then
                    If i = 1 Then    'ver esto cuand oes un solo item
                        recibos.Add tmp, CStr(tmp.id), 1
                    ElseIf (i - 1) = recibos.count Then
                        recibos.Add tmp, CStr(tmp.id), , i - 1
                    Else
                        recibos.Add tmp, CStr(tmp.id), i
                    End If
                Else
                    recibos.Add tmp, CStr(tmp.id)
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
    If archivos.escanearDocumento(OrigenArchivos.OA_Recibos, recibo.id) Then
        Set m_Archivos = DAOArchivo.GetCantidadArchivosPorReferencia(OA_Recibos)
        Me.grilla_recibos.RefreshRowIndex (recibo.id)
    End If
End Sub

Private Sub mnuAnular_Click()
    On Error GoTo err1


    If MsgBox("¿Desera anular el recibo número " & recibo.id & " ?" & Chr(10) & "Esta acción no tiene rollback", vbYesNo, "Confirmación") = vbYes Then
        DAOReciboAnticipo.Anular recibo
        MsgBox "Recibo anulado con éxito!", vbInformation, "Información"
    End If

    Exit Sub
err1:
    MsgBox Err.Description
End Sub

Private Sub mnuArchivos_Click()
    Dim frmarchi1 As New frmArchivos2
    frmarchi1.Origen = OA_Recibos
    frmarchi1.ObjetoId = recibo.id
    frmarchi1.caption = "Recibo Nº " & recibo.id
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
    Dim F As New frmAdminCobranzasNuevoReciboAnticipo
    F.editar = False
    F.reciboId = recibo.id
    F.Show
    Exit Sub
err1:


End Sub


