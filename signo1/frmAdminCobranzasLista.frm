VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmAdminCobranzasLista 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recibos"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13635
   Icon            =   "frmAdminCobranzasLista.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   13635
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   12330
      TabIndex        =   6
      Top             =   240
      Visible         =   0   'False
      Width           =   1140
   End
   Begin XtremeSuiteControls.ComboBox cboCliente 
      Height          =   315
      Left            =   1395
      TabIndex        =   5
      Top             =   240
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
      TabIndex        =   4
      Top             =   660
      Width           =   1515
      _Version        =   786432
      _ExtentX        =   2672
      _ExtentY        =   609
      _StockProps     =   79
      Caption         =   "Buscar"
      UseVisualStyle  =   -1  'True
   End
   Begin VB.TextBox txtNroRecibo 
      Height          =   285
      Left            =   1395
      TabIndex        =   1
      Top             =   675
      Width           =   1500
   End
   Begin GridEX20.GridEX grilla_recibos 
      Height          =   5640
      Left            =   0
      TabIndex        =   0
      Top             =   1185
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
      ImagePicture1   =   "frmAdminCobranzasLista.frx":000C
      ImagePicture2   =   "frmAdminCobranzasLista.frx":0326
      RowHeaders      =   -1  'True
      DataMode        =   99
      ColumnHeaderHeight=   285
      FrozenColumns   =   1
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   10
      Column(1)       =   "frmAdminCobranzasLista.frx":0640
      Column(2)       =   "frmAdminCobranzasLista.frx":07CC
      Column(3)       =   "frmAdminCobranzasLista.frx":0920
      Column(4)       =   "frmAdminCobranzasLista.frx":0A40
      Column(5)       =   "frmAdminCobranzasLista.frx":0B8C
      Column(6)       =   "frmAdminCobranzasLista.frx":0CD4
      Column(7)       =   "frmAdminCobranzasLista.frx":0E44
      Column(8)       =   "frmAdminCobranzasLista.frx":0FA0
      Column(9)       =   "frmAdminCobranzasLista.frx":10F8
      Column(10)      =   "frmAdminCobranzasLista.frx":1240
      FormatStylesCount=   10
      FormatStyle(1)  =   "frmAdminCobranzasLista.frx":13C0
      FormatStyle(2)  =   "frmAdminCobranzasLista.frx":14F8
      FormatStyle(3)  =   "frmAdminCobranzasLista.frx":15A8
      FormatStyle(4)  =   "frmAdminCobranzasLista.frx":165C
      FormatStyle(5)  =   "frmAdminCobranzasLista.frx":1734
      FormatStyle(6)  =   "frmAdminCobranzasLista.frx":17EC
      FormatStyle(7)  =   "frmAdminCobranzasLista.frx":18CC
      FormatStyle(8)  =   "frmAdminCobranzasLista.frx":1978
      FormatStyle(9)  =   "frmAdminCobranzasLista.frx":1A28
      FormatStyle(10) =   "frmAdminCobranzasLista.frx":1B08
      ImageCount      =   2
      ImagePicture(1) =   "frmAdminCobranzasLista.frx":1BB4
      ImagePicture(2) =   "frmAdminCobranzasLista.frx":1ECE
      PrinterProperties=   "frmAdminCobranzasLista.frx":21E8
   End
   Begin XtremeSuiteControls.PushButton cmdImprimir 
      Height          =   345
      Left            =   5760
      TabIndex        =   7
      Top             =   660
      Width           =   1515
      _Version        =   786432
      _ExtentX        =   2672
      _ExtentY        =   609
      _StockProps     =   79
      Caption         =   "Imprimir"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.GroupBox GroupBox2 
      Height          =   1050
      Left            =   7395
      TabIndex        =   8
      Top             =   75
      Width           =   4695
      _Version        =   786432
      _ExtentX        =   8281
      _ExtentY        =   1852
      _StockProps     =   79
      Caption         =   "Fecha Emision"
      BackColor       =   -2147483633
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.DateTimePicker dtpDesde 
         Height          =   315
         Left            =   825
         TabIndex        =   9
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
         TabIndex        =   10
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
         TabIndex        =   11
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
         TabIndex        =   14
         Top             =   285
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
         Left            =   255
         TabIndex        =   13
         Top             =   660
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
         Left            =   2430
         TabIndex        =   12
         Top             =   675
         Width           =   420
         _Version        =   786432
         _ExtentX        =   741
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Hasta"
         AutoSize        =   -1  'True
      End
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
      Left            =   690
      TabIndex        =   3
      Top             =   300
      Width           =   600
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
      Left            =   315
      TabIndex        =   2
      Top             =   705
      Width           =   975
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

        Set recibo = DAORecibo.FindById(recibo.Id, True, True, True, True, True)

        If DAORecibo.aprobar(recibo) Then
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
    T = "SELECT * FROM AdminRecibos WHERE tot_estatico_recibo = 0"

    Dim rs As Collection

    Set rs = DAORecibo.FindAll("tot_estatico_recibo = 0", , , , , True)

    Dim rc As recibo


    For Each rc In rs
        If rc.estado = EstadoRecibo.Aprobado Then
            rc.TotalEstatico.TotalReciboEstatico = rc.Total
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
    F.reciboId = recibo.Id
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

    recibo = recibos.item(RowIndex)

    If ColIndex = 6 And tmpArchivos.item(recibo.Id) > 0 Then
        IconIndex = 1
    End If
End Sub

Private Sub grilla_recibos_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
    SeleccionarRecibo
        Me.nro.caption = "[ Nro. " & Format(recibo.Id, "0000") & " ]"
            
            If recibo.estado = EstadoRecibo.Pendiente Then   'pendiente
                    Me.editarRecibo.Enabled = True
                    Me.imprimirRecibo.Enabled = False
                    Me.mnuAnular.Enabled = False
                    Me.aprobarRecibo.Enabled = True
                    Me.mnuAplicarComprobante.Enabled = False
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
                        
                        If recibo.ACuentaDisponible > 0 Then
                                Me.mnuAplicarComprobante.Enabled = True
                        End If
                
                        If recibo.ACuentaDisponible <= 0 Then
                            Me.mnuAplicarComprobante.Enabled = False
                        End If
    
                ElseIf recibo.estado = EstadoRecibo.Pendiente Then    'pendiente
                    Me.editarRecibo.Enabled = False
                    Me.imprimirRecibo.Enabled = True
                    Me.aprobarRecibo.Enabled = False
                    Me.mnuAnular.Enabled = False
    
                ElseIf recibo.estado = EstadoRecibo.Reciboanulado Then    'anulado
                    Me.editarRecibo.Enabled = False
                    Me.imprimirRecibo.Enabled = True
                    Me.aprobarRecibo.Enabled = False
                    Me.mnuAnular.Enabled = False
        
            'End If
        
       End If
       
        Me.PopupMenu Me.mnu
        
    End If

End Sub

Private Sub grilla_recibos_RowFormat(RowBuffer As GridEX20.JSRowData)
    If RowBuffer.RowIndex = 0 Then Exit Sub
    If RowBuffer.value(9) = "Aprobado" Then
        RowBuffer.CellStyle(9) = "Verde"
    ElseIf RowBuffer.value(9) = "Anulado" Then

        RowBuffer.CellStyle(9) = "Anulado"
    Else
        RowBuffer.CellStyle(9) = "Rojo"
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
    Set recibo = recibos.item(Me.grilla_recibos.RowIndex(Me.grilla_recibos.row))

End Sub

Private Sub grilla_recibos_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Set recibo = recibos.item(RowIndex)
    Values(1) = recibo.Id
    Values(2) = Format(recibo.FEcha, "yyyy/mm/dd", vbSunday)
          
    Values(3) = recibo.cliente.razon
    Values(4) = recibo.FechaCreacion
    Values(5) = recibo.moneda.NombreCorto
    
    'VALORES POR PAGAR
    Values(6) = Replace(FormatCurrency(funciones.FormatearDecimales(recibo.TotalEstatico.TotalReciboEstatico)), "$", "")
    
    'IMPORTE TOTAL PAGADO
    Values(7) = Replace(FormatCurrency(funciones.FormatearDecimales(recibo.TotalEstatico.TotalRecibidoEstatico)), "$", "")
    
    'IMPORTE SALDO TOTAL A CUENTA
    Values(8) = Replace(FormatCurrency(funciones.FormatearDecimales(recibo.ACuentaDisponible)), "$", "")
    
    Values(9) = enums.EnumEstadoRecibo(recibo.estado)
    Values(10) = IIf(IsEmpty(tmpArchivos(recibo.Id)), 0, tmpArchivos(recibo.Id))
End Sub


Private Sub imprimirRecibo_Click()
    If MsgBox("¿Desea imprimir el recibo?", vbYesNo, "Confirmación") = vbYes Then
        DAORecibo.Imprimir recibo.Id

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
    If archivos.escanearDocumento(OrigenArchivos.OA_Recibos, recibo.Id) Then
        Set m_Archivos = DAOArchivo.GetCantidadArchivosPorReferencia(OA_Recibos)
        Me.grilla_recibos.RefreshRowIndex (recibo.Id)
    End If
End Sub

Private Sub mnuAnular_Click()
    On Error GoTo err1


    If MsgBox("¿Desera anular el recibo número " & recibo.Id & " ?" & Chr(10) & "Esta acción no tiene rollback", vbYesNo, "Confirmación") = vbYes Then
        DAORecibo.Anular recibo
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
    F.reciboId = recibo.Id
    F.Show
    
End Sub

Private Sub mnuArchivos_Click()
    Dim frmarchi1 As New frmArchivos2
    frmarchi1.Origen = OA_Recibos
    frmarchi1.ObjetoId = recibo.Id
    frmarchi1.caption = "Recibo Nº " & recibo.Id
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
    F.reciboId = recibo.Id
    F.Show
    Exit Sub
err1:


End Sub

