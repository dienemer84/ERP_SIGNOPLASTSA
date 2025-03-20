VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmPlaneamientoRemitosLista 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Remitos"
   ClientHeight    =   8805
   ClientLeft      =   1455
   ClientTop       =   2130
   ClientWidth     =   12825
   Icon            =   "frmListaRemitos.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8805
   ScaleWidth      =   12825
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.ProgressBar ProgressBar 
      Height          =   495
      Left            =   11760
      TabIndex        =   27
      Top             =   1800
      Visible         =   0   'False
      Width           =   375
      _Version        =   786432
      _ExtentX        =   661
      _ExtentY        =   873
      _StockProps     =   93
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   19
      Top             =   8460
      Width           =   12825
      _ExtentX        =   22622
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Contadores"
            TextSave        =   "Contadores"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Contadores"
            TextSave        =   "Contadores"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Contadores"
            TextSave        =   "Contadores"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   3105
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   16965
      _Version        =   786432
      _ExtentX        =   29924
      _ExtentY        =   5477
      _StockProps     =   79
      Caption         =   "Parámetros de búsqueda"
      UseVisualStyle  =   -1  'True
      Begin VB.TextBox txtFacturas 
         Height          =   285
         Left            =   1080
         TabIndex        =   28
         Top             =   2640
         Visible         =   0   'False
         Width           =   1695
      End
      Begin XtremeSuiteControls.PushButton btnClearEstado 
         Height          =   285
         Left            =   2880
         TabIndex        =   25
         Top             =   2195
         Width           =   420
         _Version        =   786432
         _ExtentX        =   741
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "X"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cboEstado 
         Height          =   315
         Left            =   1080
         TabIndex        =   23
         Top             =   2160
         Width           =   1695
         _Version        =   786432
         _ExtentX        =   2990
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.GroupBox GroupBox 
         Height          =   855
         Index           =   1
         Left            =   6360
         TabIndex        =   20
         Top             =   1680
         Width           =   4815
         _Version        =   786432
         _ExtentX        =   8493
         _ExtentY        =   1508
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.PushButton PusExportar 
            Height          =   495
            Left            =   3480
            TabIndex        =   26
            Top             =   240
            Width           =   1215
            _Version        =   786432
            _ExtentX        =   2143
            _ExtentY        =   873
            _StockProps     =   79
            Caption         =   "Exportar"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton cmdBuscar 
            Default         =   -1  'True
            Height          =   495
            Left            =   240
            TabIndex        =   21
            Top             =   240
            Width           =   1575
            _Version        =   786432
            _ExtentX        =   2778
            _ExtentY        =   873
            _StockProps     =   79
            Caption         =   "Buscar"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton cmdImprimir 
            Height          =   495
            Left            =   2160
            TabIndex        =   22
            Top             =   240
            Width           =   1200
            _Version        =   786432
            _ExtentX        =   2117
            _ExtentY        =   873
            _StockProps     =   79
            Caption         =   "Imprimir"
            UseVisualStyle  =   -1  'True
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox 
         Height          =   1335
         Index           =   0
         Left            =   6360
         TabIndex        =   12
         Top             =   240
         Width           =   4815
         _Version        =   786432
         _ExtentX        =   8493
         _ExtentY        =   2355
         _StockProps     =   79
         Caption         =   "Fecha"
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.DateTimePicker dtpDesde 
            Height          =   315
            Left            =   855
            TabIndex        =   13
            Top             =   735
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
            Left            =   3030
            TabIndex        =   14
            Top             =   735
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
            TabIndex        =   15
            Top             =   240
            Width           =   2190
            _Version        =   786432
            _ExtentX        =   3863
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Style           =   2
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.Label Label6 
            Height          =   195
            Left            =   2460
            TabIndex        =   18
            Top             =   795
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
            Left            =   285
            TabIndex        =   17
            Top             =   780
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
            TabIndex        =   16
            Top             =   300
            Width           =   480
            _Version        =   786432
            _ExtentX        =   847
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Rango"
            AutoSize        =   -1  'True
         End
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   285
         Left            =   1080
         TabIndex        =   3
         Top             =   1305
         Width           =   4575
      End
      Begin VB.TextBox txtNroRemito 
         Height          =   285
         Left            =   1080
         TabIndex        =   2
         Top             =   480
         Width           =   1695
      End
      Begin XtremeSuiteControls.ComboBox cboClientes 
         Height          =   315
         Left            =   1080
         TabIndex        =   7
         Top             =   840
         Width           =   4575
         _Version        =   786432
         _ExtentX        =   8070
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   285
         Left            =   5760
         TabIndex        =   8
         Top             =   855
         Width           =   420
         _Version        =   786432
         _ExtentX        =   741
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "X"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cboEstadoFacturado 
         Height          =   315
         Left            =   1080
         TabIndex        =   9
         Top             =   1725
         Width           =   1695
         _Version        =   786432
         _ExtentX        =   2990
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.PushButton cmdLimpiaEstado 
         Height          =   285
         Left            =   2880
         TabIndex        =   11
         Top             =   1740
         Width           =   420
         _Version        =   786432
         _ExtentX        =   741
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "X"
         UseVisualStyle  =   -1  'True
      End
      Begin VB.Label lblFacturas 
         Caption         =   "Facturas:"
         Height          =   255
         Left            =   360
         TabIndex        =   29
         Top             =   2640
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lblEstado 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Estado:"
         Height          =   165
         Left            =   240
         TabIndex        =   24
         Top             =   2235
         Width           =   735
      End
      Begin XtremeSuiteControls.Label Label4 
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   1785
         Width           =   765
         _Version        =   786432
         _ExtentX        =   1349
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Facturado:"
         AutoSize        =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente:"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   870
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción:"
         Height          =   255
         Left            =   -120
         TabIndex        =   5
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Número:"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   495
         Width           =   735
      End
   End
   Begin GridEX20.GridEX grilla 
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   3360
      Width           =   16935
      _ExtentX        =   29871
      _ExtentY        =   8070
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
Dim facturasRemitos As Dictionary
Dim m_Archivos As Dictionary


Private Sub AnularRto_Click()
    Dim A As Long
    A = Me.grilla.RowIndex(Me.grilla.row)
    If MsgBox("¿Está seguro de anular el remito?", vbYesNo, "Confirmación") = vbYes Then
        If DAORemitoS.Anular(tmpRto) Then
            MsgBox "Remito anulado con éxito!", vbExclamation, "Información"
            Me.grilla.RefreshRowIndex A

            llenarContadoresStatusBar

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

Private Sub cmdImprimir_Click()
    Dim pro As String
    Dim q As String
    If Me.cboClientes.ListIndex > -1 Then
        pro = " Cliente: " & Me.cboClientes.Text
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

Private Sub btnClearEstado_Click()
    Me.cboEstado.ListIndex = -1
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
    If LenB(Me.txtDescripcion.Text) > 0 Then
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
        filtro = filtro & " and " & DAORemitoS.TABLA_REMITO & "." & DAORemitoS.CAMPO_ESTADO_FACTURADO & "=" & (Me.cboEstadoFacturado.ItemData(Me.cboEstadoFacturado.ListIndex))
    End If

    If Me.cboEstado.ListIndex > -1 Then
        filtro = filtro & " and " & DAORemitoS.TABLA_REMITO & "." & DAORemitoS.CAMPO_ESTADO & "=" & (Me.cboEstado.ItemData(Me.cboEstado.ListIndex))
    End If
    

    

    Set remitos = DAORemitoS.FindAll("and " & filtro)

    Dim remi As Remito
    Dim remitosId As New Collection
    
    For Each remi In remitos
        remitosId.Add remi.Id
    Next
    
    Set facturasRemitos = New Dictionary

'    If Not IsEmpty(Me.txtFacturas) And IsNumeric(txtFacturas) Then
'        filtro = filtro & "  and " & DAORemitoS.TABLA_REMITO & "." & DAORemitoS.CAMPO_NUMERO & "=" & Me.txtFacturas
'   End If
   
   If remitosId.count > 0 Then
       Set facturasRemitos = DAOFactura.FindAllByRemitos(remitosId)

    End If




    Me.grilla.ItemCount = 0
    Me.grilla.ItemCount = remitos.count
    Me.grilla.Update

    llenarContadoresStatusBar

    Me.caption = "Remitos [Cantidad: " & remitos.count & "]"

End Sub


Private Sub Form_Load()
    FormHelper.Customize Me

    DAOCliente.llenarComboXtremeSuite Me.cboClientes, False, True, False
    Me.cboClientes.ListIndex = -1
    GridEXHelper.CustomizeGrid Me.grilla, True
    id_suscriber = funciones.CreateGUID
    Me.grilla.Columns(8).Visible = VerInfoAdministracion
    Me.cboEstadoFacturado.Visible = VerInfoAdministracion
    Me.cboEstado.Visible = VerInfoAdministracion
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

    Me.cboEstado.Clear
    Me.cboEstado.AddItem enums.EnumEstadoRemito(EstadoRemito.RemitoAprobado)
    Me.cboEstado.ItemData(Me.cboEstado.NewIndex) = EstadoRemito.RemitoAprobado
    Me.cboEstado.AddItem enums.EnumEstadoRemito(EstadoRemito.RemitoPendiente)
    Me.cboEstado.ItemData(Me.cboEstado.NewIndex) = EstadoRemito.RemitoPendiente
    Me.cboEstado.AddItem enums.EnumEstadoRemito(EstadoRemito.RemitoAnulado)
    Me.cboEstado.ItemData(Me.cboEstado.NewIndex) = EstadoRemito.RemitoAnulado

    listaRemitos

    llenarContadoresStatusBar

End Sub

Public Sub llenarContadoresStatusBar()

    ContarPendientes
    ContarAprobados
    ContarAnulados

    ContarNoFacturados
    ContarNoFacturables
    ContarFacturadosParcial
    ContarFacturadosTotal

    MostrarCantidadPendientes
    MostrarCantidadAnulados
    MostrarCantidadAprobados

    MostrarFacturadosTotal
    MostrarNoFacturados
    MostrarNoFacturables
    MostrarFacturadosParcial

    StatusBar1.Height = 350
    StatusBar1.Panels(1).Width = 2000
    StatusBar1.Panels(2).Width = 2000
    StatusBar1.Panels(3).Width = 2000

    StatusBar1.Panels(4).Width = 2000
    StatusBar1.Panels(5).Width = 2000
    StatusBar1.Panels(6).Width = 2000
    StatusBar1.Panels(7).Width = 2000
End Sub


Private Sub MostrarCantidadPendientes()
    Dim Cantidad As Integer
    Cantidad = ContarPendientes()
    StatusBar1.Panels(1).Text = "Pendientes Total: " & Cantidad
End Sub


Private Sub MostrarCantidadAnulados()
    Dim Cantidad As Integer
    Cantidad = ContarAnulados()
    StatusBar1.Panels(2).Text = "Anulados Total: " & Cantidad
End Sub


Private Sub MostrarCantidadAprobados()
    Dim Cantidad As Integer
    Cantidad = ContarAprobados()
    StatusBar1.Panels(3).Text = "Aprobados Total: " & Cantidad
End Sub


Private Sub MostrarFacturadosTotal()
    Dim Cantidad As Integer
    Cantidad = ContarFacturadosTotal()
    StatusBar1.Panels(4).Text = "Facturados Total: " & Cantidad
End Sub


Private Sub MostrarNoFacturados()
    Dim Cantidad As Integer
    Cantidad = ContarNoFacturados()
    StatusBar1.Panels(5).Text = "No Facturados: " & Cantidad
End Sub
Private Sub MostrarNoFacturables()
    Dim Cantidad As Integer
    Cantidad = ContarNoFacturables()
    StatusBar1.Panels(6).Text = "No Facturables: " & Cantidad
End Sub
Private Sub MostrarFacturadosParcial()
    Dim Cantidad As Integer
    Cantidad = ContarFacturadosParcial()
    StatusBar1.Panels(7).Text = "Facturados Parcial: " & Cantidad
End Sub

Private Function ContarAprobados() As Integer
    Dim rs As Recordset
    Dim strsql As String
    ContarAprobados = 0
    strsql = "select count(id) as cantidad from remitos where estado= " & EstadoRemito.RemitoAprobado
    Set rs = conectar.RSFactory(strsql)
    If Not rs.EOF And Not rs.BOF Then
        ContarAprobados = rs!Cantidad
    End If
End Function

Private Function ContarPendientes() As Integer
    Dim rs As Recordset
    Dim strsql As String
    ContarPendientes = 0
    strsql = "select count(id) as cantidad from remitos where estado= " & EstadoRemito.RemitoPendiente
    Set rs = conectar.RSFactory(strsql)
    If Not rs.EOF And Not rs.BOF Then
        ContarPendientes = rs!Cantidad
    End If
End Function

Private Function ContarAnulados() As Integer
    Dim rs As Recordset
    Dim strsql As String
    ContarAnulados = 0
    strsql = "select count(id) as cantidad from remitos where estado= " & EstadoRemito.RemitoAnulado
    Set rs = conectar.RSFactory(strsql)
    If Not rs.EOF And Not rs.BOF Then
        ContarAnulados = rs!Cantidad
    End If
End Function

Private Function ContarNoFacturados() As Integer
    Dim rs As Recordset
    Dim strsql As String
    ContarNoFacturados = 0
    strsql = "select count(id) as cantidad from remitos where estadoFacturado= " & EstadoRemitoFacturado.RemitoNoFacturado
    Set rs = conectar.RSFactory(strsql)
    If Not rs.EOF And Not rs.BOF Then
        ContarNoFacturados = rs!Cantidad
    End If
End Function

Private Function ContarNoFacturables() As Integer
    Dim rs As Recordset
    Dim strsql As String
    ContarNoFacturables = 0
    strsql = "select count(id) as cantidad from remitos where estadoFacturado= " & EstadoRemitoFacturado.RemitoNoFacturable
    Set rs = conectar.RSFactory(strsql)
    If Not rs.EOF And Not rs.BOF Then
        ContarNoFacturables = rs!Cantidad
    End If
End Function

Private Function ContarFacturadosParcial() As Integer
    Dim rs As Recordset
    Dim strsql As String
    ContarFacturadosParcial = 0
    strsql = "select count(id) as cantidad from remitos where estadoFacturado= " & EstadoRemitoFacturado.RemitoFacturadoParcial
    Set rs = conectar.RSFactory(strsql)
    If Not rs.EOF And Not rs.BOF Then
        ContarFacturadosParcial = rs!Cantidad
    End If
End Function

Private Function ContarFacturadosTotal() As Integer
    Dim rs As Recordset
    Dim strsql As String
    ContarFacturadosTotal = 0
    strsql = "select count(id) as cantidad from remitos where estadoFacturado= " & EstadoRemitoFacturado.RemitoFacturadoTotal
    Set rs = conectar.RSFactory(strsql)
    If Not rs.EOF And Not rs.BOF Then
        ContarFacturadosTotal = rs!Cantidad
    End If
End Function



Private Sub Form_Resize()
    On Error Resume Next
    Me.grilla.Width = Me.ScaleWidth - 180
    Me.grilla.Height = Me.ScaleHeight - 4000
    Me.GroupBox1.Width = Me.ScaleWidth - 180

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
            
            If tmpRto.estado = RemitoPendiente Then
                Me.mnuEditar.Enabled = True
                Me.mnuEditar.caption = "Editar..."
            ElseIf tmpRto.estado = RemitoAprobado Then
                Me.mnuEditar.caption = "Editar detalles..."
                Me.mnuEditar.Enabled = True
            End If
            
            
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


Private Sub grilla_UnboundReadData(ByVal RowIndex As Long, _
                                   ByVal Bookmark As Variant, _
                                   ByVal Values As GridEX20.JSRowData)

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
                .value(9) = Left$(facturasRemitos.item(CStr(tmpRto.numero)), Len(facturasRemitos.item(CStr(tmpRto.numero))) - 2)

            End If

        End If

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

    llenarContadoresStatusBar

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


Private Sub PusExportar_Click()
'FUNCIÓN PARA EXPORTAR A EXCEL


If (remitos.count > 0) Then


'INICIA EL PROGRESSBAR Y LO MUESTRA
Me.ProgressBar.Visible = True
'    Me.lblExportando.Visible = True


'DEFINE EL VALOR MINIMO Y EL MAXIMO DEL PROGRESSBAR (CANTIDAD DE DATOS EN LA COLECCIÓN COL)
Me.ProgressBar.min = 0
Me.ProgressBar.max = remitos.count


'Dim xlApplication As New Excel.Application
    Dim xlApplication As Object
    Set xlApplication = CreateObject("Excel.Application")

    'Dim xlWorkbook As New Excel.Workbook
    Dim xlWorkbook As Object
    Set xlWorkbook = CreateObject("Excel.Application")

    'Dim xlWorksheet As New Excel.Worksheet
    Dim xlWorksheet As Object
    Set xlWorksheet = CreateObject("Excel.Application")


    Set xlWorkbook = xlApplication.Workbooks.Add

    Set xlWorksheet = xlWorkbook.Worksheets.item(1)

    xlWorksheet.Activate

    xlWorksheet.Cells(1, 1).value = "Reporte de Remitos"

    xlWorksheet.Columns(4).HorizontalAlignment = xlLeft
    xlWorksheet.Columns(7).HorizontalAlignment = xlLeft

    xlWorksheet.Cells(2, 1).value = "Número"
    xlWorksheet.Cells(2, 2).value = "Cliente"
    xlWorksheet.Cells(2, 3).value = "Detalle"
    xlWorksheet.Cells(2, 4).value = "Fecha"
    xlWorksheet.Cells(2, 5).value = "Creador"
    xlWorksheet.Cells(2, 6).value = "Aprobador"
    xlWorksheet.Cells(2, 7).value = "Facturado"
    xlWorksheet.Cells(2, 8).value = "Facturas"
    xlWorksheet.Cells(2, 9).value = "Archivos"
    
    Dim idx As Integer
    idx = 3

    Dim Remito As Remito

    'DEFINE EL CONTADOR DEL PROGRESSBAR Y LO INICIA EN 0
    Dim d As Long
    d = 0


    For Each Remito In remitos

            xlWorksheet.Cells(idx, 1).value = Remito.numero
            
            xlWorksheet.Cells(idx, 2).value = Remito.cliente.razon
            xlWorksheet.Cells(idx, 3).value = Remito.detalle
            
            xlWorksheet.Cells(idx, 4).value = Remito.FEcha

            xlWorksheet.Cells(idx, 5).value = Remito.usuarioCreador.usuario
            
            If IsSomething(Remito.usuarioAprobador) Then
                xlWorksheet.Cells(idx, 6).value = Remito.usuarioAprobador.usuario
            Else
                xlWorksheet.Cells(idx, 6).value = vbNullString
            End If
            

            xlWorksheet.Cells(idx, 7).value = Remito.VerEstadoFacturado

            If facturasRemitos.Exists(CStr(Remito.numero)) Then
                If LenB(facturasRemitos.item(CStr(Remito.numero))) >= 0 Then
                    xlWorksheet.Cells(idx, 8).value = Left$(facturasRemitos.item(CStr(Remito.numero)), Len(facturasRemitos.item(CStr(Remito.numero))) - 2)
                End If
            End If

            xlWorksheet.Cells(idx, 9).value = "(" & Val(m_Archivos.item(Remito.Id)) & ")"

        
        idx = idx + 1

        'POR CADA ITERACION SUMA UN VALOR A LA VARIABLE D DEL PROGRESSBAR
        d = d + 1
        Me.ProgressBar.value = d


    Next

    'AUTOSIZE
    xlApplication.ScreenUpdating = False

    Dim wkSt As String

    wkSt = xlWorksheet.Name

    xlWorksheet.Cells.EntireColumn.AutoFit

    xlWorkbook.Sheets(wkSt).Select

    xlApplication.ScreenUpdating = True

    xlWorksheet.PageSetup.Orientation = xlLandscape
    xlWorksheet.PageSetup.BottomMargin = xlApplication.CentimetersToPoints(1)
    xlWorksheet.PageSetup.TopMargin = xlApplication.CentimetersToPoints(1)
    xlWorksheet.PageSetup.LeftMargin = xlApplication.CentimetersToPoints(1)
    xlWorksheet.PageSetup.RightMargin = xlApplication.CentimetersToPoints(1)

    Dim filename As String
    filename = funciones.GetTmpPath() & "tmp_info " & Hour(Now) & Minute(Now) & Second(Now) & " .xlsx"

    If Dir(filename) <> vbNullString Then Kill filename

    xlWorkbook.SaveAs filename

    xlWorkbook.Saved = True
    xlWorkbook.Close
    xlApplication.Quit

    funciones.ShellExecute 0, "open", filename, "", "", 0

    Set xlWorksheet = Nothing
    Set xlWorkbook = Nothing
    Set xlApplication = Nothing

    'REINICIA EL PROGRESSBAR Y LO OCULTA
    Me.ProgressBar.value = 0
    Me.ProgressBar.Visible = False
    '    Me.lblExportando.Visible = False
    
    Else
    MsgBox ("No hay resultados para exportar!")
    
    End If
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

    Me.WindowState = vbMaximized
    
    Dim frm As frmPlaneamientoRemitoVer
    Set frm = New frmPlaneamientoRemitoVer
    Set frm.Remito = tmpRto
    frm.MostrarInfoAdministracion = VerInfoAdministracion
    frm.Show
    

End Sub
