VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmPlaneamientoOELista 
   Caption         =   "Ordenes de entrega"
   ClientHeight    =   10620
   ClientLeft      =   60
   ClientTop       =   3120
   ClientWidth     =   11565
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10620
   ScaleMode       =   0  'User
   ScaleWidth      =   18735
   Begin GridEX20.GridEX gridEntregas 
      Height          =   4455
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   7858
      Version         =   "2.0"
      AllowRowSizing  =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      MethodHoldFields=   -1  'True
      AllowEdit       =   0   'False
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   7
      Column(1)       =   "frmListaOE.frx":0000
      Column(2)       =   "frmListaOE.frx":0110
      Column(3)       =   "frmListaOE.frx":01FC
      Column(4)       =   "frmListaOE.frx":02E8
      Column(5)       =   "frmListaOE.frx":03DC
      Column(6)       =   "frmListaOE.frx":04D0
      Column(7)       =   "frmListaOE.frx":05C0
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmListaOE.frx":06B4
      FormatStyle(2)  =   "frmListaOE.frx":07EC
      FormatStyle(3)  =   "frmListaOE.frx":089C
      FormatStyle(4)  =   "frmListaOE.frx":0950
      FormatStyle(5)  =   "frmListaOE.frx":0A28
      FormatStyle(6)  =   "frmListaOE.frx":0AE0
      ImageCount      =   0
      PrinterProperties=   "frmListaOE.frx":0BC0
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "Volver"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4680
      Width           =   1095
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1545
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11445
      _Version        =   786432
      _ExtentX        =   20188
      _ExtentY        =   2725
      _StockProps     =   79
      Caption         =   "Parámetros de búsqueda"
      BackColor       =   12632256
      UseVisualStyle  =   -1  'True
      Begin VB.TextBox txtNroRemito 
         Height          =   285
         Left            =   1425
         TabIndex        =   4
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   285
         Left            =   1425
         TabIndex        =   3
         Top             =   645
         Width           =   6855
      End
      Begin XtremeSuiteControls.ComboBox cboClientes 
         Height          =   315
         Left            =   4530
         TabIndex        =   5
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
         TabIndex        =   6
         Top             =   225
         Width           =   495
         _Version        =   786432
         _ExtentX        =   873
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "X"
         BackColor       =   12632256
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton cmdBuscar 
         Default         =   -1  'True
         Height          =   495
         Left            =   9960
         TabIndex        =   7
         Top             =   960
         Width           =   1335
         _Version        =   786432
         _ExtentX        =   2355
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Buscar"
         BackColor       =   12632256
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
      Begin XtremeSuiteControls.DateTimePicker dtpDesde 
         Height          =   315
         Left            =   5730
         TabIndex        =   8
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
         TabIndex        =   9
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
         TabIndex        =   10
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
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Número"
         Height          =   255
         Left            =   270
         TabIndex        =   16
         Top             =   270
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   645
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         Height          =   255
         Left            =   3315
         TabIndex        =   14
         Top             =   240
         Width           =   1095
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
         BackColor       =   12632256
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label5 
         Height          =   195
         Left            =   5160
         TabIndex        =   12
         Top             =   1110
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
         Left            =   795
         TabIndex        =   11
         Top             =   1110
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
   Begin VB.Menu entrega 
      Caption         =   "entregas"
      Visible         =   0   'False
      Begin VB.Menu OENumero 
         Caption         =   "Numero"
         Enabled         =   0   'False
      End
      Begin VB.Menu vereditar 
         Caption         =   "vereditar"
      End
      Begin VB.Menu as234 
         Caption         =   "-"
      End
      Begin VB.Menu AprobarOE 
         Caption         =   "Aprobar"
      End
      Begin VB.Menu remitar 
         Caption         =   "Remitar..."
      End
      Begin VB.Menu cerrarOE 
         Caption         =   "Cerrar..."
      End
      Begin VB.Menu verHistorialOE 
         Caption         =   "Ver historial..."
      End
      Begin VB.Menu nada 
         Caption         =   "-"
      End
      Begin VB.Menu RtosEntregados 
         Caption         =   "Remitos entregados...."
      End
      Begin VB.Menu printOrder 
         Caption         =   "Imprimir..."
      End
   End
End
Attribute VB_Name = "frmPlaneamientoOELista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim tmpOe As OrdenDeEntrega
Dim ordenes As New Collection

Private mOESeleccionada As OrdenDeEntrega
Private mModoVerEditar As Integer

Private Const MODO_EDITAR As Integer = 1
Private Const MODO_VER As Integer = 2


Private Sub cboRangos_Click()

    funciones.CalculateDateRange _
        Me.cboRangos, _
        Me.dtpDesde, _
        Me.dtpHasta

End Sub


Public Sub LlenarListaOE(Optional ByVal filtro As String = vbNullString)

    On Error GoTo errHandler

    Set ordenes = DAOOrdenDeEntrega.GetAll(filtro)

    Me.gridEntregas.ItemCount = 0

    If Not ordenes Is Nothing Then
        Me.gridEntregas.ItemCount = ordenes.count
    End If

    Exit Sub

errHandler:

    MsgBox "Error al cargar las Ordenes de Entrega." & vbCrLf & _
           "Error " & Err.Number & vbCrLf & _
           Err.Description & vbCrLf & _
           "Origen: " & Err.Source, _
           vbCritical, "Ordenes de Entrega"

End Sub


Private Sub AprobarOE_Click()
'Dim vidOe As Long
'vidOe = CLng(Me.lstOE.selectedItem)
'    If MsgBox("¿Está seguro de aprobar la O/E?", vbYesNo, "Confirmación") = vbYes Then
'        If claseP.AprobarOrdenEntrega(vidOe) Then
'            MsgBox "Orden de Entrega aprobada con éxito!", vbInformation, "Información"
'        Else
'            MsgBox "Se produjo un error al aprobar la OE!", vbCritical, "Error"
'        End If
'    End If

End Sub


Private Sub cmdBuscar_Click()

    On Error GoTo errHandler

    Dim filtro As String
    Dim texto As String
    Dim idCliente As Long

    filtro = vbNullString


    '--------------------------------------------------
    ' NUMERO DE ORDEN DE ENTREGA
    '--------------------------------------------------
    texto = Trim$(Me.txtNroRemito.Text)

    If Len(texto) > 0 Then

        If Not IsNumeric(texto) Then
            MsgBox "El número de Orden de Entrega debe ser numérico.", _
                   vbExclamation, "Ordenes de Entrega"

            Me.txtNroRemito.SetFocus
            Exit Sub
        End If

        AgregarFiltro filtro, _
            "pe.id = " & CLng(texto)

    End If


    '--------------------------------------------------
    ' DESCRIPCION / REFERENCIA
    '--------------------------------------------------
    texto = Trim$(Me.txtDescripcion.Text)

    If Len(texto) > 0 Then

        AgregarFiltro filtro, _
            "pe.referencia LIKE " & _
            conectar.Escape("%" & texto & "%")

    End If


    '--------------------------------------------------
    ' CLIENTE
    '--------------------------------------------------
    If Me.cboClientes.ListIndex >= 0 Then

        idCliente = Me.cboClientes.ItemData(Me.cboClientes.ListIndex)

        If idCliente > 0 Then

            AgregarFiltro filtro, _
                "pe.IdCliente = " & idCliente

        End If

    End If


    '--------------------------------------------------
    ' FECHA DESDE
    '--------------------------------------------------
    If Not IsNull(Me.dtpDesde.value) Then

        AgregarFiltro filtro, _
            "pe.fecha >= " & _
            conectar.Escape( _
                Format$(Me.dtpDesde.value, _
                        "yyyy-mm-dd 00:00:00"))

    End If


    '--------------------------------------------------
    ' FECHA HASTA
    '--------------------------------------------------
    If Not IsNull(Me.dtpHasta.value) Then

        AgregarFiltro filtro, _
            "pe.fecha <= " & _
            conectar.Escape( _
                Format$(Me.dtpHasta.value, _
                        "yyyy-mm-dd 23:59:59"))

    End If


    '--------------------------------------------------
    ' BUSCAR
    '--------------------------------------------------
    LlenarListaOE filtro

    Exit Sub


errHandler:

    MsgBox "Error al buscar Ordenes de Entrega." & vbCrLf & _
           "Error " & Err.Number & vbCrLf & _
           Err.Description, _
           vbCritical, "Ordenes de Entrega"

End Sub


Private Sub Command1_Click()
    Unload Me
End Sub


Private Sub Form_Load()

    On Error GoTo errHandler

    FormHelper.Customize Me

    GridEXHelper.CustomizeGrid Me.gridEntregas, True

    'Clientes
    DAOCliente.llenarComboXtremeSuite Me.cboClientes, True

    'Rangos de fecha
    Dim i As Integer

    funciones.FillComboBoxDateRanges Me.cboRangos

    For i = 0 To Me.cboRangos.ListCount - 1
        If Me.cboRangos.ItemData(i) = DateRangeValue.DRV_YearCurrent Then
            Exit For
        End If
    Next i

    If i < Me.cboRangos.ListCount Then
        Me.cboRangos.ListIndex = i
    End If

    'Carga de Ordenes de Entrega
    LlenarListaOE

    'NO hacer AutoSize acá.
    'GridEXHelper.AutoSizeColumns Me.gridEntregas, True

    Exit Sub

errHandler:

    MsgBox "Error al inicializar el listado de Ordenes de Entrega." & _
           vbCrLf & _
           "Error " & Err.Number & vbCrLf & _
           Err.Description & vbCrLf & _
           "Origen: " & Err.Source, _
           vbCritical, "Ordenes de Entrega"

End Sub


Private Sub gridEntregas_MouseUp( _
    Button As Integer, _
    Shift As Integer, _
    x As Single, _
    y As Single)

    On Error GoTo errHandler

    If Button <> 2 Then Exit Sub

    If Me.gridEntregas.ItemCount = 0 Then Exit Sub

    Set mOESeleccionada = ObtenerOESeleccionada()

    If mOESeleccionada Is Nothing Then Exit Sub

    ConfigurarMenuContextualOE

    Me.PopupMenu Me.entrega

    Exit Sub


errHandler:

    MsgBox "Error al abrir el menú de la Orden de Entrega." & _
           vbCrLf & _
           "Error " & Err.Number & vbCrLf & _
           Err.Description, _
           vbCritical, "Ordenes de Entrega"

End Sub


Private Sub gridEntregas_UnboundReadData( _
    ByVal RowIndex As Long, _
    ByVal Bookmark As Variant, _
    ByVal Values As GridEX20.JSRowData)

    On Error GoTo errHandler

    If RowIndex <= 0 Then Exit Sub
    If ordenes Is Nothing Then Exit Sub
    If RowIndex > ordenes.count Then Exit Sub

    Set tmpOe = ordenes.item(RowIndex)

    If tmpOe Is Nothing Then Exit Sub

    With Values

        .value(1) = tmpOe.Id
        .value(2) = tmpOe.FEcha

        'Cliente
        If Not tmpOe.Cliente Is Nothing Then
            .value(3) = tmpOe.Cliente.razon
        Else
            .value(3) = ""
        End If

        .value(4) = tmpOe.referencia

        'Usuario creador
        If Not tmpOe.usuarioCreador Is Nothing Then
            .value(5) = tmpOe.usuarioCreador.Usuario
        Else
            .value(5) = ""
        End If

        'Usuario aprobador
        If Not tmpOe.usuarioAprobador Is Nothing Then
            .value(6) = tmpOe.usuarioAprobador.Usuario
        Else
            .value(6) = ""
        End If

        .value(7) = enumEstadoOrdenEntrega(tmpOe.estado)

    End With

    Exit Sub

errHandler:

    Debug.Print "gridEntregas_UnboundReadData - Fila: " & RowIndex & _
                " - Error " & Err.Number & _
                " - " & Err.Description

End Sub

Private Sub printOrder_Click()
' If Me.lstOE.ListItems.count > 0 Then
'claseP.imprimirOrdenEntrega (CLng(Me.lstOE.selectedItem))
'  End If


End Sub

Private Sub PushButton1_Click()

    On Error GoTo errHandler

    'Dejar cliente en "Todos"
    If Me.cboClientes.ListCount > 0 Then
        Me.cboClientes.ListIndex = 0
    End If

    'Volver a ejecutar búsqueda con los demás filtros
    cmdBuscar_Click

    Exit Sub

errHandler:

    MsgBox "Error al limpiar el cliente." & vbCrLf & _
           Err.Description, _
           vbCritical, "Ordenes de Entrega"

End Sub

Private Sub remitar_Click()
'If Me.lstOE.ListItems.count > 0 Then
'frmRemitar.idPedidoEntrega = CLng(Me.lstOE.selectedItem)
'        frmRemitar.idPe = CLng(Me.lstOE.selectedItem)
'        frmRemitar.Frame1.caption = "[ Nro." & Me.lstOE.selectedItem & " ]"
'        frmRemitar.Show
' End If
End Sub


Private Sub RtosEntregados_Click()
'  If Me.lstOE.ListItems.count > 0 Then
    frmRemitosEntregados.Origen = 2
    'frmRemitosEntregados.idPedidoEntrega = Me.lstOE.selectedItem
    'frmRemitosEntregados.caption = "Nro." & Me.lstOE.selectedItem
    '        frmRemitosEntregados.Show
    ' End If
End Sub


Private Sub vereditar_Click()

    On Error GoTo errHandler

    If mOESeleccionada Is Nothing Then

        MsgBox "No hay una Orden de Entrega seleccionada.", _
               vbExclamation, _
               "Ordenes de Entrega"

        Exit Sub

    End If


    Select Case mModoVerEditar


        Case MODO_EDITAR

            frmPlaneamientoOEEditar.IDOE = mOESeleccionada.Id
            frmPlaneamientoOEEditar.Show


        Case MODO_VER

            frmPlaneamientoOEVer.IDOE = mOESeleccionada.Id
            frmPlaneamientoOEVer.Show


        Case Else

            MsgBox "No se pudo determinar la acción para la O/E.", _
                   vbExclamation, _
                   "Ordenes de Entrega"

    End Select


    Exit Sub


errHandler:

    MsgBox "Error al abrir la Orden de Entrega Nro. " & _
           CStr(mOESeleccionada.Id) & "." & vbCrLf & _
           Err.Description, _
           vbCritical, "Ordenes de Entrega"

End Sub


Private Sub AgregarFiltro(ByRef filtro As String, ByVal condicion As String)

    If Len(Trim$(condicion)) = 0 Then Exit Sub

    If Len(Trim$(filtro)) > 0 Then
        filtro = filtro & " AND "
    End If

    filtro = filtro & condicion

End Sub


Private Function ObtenerOESeleccionada() As OrdenDeEntrega

    On Error GoTo errHandler

    Dim indice As Long

    Set ObtenerOESeleccionada = Nothing

    If Me.gridEntregas.ItemCount = 0 Then Exit Function
    If ordenes Is Nothing Then Exit Function

    indice = Me.gridEntregas.RowIndex(Me.gridEntregas.row)

    If indice <= 0 Then Exit Function
    If indice > ordenes.count Then Exit Function

    Set ObtenerOESeleccionada = ordenes.item(indice)

    Exit Function

errHandler:

    Set ObtenerOESeleccionada = Nothing

End Function


Private Sub ConfigurarMenuContextualOE()

    On Error GoTo errHandler

    If mOESeleccionada Is Nothing Then Exit Sub

    '-----------------------------------------
    ' Cabecera
    '-----------------------------------------
    Me.OENumero.caption = _
        "[ O/E Nro. " & CStr(mOESeleccionada.Id) & " ]"


    '-----------------------------------------
    ' Reset general
    '-----------------------------------------
    Me.vereditar.Enabled = False
    Me.AprobarOE.Enabled = False
    Me.remitar.Enabled = False
    Me.cerrarOE.Enabled = False
    Me.verHistorialOE.Enabled = False
    Me.RtosEntregados.Enabled = False
    Me.printOrder.Enabled = False

    mModoVerEditar = 0


    '-----------------------------------------
    ' Estado de la OE
    '-----------------------------------------
    Select Case mOESeleccionada.estado


        '=====================================
        ' PENDIENTE
        '=====================================
        Case EstadoOrdenEntrega.Pendiente

            Me.vereditar.caption = "Editar..."
            Me.vereditar.Enabled = True
            mModoVerEditar = MODO_EDITAR

            If Permisos.planOEaprobaciones Then
                Me.AprobarOE.Enabled = True
            End If

            Me.verHistorialOE.Enabled = True


        '=====================================
        ' APROBADA
        '=====================================
        Case EstadoOrdenEntrega.Aprobado

            Me.vereditar.caption = "Ver..."
            Me.vereditar.Enabled = True
            mModoVerEditar = MODO_VER

            Me.remitar.Enabled = True
            Me.RtosEntregados.Enabled = True
            Me.printOrder.Enabled = True
            Me.verHistorialOE.Enabled = True


        '=====================================
        ' FINALIZADA
        '=====================================
        Case EstadoOrdenEntrega.FINALIZADO

            Me.vereditar.caption = "Ver..."
            Me.vereditar.Enabled = True
            mModoVerEditar = MODO_VER

            Me.RtosEntregados.Enabled = True
            Me.printOrder.Enabled = True
            Me.verHistorialOE.Enabled = True


        '=====================================
        ' ESTADO DESCONOCIDO
        '=====================================
        Case Else

            Me.vereditar.caption = "Ver..."
            Me.vereditar.Enabled = True
            mModoVerEditar = MODO_VER

    End Select

    Exit Sub


errHandler:

    MsgBox "Error al configurar el menú de la Orden de Entrega." & _
           vbCrLf & _
           "Error " & Err.Number & vbCrLf & _
           Err.Description, _
           vbCritical, "Ordenes de Entrega"

End Sub

