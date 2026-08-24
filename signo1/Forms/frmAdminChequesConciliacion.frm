VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmAdminChequesConciliacion 
   Caption         =   "Conciliación de Cheques"
   ClientHeight    =   10425
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14550
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10425
   ScaleWidth      =   14550
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.GroupBox GroupBox4 
      Height          =   2700
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   18495
      _Version        =   786432
      _ExtentX        =   32623
      _ExtentY        =   4762
      _StockProps     =   79
      Caption         =   "Parámetros de búsqueda"
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
      Begin VB.CheckBox chkMostrarIngresados 
         Caption         =   "Mostrar cheques ya ingresados"
         Height          =   495
         Left            =   120
         TabIndex        =   23
         Top             =   840
         Width           =   2775
      End
      Begin VB.TextBox TxtNumeroChequeEnChequera 
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   2535
      End
      Begin XtremeSuiteControls.GroupBox GroupBox5 
         Height          =   2415
         Left            =   8160
         TabIndex        =   1
         Top             =   120
         Width           =   4095
         _Version        =   786432
         _ExtentX        =   7223
         _ExtentY        =   4260
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.PushButton btnExportarEnChequera 
            Height          =   495
            Index           =   0
            Left            =   2160
            TabIndex        =   2
            Top             =   1800
            Width           =   1815
            _Version        =   786432
            _ExtentX        =   3201
            _ExtentY        =   873
            _StockProps     =   79
            Caption         =   "Exportar"
            Enabled         =   0   'False
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton btnBuscarEnChequera 
            Height          =   495
            Index           =   1
            Left            =   120
            TabIndex        =   3
            Top             =   1800
            Width           =   1815
            _Version        =   786432
            _ExtentX        =   3201
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
      End
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   315
         Left            =   2760
         TabIndex        =   4
         Top             =   480
         Width           =   375
         _Version        =   786432
         _ExtentX        =   661
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "X"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.GroupBox GroFechaComprobante 
         Height          =   1215
         Index           =   7
         Left            =   3360
         TabIndex        =   6
         Top             =   120
         Width           =   4695
         _Version        =   786432
         _ExtentX        =   8281
         _ExtentY        =   2143
         _StockProps     =   79
         Caption         =   "Fecha Vencimiento"
         BackColor       =   16744576
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   4
         Begin XtremeSuiteControls.ComboBox cboRangosVtoTerceros 
            Height          =   315
            Index           =   0
            Left            =   720
            TabIndex        =   7
            Top             =   300
            Width           =   3675
            _Version        =   786432
            _ExtentX        =   6482
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Style           =   2
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.DateTimePicker dtpDesdeVtoTerceros 
            Height          =   315
            Index           =   0
            Left            =   720
            TabIndex        =   8
            Top             =   720
            Width           =   1470
            _Version        =   786432
            _ExtentX        =   2593
            _ExtentY        =   556
            _StockProps     =   68
            CheckBox        =   -1  'True
            Format          =   1
            CurrentDate     =   45190.4376157407
         End
         Begin XtremeSuiteControls.DateTimePicker dtpHastaVtoTerceros 
            Height          =   315
            Index           =   0
            Left            =   2925
            TabIndex        =   9
            Top             =   720
            Width           =   1470
            _Version        =   786432
            _ExtentX        =   2593
            _ExtentY        =   556
            _StockProps     =   68
            CheckBox        =   -1  'True
            Format          =   1
            CurrentDate     =   45190.4375810185
         End
         Begin XtremeSuiteControls.Label lblRango 
            Height          =   195
            Index           =   7
            Left            =   120
            TabIndex        =   12
            Top             =   360
            Width           =   480
            _Version        =   786432
            _ExtentX        =   847
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Rango"
            BackColor       =   12632256
            AutoSize        =   -1  'True
         End
         Begin XtremeSuiteControls.Label lblDesde 
            Height          =   195
            Index           =   7
            Left            =   165
            TabIndex        =   11
            Top             =   780
            Width           =   465
            _Version        =   786432
            _ExtentX        =   820
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Desde"
            BackColor       =   12632256
            AutoSize        =   -1  'True
         End
         Begin XtremeSuiteControls.Label lblHasta 
            Height          =   195
            Index           =   7
            Left            =   2400
            TabIndex        =   10
            Top             =   780
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
      Begin XtremeSuiteControls.GroupBox GroFechaComprobante 
         Height          =   1215
         Index           =   8
         Left            =   3360
         TabIndex        =   13
         Top             =   1320
         Width           =   4695
         _Version        =   786432
         _ExtentX        =   8281
         _ExtentY        =   2143
         _StockProps     =   79
         Caption         =   "Fecha Emitido"
         BackColor       =   16744576
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   4
         Begin XtremeSuiteControls.ComboBox cboRangosRboEmitido 
            Height          =   315
            Index           =   0
            Left            =   720
            TabIndex        =   14
            Top             =   300
            Width           =   3675
            _Version        =   786432
            _ExtentX        =   6482
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Style           =   2
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.DateTimePicker dtpHastaRboEmitido 
            Height          =   315
            Index           =   0
            Left            =   2925
            TabIndex        =   15
            Top             =   720
            Width           =   1470
            _Version        =   786432
            _ExtentX        =   2593
            _ExtentY        =   556
            _StockProps     =   68
            CheckBox        =   -1  'True
            Format          =   1
            CurrentDate     =   45190.4375578704
         End
         Begin XtremeSuiteControls.DateTimePicker dtpDesdeRboEmitido 
            Height          =   315
            Index           =   0
            Left            =   720
            TabIndex        =   16
            Top             =   720
            Width           =   1470
            _Version        =   786432
            _ExtentX        =   2593
            _ExtentY        =   556
            _StockProps     =   68
            CheckBox        =   -1  'True
            Format          =   1
            CurrentDate     =   45190.4375231481
         End
         Begin XtremeSuiteControls.Label lblRango 
            Height          =   195
            Index           =   8
            Left            =   120
            TabIndex        =   19
            Top             =   360
            Width           =   480
            _Version        =   786432
            _ExtentX        =   847
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Rango"
            BackColor       =   12632256
            AutoSize        =   -1  'True
         End
         Begin XtremeSuiteControls.Label lblDesde 
            Height          =   195
            Index           =   8
            Left            =   165
            TabIndex        =   18
            Top             =   780
            Width           =   465
            _Version        =   786432
            _ExtentX        =   820
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Desde"
            BackColor       =   12632256
            AutoSize        =   -1  'True
         End
         Begin XtremeSuiteControls.Label lblHasta 
            Height          =   195
            Index           =   8
            Left            =   2400
            TabIndex        =   17
            Top             =   780
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
      Begin VB.Label Label6 
         Caption         =   "Número:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   2535
      End
   End
   Begin GridEX20.GridEX grid_cheques 
      Height          =   6015
      Left            =   120
      TabIndex        =   21
      Top             =   3120
      Width           =   18645
      _ExtentX        =   32888
      _ExtentY        =   10610
      Version         =   "2.0"
      PreviewRowIndent=   200
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      EmptyRows       =   -1  'True
      PreviewColumn   =   "destino_campo_origen"
      PreviewRowLines =   1
      ColumnAutoResize=   -1  'True
      MethodHoldFields=   -1  'True
      RowHeaders      =   -1  'True
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   11
      Column(1)       =   "frmAdminChequesConciliacion.frx":0000
      Column(2)       =   "frmAdminChequesConciliacion.frx":0164
      Column(3)       =   "frmAdminChequesConciliacion.frx":02BC
      Column(4)       =   "frmAdminChequesConciliacion.frx":040C
      Column(5)       =   "frmAdminChequesConciliacion.frx":0554
      Column(6)       =   "frmAdminChequesConciliacion.frx":0694
      Column(7)       =   "frmAdminChequesConciliacion.frx":07F8
      Column(8)       =   "frmAdminChequesConciliacion.frx":094C
      Column(9)       =   "frmAdminChequesConciliacion.frx":0A88
      Column(10)      =   "frmAdminChequesConciliacion.frx":0B48
      Column(11)      =   "frmAdminChequesConciliacion.frx":0CDC
      FormatStylesCount=   7
      FormatStyle(1)  =   "frmAdminChequesConciliacion.frx":0E5C
      FormatStyle(2)  =   "frmAdminChequesConciliacion.frx":0F94
      FormatStyle(3)  =   "frmAdminChequesConciliacion.frx":1044
      FormatStyle(4)  =   "frmAdminChequesConciliacion.frx":10F8
      FormatStyle(5)  =   "frmAdminChequesConciliacion.frx":11D0
      FormatStyle(6)  =   "frmAdminChequesConciliacion.frx":1288
      FormatStyle(7)  =   "frmAdminChequesConciliacion.frx":1368
      ImageCount      =   0
      PrinterProperties=   "frmAdminChequesConciliacion.frx":1424
   End
   Begin VB.Label Label13 
      Caption         =   "Label13"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   2835
      Width           =   8055
   End
End
Attribute VB_Name = "frmAdminChequesConciliacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private chequesConciliacion As Collection
Private cargandoDatos As Boolean


Private Sub cmdExportar_Click()

End Sub

Private Sub Form_Load()

    On Error GoTo err1

    cargandoDatos = True

    FormHelper.Customize Me
    GridEXHelper.CustomizeGrid Me.grid_cheques, True, True

    funciones.FillComboBoxDateRanges _
        Me.cboRangosVtoTerceros(0)

    funciones.FillComboBoxDateRanges _
        Me.cboRangosRboEmitido(0)

    Me.cboRangosVtoTerceros(0).ListIndex = -1
    Me.cboRangosRboEmitido(0).ListIndex = -1

    'Al abrir no se aplican filtros de fechas
    Me.dtpDesdeVtoTerceros(0).value = Null
    Me.dtpHastaVtoTerceros(0).value = Null
    Me.dtpDesdeRboEmitido(0).value = Null
    Me.dtpHastaRboEmitido(0).value = Null

    'Por defecto se muestran solamente los pendientes
    Me.chkMostrarIngresados.value = vbUnchecked

    cargandoDatos = False

    CargarCheques

    Exit Sub

err1:
    cargandoDatos = False

    MsgBox "No se pudo abrir la conciliación de cheques." & _
           vbCrLf & Err.Description, _
           vbExclamation, _
           "Conciliación de cheques"

End Sub


Private Sub CargarCheques( _
    Optional ByVal AvisarChequeNoEncontrado As Boolean = False)

    On Error GoTo err1

    If cargandoDatos Then Exit Sub

    cargandoDatos = True

    Dim filter As String
    Dim numeroBuscado As String
    Dim mostrarIngresados As Boolean

    filter = "1 = 1"

    numeroBuscado = _
        Trim$(Me.TxtNumeroChequeEnChequera.Text)

    If LenB(numeroBuscado) > 0 Then

        filter = filter & _
                 " AND cheq.numero = " & _
                 conectar.Escape(numeroBuscado)

    End If

    'Fecha de vencimiento desde
    If Not IsNull(Me.dtpDesdeVtoTerceros(0).value) Then

        filter = filter & _
                 " AND cheq.fecha_vencimiento >= " & _
                 conectar.Escape( _
                    Format$( _
                        Me.dtpDesdeVtoTerceros(0).value, _
                        "yyyy-mm-dd"))

    End If

    'Fecha de vencimiento hasta
    If Not IsNull(Me.dtpHastaVtoTerceros(0).value) Then

        filter = filter & _
                 " AND cheq.fecha_vencimiento <= " & _
                 conectar.Escape( _
                    Format$( _
                        Me.dtpHastaVtoTerceros(0).value, _
                        "yyyy-mm-dd"))

    End If

    'Fecha de emisión desde
    If Not IsNull(Me.dtpDesdeRboEmitido(0).value) Then

        filter = filter & _
                 " AND cheq.fecha_emision >= " & _
                 conectar.Escape( _
                    Format$( _
                        Me.dtpDesdeRboEmitido(0).value, _
                        "yyyy-mm-dd"))

    End If

    'Fecha de emisión hasta
    If Not IsNull(Me.dtpHastaRboEmitido(0).value) Then

        filter = filter & _
                 " AND cheq.fecha_emision <= " & _
                 conectar.Escape( _
                    Format$( _
                        Me.dtpHastaRboEmitido(0).value, _
                        "yyyy-mm-dd"))

    End If

    mostrarIngresados = _
        (Me.chkMostrarIngresados.value = vbChecked)

    Set chequesConciliacion = _
        DAOCheques.FindAllPropiosConciliacion( _
            filter, _
            mostrarIngresados)

    If chequesConciliacion Is Nothing Then
        Set chequesConciliacion = New Collection
    End If

    Me.grid_cheques.ItemCount = 0
    Me.grid_cheques.ItemCount = chequesConciliacion.count
    Me.grid_cheques.Refresh

    Me.Label13.caption = _
        "Cheques mostrados: [ " & _
        chequesConciliacion.count & " ]"

    If AvisarChequeNoEncontrado Then

        If LenB(numeroBuscado) > 0 And _
           chequesConciliacion.count = 0 Then

            MsgBox "Cheque no encontrado.", _
                   vbInformation, _
                   "Conciliación de cheques"

        End If

    End If

salir:
    cargandoDatos = False
    Exit Sub

err1:
    Me.grid_cheques.ItemCount = 0
    Me.Label13.caption = "Cheques mostrados: [ 0 ]"

    MsgBox "No se pudieron cargar los cheques." & _
           vbCrLf & Err.Description, _
           vbExclamation, _
           "Conciliación de cheques"

    Resume salir

End Sub


Private Sub grid_cheques_UnboundReadData( _
    ByVal RowIndex As Long, _
    ByVal Bookmark As Variant, _
    ByVal Values As GridEX20.JSRowData)

    On Error GoTo err1

    If chequesConciliacion Is Nothing Then Exit Sub

    If RowIndex < 1 Or _
       RowIndex > chequesConciliacion.count Then
        Exit Sub
    End If

    Dim ch As cheque

    Set ch = chequesConciliacion.item(RowIndex)

    With Values

        '1 - Banco
        If Not ch.chequera Is Nothing Then
            If Not ch.chequera.Banco Is Nothing Then
                .value(1) = ch.chequera.Banco.nombre
            Else
                .value(1) = "SIN BANCO"
            End If
        End If

        '2 - Cuenta bancaria
        If Not ch.chequera Is Nothing Then
            If Not ch.chequera.CuentaBancaria Is Nothing Then
                .value(2) = ch.chequera.CuentaBancaria.numero
            Else
                .value(2) = "SIN CUENTA ASIGNADA"
            End If
        End If

        '3 - Chequera
        If Not ch.chequera Is Nothing Then
            .value(3) = ch.chequera.numero
        End If

        '4 - Número
        .value(4) = ch.numero

        '5 - Monto
        If ch.Utilizado Then
            .value(5) = funciones.FormatearDecimales(ch.Monto)
        Else
            .value(5) = Empty
        End If

        '6 - Vencimiento
        If CDbl(ch.FechaVencimiento) > 0 Then
            .value(6) = ch.FechaVencimiento
        Else
            .value(6) = Empty
        End If

        '7 - Emisión
        If CDbl(ch.FechaEmision) > 0 Then
            .value(7) = ch.FechaEmision
        Else
            .value(7) = Empty
        End If

        '8 - Destino y utilización
        .value(8) = DescripcionUsoCheque(ch)

        '9 - ID oculto
        .value(9) = ch.Id

        '10 - Ingresado
        .value(10) = Abs(CInt(ch.entro))

        '11 - Fecha de ingreso
        If CDbl(ch.FechaIngresoBanco) > 0 Then
            .value(11) = ch.FechaIngresoBanco
        Else
            .value(11) = Empty
        End If

    End With

    Exit Sub

err1:
    Debug.Print "grid_cheques_UnboundReadData: " & _
                Err.Number & " - " & Err.Description

End Sub


Private Sub grid_cheques_UnboundUpdate( _
    ByVal RowIndex As Long, _
    ByVal Bookmark As Variant, _
    ByVal Values As GridEX20.JSRowData)

    On Error GoTo err1

    If cargandoDatos Then Exit Sub
    If chequesConciliacion Is Nothing Then Exit Sub
    If Not IsNumeric(Values(9)) Then Exit Sub

    Dim ch As cheque
    Dim idCheque As Long
    Dim nuevoIngresado As Boolean
    Dim nuevaFecha As Date
    Dim cambioEstado As Boolean
    Dim cambioFecha As Boolean
    Dim respuesta As VbMsgBoxResult
    Dim mensaje As String

    idCheque = CLng(Values(9))

    Set ch = BuscarChequePorId(idCheque)

    If ch Is Nothing Then Exit Sub

    If IsNull(Values(10)) Or _
       IsEmpty(Values(10)) Then

        nuevoIngresado = False

    Else
        nuevoIngresado = CBool(Values(10))
    End If

    If nuevoIngresado Then

        If IsDate(Values(11)) Then
            nuevaFecha = CDate(Values(11))
        Else
            nuevaFecha = Date
        End If

    Else
        nuevaFecha = 0
    End If

    'Solamente se impide marcar un cheque disponible.
    'Un cheque histórico ya marcado sí puede modificar su fecha.
    If nuevoIngresado And _
       Not ch.entro And _
       Not ch.Utilizado Then

        MsgBox "El cheque N° " & ch.numero & _
               " todavía no fue utilizado." & vbCrLf & _
               "No se puede registrar su ingreso al banco.", _
               vbExclamation, _
               "Conciliación de cheques"

        CargarCheques
        Exit Sub

    End If

    If nuevoIngresado Then

        If nuevaFecha > Date Then

            MsgBox "La fecha de ingreso no puede ser posterior a hoy.", _
                   vbExclamation, _
                   "Fecha incorrecta"

            CargarCheques
            Exit Sub

        End If

    End If

    cambioEstado = (nuevoIngresado <> ch.entro)

    cambioFecha = _
        (Int(CDbl(nuevaFecha)) <> _
         Int(CDbl(ch.FechaIngresoBanco)))

    If Not cambioEstado And Not cambioFecha Then
        Exit Sub
    End If

    If nuevoIngresado And Not ch.entro Then

        mensaje = _
            "¿Confirma que el cheque N° " & ch.numero & _
            " ingresó al banco?" & vbCrLf & vbCrLf & _
            "Fecha de ingreso: " & _
            Format$(nuevaFecha, "dd/mm/yyyy")

    ElseIf Not nuevoIngresado And ch.entro Then

        mensaje = _
            "¿Confirma que desea quitar la marca de ingresado " & _
            "del cheque N° " & ch.numero & "?" & _
            vbCrLf & vbCrLf & _
            "También se eliminará la fecha de ingreso."

    Else

        mensaje = _
            "¿Confirma cambiar la fecha de ingreso del cheque N° " & _
            ch.numero & "?" & vbCrLf & vbCrLf & _
            "Nueva fecha: " & _
            Format$(nuevaFecha, "dd/mm/yyyy")

    End If

    respuesta = MsgBox( _
                    mensaje, _
                    vbQuestion + vbYesNo + vbDefaultButton2, _
                    "Confirmar modificación")

    If respuesta <> vbYes Then
        CargarCheques
        Exit Sub
    End If

    If Not DAOCheques.ActualizarIngresoBanco( _
                ch.Id, _
                nuevoIngresado, _
                nuevaFecha) Then

        Err.Raise vbObjectError + 1001, _
                  "grid_cheques_UnboundUpdate", _
                  "No se pudo actualizar el cheque."

    End If

    ch.entro = nuevoIngresado
    ch.FechaIngresoBanco = nuevaFecha

    'Si no se muestran ingresados, el cheque desaparecerá
    'inmediatamente después de ser marcado.
    CargarCheques

    Exit Sub

err1:
    MsgBox "No se pudo guardar el ingreso del cheque." & _
           vbCrLf & Err.Description, _
           vbExclamation, _
           "Conciliación de cheques"

    CargarCheques

End Sub


Private Function BuscarChequePorId( _
    ByVal idCheque As Long) As cheque

    Dim ch As cheque

    If chequesConciliacion Is Nothing Then Exit Function

    For Each ch In chequesConciliacion

        If ch.Id = idCheque Then
            Set BuscarChequePorId = ch
            Exit Function
        End If

    Next ch

End Function


Private Function DescripcionUsoCheque( _
    ByVal ch As cheque) As String

    Dim uso As String
    Dim destino As String
    Dim numeroLiquidacion As Long

    If ch Is Nothing Then Exit Function

    destino = Trim$(ch.OrigenDestino)

    If ch.estado = ChequeAnulado Then

        uso = "ANULADO"

    ElseIf ch.NumeroMovimiento > 0 Then

        uso = "Movimiento de Caja y Bancos N° " & _
              ch.NumeroMovimiento

    ElseIf ch.NumeroPagoACuenta > 0 Then

        uso = "Pago a Cuenta N° " & _
              ch.NumeroPagoACuenta

    ElseIf ch.NumeroLiquidacionCaja > 0 Or _
           ch.IdLiquidacionCajaOrigen > 0 Then

        numeroLiquidacion = ch.NumeroLiquidacionCaja

        If numeroLiquidacion = 0 Then
            numeroLiquidacion = ch.IdLiquidacionCajaOrigen
        End If

        uso = "Liquidación de Caja N° " & _
              numeroLiquidacion

    ElseIf ch.IdOrdenPagoOrigen > 0 Then

        uso = "Orden de Pago N° " & _
              ch.IdOrdenPagoOrigen

    ElseIf ch.Utilizado Then

        uso = "UTILIZADO - ORIGEN NO IDENTIFICADO"

    Else
        uso = "DISPONIBLE"
    End If

    If LenB(destino) > 0 Then

        If uso = "DISPONIBLE" Then
            DescripcionUsoCheque = destino
        Else
            DescripcionUsoCheque = destino & " - " & uso
        End If

    Else
        DescripcionUsoCheque = uso
    End If

End Function


Private Sub btnBuscarEnChequera_Click(Index As Integer)

    CargarCheques True

End Sub


Private Sub TxtNumeroChequeEnChequera_KeyPress( _
    KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then

        KeyAscii = 0
        CargarCheques True

    End If

End Sub


Private Sub PushButton1_Click()

    Me.TxtNumeroChequeEnChequera.Text = vbNullString
    CargarCheques

    Me.TxtNumeroChequeEnChequera.SetFocus

End Sub


Private Sub chkMostrarIngresados_Click()

    If cargandoDatos Then Exit Sub

    CargarCheques

End Sub


Private Sub cboRangosVtoTerceros_Click(Index As Integer)

    funciones.CalculateDateRange _
        Me.cboRangosVtoTerceros(0), _
        Me.dtpDesdeVtoTerceros(0), _
        Me.dtpHastaVtoTerceros(0)

End Sub


Private Sub cboRangosRboEmitido_Click(Index As Integer)

    funciones.CalculateDateRange _
        Me.cboRangosRboEmitido(0), _
        Me.dtpDesdeRboEmitido(0), _
        Me.dtpHastaRboEmitido(0)

End Sub


Private Sub grid_cheques_ColumnHeaderClick( _
    ByVal Column As GridEX20.JSColumn)

    GridEXHelper.ColumnHeaderClick _
        Me.grid_cheques, Column

End Sub

