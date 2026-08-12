VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmDepositarCheque 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Boleta de Deposito"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   9795
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   9795
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.GroupBox GroupBox3 
      Height          =   2610
      Left            =   60
      TabIndex        =   11
      Top             =   3840
      Width           =   9645
      _Version        =   786432
      _ExtentX        =   17013
      _ExtentY        =   4604
      _StockProps     =   79
      Caption         =   "Contenido"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.TabControl TabControl2 
         Height          =   2070
         Left            =   210
         TabIndex        =   12
         Top             =   360
         Width           =   9330
         _Version        =   786432
         _ExtentX        =   16457
         _ExtentY        =   3651
         _StockProps     =   68
         Appearance      =   10
         Color           =   32
         ItemCount       =   2
         Item(0).Caption =   "Cheques"
         Item(0).ControlCount=   2
         Item(0).Control(0)=   "PushButton3"
         Item(0).Control(1)=   "gridCheques"
         Item(1).Caption =   "Caja"
         Item(1).ControlCount=   1
         Item(1).Control(0)=   "GridCajas"
         Begin GridEX20.GridEX gridCheques 
            Height          =   1350
            Left            =   195
            TabIndex        =   14
            Top             =   510
            Width           =   8985
            _ExtentX        =   15849
            _ExtentY        =   2381
            Version         =   "2.0"
            BoundColumnIndex=   ""
            ReplaceColumnIndex=   ""
            ColumnAutoResize=   -1  'True
            MethodHoldFields=   -1  'True
            DataMode        =   99
            ColumnHeaderHeight=   285
            IntProp1        =   0
            IntProp2        =   0
            IntProp7        =   0
            ColumnsCount    =   6
            Column(1)       =   "frmDepositarCheque.frx":0000
            Column(2)       =   "frmDepositarCheque.frx":0118
            Column(3)       =   "frmDepositarCheque.frx":021C
            Column(4)       =   "frmDepositarCheque.frx":0308
            Column(5)       =   "frmDepositarCheque.frx":03F4
            Column(6)       =   "frmDepositarCheque.frx":04F0
            FormatStylesCount=   6
            FormatStyle(1)  =   "frmDepositarCheque.frx":05DC
            FormatStyle(2)  =   "frmDepositarCheque.frx":0714
            FormatStyle(3)  =   "frmDepositarCheque.frx":07C4
            FormatStyle(4)  =   "frmDepositarCheque.frx":0878
            FormatStyle(5)  =   "frmDepositarCheque.frx":0950
            FormatStyle(6)  =   "frmDepositarCheque.frx":0A08
            ImageCount      =   0
            PrinterProperties=   "frmDepositarCheque.frx":0AE8
         End
         Begin XtremeSuiteControls.PushButton PushButton3 
            Height          =   300
            Left            =   -66250
            TabIndex        =   13
            Top             =   1065
            Width           =   1485
            _Version        =   786432
            _ExtentX        =   2619
            _ExtentY        =   529
            _StockProps     =   79
            Caption         =   "Agregar a Boleta"
            UseVisualStyle  =   -1  'True
         End
         Begin GridEX20.GridEX GridCajas 
            Height          =   1350
            Left            =   -69805
            TabIndex        =   15
            Top             =   510
            Visible         =   0   'False
            Width           =   7065
            _ExtentX        =   12462
            _ExtentY        =   2381
            Version         =   "2.0"
            BoundColumnIndex=   ""
            ReplaceColumnIndex=   ""
            DataMode        =   99
            ColumnHeaderHeight=   285
            IntProp1        =   0
            IntProp2        =   0
            IntProp7        =   0
            ColumnsCount    =   2
            Column(1)       =   "frmDepositarCheque.frx":0CC0
            Column(2)       =   "frmDepositarCheque.frx":0D88
            FormatStylesCount=   6
            FormatStyle(1)  =   "frmDepositarCheque.frx":0E2C
            FormatStyle(2)  =   "frmDepositarCheque.frx":0F64
            FormatStyle(3)  =   "frmDepositarCheque.frx":1014
            FormatStyle(4)  =   "frmDepositarCheque.frx":10C8
            FormatStyle(5)  =   "frmDepositarCheque.frx":11A0
            FormatStyle(6)  =   "frmDepositarCheque.frx":1258
            ImageCount      =   0
            PrinterProperties=   "frmDepositarCheque.frx":1338
         End
      End
   End
   Begin XtremeSuiteControls.PushButton PushButton1 
      Height          =   465
      Left            =   7920
      TabIndex        =   4
      Top             =   6480
      Width           =   1620
      _Version        =   786432
      _ExtentX        =   2857
      _ExtentY        =   820
      _StockProps     =   79
      Caption         =   "Depositar"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   2070
      Left            =   105
      TabIndex        =   0
      Top             =   1605
      Width           =   9600
      _Version        =   786432
      _ExtentX        =   16933
      _ExtentY        =   3651
      _StockProps     =   79
      Caption         =   "Origenes"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.TabControl TabControl1 
         Height          =   1710
         Left            =   135
         TabIndex        =   9
         Top             =   255
         Width           =   9330
         _Version        =   786432
         _ExtentX        =   16457
         _ExtentY        =   3016
         _StockProps     =   68
         Appearance      =   10
         Color           =   32
         ItemCount       =   2
         Item(0).Caption =   "Cheques"
         Item(0).ControlCount=   4
         Item(0).Control(0)=   "cmdAgregarCheque"
         Item(0).Control(1)=   "Label1"
         Item(0).Control(2)=   "txtNroCheque"
         Item(0).Control(3)=   "cboCheques"
         Item(1).Caption =   "Caja"
         Item(1).ControlCount=   6
         Item(1).Control(0)=   "cmdAgregarCaja"
         Item(1).Control(1)=   "Label3"
         Item(1).Control(2)=   "cboCaja"
         Item(1).Control(3)=   "Label4"
         Item(1).Control(4)=   "txtImporte"
         Item(1).Control(5)=   "cboMoneda"
         Begin VB.TextBox txtImporte 
            Height          =   285
            Left            =   -64435
            TabIndex        =   23
            Top             =   600
            Visible         =   0   'False
            Width           =   1680
         End
         Begin XtremeSuiteControls.ComboBox cboCheques 
            Height          =   315
            Left            =   2640
            TabIndex        =   19
            Top             =   585
            Width           =   6480
            _Version        =   786432
            _ExtentX        =   11430
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Text            =   "ComboBox1"
         End
         Begin VB.TextBox txtNroCheque 
            Height          =   285
            Left            =   735
            TabIndex        =   18
            Top             =   600
            Width           =   1770
         End
         Begin XtremeSuiteControls.PushButton cmdAgregarCaja 
            Height          =   300
            Left            =   -64210
            TabIndex        =   10
            Top             =   1260
            Visible         =   0   'False
            Width           =   1485
            _Version        =   786432
            _ExtentX        =   2619
            _ExtentY        =   529
            _StockProps     =   79
            Caption         =   "Agregar a Boleta"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton cmdAgregarCheque 
            Height          =   420
            Left            =   7560
            TabIndex        =   16
            Top             =   1140
            Width           =   1605
            _Version        =   786432
            _ExtentX        =   2831
            _ExtentY        =   741
            _StockProps     =   79
            Caption         =   "Agregar a Boleta"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.ComboBox cboCaja 
            Height          =   315
            Left            =   -69175
            TabIndex        =   21
            Top             =   600
            Visible         =   0   'False
            Width           =   2865
            _Version        =   786432
            _ExtentX        =   5054
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.ComboBox cboMoneda 
            Height          =   315
            Left            =   -65485
            TabIndex        =   24
            Top             =   600
            Visible         =   0   'False
            Width           =   1050
            _Version        =   786432
            _ExtentX        =   1852
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Text            =   "ComboBox1"
         End
         Begin VB.Label Label4 
            Caption         =   "Importe"
            Height          =   225
            Left            =   -66085
            TabIndex        =   22
            Top             =   630
            Visible         =   0   'False
            Width           =   810
         End
         Begin VB.Label Label3 
            Caption         =   "Nombre"
            Height          =   225
            Left            =   -69850
            TabIndex        =   20
            Top             =   630
            Visible         =   0   'False
            Width           =   810
         End
         Begin VB.Label Label1 
            Caption         =   "Número"
            Height          =   225
            Left            =   135
            TabIndex        =   17
            Top             =   630
            Width           =   600
         End
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox2 
      Height          =   1530
      Left            =   90
      TabIndex        =   1
      Top             =   60
      Width           =   7695
      _Version        =   786432
      _ExtentX        =   13573
      _ExtentY        =   2699
      _StockProps     =   79
      Caption         =   "Datos de la boleta"
      UseVisualStyle  =   -1  'True
      Begin VB.TextBox txtBoletaDeposito 
         Height          =   285
         Left            =   885
         TabIndex        =   8
         Top             =   330
         Width           =   1635
      End
      Begin XtremeSuiteControls.ComboBox cboCuentasBancarias 
         Height          =   315
         Left            =   885
         TabIndex        =   3
         Top             =   930
         Width           =   3510
         _Version        =   786432
         _ExtentX        =   6191
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.DateTimePicker DateTimePicker1 
         Height          =   255
         Left            =   885
         TabIndex        =   7
         Top             =   645
         Width           =   3480
         _Version        =   786432
         _ExtentX        =   6138
         _ExtentY        =   450
         _StockProps     =   68
         CurrentDate     =   40801.6882407407
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Fecha "
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
         Left            =   135
         TabIndex        =   6
         Top             =   660
         Width           =   660
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0E0FF&
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
         Left            =   135
         TabIndex        =   5
         Top             =   345
         Width           =   825
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Cuenta"
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
         Left            =   135
         TabIndex        =   2
         Top             =   975
         Width           =   750
      End
   End
   Begin VB.Label lblTotalBoleta 
      Caption         =   "Label5"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   6600
      Width           =   3855
   End
End
Attribute VB_Name = "frmDepositarCheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim col As New Collection
Public cheque As cheque
Dim Cheques As New Collection
Dim Cajas As New Collection
Dim OpCaja As operacion


Private Sub cmdAgregarCaja_Click()
    Set OpCaja = New operacion
    'Set caja = DAOCaja.FindById(Me.cboCaja.ItemData(Me.cboCaja.ListIndex))



End Sub


Private Sub cmdAgregarCheque_Click()

    On Error GoTo err1

    Dim agregado As Boolean

    agregado = False


    If Me.cboCheques.ListIndex < 0 Then

        MsgBox "Debe seleccionar un cheque.", _
               vbExclamation, "Boleta de depósito"

        Exit Sub

    End If


    Set cheque = DAOCheques.FindById( _
                    Me.cboCheques.ItemData( _
                        Me.cboCheques.ListIndex))


    If cheque Is Nothing Then

        MsgBox "No se pudo obtener el cheque seleccionado.", _
               vbExclamation, "Boleta de depósito"

        Exit Sub

    End If


    If Not cheque.EnCartera Then

        MsgBox "El cheque seleccionado ya no se encuentra en cartera.", _
               vbExclamation, "Boleta de depósito"

        Exit Sub

    End If


    If cheque.Depositado Then

        MsgBox "El cheque seleccionado ya figura como depositado.", _
               vbExclamation, "Boleta de depósito"

        Exit Sub

    End If


    '-------------------------------------------------------
    ' AGREGAR A LA BOLETA
    '-------------------------------------------------------

    If Not BuscarEnColeccion(Cheques, cheque.Id) Then

        Cheques.Add cheque, CStr(cheque.Id)

        agregado = True

    Else

        MsgBox "El cheque ya fue agregado a esta boleta.", _
               vbInformation, "Boleta de depósito"

    End If


    '-------------------------------------------------------
    ' ACTUALIZAR GRILLA
    '-------------------------------------------------------

    Me.gridCheques.ItemCount = 0

    Me.gridCheques.ItemCount = Cheques.count

    Me.gridCheques.Update

    GridEXHelper.AutoSizeColumns Me.gridCheques, True


    '-------------------------------------------------------
    ' SI SE AGREGÓ CORRECTAMENTE:
    '
    ' 1 - RECALCULAR TOTAL
    ' 2 - LIMPIAR BUSQUEDA
    ' 3 - DEJAR CURSOR PARA EL SIGUIENTE CHEQUE
    '-------------------------------------------------------

    If agregado Then

        ActualizarTotalBoleta

        LimpiarBusquedaCheque

    End If


    Exit Sub


err1:

    MsgBox "No se pudo agregar el cheque." & vbCrLf & _
           Err.Description, _
           vbCritical, "Boleta de depósito"

End Sub

Private Sub Form_Load()
    Customize Me
    GridEXHelper.CustomizeGrid Me.GridCajas, False, True
    GridEXHelper.CustomizeGrid Me.gridCheques, False, False

    Me.DateTimePicker1.value = Now
    DAOCuentaBancaria.llenarComboXtremeSuite Me.cboCuentasBancarias
    DAOMoneda.llenarComboXtremeSuite Me.cboMoneda
    DAOCaja.llenarComboXtremeSuite Me.cboCaja

    Me.gridCheques.ItemCount = 0
    Me.GridCajas.ItemCount = 0
    
    ActualizarTotalBoleta
    
End Sub


Private Sub gridCheques_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Set cheque = Cheques(RowIndex)
    Values(1) = cheque.numero
    Values(2) = cheque.FechaVencimiento
    Values(3) = cheque.moneda.NombreCorto & " " & cheque.Monto
    Values(4) = cheque.Banco.nombre
    Values(5) = cheque.FechaRecibido
    Values(6) = cheque.OrigenCheque

End Sub


Private Sub PushButton1_Click()

    On Error GoTo err1


    Dim cuenta As CuentaBancaria
    Dim boleta As BoletaDeposito
    Dim ch As cheque

    Dim numeroBoleta As Double
    Dim detalleError As String


    '-------------------------------------------------------
    ' NUMERO DE BOLETA
    '-------------------------------------------------------

    If LenB(Trim$(Me.txtBoletaDeposito.Text)) = 0 Then

        MsgBox "Debe ingresar el número de la boleta de depósito.", _
               vbExclamation, "Boleta de depósito"

        Me.txtBoletaDeposito.SetFocus
        Exit Sub

    End If


    If Not IsNumeric(Me.txtBoletaDeposito.Text) Then

        MsgBox "El número de boleta debe ser numérico.", _
               vbExclamation, "Boleta de depósito"

        Me.txtBoletaDeposito.SetFocus
        Exit Sub

    End If


    numeroBoleta = CDbl(Me.txtBoletaDeposito.Text)


    If numeroBoleta <= 0 Or _
       numeroBoleta <> Fix(numeroBoleta) Or _
       numeroBoleta > 2147483647# Then

        MsgBox "Ingrese un número de boleta válido.", _
               vbExclamation, "Boleta de depósito"

        Me.txtBoletaDeposito.SetFocus
        Exit Sub

    End If


    '-------------------------------------------------------
    ' CUENTA BANCARIA
    '-------------------------------------------------------

    If Me.cboCuentasBancarias.ListIndex < 0 Then

        MsgBox "Debe seleccionar la cuenta bancaria destino.", _
               vbExclamation, "Boleta de depósito"

        Exit Sub

    End If


    Set cuenta = DAOCuentaBancaria.FindById( _
                    Me.cboCuentasBancarias.ItemData( _
                        Me.cboCuentasBancarias.ListIndex))


    If cuenta Is Nothing Then

        MsgBox "No se pudo obtener la cuenta bancaria seleccionada.", _
               vbCritical, "Boleta de depósito"

        Exit Sub

    End If


    If cuenta.moneda Is Nothing Then

        MsgBox "La cuenta bancaria seleccionada no tiene moneda definida.", _
               vbExclamation, "Boleta de depósito"

        Exit Sub

    End If


    '-------------------------------------------------------
    ' CHEQUES
    '-------------------------------------------------------

    If Cheques.count = 0 Then

        MsgBox "Debe agregar por lo menos un cheque a la boleta.", _
               vbExclamation, "Boleta de depósito"

        Exit Sub

    End If


    'Validar moneda antes de preguntarle al usuario
    For Each ch In Cheques

        If ch.moneda Is Nothing Then

            MsgBox "El cheque Nº " & ch.numero & _
                   " no tiene moneda definida.", _
                   vbExclamation, "Boleta de depósito"

            Exit Sub

        End If


        If ch.moneda.Id <> cuenta.moneda.Id Then

            MsgBox "El cheque Nº " & ch.numero & _
                   " está expresado en " & ch.moneda.NombreCorto & _
                   " y la cuenta seleccionada está expresada en " & _
                   cuenta.moneda.NombreCorto & "." & vbCrLf & vbCrLf & _
                   "No se puede realizar el depósito.", _
                   vbExclamation, "Boleta de depósito"

            Exit Sub

        End If

    Next ch


    '-------------------------------------------------------
    ' CONFIRMACION
    '-------------------------------------------------------

    If MsgBox( _
        "¿Confirma el depósito de " & _
        Cheques.count & " cheque(s) en la cuenta seleccionada?", _
        vbQuestion + vbYesNo, _
        "Boleta de depósito") <> vbYes Then

        Exit Sub

    End If


    '-------------------------------------------------------
    ' ARMAR BOLETA
    '-------------------------------------------------------

    Set boleta = New BoletaDeposito

    boleta.numero = CLng(numeroBoleta)

    boleta.fechaDeposito = Me.DateTimePicker1.value

    Set boleta.CuentaDestino = cuenta

    boleta.TipoDeposito = DepositoCheque


    'Pasar los cheques del formulario al objeto BoletaDeposito
    For Each ch In Cheques

        boleta.Cheques.Add ch, CStr(ch.Id)

    Next ch


    '-------------------------------------------------------
    ' GUARDAR
    '-------------------------------------------------------

    If DAOBoletaDeposito.Save(boleta) Then

        MsgBox "El depósito se registró correctamente.", _
               vbInformation, "Boleta de depósito"

        'Se cierra para no permitir volver a depositar
        'los mismos cheques desde la colección que quedó cargada.
        Unload Me

    Else

        detalleError = DAOBoletaDeposito.UltimoError

        If LenB(detalleError) > 0 Then
            detalleError = vbCrLf & vbCrLf & detalleError
        End If


        MsgBox "No se pudo efectuar el depósito." & _
               detalleError, _
               vbCritical, "Boleta de depósito"

    End If


    Exit Sub


err1:

    MsgBox "Se produjo un error al realizar el depósito." & vbCrLf & _
           Err.Description, _
           vbCritical, "Boleta de depósito"

End Sub


Private Sub txtNroCheque_Change()
    On Error Resume Next
    Dim mostrar As String
    Set col = DAOCheques.FindAll(DAOCheques.CAMPO_EN_CARTERA & "=1 and  " & DAOCheques.TABLA_CHEQUE & "." & DAOCheques.CAMPO_NUMERO & "=" & val(Me.txtNroCheque))
    If col.count >= 1 Then
        cboCheques.Clear
        For Each cheque In col
            mostrar = cheque.Banco.nombre

            If LenB(cheque.OrigenDestino) > 0 Then mostrar = mostrar & " | " & cheque.OrigenDestino

            mostrar = mostrar & " | " & cheque.moneda.NombreCorto & " " & cheque.Monto & " | " & cheque.FechaVencimiento
            Me.cboCheques.AddItem mostrar
            cboCheques.ItemData(cboCheques.NewIndex) = cheque.Id
        Next cheque

        If cboCheques.ListCount > 0 Then cboCheques.ListIndex = 0
    Else
        cboCheques.Clear
    End If


End Sub


Private Sub txtNroCheque_GotFocus()

    foco Me.txtNroCheque
    
End Sub


Private Sub ActualizarTotalBoleta()

    Dim ch As cheque
    Dim total As Double
    Dim moneda As String

    total = 0
    moneda = vbNullString

    For Each ch In Cheques

        total = total + ch.Monto

        If LenB(moneda) = 0 Then
            If Not ch.moneda Is Nothing Then
                moneda = ch.moneda.NombreCorto
            End If
        End If

    Next ch

    If Cheques.count = 0 Then

        Me.lblTotalBoleta.caption = _
            "TOTAL BOLETA: $ 0,00"

    Else

        Me.lblTotalBoleta.caption = _
            "TOTAL BOLETA: " & _
            moneda & " " & _
             Replace(FormatCurrency(funciones.FormatearDecimales(total)), "$", "")

    End If

End Sub


Private Sub LimpiarBusquedaCheque()

    Me.txtNroCheque.Text = vbNullString

    Me.cboCheques.Clear

    Set cheque = Nothing

    Me.txtNroCheque.SetFocus

End Sub

