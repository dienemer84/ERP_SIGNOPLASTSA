VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmDepositarCheque 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Boleta de Deposito"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   7920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   7920
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.GroupBox GroupBox3 
      Height          =   2610
      Left            =   60
      TabIndex        =   11
      Top             =   3840
      Width           =   7725
      _Version        =   786432
      _ExtentX        =   13626
      _ExtentY        =   4604
      _StockProps     =   79
      Caption         =   "Contenido"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.TabControl TabControl2 
         Height          =   2070
         Left            =   210
         TabIndex        =   12
         Top             =   360
         Width           =   7410
         _Version        =   786432
         _ExtentX        =   13070
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
            Width           =   7065
            _ExtentX        =   12462
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
      Height          =   345
      Left            =   6390
      TabIndex        =   4
      Top             =   6600
      Width           =   1380
      _Version        =   786432
      _ExtentX        =   2434
      _ExtentY        =   609
      _StockProps     =   79
      Caption         =   "Depositar"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   2070
      Left            =   105
      TabIndex        =   0
      Top             =   1605
      Width           =   7680
      _Version        =   786432
      _ExtentX        =   13547
      _ExtentY        =   3651
      _StockProps     =   79
      Caption         =   "Origenes"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.TabControl TabControl1 
         Height          =   1710
         Left            =   135
         TabIndex        =   9
         Top             =   255
         Width           =   7410
         _Version        =   786432
         _ExtentX        =   13070
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
            Left            =   1560
            TabIndex        =   19
            Top             =   585
            Width           =   5760
            _Version        =   786432
            _ExtentX        =   10160
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
            Width           =   810
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
            Height          =   300
            Left            =   5790
            TabIndex        =   16
            Top             =   1260
            Width           =   1485
            _Version        =   786432
            _ExtentX        =   2619
            _ExtentY        =   529
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
End
Attribute VB_Name = "frmDepositarCheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim col As New Collection
Public cheque As cheque
Dim cheques As New Collection
Dim Cajas As New Collection
Dim OpCaja As operacion


Private Sub cmdAgregarCaja_Click()
    Set OpCaja = New operacion
    'Set caja = DAOCaja.FindById(Me.cboCaja.ItemData(Me.cboCaja.ListIndex))



End Sub

Private Sub cmdAgregarCheque_Click()
    Set cheque = DAOCheques.FindById(Me.cboCheques.ItemData(Me.cboCheques.ListIndex))
    If IsSomething(cheque) Then

        If Not BuscarEnColeccion(cheques, cheque.Id) Then
            cheques.Add cheque, CStr(cheque.Id)
        End If

    End If

    Me.gridCheques.ItemCount = cheques.count
    GridEXHelper.AutoSizeColumns Me.gridCheques, True
End Sub

Private Sub Form_Load()
    Customize Me
    GridEXHelper.CustomizeGrid Me.gridCajas, False, True
    GridEXHelper.CustomizeGrid Me.gridCheques, False, False

    Me.DateTimePicker1.value = Now
    DAOCuentaBancaria.llenarComboXtremeSuite Me.cboCuentasBancarias
    DAOMoneda.llenarComboXtremeSuite Me.cboMoneda
    DAOCaja.llenarComboXtremeSuite Me.cboCaja

    Me.gridCheques.ItemCount = 0
    Me.gridCajas.ItemCount = 0
End Sub

Private Sub gridCheques_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Set cheque = cheques(rowIndex)
    Values(1) = cheque.numero
    Values(2) = cheque.FechaVencimiento
    Values(3) = cheque.moneda.NombreCorto & " " & cheque.Monto
    Values(4) = cheque.Banco.nombre
    Values(5) = cheque.FechaRecibido
    Values(6) = cheque.OrigenCheque

End Sub

Private Sub PushButton1_Click()



    If MsgBox("¿Confirma el deposito?", vbYesNo, "Consulta") = vbYes Then

        Dim cuenta As New CuentaBancaria
        Set cuenta = DAOCuentaBancaria.FindById(Me.cboCuentasBancarias.ItemData(Me.cboCuentasBancarias.ListIndex))

        Dim boleta As New BoletaDeposito
        Set boleta.CuentaDestino = cuenta
        boleta.FechaDeposito = Me.DateTimePicker1.value
        boleta.numero = Me.txtBoletaDeposito
        boleta.TipoDeposito = DepositoCheque



        If Not IsSomething(cuenta) Then Exit Sub
        If DAOBoletaDeposito.Save() Then
            MsgBox "Depósito exitoso!", vbInformation, "Información"
        Else
            MsgBox "No se pudo efectuar el depósito", vbCritical, "Error"
        End If


    End If
End Sub

Private Sub txtNroCheque_Change()
    On Error Resume Next
    Dim mostrar As String
    Set col = DAOCheques.FindAll(DAOCheques.CAMPO_EN_CARTERA & "=1 and  " & DAOCheques.TABLA_CHEQUE & "." & DAOCheques.CAMPO_NUMERO & "=" & Val(Me.txtNroCheque))
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
