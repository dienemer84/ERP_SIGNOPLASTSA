VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~3.OCX"
Begin VB.Form frmAdminCheques 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administración de cheques"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15390
   Icon            =   "frmAdminCheques.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   15390
   Begin GridEX20.GridEX gridBancos 
      Height          =   1845
      Left            =   345
      TabIndex        =   0
      Top             =   8160
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   3254
      Version         =   "2.0"
      BoundColumnIndex=   "id"
      ReplaceColumnIndex=   "nombre"
      ActAsDropDown   =   -1  'True
      ColumnAutoResize=   -1  'True
      HideSelection   =   2
      MethodHoldFields=   -1  'True
      ContScroll      =   -1  'True
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      NewRowPos       =   1
      RowHeaders      =   -1  'True
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   2
      Column(1)       =   "frmAdminCheques.frx":000C
      Column(2)       =   "frmAdminCheques.frx":010C
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminCheques.frx":01FC
      FormatStyle(2)  =   "frmAdminCheques.frx":0334
      FormatStyle(3)  =   "frmAdminCheques.frx":03E4
      FormatStyle(4)  =   "frmAdminCheques.frx":0498
      FormatStyle(5)  =   "frmAdminCheques.frx":0570
      FormatStyle(6)  =   "frmAdminCheques.frx":0628
      ImageCount      =   0
      PrinterProperties=   "frmAdminCheques.frx":0708
   End
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   7875
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   15375
      _Version        =   786432
      _ExtentX        =   27120
      _ExtentY        =   13891
      _StockProps     =   68
      Appearance      =   10
      Color           =   32
      PaintManager.BoldSelected=   -1  'True
      PaintManager.ShowIcons=   -1  'True
      ItemCount       =   3
      Item(0).Caption =   "Cartera"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "grid_cartera_cheques"
      Item(1).Caption =   "Administrar Chequeras"
      Item(1).ControlCount=   3
      Item(1).Control(0)=   "grid_chequeras"
      Item(1).Control(1)=   "grid_cheques"
      Item(1).Control(2)=   "GroupBox1"
      Item(2).Caption =   "Emitidos por Banco"
      Item(2).ControlCount=   2
      Item(2).Control(0)=   "gridChequesEmitidos"
      Item(2).Control(1)=   "GroupBox2"
      Begin XtremeSuiteControls.GroupBox GroupBox2 
         Height          =   1305
         Left            =   -69700
         TabIndex        =   19
         Top             =   450
         Visible         =   0   'False
         Width           =   14715
         _Version        =   786432
         _ExtentX        =   25956
         _ExtentY        =   2302
         _StockProps     =   79
         Caption         =   "Parámetros de búsqueda"
         UseVisualStyle  =   -1  'True
         Begin VB.TextBox txtIdOP 
            Height          =   285
            Left            =   8385
            TabIndex        =   36
            Top             =   810
            Width           =   1425
         End
         Begin VB.TextBox txtNroCheque 
            Height          =   285
            Left            =   6150
            TabIndex        =   34
            Top             =   810
            Width           =   1425
         End
         Begin XtremeSuiteControls.CheckBox chkIngresados 
            Height          =   195
            Left            =   10095
            TabIndex        =   20
            Top             =   465
            Width           =   1395
            _Version        =   786432
            _ExtentX        =   2461
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Ingresados"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton PushButton1 
            Height          =   345
            Left            =   13095
            TabIndex        =   21
            Top             =   315
            Width           =   1365
            _Version        =   786432
            _ExtentX        =   2408
            _ExtentY        =   609
            _StockProps     =   79
            Caption         =   "Buscar"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.ComboBox cboBancos1 
            Height          =   315
            Left            =   915
            TabIndex        =   22
            Top             =   375
            Width           =   3765
            _Version        =   786432
            _ExtentX        =   6641
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.PushButton CMDsINCliente 
            Height          =   255
            Left            =   4740
            TabIndex        =   23
            Top             =   405
            Width           =   420
            _Version        =   786432
            _ExtentX        =   741
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "X"
            BackColor       =   12632256
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.DateTimePicker dtpDesde 
            Height          =   315
            Left            =   6150
            TabIndex        =   24
            Top             =   375
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
            Left            =   8385
            TabIndex        =   25
            Top             =   375
            Width           =   1470
            _Version        =   786432
            _ExtentX        =   2593
            _ExtentY        =   556
            _StockProps     =   68
            CheckBox        =   -1  'True
            Format          =   1
         End
         Begin XtremeSuiteControls.PushButton PushButton2 
            Height          =   345
            Left            =   13095
            TabIndex        =   26
            Top             =   690
            Width           =   1365
            _Version        =   786432
            _ExtentX        =   2408
            _ExtentY        =   609
            _StockProps     =   79
            Caption         =   "Imprimir"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.ComboBox cboChequera2 
            Height          =   315
            Left            =   900
            TabIndex        =   31
            Top             =   810
            Width           =   3765
            _Version        =   786432
            _ExtentX        =   6641
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.PushButton PushButton3 
            Height          =   255
            Left            =   4740
            TabIndex        =   32
            Top             =   810
            Width           =   420
            _Version        =   786432
            _ExtentX        =   741
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "X"
            BackColor       =   12632256
            UseVisualStyle  =   -1  'True
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "O.P."
            Height          =   300
            Left            =   7980
            TabIndex        =   35
            Top             =   840
            Width           =   945
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Número"
            Height          =   300
            Left            =   5550
            TabIndex        =   33
            Top             =   840
            Width           =   945
         End
         Begin VB.Label Label8 
            Caption         =   "Chequera"
            Height          =   240
            Left            =   105
            TabIndex        =   30
            Top             =   840
            Width           =   690
         End
         Begin VB.Label lblBanco 
            AutoSize        =   -1  'True
            Caption         =   "Banco"
            Height          =   195
            Left            =   345
            TabIndex        =   29
            Top             =   435
            Width           =   465
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            Height          =   195
            Left            =   5625
            TabIndex        =   28
            Top             =   450
            Width           =   465
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            Height          =   195
            Left            =   7875
            TabIndex        =   27
            Top             =   450
            Width           =   420
         End
      End
      Begin GridEX20.GridEX gridChequesEmitidos 
         Height          =   5655
         Left            =   -69715
         TabIndex        =   2
         Top             =   1845
         Visible         =   0   'False
         Width           =   14730
         _ExtentX        =   25982
         _ExtentY        =   9975
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
         ColumnsCount    =   8
         Column(1)       =   "frmAdminCheques.frx":08E0
         Column(2)       =   "frmAdminCheques.frx":0A20
         Column(3)       =   "frmAdminCheques.frx":0B2C
         Column(4)       =   "frmAdminCheques.frx":0C30
         Column(5)       =   "frmAdminCheques.frx":0D44
         Column(6)       =   "frmAdminCheques.frx":0EAC
         Column(7)       =   "frmAdminCheques.frx":0FA0
         Column(8)       =   "frmAdminCheques.frx":1088
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmAdminCheques.frx":1174
         FormatStyle(2)  =   "frmAdminCheques.frx":12AC
         FormatStyle(3)  =   "frmAdminCheques.frx":135C
         FormatStyle(4)  =   "frmAdminCheques.frx":1410
         FormatStyle(5)  =   "frmAdminCheques.frx":14E8
         FormatStyle(6)  =   "frmAdminCheques.frx":15A0
         ImageCount      =   0
         PrinterProperties=   "frmAdminCheques.frx":1680
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   1860
         Left            =   -69715
         TabIndex        =   3
         Top             =   5835
         Visible         =   0   'False
         Width           =   7365
         _Version        =   786432
         _ExtentX        =   12991
         _ExtentY        =   3281
         _StockProps     =   79
         Caption         =   "Crear Chequera"
         UseVisualStyle  =   -1  'True
         Begin VB.TextBox txtDesde 
            Height          =   285
            Left            =   990
            TabIndex        =   9
            Text            =   "0"
            Top             =   630
            Width           =   1035
         End
         Begin VB.TextBox txtHasta 
            Height          =   285
            Left            =   2910
            TabIndex        =   8
            Text            =   "0"
            Top             =   615
            Width           =   1020
         End
         Begin VB.TextBox txtNumero 
            Height          =   285
            Left            =   1005
            TabIndex        =   7
            Text            =   "0"
            Top             =   300
            Width           =   1035
         End
         Begin VB.TextBox txtObservaciones 
            Height          =   1080
            Left            =   4065
            MultiLine       =   -1  'True
            TabIndex        =   4
            Top             =   225
            Width           =   3120
         End
         Begin XtremeSuiteControls.ComboBox cboMonedas 
            Height          =   315
            Left            =   975
            TabIndex        =   5
            Top             =   1380
            Width           =   1515
            _Version        =   786432
            _ExtentX        =   2672
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Appearance      =   6
            Text            =   "ComboBox1"
            AutoComplete    =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton cmdCrear 
            Height          =   390
            Left            =   5745
            TabIndex        =   6
            Top             =   1395
            Width           =   1470
            _Version        =   786432
            _ExtentX        =   2593
            _ExtentY        =   688
            _StockProps     =   79
            Caption         =   "Crear"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.ComboBox cboBancos 
            Height          =   315
            Left            =   975
            TabIndex        =   10
            Top             =   990
            Width           =   2970
            _Version        =   786432
            _ExtentX        =   5239
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Appearance      =   6
            Text            =   "ComboBox1"
            AutoComplete    =   -1  'True
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Desde"
            Height          =   270
            Left            =   315
            TabIndex        =   15
            Top             =   660
            Width           =   570
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Numero"
            Height          =   270
            Left            =   -45
            TabIndex        =   14
            Top             =   330
            Width           =   945
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Hasta"
            Height          =   240
            Left            =   2115
            TabIndex        =   13
            Top             =   645
            Width           =   675
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Bancos"
            Height          =   180
            Left            =   -30
            TabIndex        =   12
            Top             =   1035
            Width           =   945
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Moneda"
            Height          =   165
            Left            =   150
            TabIndex        =   11
            Top             =   1470
            Width           =   750
         End
      End
      Begin GridEX20.GridEX grid_cheques 
         Height          =   7110
         Left            =   -62215
         TabIndex        =   16
         Top             =   615
         Visible         =   0   'False
         Width           =   7485
         _ExtentX        =   13203
         _ExtentY        =   12541
         Version         =   "2.0"
         PreviewRowIndent=   200
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         PreviewColumn   =   5
         PreviewRowLines =   1
         ColumnAutoResize=   -1  'True
         MethodHoldFields=   -1  'True
         DataMode        =   99
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   5
         Column(1)       =   "frmAdminCheques.frx":1858
         Column(2)       =   "frmAdminCheques.frx":1970
         Column(3)       =   "frmAdminCheques.frx":1A84
         Column(4)       =   "frmAdminCheques.frx":1BBC
         Column(5)       =   "frmAdminCheques.frx":1CCC
         FormatStylesCount=   7
         FormatStyle(1)  =   "frmAdminCheques.frx":1D8C
         FormatStyle(2)  =   "frmAdminCheques.frx":1EC4
         FormatStyle(3)  =   "frmAdminCheques.frx":1F74
         FormatStyle(4)  =   "frmAdminCheques.frx":2028
         FormatStyle(5)  =   "frmAdminCheques.frx":2100
         FormatStyle(6)  =   "frmAdminCheques.frx":21B8
         FormatStyle(7)  =   "frmAdminCheques.frx":2298
         ImageCount      =   0
         PrinterProperties=   "frmAdminCheques.frx":2354
      End
      Begin GridEX20.GridEX grid_chequeras 
         Height          =   5130
         Left            =   -69745
         TabIndex        =   17
         Top             =   630
         Visible         =   0   'False
         Width           =   7440
         _ExtentX        =   13123
         _ExtentY        =   9049
         Version         =   "2.0"
         HoldSortSettings=   -1  'True
         DefaultGroupMode=   1
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         PreviewColumn   =   "observaciones"
         PreviewRowLines =   1
         ColumnAutoResize=   -1  'True
         MethodHoldFields=   -1  'True
         DataMode        =   99
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   6
         Column(1)       =   "frmAdminCheques.frx":252C
         Column(2)       =   "frmAdminCheques.frx":2644
         Column(3)       =   "frmAdminCheques.frx":2740
         Column(4)       =   "frmAdminCheques.frx":282C
         Column(5)       =   "frmAdminCheques.frx":2928
         Column(6)       =   "frmAdminCheques.frx":2A24
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmAdminCheques.frx":2B4C
         FormatStyle(2)  =   "frmAdminCheques.frx":2C84
         FormatStyle(3)  =   "frmAdminCheques.frx":2D34
         FormatStyle(4)  =   "frmAdminCheques.frx":2DE8
         FormatStyle(5)  =   "frmAdminCheques.frx":2EC0
         FormatStyle(6)  =   "frmAdminCheques.frx":2F78
         ImageCount      =   0
         PrinterProperties=   "frmAdminCheques.frx":3058
      End
      Begin GridEX20.GridEX grid_cartera_cheques 
         Height          =   6705
         Left            =   285
         TabIndex        =   18
         Top             =   570
         Width           =   14835
         _ExtentX        =   26167
         _ExtentY        =   11827
         Version         =   "2.0"
         DefaultGroupMode=   1
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         PreviewColumn   =   "observaciones"
         PreviewRowLines =   1
         ColumnAutoResize=   -1  'True
         MethodHoldFields=   -1  'True
         ContScroll      =   -1  'True
         AllowEdit       =   0   'False
         RowHeaders      =   -1  'True
         DataMode        =   99
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   9
         Column(1)       =   "frmAdminCheques.frx":3230
         Column(2)       =   "frmAdminCheques.frx":3358
         Column(3)       =   "frmAdminCheques.frx":344C
         Column(4)       =   "frmAdminCheques.frx":35A0
         Column(5)       =   "frmAdminCheques.frx":36F0
         Column(6)       =   "frmAdminCheques.frx":3800
         Column(7)       =   "frmAdminCheques.frx":390C
         Column(8)       =   "frmAdminCheques.frx":3A20
         Column(9)       =   "frmAdminCheques.frx":3B68
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmAdminCheques.frx":3C90
         FormatStyle(2)  =   "frmAdminCheques.frx":3DC8
         FormatStyle(3)  =   "frmAdminCheques.frx":3E78
         FormatStyle(4)  =   "frmAdminCheques.frx":3F2C
         FormatStyle(5)  =   "frmAdminCheques.frx":4004
         FormatStyle(6)  =   "frmAdminCheques.frx":40BC
         ImageCount      =   0
         PrinterProperties=   "frmAdminCheques.frx":419C
      End
   End
   Begin VB.Menu mnuOpcionesChequeChequera 
      Caption         =   "mnuOpcionesChequeChequera"
      Visible         =   0   'False
      Begin VB.Menu mnuPasarCartera 
         Caption         =   "Pasar a cartera..."
      End
      Begin VB.Menu mnuAnularCheque 
         Caption         =   "Anular..."
      End
   End
End
Attribute VB_Name = "frmAdminCheques"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As Recordset
Dim cartera As Collection
Dim tmpChequera As chequera
Dim cheques1 As New Collection
Dim chequeras As Collection
Dim cheques2 As New Collection
Dim tmpCheque As cheque
Dim bancos As Collection
Dim Banco As Banco
Private Sub cmdCrear_Click()
    Dim x As Long
    Dim col As Collection
    Dim id_banco As Long


    If MsgBox("¿Está seguro de crear la chequera?", vbQuestion + vbYesNo) = vbYes Then
        Dim chequera As New chequera
        If Me.cboBancos.ListIndex = -1 Then
            MsgBox "Seleccione un banco Correcto!", vbCritical, "Error"
            Exit Sub
        End If
        If Not IsNumeric(Me.txtNumero) Or Not IsNumeric(Me.txtDesde) Or Not IsNumeric(Me.txtHasta) Then
            MsgBox "Ingrese números válidos!", vbCritical, "Error"
            Exit Sub
        End If
        id_banco = Me.cboBancos.ItemData(Me.cboBancos.ListIndex)
        Set col = DAOChequeras.GetAll(DAOChequeras.CAMPO_NUMERO & "=" & Me.txtNumero & " AND id_banco=" & id_banco)
        If col.count > 0 Then
            MsgBox "El número de chequera de ese banco ya existe!", vbCritical, "Error"
            Exit Sub
        End If

        Set chequera.Banco = DAOBancos.GetById(id_banco)
        chequera.FechaCreacion = Now
        Set chequera.Moneda = DAOMoneda.GetById(Me.cboMonedas.ItemData(Me.cboMonedas.ListIndex))
        chequera.numero = CLng(Me.txtNumero)
        chequera.NumeroDesde = CLng(Me.txtDesde)
        chequera.NumeroHasta = CLng(Me.txtHasta)
        chequera.Observaciones = UCase(Me.txtObservaciones)
        Dim cheque As cheque
        For x = chequera.NumeroDesde To chequera.NumeroHasta
            Set cheque = New cheque
            cheque.numero = x
            cheque.EnCartera = False
            cheque.Propio = True
            cheque.id = 0
            Set cheque.Banco = chequera.Banco
            Set cheque.Moneda = chequera.Moneda
            chequera.Cheques.Add cheque

        Next
        If DAOChequeras.Guardar(chequera) Then
            MsgBox "Guardado Correctamente!", vbInformation, "Información"
            MostrarChequeras
        End If
    End If



End Sub

Private Sub CMDsINCliente_Click()
    Me.cboBancos1.ListIndex = -1
End Sub

Private Sub Form_Load()
    FormHelper.Customize Me
    GridEXHelper.CustomizeGrid Me.grid_chequeras, True, False
    GridEXHelper.CustomizeGrid Me.grid_cartera_cheques, True, True
    GridEXHelper.CustomizeGrid Me.grid_cheques, False, False
    GridEXHelper.CustomizeGrid Me.gridBancos, False, False
    GridEXHelper.CustomizeGrid Me.gridChequesEmitidos, False, False

    DAOBancos.llenarComboXtremeSuite Me.cboBancos
    DAOBancos.llenarComboXtremeSuite Me.cboBancos1
    DAOMoneda.llenarComboXtremeSuite Me.cboMonedas

    DAOChequeras.llenarComboXtremeSuite Me.cboChequera2


    Set bancos = DAOBancos.GetAll("id in (select idBanco from AdminConfigCuentas group by idBanco) ")

    cboBancos1.Clear
    For Each Banco In bancos
        cboBancos1.AddItem Banco.nombre
        cboBancos1.ItemData(cboBancos1.NewIndex) = Banco.id
    Next

    cboBancos1.ListIndex = -1

    Me.cboChequera2.ListIndex = -1
    Set bancos = DAOBancos.GetAll()
    Me.grid_cheques.ItemCount = 0
    Me.gridBancos.ItemCount = 0
    Me.gridChequesEmitidos.ItemCount = 0


    MostrarChequeras
    MostrarCartera

    Set Me.grid_cartera_cheques.Columns("banco").DropDownControl = Me.gridBancos
    Me.gridBancos.ItemCount = bancos.count


    Dim idc As Long
    idc = chequeras.item(Me.grid_chequeras.RowIndex(Me.grid_chequeras.row)).id

    Set tmpChequera = DAOChequeras.GetById(idc)
    Set tmpChequera.Cheques = DAOCheques.FindAllByChequeraId(idc)


End Sub
Private Sub Form_Resize()
    Me.TabControl1.Width = Me.ScaleWidth
    Me.TabControl1.Height = Me.ScaleHeight
End Sub
Private Sub MostrarCartera()
    Set cartera = DAOCheques.FindAllEnCartera()
    Me.grid_cartera_cheques.ItemCount = 0
    Me.grid_cartera_cheques.ItemCount = cartera.count
End Sub

Private Sub MostrarChequeras()
    Set chequeras = DAOChequeras.GetAll
    Me.grid_chequeras.ItemCount = 0
    Me.grid_chequeras.ItemCount = chequeras.count
End Sub

Private Sub grid_cartera_cheques_BeforeUpdate(ByVal Cancel As GridEX20.JSRetBoolean)
    'validar


    Dim cond1 As Boolean, cond2 As Boolean
    Dim cond3 As Boolean
    cond1 = Not (IsDate(Me.grid_cartera_cheques.value(7)) And IsDate(Me.grid_cartera_cheques.value(3)))
    cond2 = Not (IsNumeric(Me.grid_cartera_cheques.value(2)) And IsNumeric(Me.grid_cartera_cheques.value(1)))
    cond3 = False    ' Not (IsNumeric(Me.grid_cartera_cheques.value(5)) And Val(Me.grid_cartera_cheques.value(5)) > 0)
    Cancel = cond1 Or cond2 Or cond3
End Sub

Private Sub grid_cartera_cheques_ColumnHeaderClick(ByVal Column As GridEX20.JSColumn)
    GridEXHelper.ColumnHeaderClick Me.grid_cartera_cheques, Column
End Sub


Private Sub grid_cartera_cheques_UnboundAddNew(ByVal NewRowBookmark As GridEX20.JSRetVariant, ByVal Values As GridEX20.JSRowData)
    On Error GoTo err1
    Set tmpCheque = New cheque
    Set tmpCheque.Banco = DAOBancos.GetById(Values(5))
    tmpCheque.EnCartera = True
    tmpCheque.FechaRecibido = Values(7)
    tmpCheque.FechaVencimiento = Values(3)
    Set tmpCheque.Moneda = DAOMoneda.GetById(0)       ' reemplazar x un combo
    tmpCheque.Monto = Values(2)
    tmpCheque.numero = Values(1)
    tmpCheque.OrigenDestino = Values(4)
    tmpCheque.Propio = False

    If Not DAOCheques.Guardar(tmpCheque) Then GoTo err1
    cartera.Add tmpCheque, CStr(tmpCheque.id)

    Exit Sub
err1:

End Sub

Private Sub grid_cartera_cheques_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    On Error GoTo err1
    Set tmpCheque = cartera.item(RowIndex)
    With Values
        .value(1) = tmpCheque.id
        .value(2) = tmpCheque.numero
        .value(3) = funciones.FormatearDecimales(tmpCheque.Monto)
        .value(4) = tmpCheque.FechaVencimiento
        .value(5) = tmpCheque.OrigenDestino
        .value(6) = tmpCheque.Banco.nombre
        .value(7) = tmpCheque.OrigenCheque
        .value(8) = tmpCheque.FechaRecibido
        .value(9) = tmpCheque.Observaciones

    End With

    Exit Sub
err1:


End Sub

Private Sub grid_cartera_cheques_UnboundUpdate(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    On Error GoTo err1
    Dim ant As String

    ant = tmpCheque.OrigenDestino

    Set tmpCheque = cartera.item(RowIndex)
    tmpCheque.OrigenDestino = Values(4)
    If Not DAOCheques.Guardar(tmpCheque) Then GoTo err1
    Exit Sub
err1:
    tmpCheque.OrigenDestino = ant
End Sub

Private Sub grid_chequeras_SelectionChange()
    On Error Resume Next
    Set tmpChequera.Cheques = DAOCheques.FindAllByChequeraId(tmpChequera.id)
    mostrarCheques
End Sub
Private Sub mostrarCheques()
    Me.grid_cheques.ItemCount = 0
    Me.grid_cheques.ItemCount = tmpChequera.Cheques.count
End Sub
Private Sub grid_chequeras_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Set tmpChequera = chequeras.item(RowIndex)
    With Values
        .value(1) = tmpChequera.numero
        .value(2) = tmpChequera.FechaCreacion
        .value(3) = tmpChequera.Banco.nombre
        .value(4) = tmpChequera.NumeroDesde
        .value(5) = tmpChequera.NumeroHasta

    End With

End Sub


Private Sub grid_cheques_DblClick()
    If Me.grid_cheques.RowIndex(Me.grid_cheques.row) > 0 Then
        Set tmpCheque = tmpChequera.Cheques(Me.grid_cheques.RowIndex(Me.grid_cheques.row))
        PasarACartera tmpCheque
    End If
End Sub

Private Sub PasarACartera(ch As cheque)
    If ch.EnCartera Then
        MsgBox "El cheque ya se encuentra en cartera.", vbInformation
    Else
        If ch.Utilizado Then
            MsgBox "El cheque ya fue utilizado.", vbInformation
        Else
            Dim f000 As New frmChequePropioACartera
            Set f000.cheque = ch
            Load f000
            f000.Show 1
            mostrarCheques
        End If
    End If
End Sub

Private Sub grid_cheques_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Set tmpCheque = tmpChequera.Cheques(Me.grid_cheques.RowIndex(Me.grid_cheques.row))
        Me.mnuAnularCheque.Enabled = (tmpCheque.IdOrdenPagoOrigen <= 0) Or tmpCheque.estado = ChequeAnulado
        Me.PopupMenu Me.mnuOpcionesChequeChequera
    End If
End Sub


Private Sub grid_cheques_RowFormat(RowBuffer As GridEX20.JSRowData)
    On Error GoTo err1

    If tmpCheque.estado = ChequeAnulado Then RowBuffer.RowStyle = "anulado"
    Exit Sub
err1:

End Sub

Private Sub grid_cheques_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    On Error GoTo err1
    If Not IsSomething(tmpChequera.Cheques) Or tmpChequera.Cheques.count = 0 Then Set tmpChequera.Cheques = DAOCheques.FindAllByChequeraId(tmpChequera.id)
    Set tmpCheque = tmpChequera.Cheques(RowIndex)
    With Values
        .value(1) = tmpCheque.numero
        .value(2) = IIf(tmpCheque.Utilizado, funciones.FormatearDecimales(tmpCheque.Monto), Empty)
        .value(3) = IIf(tmpCheque.Utilizado, tmpCheque.FechaVencimiento, Empty)
        .value(4) = IIf(tmpCheque.Utilizado, tmpCheque.OrigenDestino, Empty)
        '.value(5) = IIf(tmpCheque.Utilizado, tmpCheque.Observaciones, "DISPONIBLE")
        .value(5) = IIf(tmpCheque.estado = ChequeAnulado, "ANULADO", IIf(tmpCheque.Utilizado, "Utilizado en Orden de Pago Nº " & tmpCheque.IdOrdenPagoOrigen, "DISPONIBLE"))
    End With
    Exit Sub
err1:
End Sub

Private Sub gridBancos_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex <= bancos.count Then
        Set Banco = bancos.item(RowIndex)
        Values(1) = Banco.id
        Values(2) = Banco.nombre
    End If
End Sub

Private Sub gridChequesEmitidos_ColumnHeaderClick(ByVal Column As GridEX20.JSColumn)
    GridEXHelper.ColumnHeaderClick Me.gridChequesEmitidos, Column

End Sub

Private Function buscarOP(chequeid As Long) As String
    Set rs = conectar.RSFactory("SELECT op.FECHA,opc.id_cheque FROM ordenes_pago_cheques opc INNER JOIN ordenes_pago op ON opc.id_orden_pago=op.id WHERE opc.id_cheque=" & chequeid)
    If Not rs.EOF And Not rs.BOF Then
        buscarOP = rs!FEcha
    End If
End Function

Private Sub gridChequesEmitidos_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Set tmpCheque = cheques1.item(RowIndex)

    Values(1) = buscarOP(tmpCheque.id)
    Values(2) = tmpCheque.FechaEmision
    Values(3) = tmpCheque.FechaVencimiento
    Values(4) = tmpCheque.numero
    Values(5) = funciones.FormatearDecimales(tmpCheque.Monto)
    Values(6) = tmpCheque.OrigenDestino
    Values(7) = tmpCheque.entro
    Values(8) = tmpCheque.IdOrdenPagoOrigen



End Sub

Private Sub mnuDepositar_Click()
    Dim ff As New frmDepositarCheque

    Set ff.cheque = tmpCheque
    ff.Show
End Sub

Private Sub mnuPasarCartera_Click()
    grid_cheques_DblClick
End Sub

Private Sub PushButton1_Click()
    Dim q As String
    Set cheques1 = New Collection

    q = "ingresado=" & Abs(Me.chkIngresados.value) & " and propio=1"



    If Not IsNull(Me.dtpDesde) Then
        q = q & " and fecha_emision>=" & conectar.Escape(Format(Me.dtpDesde.value, "yyyy-mm-dd"))
    End If

    If Not IsNull(Me.dtpHasta) Then
        q = q & " and fecha_emision<=" & conectar.Escape(Format(Me.dtpHasta.value, "yyyy-mm-dd"))
    End If



    If Me.cboBancos1.ListIndex > -1 Then
        q = q & " and cheqs.id_banco=" & Me.cboBancos1.ItemData(Me.cboBancos1.ListIndex)
    End If


    If Me.cboChequera2.ListIndex > -1 Then
        q = q & " and cheq.id_chequera=" & Me.cboChequera2.ItemData(Me.cboChequera2.ListIndex)
    End If


    If LenB(Me.txtNroCheque) > 0 Then
        q = q & " and cheq.numero=" & Val(Me.txtNroCheque)
    End If

    If LenB(Me.txtIdOP) > 0 Then
        q = q & " and cheq.orden_pago_origen=" & Val(Me.txtIdOP)
    End If


    Me.gridChequesEmitidos.ItemCount = 0
    q = q & "  order by fecha_vencimiento desc"
    Set cheques2 = New Collection
    Set cheques2 = DAOCheques.FindAll(q)

    For Each tmpCheque In cheques2
        If tmpCheque.Monto > 0 Then cheques1.Add tmpCheque


    Next tmpCheque

    Me.gridChequesEmitidos.ItemCount = cheques1.count
    GridEXHelper.AutoSizeColumns Me.gridChequesEmitidos
End Sub

Private Sub PushButton2_Click()
    Dim elegidos As Boolean
    Dim q As String




    If Not IsNull(Me.dtpDesde) Then
        q = "Desde " & Format(Me.dtpDesde, "dd-mm-yyyy") & Chr(10)
    End If
    If Not IsNull(Me.dtpHasta) Then
        q = q & "Hasta " & Format(Me.dtpHasta, "dd-mm-yyyy") & Chr(10)

    End If


    If IsNull(Me.dtpHasta) And IsNull(Me.dtpDesde) Then
        q = "PERIODO SIN ESPECIFICAR"
    End If


    With Me.gridChequesEmitidos.PrinterProperties
        .FitColumns = True
        .RepeatHeaders = True
        .Orientation = jgexPPLandscape
        .HeaderString(jgexHFCenter) = "Listado de cheques"
        .FooterString(jgexHFCenter) = Now
        .HeaderString(jgexHFLeft) = q

    End With
    Load frmPrintPreview
    frmPrintPreview.Move Me.Left, Me.Top, Me.Width, Me.Height
    gridChequesEmitidos.PrintPreview frmPrintPreview.GEXPreview1, elegidos
    frmPrintPreview.Show 1

End Sub

Private Sub PushButton3_Click()
    Me.cboChequera2.ListIndex = -1
End Sub

Private Sub txtDesde_Validate(Cancel As Boolean)
    ValidarTextBox Me.txtDesde, Cancel
End Sub

Private Sub txtHasta_Validate(Cancel As Boolean)
    ValidarTextBox Me.txtHasta, Cancel
End Sub
Private Sub txtIdOP_GotFocus()
    foco Me.txtIdOP
End Sub
Private Sub txtNroCheque_GotFocus()
    foco Me.txtNroCheque
End Sub
Private Sub txtNumero_Validate(Cancel As Boolean)
    funciones.ValidarTextBox Me.txtNumero, Cancel
End Sub
