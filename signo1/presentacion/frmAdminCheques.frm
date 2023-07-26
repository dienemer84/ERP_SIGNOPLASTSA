VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmAdminCheques 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administración de cheques"
   ClientHeight    =   8130
   ClientLeft      =   8250
   ClientTop       =   2265
   ClientWidth     =   15390
   Icon            =   "frmAdminCheques.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8130
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
      ItemCount       =   4
      SelectedItem    =   3
      Item(0).Caption =   "Cartera"
      Item(0).ControlCount=   3
      Item(0).Control(0)=   "grid_cartera_cheques"
      Item(0).Control(1)=   "Frame3"
      Item(0).Control(2)=   "Frame4"
      Item(1).Caption =   "Administrar Chequeras"
      Item(1).ControlCount=   3
      Item(1).Control(0)=   "grid_chequeras"
      Item(1).Control(1)=   "grid_cheques"
      Item(1).Control(2)=   "GroupBox1"
      Item(2).Caption =   "Cheq. Propios Utilizados"
      Item(2).ControlCount=   2
      Item(2).Control(0)=   "gridChequesEmitidos"
      Item(2).Control(1)=   "GroupBox2"
      Item(3).Caption =   "Cheq. 3eros Utilizados"
      Item(3).ControlCount=   3
      Item(3).Control(0)=   "GroupBox3"
      Item(3).Control(1)=   "grpContador"
      Item(3).Control(2)=   "grdCheques3eros"
      Begin VB.Frame Frame4 
         Caption         =   "Filtros"
         Height          =   735
         Left            =   -69880
         TabIndex        =   50
         Top             =   2040
         Visible         =   0   'False
         Width           =   15135
      End
      Begin VB.Frame Frame3 
         Caption         =   "Búsqueda"
         Height          =   1695
         Left            =   -69880
         TabIndex        =   49
         Top             =   360
         Visible         =   0   'False
         Width           =   15135
         Begin VB.TextBox Text1 
            Enabled         =   0   'False
            Height          =   375
            Left            =   480
            TabIndex        =   52
            Top             =   360
            Visible         =   0   'False
            Width           =   2535
         End
         Begin XtremeSuiteControls.PushButton btnBuscarEnCartera 
            Height          =   495
            Left            =   13440
            TabIndex        =   51
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
            Enabled         =   0   'False
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton btnExportarCartera 
            Height          =   495
            Index           =   1
            Left            =   13440
            TabIndex        =   53
            Top             =   960
            Width           =   1575
            _Version        =   786432
            _ExtentX        =   2778
            _ExtentY        =   873
            _StockProps     =   79
            Caption         =   "Exportar"
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
      Begin XtremeSuiteControls.GroupBox grpContador 
         Height          =   615
         Left            =   120
         TabIndex        =   44
         Top             =   2040
         Width           =   15135
         _Version        =   786432
         _ExtentX        =   26696
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "Contador de resultados"
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.Label lblContador 
            Height          =   255
            Left            =   240
            TabIndex        =   46
            Top             =   240
            Width           =   5415
            _Version        =   786432
            _ExtentX        =   9551
            _ExtentY        =   450
            _StockProps     =   79
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin GridEX20.GridEX grdCheques3eros 
         Height          =   5055
         Left            =   120
         TabIndex        =   38
         Top             =   2640
         Width           =   15135
         _ExtentX        =   26696
         _ExtentY        =   8916
         Version         =   "2.0"
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         MethodHoldFields=   -1  'True
         AllowColumnDrag =   0   'False
         AllowEdit       =   0   'False
         DataMode        =   99
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   8
         Column(1)       =   "frmAdminCheques.frx":08E0
         Column(2)       =   "frmAdminCheques.frx":0A24
         Column(3)       =   "frmAdminCheques.frx":0B3C
         Column(4)       =   "frmAdminCheques.frx":0C9C
         Column(5)       =   "frmAdminCheques.frx":0DF0
         Column(6)       =   "frmAdminCheques.frx":0F38
         Column(7)       =   "frmAdminCheques.frx":10A8
         Column(8)       =   "frmAdminCheques.frx":11E4
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmAdminCheques.frx":1328
         FormatStyle(2)  =   "frmAdminCheques.frx":1460
         FormatStyle(3)  =   "frmAdminCheques.frx":1510
         FormatStyle(4)  =   "frmAdminCheques.frx":15C4
         FormatStyle(5)  =   "frmAdminCheques.frx":169C
         FormatStyle(6)  =   "frmAdminCheques.frx":1754
         ImageCount      =   0
         PrinterProperties=   "frmAdminCheques.frx":1834
      End
      Begin XtremeSuiteControls.GroupBox GroupBox3 
         Height          =   1695
         Left            =   120
         TabIndex        =   37
         Top             =   360
         Width           =   15135
         _Version        =   786432
         _ExtentX        =   26696
         _ExtentY        =   2990
         _StockProps     =   79
         Caption         =   "Parámetros de búsqueda"
         UseVisualStyle  =   -1  'True
         Begin VB.Frame Frame2 
            Caption         =   "Fecha"
            Height          =   1215
            Left            =   7680
            TabIndex        =   48
            Top             =   240
            Visible         =   0   'False
            Width           =   3735
         End
         Begin VB.Frame Frame1 
            Caption         =   "Fecha"
            Height          =   1215
            Left            =   4080
            TabIndex        =   47
            Top             =   240
            Visible         =   0   'False
            Width           =   3495
         End
         Begin VB.TextBox txtNumeroOP 
            Height          =   285
            Left            =   1320
            TabIndex        =   42
            Top             =   840
            Width           =   2175
         End
         Begin VB.TextBox txtNumeroCheq 
            Height          =   315
            Left            =   1320
            TabIndex        =   41
            Top             =   360
            Width           =   2175
         End
         Begin XtremeSuiteControls.PushButton btnExportar 
            Height          =   495
            Index           =   0
            Left            =   13320
            TabIndex        =   40
            Top             =   960
            Width           =   1575
            _Version        =   786432
            _ExtentX        =   2778
            _ExtentY        =   873
            _StockProps     =   79
            Caption         =   "Exportar"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton btnBuscar 
            Height          =   495
            Left            =   13320
            TabIndex        =   39
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
         Begin XtremeSuiteControls.Label lblOP 
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   840
            Width           =   1095
            _Version        =   786432
            _ExtentX        =   1931
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "OP"
            Alignment       =   1
         End
         Begin XtremeSuiteControls.Label lblNumero 
            Height          =   375
            Left            =   240
            TabIndex        =   43
            Top             =   330
            Width           =   975
            _Version        =   786432
            _ExtentX        =   1720
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Número"
            Alignment       =   1
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox2 
         Height          =   1425
         Left            =   -69880
         TabIndex        =   19
         Top             =   360
         Visible         =   0   'False
         Width           =   15135
         _Version        =   786432
         _ExtentX        =   26696
         _ExtentY        =   2514
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
            Height          =   465
            Left            =   13440
            TabIndex        =   21
            Top             =   240
            Width           =   1485
            _Version        =   786432
            _ExtentX        =   2619
            _ExtentY        =   820
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
            Height          =   465
            Left            =   13440
            TabIndex        =   26
            Top             =   840
            Width           =   1485
            _Version        =   786432
            _ExtentX        =   2619
            _ExtentY        =   820
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
         Left            =   -69880
         TabIndex        =   2
         Top             =   1845
         Visible         =   0   'False
         Width           =   15090
         _ExtentX        =   26617
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
         Column(1)       =   "frmAdminCheques.frx":1A0C
         Column(2)       =   "frmAdminCheques.frx":1B78
         Column(3)       =   "frmAdminCheques.frx":1CB0
         Column(4)       =   "frmAdminCheques.frx":1DE0
         Column(5)       =   "frmAdminCheques.frx":1F20
         Column(6)       =   "frmAdminCheques.frx":2088
         Column(7)       =   "frmAdminCheques.frx":21A8
         Column(8)       =   "frmAdminCheques.frx":22BC
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmAdminCheques.frx":23D4
         FormatStyle(2)  =   "frmAdminCheques.frx":250C
         FormatStyle(3)  =   "frmAdminCheques.frx":25BC
         FormatStyle(4)  =   "frmAdminCheques.frx":2670
         FormatStyle(5)  =   "frmAdminCheques.frx":2748
         FormatStyle(6)  =   "frmAdminCheques.frx":2800
         ImageCount      =   0
         PrinterProperties=   "frmAdminCheques.frx":28E0
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
         Column(1)       =   "frmAdminCheques.frx":2AB8
         Column(2)       =   "frmAdminCheques.frx":2BD0
         Column(3)       =   "frmAdminCheques.frx":2CE4
         Column(4)       =   "frmAdminCheques.frx":2E1C
         Column(5)       =   "frmAdminCheques.frx":2F2C
         FormatStylesCount=   7
         FormatStyle(1)  =   "frmAdminCheques.frx":2FEC
         FormatStyle(2)  =   "frmAdminCheques.frx":3124
         FormatStyle(3)  =   "frmAdminCheques.frx":31D4
         FormatStyle(4)  =   "frmAdminCheques.frx":3288
         FormatStyle(5)  =   "frmAdminCheques.frx":3360
         FormatStyle(6)  =   "frmAdminCheques.frx":3418
         FormatStyle(7)  =   "frmAdminCheques.frx":34F8
         ImageCount      =   0
         PrinterProperties=   "frmAdminCheques.frx":35B4
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
         Column(1)       =   "frmAdminCheques.frx":378C
         Column(2)       =   "frmAdminCheques.frx":38A4
         Column(3)       =   "frmAdminCheques.frx":39A0
         Column(4)       =   "frmAdminCheques.frx":3A8C
         Column(5)       =   "frmAdminCheques.frx":3B88
         Column(6)       =   "frmAdminCheques.frx":3C84
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmAdminCheques.frx":3DAC
         FormatStyle(2)  =   "frmAdminCheques.frx":3EE4
         FormatStyle(3)  =   "frmAdminCheques.frx":3F94
         FormatStyle(4)  =   "frmAdminCheques.frx":4048
         FormatStyle(5)  =   "frmAdminCheques.frx":4120
         FormatStyle(6)  =   "frmAdminCheques.frx":41D8
         ImageCount      =   0
         PrinterProperties=   "frmAdminCheques.frx":42B8
      End
      Begin GridEX20.GridEX grid_cartera_cheques 
         Height          =   4785
         Left            =   -69880
         TabIndex        =   18
         Top             =   2880
         Visible         =   0   'False
         Width           =   15075
         _ExtentX        =   26591
         _ExtentY        =   8440
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
         Column(1)       =   "frmAdminCheques.frx":4490
         Column(2)       =   "frmAdminCheques.frx":45B8
         Column(3)       =   "frmAdminCheques.frx":46AC
         Column(4)       =   "frmAdminCheques.frx":4800
         Column(5)       =   "frmAdminCheques.frx":4950
         Column(6)       =   "frmAdminCheques.frx":4A60
         Column(7)       =   "frmAdminCheques.frx":4B6C
         Column(8)       =   "frmAdminCheques.frx":4C8C
         Column(9)       =   "frmAdminCheques.frx":4DD4
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmAdminCheques.frx":4EFC
         FormatStyle(2)  =   "frmAdminCheques.frx":5034
         FormatStyle(3)  =   "frmAdminCheques.frx":50E4
         FormatStyle(4)  =   "frmAdminCheques.frx":5198
         FormatStyle(5)  =   "frmAdminCheques.frx":5270
         FormatStyle(6)  =   "frmAdminCheques.frx":5328
         ImageCount      =   0
         PrinterProperties=   "frmAdminCheques.frx":5408
      End
   End
   Begin VB.Menu veOP 
      Caption         =   "Ver OP"
      Visible         =   0   'False
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
Dim cheques3 As New Collection
Dim tmpCheque As cheque
Dim tmpCheque3eros As cheque
Dim bancos As Collection
Dim Banco As Banco


Private Sub btnBuscar_Click()

    Dim q As String
    Set cheques3 = New Collection

    q = "propio=0 and en_cartera=0 and orden_pago_origen!=0"



    '    If Not IsNull(Me.dtpDesde) Then
    '        q = q & " and fecha_emision>=" & conectar.Escape(Format(Me.dtpDesde.value, "yyyy-mm-dd"))
    '    End If
    '
    '    If Not IsNull(Me.dtpHasta) Then
    '        q = q & " and fecha_emision<=" & conectar.Escape(Format(Me.dtpHasta.value, "yyyy-mm-dd"))
    '    End If



    '    If Me.cboBancos1.ListIndex > -1 Then
    '        q = q & " and cheqs.id_banco=" & Me.cboBancos1.ItemData(Me.cboBancos1.ListIndex)
    '    End If
    '
    '
    '    If Me.cboChequera2.ListIndex > -1 Then
    '        q = q & " and cheq.id_chequera=" & Me.cboChequera2.ItemData(Me.cboChequera2.ListIndex)
    '    End If


    If LenB(Me.txtNumeroCheq) > 0 Then
        q = q & " and cheq.numero=" & Val(Me.txtNumeroCheq)
    End If

    If LenB(Me.txtNumeroOP) > 0 Then
        q = q & " and cheq.orden_pago_origen=" & Val(Me.txtNumeroOP)
    End If


    Me.grdCheques3eros.ItemCount = 0

    '    q = q & "  order by fecha_vencimiento desc"

    Set cheques3 = New Collection

    Set cheques3 = DAOCheques.FindAll(q)

    '    For Each tmpCheque3eros In cheques3
    '        If tmpCheque3eros.Monto > 0 Then cheques3.Add tmpCheque3eros
    '
    '    Next tmpCheque3eros



    Me.grdCheques3eros.ItemCount = cheques3.count

    If cheques3.count <> 0 Then

        lblContador.caption = "Cheques encontrados: " & cheques3.count
    Else
        lblContador.caption = "Sin resultados"
    End If

    GridEXHelper.AutoSizeColumns Me.grdCheques3eros

End Sub



Private Sub btnExportar1_Click()

'FUNCIÓN PARA EXPORTAR A EXCEL


'INICIA EL PROGRESSBAR Y LO MUESTRA
'    Me.progreso.Visible = True
'    Me.lblExportando.Visible = True

'DEFINE EL VALOR MINIMO Y EL MAXIMO DEL PROGRESSBAR (CANTIDAD DE DATOS EN LA COLECCIÓN COL)
'    progreso.min = 0
'    progreso.max = facturas.count


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

    xlWorksheet.Cells(1, 1).value = "Reporte de Cheques de 3ros Utilizados"

    xlWorksheet.Columns(4).HorizontalAlignment = xlLeft
    xlWorksheet.Columns(7).HorizontalAlignment = xlLeft

    xlWorksheet.Cells(2, 1).value = "Origen"
    xlWorksheet.Cells(2, 2).value = "Banco"
    xlWorksheet.Cells(2, 3).value = "Fecha Emisión"
    xlWorksheet.Cells(2, 4).value = "Número Cheque"
    xlWorksheet.Cells(2, 5).value = "Importe"
    xlWorksheet.Cells(2, 6).value = "Fecha Vencimiento"
    xlWorksheet.Cells(2, 7).value = "N° OP"


    Dim idx As Integer
    idx = 3

    Dim che As cheque

    'DEFINE EL CONTADOR DEL PROGRESSBAR Y LO INICIA EN 0
    Dim d As Long
    d = 0


    For Each che In cheques3

        Debug.Print

        xlWorksheet.Cells(idx, 1).value = che.OrigenDestino
        xlWorksheet.Cells(idx, 2).value = che.Banco.nombre
        xlWorksheet.Cells(idx, 3).value = che.FechaEmision
        xlWorksheet.Cells(idx, 4).value = che.numero
        xlWorksheet.Cells(idx, 5).value = che.Monto
        xlWorksheet.Cells(idx, 6).value = che.FechaVencimiento
        xlWorksheet.Cells(idx, 7).value = che.IdOrdenPagoOrigen



        idx = idx + 1

        'POR CADA ITERACION SUMA UN VALOR A LA VARIABLE D DEL PROGRESSBAR
        d = d + 1
        '        progreso.value = d


    Next

    xlWorksheet.Cells(idx, 5).Formula = "=SUM(E3:E" & idx - 1 & ")"

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
    '    progreso.value = 0
    '    Me.progreso.Visible = False
    '    Me.lblExportando.Visible = False
End Sub

Private Sub btnExportar_Click(Index As Integer)

'FUNCIÓN PARA EXPORTAR A EXCEL


'INICIA EL PROGRESSBAR Y LO MUESTRA
'    Me.progreso.Visible = True
'    Me.lblExportando.Visible = True

'DEFINE EL VALOR MINIMO Y EL MAXIMO DEL PROGRESSBAR (CANTIDAD DE DATOS EN LA COLECCIÓN COL)
'    progreso.min = 0
'    progreso.max = facturas.count


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

    xlWorksheet.Cells(1, 1).value = "Reporte de Cheques de 3ros Utilizados"

    xlWorksheet.Columns(4).HorizontalAlignment = xlLeft
    xlWorksheet.Columns(7).HorizontalAlignment = xlLeft

    xlWorksheet.Cells(2, 1).value = "Origen"
    xlWorksheet.Cells(2, 2).value = "Banco"
    xlWorksheet.Cells(2, 3).value = "Fecha Emisión"
    xlWorksheet.Cells(2, 4).value = "Número Cheque"
    xlWorksheet.Cells(2, 5).value = "Importe"
    xlWorksheet.Cells(2, 6).value = "Fecha Vencimiento"
    xlWorksheet.Cells(2, 7).value = "N° OP"


    Dim idx As Integer
    idx = 3

    Dim che As cheque

    'DEFINE EL CONTADOR DEL PROGRESSBAR Y LO INICIA EN 0
    Dim d As Long
    d = 0


    For Each che In cheques3

        Debug.Print

        xlWorksheet.Cells(idx, 1).value = che.OrigenDestino
        xlWorksheet.Cells(idx, 2).value = che.Banco.nombre
        xlWorksheet.Cells(idx, 3).value = che.FechaEmision
        xlWorksheet.Cells(idx, 4).value = che.numero
        xlWorksheet.Cells(idx, 5).value = che.Monto
        xlWorksheet.Cells(idx, 6).value = che.FechaVencimiento
        xlWorksheet.Cells(idx, 7).value = che.IdOrdenPagoOrigen



        idx = idx + 1

        'POR CADA ITERACION SUMA UN VALOR A LA VARIABLE D DEL PROGRESSBAR
        d = d + 1
        '        progreso.value = d


    Next

    xlWorksheet.Cells(idx, 5).Formula = "=SUM(E3:E" & idx - 1 & ")"

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
    '    progreso.value = 0
    '    Me.progreso.Visible = False
    '    Me.lblExportando.Visible = False
End Sub

Private Sub btnExportarCartera_Click(Index As Integer)

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

    xlWorksheet.Cells(1, 1).value = "Reporte de Cartera"

    xlWorksheet.Columns(4).HorizontalAlignment = xlLeft
    xlWorksheet.Columns(7).HorizontalAlignment = xlLeft

    xlWorksheet.Cells(2, 1).value = "ID"
    xlWorksheet.Cells(2, 2).value = "NÚMERO"
    xlWorksheet.Cells(2, 3).value = "MONTO"
    xlWorksheet.Cells(2, 4).value = "VENCIMIENTO"
    xlWorksheet.Cells(2, 5).value = "ORIGEN"
    xlWorksheet.Cells(2, 6).value = "BANCO NOMBRE"
    xlWorksheet.Cells(2, 7).value = "CLASIFICACION"
    xlWorksheet.Cells(2, 8).value = "RECIBIDO"
        

    Dim idx As Integer
    idx = 3

    Dim che As cheque

    'DEFINE EL CONTADOR DEL PROGRESSBAR Y LO INICIA EN 0
    Dim d As Long
    d = 0


    For Each che In cartera

        Debug.Print

        xlWorksheet.Cells(idx, 1).value = che.Id
        xlWorksheet.Cells(idx, 2).value = che.numero
        xlWorksheet.Cells(idx, 3).value = che.Monto
        xlWorksheet.Cells(idx, 4).value = che.FechaVencimiento
        xlWorksheet.Cells(idx, 5).value = che.OrigenDestino
        xlWorksheet.Cells(idx, 6).value = che.Banco.nombre
        xlWorksheet.Cells(idx, 7).value = che.OrigenCheque
        xlWorksheet.Cells(idx, 8).value = che.FechaRecibido

        idx = idx + 1

        'POR CADA ITERACION SUMA UN VALOR A LA VARIABLE D DEL PROGRESSBAR
        d = d + 1
        '        progreso.value = d


    Next

    xlWorksheet.Cells(idx, 5).Formula = "=SUM(E3:E" & idx - 1 & ")"

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
    '    progreso.value = 0
    '    Me.progreso.Visible = False
    '    Me.lblExportando.Visible = False


End Sub

Private Sub cmdCrear_Click()
    Dim x As Long
    Dim col As Collection
    Dim id_banco As Long


    If MsgBox("zEstá seguro de crear la chequera?", vbQuestion + vbYesNo) = vbYes Then
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
        Set chequera.moneda = DAOMoneda.GetById(Me.cboMonedas.ItemData(Me.cboMonedas.ListIndex))
        chequera.numero = CLng(Me.txtNumero)
        chequera.NumeroDesde = CLng(Me.txtDesde)
        chequera.NumeroHasta = CLng(Me.txtHasta)
        chequera.observaciones = UCase(Me.txtObservaciones)
        Dim cheque As cheque
        For x = chequera.NumeroDesde To chequera.NumeroHasta
            Set cheque = New cheque
            cheque.numero = x
            cheque.EnCartera = False
            cheque.Propio = True
            cheque.Id = 0
            Set cheque.Banco = chequera.Banco
            Set cheque.moneda = chequera.moneda
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
    GridEXHelper.CustomizeGrid Me.grdCheques3eros, True, True

    DAOBancos.llenarComboXtremeSuite Me.cboBancos
    DAOBancos.llenarComboXtremeSuite Me.cboBancos1
    DAOMoneda.llenarComboXtremeSuite Me.cboMonedas

    DAOChequeras.llenarComboXtremeSuite Me.cboChequera2


    Set bancos = DAOBancos.GetAll("id in (select idBanco from AdminConfigCuentas group by idBanco) ")

    cboBancos1.Clear
    For Each Banco In bancos
        cboBancos1.AddItem Banco.nombre
        cboBancos1.ItemData(cboBancos1.NewIndex) = Banco.Id
    Next

    cboBancos1.ListIndex = -1

    Me.cboChequera2.ListIndex = -1
    Set bancos = DAOBancos.GetAll()
    Me.grid_cheques.ItemCount = 0
    Me.gridBancos.ItemCount = 0
    Me.gridChequesEmitidos.ItemCount = 0
    Me.grdCheques3eros.ItemCount = 0

    MostrarChequeras
    MostrarCartera

    Set Me.grid_cartera_cheques.Columns("banco").DropDownControl = Me.gridBancos
    Me.gridBancos.ItemCount = bancos.count


    Dim idc As Long
    idc = chequeras.item(Me.grid_chequeras.rowIndex(Me.grid_chequeras.row)).Id

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

Private Sub grdCheques3eros_ColumnHeaderClick(ByVal Column As GridEX20.JSColumn)
    GridEXHelper.ColumnHeaderClick Me.grdCheques3eros, Column
End Sub

Private Sub grdCheques3eros_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)

    Set tmpCheque3eros = cheques3.item(rowIndex)

    Values(1) = tmpCheque3eros.OrigenDestino
    Values(2) = tmpCheque3eros.Banco.nombre
    Values(3) = tmpCheque3eros.FechaEmision
    Values(4) = tmpCheque3eros.numero
    Values(5) = funciones.FormatearDecimales(tmpCheque3eros.Monto)
    Values(6) = tmpCheque3eros.FechaVencimiento
    Values(7) = tmpCheque3eros.IdOrdenPagoOrigen
    Values(8) = "De Quien es LA OP"

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
    Set tmpCheque.moneda = DAOMoneda.GetById(0)       ' reemplazar x un combo
    tmpCheque.Monto = Values(2)
    tmpCheque.numero = Values(1)
    tmpCheque.OrigenDestino = Values(4)
    tmpCheque.Propio = False

    If Not DAOCheques.Guardar(tmpCheque) Then GoTo err1
    cartera.Add tmpCheque, CStr(tmpCheque.Id)

    Exit Sub
err1:

End Sub

Private Sub grid_cartera_cheques_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    On Error GoTo err1
    Set tmpCheque = cartera.item(rowIndex)
    With Values
        .value(1) = tmpCheque.Id
        .value(2) = tmpCheque.numero
        .value(3) = funciones.FormatearDecimales(tmpCheque.Monto)
        .value(4) = tmpCheque.FechaVencimiento
        .value(5) = tmpCheque.OrigenDestino
        .value(6) = tmpCheque.Banco.nombre
        .value(7) = tmpCheque.OrigenCheque
        .value(8) = tmpCheque.FechaRecibido
        .value(9) = tmpCheque.observaciones

    End With




    Exit Sub
err1:


End Sub

Private Sub grid_cartera_cheques_UnboundUpdate(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    On Error GoTo err1
    Dim ant As String

    ant = tmpCheque.OrigenDestino

    Set tmpCheque = cartera.item(rowIndex)
    tmpCheque.OrigenDestino = Values(4)
    If Not DAOCheques.Guardar(tmpCheque) Then GoTo err1
    Exit Sub
err1:
    tmpCheque.OrigenDestino = ant
End Sub

Private Sub grid_chequeras_SelectionChange()
    On Error Resume Next
    Set tmpChequera.Cheques = DAOCheques.FindAllByChequeraId(tmpChequera.Id)
    mostrarCheques
End Sub
Private Sub mostrarCheques()
    Me.grid_cheques.ItemCount = 0
    Me.grid_cheques.ItemCount = tmpChequera.Cheques.count
End Sub
Private Sub grid_chequeras_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Set tmpChequera = chequeras.item(rowIndex)
    With Values
        .value(1) = tmpChequera.numero
        .value(2) = tmpChequera.FechaCreacion
        .value(3) = tmpChequera.Banco.nombre
        .value(4) = tmpChequera.NumeroDesde
        .value(5) = tmpChequera.NumeroHasta

    End With

End Sub


Private Sub grid_cheques_DblClick()
    If Me.grid_cheques.rowIndex(Me.grid_cheques.row) > 0 Then
        Set tmpCheque = tmpChequera.Cheques(Me.grid_cheques.rowIndex(Me.grid_cheques.row))
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
        Set tmpCheque = tmpChequera.Cheques(Me.grid_cheques.rowIndex(Me.grid_cheques.row))
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

Private Sub grid_cheques_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    On Error GoTo err1
    If Not IsSomething(tmpChequera.Cheques) Or tmpChequera.Cheques.count = 0 Then Set tmpChequera.Cheques = DAOCheques.FindAllByChequeraId(tmpChequera.Id)
    Set tmpCheque = tmpChequera.Cheques(rowIndex)
    With Values
        .value(1) = tmpCheque.numero
        .value(2) = IIf(tmpCheque.Utilizado, funciones.FormatearDecimales(tmpCheque.Monto), Empty)
        .value(3) = IIf(tmpCheque.Utilizado, tmpCheque.FechaVencimiento, Empty)
        .value(4) = IIf(tmpCheque.Utilizado, tmpCheque.OrigenDestino, Empty)
        '.value(5) = IIf(tmpCheque.Utilizado, tmpCheque.Observaciones, "DISPONIBLE")
        .value(5) = IIf(tmpCheque.estado = ChequeAnulado, "ANULADO", IIf(tmpCheque.Utilizado, "Utilizado en Orden de Pago Ns " & tmpCheque.IdOrdenPagoOrigen, "DISPONIBLE"))
    End With
    Exit Sub
err1:
End Sub

Private Sub gridBancos_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If rowIndex <= bancos.count Then
        Set Banco = bancos.item(rowIndex)
        Values(1) = Banco.Id
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

Private Sub gridChequesEmitidos_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Set tmpCheque = cheques1.item(rowIndex)

    Values(1) = buscarOP(tmpCheque.Id)
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

