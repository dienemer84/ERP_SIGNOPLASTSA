VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmAdminCobranzasNuevoReciboAnticipo 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recibo de Anticipo"
   ClientHeight    =   10050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11295
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10050
   ScaleWidth      =   11295
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Totales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   8040
      TabIndex        =   20
      Top             =   120
      Width           =   3135
      Begin VB.Label lblTotalCheques 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Total Cheques:"
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
         Left            =   120
         TabIndex        =   24
         Top             =   1080
         Width           =   1305
      End
      Begin VB.Label lblTotalBanco 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Total Banco:"
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
         Left            =   135
         TabIndex        =   23
         Top             =   720
         Width           =   1110
      End
      Begin VB.Label lblTotalCaja 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Total Caja:"
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
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   945
      End
      Begin VB.Label lblTotalRecibido 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Total Recibido:"
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
         Left            =   120
         TabIndex        =   21
         Top             =   1800
         Width           =   1320
      End
   End
   Begin VB.ComboBox cboMonedas 
      Height          =   315
      Left            =   11400
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   2040
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Valores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7635
      Left            =   120
      TabIndex        =   6
      Top             =   2400
      Width           =   11040
      Begin XtremeSuiteControls.GroupBox grpCheques 
         Height          =   3225
         Left            =   120
         TabIndex        =   7
         Top             =   4200
         Width           =   10770
         _Version        =   786432
         _ExtentX        =   18997
         _ExtentY        =   5689
         _StockProps     =   79
         Caption         =   "Cheques"
         UseVisualStyle  =   -1  'True
         Begin GridEX20.GridEX gridCheques 
            Height          =   2760
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   10500
            _ExtentX        =   18521
            _ExtentY        =   4868
            Version         =   "2.0"
            BoundColumnIndex=   ""
            ReplaceColumnIndex=   ""
            ColumnAutoResize=   -1  'True
            MethodHoldFields=   -1  'True
            ContScroll      =   -1  'True
            AllowDelete     =   -1  'True
            GroupByBoxVisible=   0   'False
            RowHeaders      =   -1  'True
            DataMode        =   99
            AllowAddNew     =   -1  'True
            ColumnHeaderHeight=   285
            IntProp1        =   0
            IntProp2        =   0
            IntProp7        =   0
            ColumnsCount    =   7
            Column(1)       =   "frmAdminCobranzasNuevoReciboAnticipo.frx":0000
            Column(2)       =   "frmAdminCobranzasNuevoReciboAnticipo.frx":0160
            Column(3)       =   "frmAdminCobranzasNuevoReciboAnticipo.frx":029C
            Column(4)       =   "frmAdminCobranzasNuevoReciboAnticipo.frx":03D8
            Column(5)       =   "frmAdminCobranzasNuevoReciboAnticipo.frx":04E4
            Column(6)       =   "frmAdminCobranzasNuevoReciboAnticipo.frx":05FC
            Column(7)       =   "frmAdminCobranzasNuevoReciboAnticipo.frx":0710
            FormatStylesCount=   6
            FormatStyle(1)  =   "frmAdminCobranzasNuevoReciboAnticipo.frx":07E0
            FormatStyle(2)  =   "frmAdminCobranzasNuevoReciboAnticipo.frx":0918
            FormatStyle(3)  =   "frmAdminCobranzasNuevoReciboAnticipo.frx":09C8
            FormatStyle(4)  =   "frmAdminCobranzasNuevoReciboAnticipo.frx":0A7C
            FormatStyle(5)  =   "frmAdminCobranzasNuevoReciboAnticipo.frx":0B54
            FormatStyle(6)  =   "frmAdminCobranzasNuevoReciboAnticipo.frx":0C0C
            ImageCount      =   0
            PrinterProperties=   "frmAdminCobranzasNuevoReciboAnticipo.frx":0CEC
         End
         Begin GridEX20.GridEX gridBancos 
            Height          =   1845
            Left            =   240
            TabIndex        =   9
            Top             =   840
            Visible         =   0   'False
            Width           =   3705
            _ExtentX        =   6535
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
            Column(1)       =   "frmAdminCobranzasNuevoReciboAnticipo.frx":0EC4
            Column(2)       =   "frmAdminCobranzasNuevoReciboAnticipo.frx":0FC4
            FormatStylesCount=   6
            FormatStyle(1)  =   "frmAdminCobranzasNuevoReciboAnticipo.frx":10B8
            FormatStyle(2)  =   "frmAdminCobranzasNuevoReciboAnticipo.frx":11F0
            FormatStyle(3)  =   "frmAdminCobranzasNuevoReciboAnticipo.frx":12A0
            FormatStyle(4)  =   "frmAdminCobranzasNuevoReciboAnticipo.frx":1354
            FormatStyle(5)  =   "frmAdminCobranzasNuevoReciboAnticipo.frx":142C
            FormatStyle(6)  =   "frmAdminCobranzasNuevoReciboAnticipo.frx":14E4
            ImageCount      =   0
            PrinterProperties=   "frmAdminCobranzasNuevoReciboAnticipo.frx":15C4
         End
      End
      Begin XtremeSuiteControls.GroupBox grpBanco 
         Height          =   2040
         Left            =   120
         TabIndex        =   10
         Top             =   2160
         Width           =   10770
         _Version        =   786432
         _ExtentX        =   18997
         _ExtentY        =   3598
         _StockProps     =   79
         Caption         =   "Banco"
         UseVisualStyle  =   -1  'True
         Begin GridEX20.GridEX gridDepositosOperaciones 
            Height          =   1545
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   10545
            _ExtentX        =   18600
            _ExtentY        =   2725
            Version         =   "2.0"
            BoundColumnIndex=   ""
            ReplaceColumnIndex=   ""
            ColumnAutoResize=   -1  'True
            MethodHoldFields=   -1  'True
            ContScroll      =   -1  'True
            AllowDelete     =   -1  'True
            GroupByBoxVisible=   0   'False
            RowHeaders      =   -1  'True
            ItemCount       =   3
            DataMode        =   99
            AllowAddNew     =   -1  'True
            ColumnHeaderHeight=   285
            IntProp1        =   0
            IntProp2        =   0
            IntProp7        =   0
            ColumnsCount    =   4
            Column(1)       =   "frmAdminCobranzasNuevoReciboAnticipo.frx":179C
            Column(2)       =   "frmAdminCobranzasNuevoReciboAnticipo.frx":18FC
            Column(3)       =   "frmAdminCobranzasNuevoReciboAnticipo.frx":1A38
            Column(4)       =   "frmAdminCobranzasNuevoReciboAnticipo.frx":1B6C
            FormatStylesCount=   6
            FormatStyle(1)  =   "frmAdminCobranzasNuevoReciboAnticipo.frx":1CB0
            FormatStyle(2)  =   "frmAdminCobranzasNuevoReciboAnticipo.frx":1DE8
            FormatStyle(3)  =   "frmAdminCobranzasNuevoReciboAnticipo.frx":1E98
            FormatStyle(4)  =   "frmAdminCobranzasNuevoReciboAnticipo.frx":1F4C
            FormatStyle(5)  =   "frmAdminCobranzasNuevoReciboAnticipo.frx":2024
            FormatStyle(6)  =   "frmAdminCobranzasNuevoReciboAnticipo.frx":20DC
            ImageCount      =   0
            PrinterProperties=   "frmAdminCobranzasNuevoReciboAnticipo.frx":21BC
         End
      End
      Begin XtremeSuiteControls.GroupBox grpCaja 
         Height          =   1875
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   10770
         _Version        =   786432
         _ExtentX        =   18997
         _ExtentY        =   3307
         _StockProps     =   79
         Caption         =   "Caja"
         UseVisualStyle  =   -1  'True
         Begin GridEX20.GridEX gridCajaOperaciones 
            Height          =   1500
            Left            =   120
            TabIndex        =   13
            Top             =   225
            Width           =   10530
            _ExtentX        =   18574
            _ExtentY        =   2646
            Version         =   "2.0"
            BoundColumnIndex=   ""
            ReplaceColumnIndex=   ""
            ColumnAutoResize=   -1  'True
            MethodHoldFields=   -1  'True
            ContScroll      =   -1  'True
            AllowDelete     =   -1  'True
            GroupByBoxVisible=   0   'False
            RowHeaders      =   -1  'True
            DataMode        =   99
            AllowAddNew     =   -1  'True
            ColumnHeaderHeight=   285
            IntProp1        =   0
            IntProp2        =   0
            IntProp7        =   0
            ColumnsCount    =   4
            Column(1)       =   "frmAdminCobranzasNuevoReciboAnticipo.frx":2394
            Column(2)       =   "frmAdminCobranzasNuevoReciboAnticipo.frx":24F4
            Column(3)       =   "frmAdminCobranzasNuevoReciboAnticipo.frx":2630
            Column(4)       =   "frmAdminCobranzasNuevoReciboAnticipo.frx":2764
            FormatStylesCount=   6
            FormatStyle(1)  =   "frmAdminCobranzasNuevoReciboAnticipo.frx":2898
            FormatStyle(2)  =   "frmAdminCobranzasNuevoReciboAnticipo.frx":29D0
            FormatStyle(3)  =   "frmAdminCobranzasNuevoReciboAnticipo.frx":2A80
            FormatStyle(4)  =   "frmAdminCobranzasNuevoReciboAnticipo.frx":2B34
            FormatStyle(5)  =   "frmAdminCobranzasNuevoReciboAnticipo.frx":2C0C
            FormatStyle(6)  =   "frmAdminCobranzasNuevoReciboAnticipo.frx":2CC4
            ImageCount      =   0
            PrinterProperties=   "frmAdminCobranzasNuevoReciboAnticipo.frx":2DA4
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos del Recibo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7815
      Begin VB.TextBox txtRedondeo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1200
         TabIndex        =   26
         Text            =   "0"
         Top             =   1695
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.ComboBox cboClientes 
         Height          =   315
         Left            =   1110
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   735
         Width           =   4860
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   300
         Left            =   1110
         TabIndex        =   2
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         Format          =   58458113
         CurrentDate     =   39199
      End
      Begin XtremeSuiteControls.PushButton cmdGuardar 
         Height          =   405
         Index           =   0
         Left            =   6240
         TabIndex        =   25
         Top             =   1560
         Width           =   1425
         _Version        =   786432
         _ExtentX        =   2514
         _ExtentY        =   714
         _StockProps     =   79
         Caption         =   "Guardar"
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
      End
      Begin VB.Label lblNumeroRecibo 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Número:"
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
         Left            =   360
         TabIndex        =   5
         Top             =   360
         Width           =   705
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha:"
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
         Left            =   150
         TabIndex        =   4
         Top             =   1230
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
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
         Height          =   255
         Left            =   30
         TabIndex        =   3
         Top             =   795
         Width           =   975
      End
   End
   Begin GridEX20.GridEX gridCajas 
      Height          =   1095
      Left            =   14160
      TabIndex        =   14
      Top             =   120
      Visible         =   0   'False
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   1931
      Version         =   "2.0"
      BoundColumnIndex=   "id"
      ReplaceColumnIndex=   "caja"
      ActAsDropDown   =   -1  'True
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
      Column(1)       =   "frmAdminCobranzasNuevoReciboAnticipo.frx":2F7C
      Column(2)       =   "frmAdminCobranzasNuevoReciboAnticipo.frx":30A0
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminCobranzasNuevoReciboAnticipo.frx":318C
      FormatStyle(2)  =   "frmAdminCobranzasNuevoReciboAnticipo.frx":32C4
      FormatStyle(3)  =   "frmAdminCobranzasNuevoReciboAnticipo.frx":3374
      FormatStyle(4)  =   "frmAdminCobranzasNuevoReciboAnticipo.frx":3428
      FormatStyle(5)  =   "frmAdminCobranzasNuevoReciboAnticipo.frx":3500
      FormatStyle(6)  =   "frmAdminCobranzasNuevoReciboAnticipo.frx":35B8
      ImageCount      =   0
      PrinterProperties=   "frmAdminCobranzasNuevoReciboAnticipo.frx":3698
   End
   Begin GridEX20.GridEX gridCuentasBancarias 
      Height          =   1095
      Left            =   11400
      TabIndex        =   15
      Top             =   4680
      Visible         =   0   'False
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   1931
      Version         =   "2.0"
      BoundColumnIndex=   "id"
      ReplaceColumnIndex=   "cuenta"
      ActAsDropDown   =   -1  'True
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
      Column(1)       =   "frmAdminCobranzasNuevoReciboAnticipo.frx":3870
      Column(2)       =   "frmAdminCobranzasNuevoReciboAnticipo.frx":3994
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminCobranzasNuevoReciboAnticipo.frx":3A88
      FormatStyle(2)  =   "frmAdminCobranzasNuevoReciboAnticipo.frx":3BC0
      FormatStyle(3)  =   "frmAdminCobranzasNuevoReciboAnticipo.frx":3C70
      FormatStyle(4)  =   "frmAdminCobranzasNuevoReciboAnticipo.frx":3D24
      FormatStyle(5)  =   "frmAdminCobranzasNuevoReciboAnticipo.frx":3DFC
      FormatStyle(6)  =   "frmAdminCobranzasNuevoReciboAnticipo.frx":3EB4
      ImageCount      =   0
      PrinterProperties=   "frmAdminCobranzasNuevoReciboAnticipo.frx":3F94
   End
   Begin GridEX20.GridEX gridMonedas 
      Height          =   1095
      Left            =   12960
      TabIndex        =   16
      Top             =   3120
      Visible         =   0   'False
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   1931
      Version         =   "2.0"
      BoundColumnIndex=   "id"
      ReplaceColumnIndex=   "moneda"
      ActAsDropDown   =   -1  'True
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
      Column(1)       =   "frmAdminCobranzasNuevoReciboAnticipo.frx":416C
      Column(2)       =   "frmAdminCobranzasNuevoReciboAnticipo.frx":4290
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminCobranzasNuevoReciboAnticipo.frx":4384
      FormatStyle(2)  =   "frmAdminCobranzasNuevoReciboAnticipo.frx":44BC
      FormatStyle(3)  =   "frmAdminCobranzasNuevoReciboAnticipo.frx":456C
      FormatStyle(4)  =   "frmAdminCobranzasNuevoReciboAnticipo.frx":4620
      FormatStyle(5)  =   "frmAdminCobranzasNuevoReciboAnticipo.frx":46F8
      FormatStyle(6)  =   "frmAdminCobranzasNuevoReciboAnticipo.frx":47B0
      ImageCount      =   0
      PrinterProperties=   "frmAdminCobranzasNuevoReciboAnticipo.frx":4890
   End
   Begin XtremeSuiteControls.ComboBox ComboBox1 
      Height          =   315
      Left            =   11520
      TabIndex        =   17
      Top             =   1080
      Visible         =   0   'False
      Width           =   2460
      _Version        =   786432
      _ExtentX        =   4339
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Text            =   "ComboBox1"
   End
   Begin VB.Label lblTotalRecibo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   12360
      TabIndex        =   19
      Top             =   2040
      Visible         =   0   'False
      Width           =   1410
   End
End
Attribute VB_Name = "frmAdminCobranzasNuevoReciboAnticipo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private dataLoaded As Boolean
Private clienteChange As Boolean
Private cheque As cheque
Private recibo As recibo
Private retencionRecibo As retencionRecibo
Private retenciones As New Collection
Private Retencion As Retencion
Private facturasCliente As New Collection
Private Factura As Factura
Private bancos As New Collection
Private Banco As Banco
Private CuentaBancaria As CuentaBancaria
Private cuentasBancarias As New Collection
Private Monedas As New Collection
Private moneda As clsMoneda
Private operacion As operacion
Private Cajas As New Collection
Private caja As caja
Private RecibosACuenta As New Collection
Private Editar_ As Boolean
Public Property Let editar(nvalue As Boolean)
    Editar_ = nvalue
End Property


Public Property Let reciboId(nIdRecibo As Long)

    Set recibo = DAOReciboAnticipo.FindById(nIdRecibo, True, True, True, True, True)

    If recibo Is Nothing Then
        MsgBox "recibo no encontrado, cierre pantalla", vbCritical
    End If

    Me.dtpFecha.value = recibo.FEcha
    
    lblNumeroRecibo.caption = "Número: " & recibo.id
    
    Me.txtRedondeo.text = recibo.redondeo

    Set Cajas = DAOCaja.FindAll()
    Me.gridCajas.ItemCount = Cajas.count

    Set Monedas = DAOMoneda.GetAll()
    Me.gridMonedas.ItemCount = Monedas.count

    Set cuentasBancarias = DAOCuentaBancaria.FindAll()
    Me.gridCuentasBancarias.ItemCount = cuentasBancarias.count

    Set bancos = DAOBancos.GetAll()
    Me.gridBancos.ItemCount = bancos.count

   Set Me.gridCheques.Columns("banco").DropDownControl = Me.gridBancos
    Set Me.gridCheques.Columns("moneda").DropDownControl = Me.gridMonedas

    Set Me.gridDepositosOperaciones.Columns("moneda").DropDownControl = Me.gridMonedas
    Set Me.gridDepositosOperaciones.Columns("cuenta").DropDownControl = Me.gridCuentasBancarias

    Set Me.gridCajaOperaciones.Columns("caja").DropDownControl = Me.gridCajas
    Set Me.gridCajaOperaciones.Columns("moneda").DropDownControl = Me.gridMonedas

    DAOMoneda.LlenarCombo Me.cboMonedas
    Me.cboMonedas.ListIndex = PosIndexCbo(recibo.moneda.id, Me.cboMonedas)
    
    DAOCliente.LlenarCombo Me.cboClientes, True, True
    Me.cboClientes.ListIndex = PosIndexCbo(recibo.cliente.id, Me.cboClientes)

    Me.gridCajaOperaciones.ItemCount = recibo.operacionesCaja.count
    Me.gridDepositosOperaciones.ItemCount = recibo.operacionesBanco.count
    Me.gridCheques.ItemCount = recibo.cheques.count

    Totalizar

    Me.cboMonedas.Enabled = Editar_
    Me.txtRedondeo.Enabled = Editar_
    Me.gridCajaOperaciones.AllowEdit = Editar_
    Me.gridBancos.AllowEdit = Editar_
    Me.gridCheques.AllowEdit = Editar_
    Me.gridDepositosOperaciones.AllowEdit = Editar_
    Me.gridCajaOperaciones.AllowDelete = Editar_
    Me.gridCheques.AllowDelete = Editar_
    Me.gridDepositosOperaciones.AllowDelete = Editar_

    Me.gridCajaOperaciones.AllowAddNew = Editar_
    Me.gridCheques.AllowAddNew = Editar_
    gridDepositosOperaciones.AllowAddNew = Editar_

    Me.Frame1.Enabled = Editar_
    Me.Frame5.Enabled = Editar_
    
    Me.cmdGuardar(0).Enabled = Editar_
    
    dataLoaded = True
    
End Property


Private Sub cboClientes_Click()
    If clienteChange Then Exit Sub

    If Me.cboClientes.ListIndex <> -1 Then
        If dataLoaded Then
            If recibo.facturas.count > 0 Then
                If MsgBox("Va a cambiar de cliente y perder las facturas." & vbNewLine & "¿Desea continuar?", vbQuestion + vbYesNo) = vbNo Then
                    clienteChange = True
                    Me.cboClientes.ListIndex = funciones.PosIndexCbo(recibo.cliente.id, Me.cboClientes)
                    clienteChange = False
                    Exit Sub

                End If
            End If

            Set recibo.cliente = DAOCliente.BuscarPorID(Me.cboClientes.ItemData(Me.cboClientes.ListIndex))

        End If
    End If

    clienteChange = False
    
End Sub

Private Sub cboMonedas_Click()
    If Me.cboMonedas.ListIndex <> -1 And dataLoaded Then
        Set recibo.moneda = DAOMoneda.GetById(Me.cboMonedas.ItemData(Me.cboMonedas.ListIndex))
        Totalizar
    End If
End Sub


Private Sub cmdGuardarAnticipo_Click()
    If Not recibo.IsValid Then
        MsgBox recibo.ValidationMessages, vbExclamation
        Exit Sub
    End If

    If DAOReciboAnticipo.Save(recibo) Then
        MsgBox "Recibo guardado.", vbInformation
        'Unload Me
    Else
        MsgBox "Hubo un error al intentar guardar el recibo.", vbCritical
    End If

End Sub


Private Sub Totalizar()

    Me.lblTotalCheques.caption = "Total Cheques: " & Replace(FormatCurrency(funciones.FormatearDecimales(recibo.TotalCheques)), "$", "")
    Me.lblTotalBanco.caption = "Total Banco: " & Replace(FormatCurrency(funciones.FormatearDecimales(recibo.TotalOperacionesBanco)), "$", "")
    Me.lblTotalCaja.caption = "Total Caja: " & Replace(FormatCurrency(funciones.FormatearDecimales(recibo.TotalOperacionesCaja)), "$", "")

    Dim totalRecibo As Double
    totalRecibo = funciones.FormatearDecimales(recibo.Total)
    Dim totalCancelado As Double
    totalCancelado = funciones.FormatearDecimales(recibo.TotalRecibido)

    Me.lblTotalRecibo.caption = funciones.FormatearDecimales(totalRecibo)
    Me.lblTotalRecibido.caption = "Total Recibido: " & Replace(FormatCurrency(funciones.FormatearDecimales(totalCancelado)), "$", "")

    If totalCancelado < totalRecibo Then
        lblTotalRecibo.backColor = vbRed
    ElseIf totalCancelado = totalRecibo Then
        lblTotalRecibo.backColor = vbYellow
    ElseIf totalCancelado > totalRecibo Then
        lblTotalRecibo.backColor = vbGreen
    End If

    'Me.lblDiferencia.caption = Me.lblDiferencia.Tag & funciones.FormatearDecimales(totalCancelado - MonedaConverter.Convertir(totalRecibo, recibo.moneda.id, DAOMoneda.MONEDA_PESO_ID))
    Debug.Print MonedaConverter.Convertir(totalRecibo, recibo.moneda.id, DAOMoneda.MONEDA_PESO_ID)
    recibo.aCuenta = (totalCancelado - totalRecibo)
End Sub

Private Sub cmdGuardar_Click(Index As Integer)
    If Not recibo.IsValid Then
        MsgBox recibo.ValidationMessages, vbExclamation
        Exit Sub
    End If

    If DAOReciboAnticipo.Save(recibo) Then
        MsgBox "Recibo guardado.", vbInformation
        'Unload Me
    Else
        MsgBox "Hubo un error al intentar guardar el recibo.", vbCritical
    End If
End Sub

Private Sub dtpFecha_LostFocus()
    recibo.FEcha = Me.dtpFecha.value
End Sub

Private Sub Form_Load()
    dataLoaded = False

    FormHelper.Customize Me

    GridEXHelper.CustomizeGrid Me.gridCheques, False, True
    GridEXHelper.CustomizeGrid Me.gridBancos, False, False

    GridEXHelper.CustomizeGrid Me.gridCuentasBancarias, False, False
    GridEXHelper.CustomizeGrid Me.gridMonedas, False, False
    GridEXHelper.CustomizeGrid Me.gridCajas, False, False

    GridEXHelper.CustomizeGrid Me.gridDepositosOperaciones, False, True
    GridEXHelper.CustomizeGrid Me.gridCajaOperaciones, False, True

End Sub

'Private Sub VerRecibosConSaldoACuenta()
'    Set RecibosACuenta = New Collection
'    Set RecibosACuenta = DAOReciboAnticipo.FindAll("(a_cuenta-a_cuenta_usado) >0.02")
'End Sub

Private Sub gridBancos_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex <= bancos.count Then
        Set Banco = bancos.item(RowIndex)
        Values(1) = Banco.id
        Values(2) = Banco.nombre
    End If

End Sub

Private Sub gridCajaOperaciones_BeforeUpdate(ByVal Cancel As GridEX20.JSRetBoolean)
    Cancel = _
    Not IsNumeric(Me.gridCajaOperaciones.value(1)) Or _
             (Not IsNumeric(Me.gridCajaOperaciones.value(2)) And LenB(Me.gridCajaOperaciones.value(2)) = 0) Or _
             Not IsDate(Me.gridCajaOperaciones.value(3)) Or _
             (Not IsNumeric(Me.gridCajaOperaciones.value(4)) And LenB(Me.gridCajaOperaciones.value(4)) = 0)
End Sub

Private Sub gridCajaOperaciones_UnboundAddNew(ByVal NewRowBookmark As GridEX20.JSRetVariant, ByVal Values As GridEX20.JSRowData)
    Set operacion = New operacion
    operacion.IdPertenencia = recibo.id
    operacion.Pertenencia = OrigenOperacion.caja
    operacion.Monto = Values(1)
    If IsNumeric(Values(2)) Then
        Set operacion.moneda = DAOMoneda.GetById(Values(2))
    End If
    operacion.FechaOperacion = Values(3)
    If IsNumeric(Values(4)) Then
        Set operacion.caja = DAOCaja.FindById(Values(4))
    End If
    operacion.EntradaSalida = OPEntrada
    recibo.operacionesCaja.Add operacion

    Totalizar
End Sub

Private Sub gridCajaOperaciones_UnboundDelete(ByVal RowIndex As Long, ByVal Bookmark As Variant)
    If RowIndex > 0 And recibo.operacionesCaja.count >= RowIndex Then
        recibo.operacionesCaja.remove RowIndex
        Totalizar
    End If
End Sub

Private Sub gridCajaOperaciones_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex <= recibo.operacionesCaja.count Then
        Set operacion = recibo.operacionesCaja.item(RowIndex)
        Values(1) = funciones.FormatearDecimales(operacion.Monto)
        If IsSomething(operacion.moneda) Then
            Values(2) = operacion.moneda.NombreCorto
        End If
        Values(3) = operacion.FechaOperacion
        If IsSomething(operacion.caja) Then
            Values(4) = operacion.caja.nombre
        End If
        Totalizar
    End If
End Sub

Private Sub gridCajaOperaciones_UnboundUpdate(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex > 0 And recibo.operacionesCaja.count > 0 Then
        Set operacion = recibo.operacionesCaja.item(RowIndex)

        operacion.Monto = Values(1)
        If IsNumeric(Values(2)) Then
            Set operacion.moneda = DAOMoneda.GetById(Values(2))
        End If
        operacion.FechaOperacion = Values(3)
        If IsNumeric(Values(4)) Then
            Set operacion.caja = DAOCaja.FindById(Values(4))
        End If
        operacion.EntradaSalida = OPEntrada
        Totalizar
        
    End If
End Sub

Private Sub gridCajas_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex > 0 And Cajas.count > 0 Then
        Set caja = Cajas.item(RowIndex)
        Values(1) = caja.id
        Values(2) = caja.nombre
    End If
End Sub

Private Sub gridCheques_BeforeUpdate(ByVal Cancel As GridEX20.JSRetBoolean)
    If Me.gridCheques.row = -1 Then
        If Val(Me.gridCheques.value(6)) = -1 Then
            Me.gridCheques.value(7) = recibo.cliente.razon
        End If
    End If

    Cancel = _
    LenB(Me.gridCheques.value(1)) = 0 Or _
             Not IsNumeric(Me.gridCheques.value(2)) Or _
             (Not IsNumeric(Me.gridCheques.value(3)) And LenB(Me.gridCheques.value(3)) = 0) Or _
             (Not IsNumeric(Me.gridCheques.value(4)) And LenB(Me.gridCheques.value(4)) = 0) Or _
             Not IsDate(Me.gridCheques.value(5))
End Sub

Private Sub gridCheques_UnboundAddNew(ByVal NewRowBookmark As GridEX20.JSRetVariant, ByVal Values As GridEX20.JSRowData)
    Set cheque = New cheque
    If IsNumeric(Values(4)) Then Set cheque.Banco = DAOBancos.GetById(Values(4))
    cheque.EnCartera = True
    cheque.FechaRecibido = recibo.FEcha
    If IsNumeric(Values(3)) Then Set cheque.moneda = DAOMoneda.GetById(Values(3))
    cheque.Monto = funciones.RedondearDecimales(Val(Values(2)))
    cheque.numero = Values(1)
    cheque.Propio = False
    cheque.FechaVencimiento = Values(5)

    cheque.TercerosPropio = (Values(6) = -1)

    If cheque.TercerosPropio Then
        cheque.OrigenDestino = recibo.cliente.razon
    Else
        cheque.OrigenDestino = Values(7)
    End If
    recibo.cheques.Add cheque

    Totalizar
End Sub

Private Sub gridCheques_UnboundDelete(ByVal RowIndex As Long, ByVal Bookmark As Variant)
    If RowIndex > 0 Then
        recibo.cheques.remove RowIndex
        Totalizar
    End If
End Sub

Private Sub gridCheques_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex <= recibo.cheques.count Then
        Set cheque = recibo.cheques.item(RowIndex)
        Values(1) = cheque.numero
        Values(2) = funciones.FormatearDecimales(cheque.Monto)
        Values(6) = cheque.TercerosPropio
        Values(7) = cheque.OrigenDestino
        If IsSomething(cheque.moneda) Then Values(3) = cheque.moneda.NombreCorto
        If IsSomething(cheque.Banco) Then Values(4) = cheque.Banco.nombre
        If CDbl(cheque.FechaVencimiento) > 0 Then Values(5) = cheque.FechaVencimiento
        Totalizar
    End If
End Sub

Private Sub gridCheques_UnboundUpdate(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex > 0 And recibo.cheques.count >= RowIndex Then
        Set cheque = recibo.cheques.item(RowIndex)
        If IsNumeric(Values(4)) Then Set cheque.Banco = DAOBancos.GetById(Values(4))
        cheque.EnCartera = True
        cheque.FechaRecibido = recibo.FEcha
        If IsNumeric(Values(3)) Then Set cheque.moneda = DAOMoneda.GetById(Values(3))
        cheque.Monto = Val(Values(2))
        cheque.numero = Values(1)
        cheque.Propio = False
        cheque.FechaVencimiento = Values(5)
        cheque.OrigenDestino = Values(7)

        cheque.TercerosPropio = (Values(6) = -1)

        Totalizar
        
    End If
End Sub

Private Sub gridCuentasBancarias_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If cuentasBancarias.count >= RowIndex Then
        Set CuentaBancaria = cuentasBancarias.item(RowIndex)
        Values(1) = CuentaBancaria.id
        Values(2) = CuentaBancaria.DescripcionFormateada
    End If
End Sub

Private Sub gridDepositosOperaciones_BeforeUpdate(ByVal Cancel As GridEX20.JSRetBoolean)
    Cancel = _
    Not IsNumeric(Me.gridDepositosOperaciones.value(1)) Or _
             (Not IsNumeric(Me.gridDepositosOperaciones.value(2)) And LenB(Me.gridDepositosOperaciones.value(2)) = 0) Or _
             Not IsDate(Me.gridDepositosOperaciones.value(3)) Or _
             (Not IsNumeric(Me.gridDepositosOperaciones.value(4)) And LenB(Me.gridDepositosOperaciones.value(4)) = 0)
End Sub

Private Sub gridDepositosOperaciones_UnboundAddNew(ByVal NewRowBookmark As GridEX20.JSRetVariant, ByVal Values As GridEX20.JSRowData)
    Set operacion = New operacion
    operacion.IdPertenencia = recibo.id
    operacion.Pertenencia = OrigenOperacion.Banco
    operacion.Monto = Values(1)
    If IsNumeric(Values(2)) Then
        Set operacion.moneda = DAOMoneda.GetById(Values(2))
    End If
    operacion.FechaOperacion = Values(3)
    If IsNumeric(Values(4)) Then
        Set operacion.CuentaBancaria = DAOCuentaBancaria.FindById(Values(4))
    End If
    operacion.EntradaSalida = OPEntrada


    recibo.operacionesBanco.Add operacion

    Totalizar
End Sub

Private Sub gridDepositosOperaciones_UnboundDelete(ByVal RowIndex As Long, ByVal Bookmark As Variant)
    If RowIndex > 0 And recibo.operacionesBanco.count >= RowIndex Then
        recibo.operacionesBanco.remove RowIndex
        Totalizar
    End If
End Sub

Private Sub gridDepositosOperaciones_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex <= recibo.operacionesBanco.count Then
        Set operacion = recibo.operacionesBanco.item(RowIndex)
        Values(1) = funciones.FormatearDecimales(operacion.Monto)
        If IsSomething(operacion.moneda) Then
            Values(2) = operacion.moneda.NombreCorto
        End If
        Values(3) = operacion.FechaOperacion
        If IsSomething(operacion.CuentaBancaria) Then
            Values(4) = operacion.CuentaBancaria.DescripcionFormateada
        End If
        Totalizar
    End If
End Sub

Private Sub gridDepositosOperaciones_UnboundUpdate(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex > 0 And recibo.operacionesBanco.count > 0 Then
        Set operacion = recibo.operacionesBanco.item(RowIndex)

        operacion.Monto = Values(1)
        If IsNumeric(Values(2)) Then
            Set operacion.moneda = DAOMoneda.GetById(Values(2))
        End If
        operacion.FechaOperacion = Values(3)
        If IsNumeric(Values(4)) Then
            Set operacion.CuentaBancaria = DAOCuentaBancaria.FindById(Values(4))
        End If
        operacion.EntradaSalida = OPEntrada
        Totalizar
    End If
End Sub


Private Sub gridMonedas_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex > 0 And Monedas.count > 0 Then
        Set moneda = Monedas.item(RowIndex)
        Values(1) = moneda.id
        Values(2) = moneda.NombreCorto
    End If
End Sub


Private Sub txtRedondeo_Change()
    recibo.redondeo = Val(Me.txtRedondeo.text)
    Totalizar
End Sub


