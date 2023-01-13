VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmPlaneamientoOTNueva 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modificación Orden de Trabajo"
   ClientHeight    =   8865
   ClientLeft      =   2235
   ClientTop       =   2880
   ClientWidth     =   13545
   ClipControls    =   0   'False
   Icon            =   "frmNuevaOtManual.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   13545
   Begin XtremeSuiteControls.GroupBox GroupBoxCotizacionMoneda 
      Height          =   615
      Left            =   120
      TabIndex        =   44
      Top             =   8160
      Width           =   13335
      _Version        =   786432
      _ExtentX        =   23521
      _ExtentY        =   1085
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.ComboBox ComboBoxValorMoneda 
         Height          =   360
         Left            =   2880
         TabIndex        =   45
         Top             =   180
         Width           =   1935
         _Version        =   786432
         _ExtentX        =   3413
         _ExtentY        =   635
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.Label Label1ValorDolar 
         Height          =   255
         Left            =   240
         TabIndex        =   46
         Top             =   240
         Width           =   2895
         _Version        =   786432
         _ExtentX        =   5106
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Cotización Moneda Actual:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         WordWrap        =   -1  'True
      End
   End
   Begin XtremeSuiteControls.GroupBox fraCondiciones 
      Height          =   1710
      Left            =   5145
      TabIndex        =   26
      Top             =   195
      Width           =   8250
      _Version        =   786432
      _ExtentX        =   14552
      _ExtentY        =   3016
      _StockProps     =   79
      Caption         =   "Condiciones Comerciales"
      UseVisualStyle  =   -1  'True
      Begin VB.TextBox txtAnticipo 
         Height          =   285
         Left            =   1200
         TabIndex        =   33
         Top             =   810
         Width           =   540
      End
      Begin VB.TextBox txtFormaPagoAnticipo 
         Height          =   285
         Left            =   5820
         TabIndex        =   32
         Top             =   780
         Width           =   2175
      End
      Begin VB.ComboBox cboMoneda 
         Height          =   315
         ItemData        =   "frmNuevaOtManual.frx":000C
         Left            =   990
         List            =   "frmNuevaOtManual.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   345
         Width           =   750
      End
      Begin VB.TextBox txtDto 
         Height          =   285
         Left            =   1275
         TabIndex        =   30
         Text            =   "0"
         Top             =   1305
         Width           =   480
      End
      Begin VB.TextBox txtCantDiasAnticipo 
         Height          =   285
         Left            =   3375
         TabIndex        =   29
         Top             =   780
         Width           =   480
      End
      Begin VB.TextBox txtFormaPagoSaldo 
         Height          =   285
         Left            =   5820
         TabIndex        =   28
         Top             =   1245
         Width           =   2175
      End
      Begin VB.TextBox txtCantDiasSaldo 
         Height          =   285
         Left            =   3375
         TabIndex        =   27
         Top             =   1245
         Width           =   480
      End
      Begin XtremeSuiteControls.ComboBox cboCliente2 
         Height          =   315
         Left            =   3375
         TabIndex        =   34
         Top             =   300
         Width           =   4650
         _Version        =   786432
         _ExtentX        =   8202
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         UseVisualStyle  =   -1  'True
         Text            =   "ComboBox1"
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "F.P. Anticipo"
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
         Left            =   4635
         TabIndex        =   42
         Top             =   825
         Width           =   1110
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "F.P. Saldo"
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
         Left            =   4845
         TabIndex        =   41
         Top             =   1290
         Width           =   900
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "% Anticipo"
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
         Left            =   225
         TabIndex        =   40
         Top             =   825
         Width           =   900
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "Moneda"
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
         Left            =   240
         TabIndex        =   39
         Top             =   390
         Width           =   690
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "% Descuento"
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
         Left            =   75
         TabIndex        =   38
         Top             =   1335
         Width           =   1125
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "Anticipo a          días"
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
         Left            =   2445
         TabIndex        =   37
         Top             =   810
         Width           =   1845
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "Saldo a          días"
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
         Left            =   2655
         TabIndex        =   36
         Top             =   1290
         Width           =   1635
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFC0C0&
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
         Height          =   255
         Left            =   2715
         TabIndex        =   35
         Top             =   345
         Width           =   675
      End
   End
   Begin XtremeSuiteControls.GroupBox Frame1 
      Height          =   2070
      Left            =   165
      TabIndex        =   16
      Top             =   30
      Width           =   13335
      _Version        =   786432
      _ExtentX        =   23521
      _ExtentY        =   3651
      _StockProps     =   79
      Caption         =   "Datos Generales"
      UseVisualStyle  =   -1  'True
      Begin VB.TextBox txtReferencia 
         Height          =   435
         Left            =   1125
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Top             =   675
         Width           =   3540
      End
      Begin VB.CheckBox chkMismaFecha 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Usar esta fecha para todos los detalles de la orden"
         Height          =   345
         Left            =   2460
         TabIndex        =   18
         Top             =   1590
         Width           =   2370
      End
      Begin XtremeSuiteControls.ComboBox cboCliente 
         Height          =   315
         Left            =   1140
         TabIndex        =   17
         Top             =   255
         Width           =   3540
         _Version        =   786432
         _ExtentX        =   6244
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         UseVisualStyle  =   -1  'True
         Text            =   "ComboBox1"
      End
      Begin MSComCtl2.DTPicker DTVencimiento 
         Height          =   300
         Left            =   1125
         TabIndex        =   20
         Top             =   1590
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   529
         _Version        =   393216
         Format          =   58982401
         CurrentDate     =   38926
      End
      Begin MSComCtl2.DTPicker dtpInicio 
         Height          =   300
         Left            =   1125
         TabIndex        =   21
         Top             =   1200
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   529
         _Version        =   393216
         Format          =   58982401
         CurrentDate     =   38926
      End
      Begin VB.Label Re 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Referencia"
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
         Left            =   120
         TabIndex        =   25
         Top             =   690
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "C. Costos"
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
         Left            =   210
         TabIndex        =   24
         Top             =   330
         Width           =   1005
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "Entrega O/T"
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
         Left            =   0
         TabIndex        =   23
         Top             =   1635
         Width           =   1080
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "Inicio"
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
         Left            =   585
         TabIndex        =   22
         Top             =   1245
         Width           =   480
      End
   End
   Begin XtremeSuiteControls.PushButton Command9 
      Height          =   345
      Left            =   5985
      TabIndex        =   15
      Top             =   7560
      Width           =   2565
      _Version        =   786432
      _ExtentX        =   4524
      _ExtentY        =   609
      _StockProps     =   79
      Caption         =   "Estadísticas"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton Command10 
      Height          =   345
      Left            =   5985
      TabIndex        =   14
      Top             =   7200
      Width           =   2565
      _Version        =   786432
      _ExtentX        =   4524
      _ExtentY        =   609
      _StockProps     =   79
      Caption         =   "Materializacion"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton Command8 
      Height          =   345
      Left            =   5985
      TabIndex        =   13
      Top             =   6840
      Width           =   2565
      _Version        =   786432
      _ExtentX        =   4524
      _ExtentY        =   609
      _StockProps     =   79
      Caption         =   "Resúmen de materiales"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton cmdDefinirPrecios 
      Height          =   300
      Left            =   120
      TabIndex        =   12
      Top             =   7440
      Width           =   2565
      _Version        =   786432
      _ExtentX        =   4524
      _ExtentY        =   529
      _StockProps     =   79
      Caption         =   "Definir Precios de Detalles Sel."
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton Command7 
      Height          =   300
      Left            =   120
      TabIndex        =   11
      Top             =   7140
      Width           =   2565
      _Version        =   786432
      _ExtentX        =   4524
      _ExtentY        =   529
      _StockProps     =   79
      Caption         =   "Renumerar Detalles"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton Qui 
      Height          =   300
      Left            =   120
      TabIndex        =   10
      Top             =   6840
      Width           =   2565
      _Version        =   786432
      _ExtentX        =   4524
      _ExtentY        =   529
      _StockProps     =   79
      Caption         =   "Eliminar Detalles Seleccionados"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton Command5 
      Height          =   405
      Left            =   4350
      TabIndex        =   9
      Top             =   7380
      Width           =   1140
      _Version        =   786432
      _ExtentX        =   2011
      _ExtentY        =   714
      _StockProps     =   79
      Caption         =   "Salir"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton CommandGuardar 
      Height          =   405
      Left            =   4350
      TabIndex        =   8
      Top             =   6945
      Width           =   1140
      _Version        =   786432
      _ExtentX        =   2011
      _ExtentY        =   714
      _StockProps     =   79
      Caption         =   "Guardar"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton Command6 
      Height          =   405
      Left            =   3150
      TabIndex        =   7
      Top             =   7380
      Width           =   1140
      _Version        =   786432
      _ExtentX        =   2011
      _ExtentY        =   714
      _StockProps     =   79
      Caption         =   "Imprimir"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton Command2 
      Height          =   405
      Left            =   3150
      TabIndex        =   6
      Top             =   6945
      Width           =   1140
      _Version        =   786432
      _ExtentX        =   2011
      _ExtentY        =   714
      _StockProps     =   79
      Caption         =   "Recalcular"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton cmdAgregarPieza 
      Height          =   390
      Left            =   120
      TabIndex        =   4
      Top             =   2130
      Width           =   2775
      _Version        =   786432
      _ExtentX        =   4895
      _ExtentY        =   688
      _StockProps     =   79
      Caption         =   "Agregar Pieza..."
      UseVisualStyle  =   -1  'True
   End
   Begin GridEX20.GridEX grid 
      Height          =   4185
      Left            =   120
      TabIndex        =   2
      Top             =   2580
      Width           =   13335
      _ExtentX        =   23521
      _ExtentY        =   7382
      Version         =   "2.0"
      PreviewRowIndent=   100
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      PreviewColumn   =   "pieza"
      PreviewRowLines =   1
      CalendarTodayText=   "Hoy"
      CalendarNoneText=   "Vacío"
      ColumnAutoResize=   -1  'True
      MultiSelect     =   -1  'True
      MethodHoldFields=   -1  'True
      ContScroll      =   -1  'True
      AllowColumnDrag =   0   'False
      GroupByBoxVisible=   0   'False
      BackColorHeader =   16744576
      ImageCount      =   1
      ImagePicture1   =   "frmNuevaOtManual.frx":0010
      RowHeaders      =   -1  'True
      ItemCount       =   3
      DataMode        =   99
      HeaderFontName  =   "Tahoma"
      HeaderFontBold  =   -1  'True
      HeaderFontWeight=   700
      GridLines       =   1
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   12
      Column(1)       =   "frmNuevaOtManual.frx":032A
      Column(2)       =   "frmNuevaOtManual.frx":0446
      Column(3)       =   "frmNuevaOtManual.frx":05B2
      Column(4)       =   "frmNuevaOtManual.frx":06DA
      Column(5)       =   "frmNuevaOtManual.frx":083E
      Column(6)       =   "frmNuevaOtManual.frx":09BA
      Column(7)       =   "frmNuevaOtManual.frx":0B46
      Column(8)       =   "frmNuevaOtManual.frx":0C3E
      Column(9)       =   "frmNuevaOtManual.frx":0DD6
      Column(10)      =   "frmNuevaOtManual.frx":0F4A
      Column(11)      =   "frmNuevaOtManual.frx":104E
      Column(12)      =   "frmNuevaOtManual.frx":1112
      FormatStylesCount=   8
      FormatStyle(1)  =   "frmNuevaOtManual.frx":11D6
      FormatStyle(2)  =   "frmNuevaOtManual.frx":130E
      FormatStyle(3)  =   "frmNuevaOtManual.frx":13BE
      FormatStyle(4)  =   "frmNuevaOtManual.frx":1472
      FormatStyle(5)  =   "frmNuevaOtManual.frx":154A
      FormatStyle(6)  =   "frmNuevaOtManual.frx":1602
      FormatStyle(7)  =   "frmNuevaOtManual.frx":16E2
      FormatStyle(8)  =   "frmNuevaOtManual.frx":172E
      ImageCount      =   1
      ImagePicture(1) =   "frmNuevaOtManual.frx":17BA
      PrinterProperties=   "frmNuevaOtManual.frx":1AD4
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9165
      Top             =   7095
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton PushButton1 
      Height          =   300
      Left            =   120
      TabIndex        =   43
      Top             =   7755
      Width           =   2565
      _Version        =   786432
      _ExtentX        =   4524
      _ExtentY        =   529
      _StockProps     =   79
      Caption         =   "Mayor o Último Precio"
      UseVisualStyle  =   -1  'True
   End
   Begin VB.Label lblMarco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Marco"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3075
      TabIndex        =   5
      Top             =   2190
      Width           =   600
   End
   Begin VB.Label lblModoEdicion 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "[ - MODO EDICION - ]"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   11670
      TabIndex        =   3
      ToolTipText     =   "Presione <ENTER> para terminar de editar el campo"
      Top             =   2205
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   10200
      TabIndex        =   1
      Top             =   7200
      Width           =   615
   End
   Begin VB.Label lbltot 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   10920
      TabIndex        =   0
      Top             =   7200
      Width           =   540
   End
   Begin VB.Menu m1 
      Caption         =   "Menu1"
      Visible         =   0   'False
      Begin VB.Menu Ver 
         Caption         =   "Ver Desarrollo..."
      End
      Begin VB.Menu archivos 
         Caption         =   "Archivos Asociados..."
         Begin VB.Menu mnuArchivoAsociadoPieza 
            Caption         =   "De la Pieza..."
         End
         Begin VB.Menu mnuArchivoAsociadoDetalle 
            Caption         =   "Del Detalle..."
         End
      End
      Begin VB.Menu scanear 
         Caption         =   "Adquirir..."
         Begin VB.Menu mnuAdquirirPieza 
            Caption         =   "Adquirir a Pieza"
         End
         Begin VB.Menu mnuAdquirirDetalle 
            Caption         =   "Adquirir al Detalle"
         End
      End
      Begin VB.Menu verIncidencias 
         Caption         =   "Ver Incidencias..."
         Begin VB.Menu mnuIncidenciasPieza 
            Caption         =   "Incidencias de Pieza"
         End
         Begin VB.Menu mnuIncidenciasDetalle 
            Caption         =   "Incidencias del Detalle"
         End
      End
   End
End
Attribute VB_Name = "frmPlaneamientoOTNueva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements ISuscriber
Dim CantArchivos As New Dictionary
Dim CantArchivosDetalle As New Dictionary
Private idSuscriber As String
Dim baseC As New classConfigurar
Dim idCliente As Long
Dim baseP As New classPlaneamiento
Dim baseS As New classStock
Dim vNroPresu As Long
Dim col As Collection
Dim claseS As New classStock
Private m_ot As OrdenTrabajo
Private tmpDetalle As DetalleOrdenTrabajo
Private formLoaded As Boolean
Public ActualizacionPrecios As Boolean

Private Monedas As New Collection

Public Property Let OrdenTrabajoId(value As Long)
    Set m_ot = DAOOrdenTrabajo.FindById(value)       'me la recargo por las dudas
    Set m_ot.Detalles = DAODetalleOrdenTrabajo.FindAllByOrdenTrabajo(m_ot.id)

    Me.fraCondiciones.Enabled = m_ot.NoEsMarcoNiHija
    Me.cboCliente.Enabled = m_ot.NoEsMarcoNiHija

    Me.chkMismaFecha.Enabled = Not m_ot.EsMarco
    Me.chkMismaFecha.value = CInt(m_ot.EsMarco) * -1
    Me.chkMismaFecha.Visible = Not m_ot.EsMarco

    Me.Label8.Visible = m_ot.EsMarco
    Me.dtpInicio.Visible = m_ot.EsMarco

    If m_ot.EsMarco Then
        Me.Label23.caption = "Fin"
        Me.lblMarco.caption = "La OT es Contrato Abierto"

        If CDbl(m_ot.FechaEntrega) = 0 Then m_ot.FechaEntrega = m_ot.FechaFinMarco
    ElseIf m_ot.EsHija Then
        Me.lblMarco.caption = "La OT forma parte de la OT Contrato Abierto Nº " & m_ot.OTMarcoIdPadre
    ElseIf m_ot.NoEsMarcoNiHija Then
        Me.lblMarco.caption = vbNullString
    End If


End Property

Private Sub CargarOrdenTrabajo()
    If m_ot Is Nothing Then Exit Sub

    Dim i As Long
    If (Me.cboMoneda.ListIndex <> -1) Then Me.cboCliente2.ListIndex = funciones.PosIndexCbo(m_ot.ClienteFacturar.id, cboCliente2)


    Me.cboCliente.ListIndex = funciones.PosIndexCbo(m_ot.cliente.id, cboCliente)
    Me.txtReferencia.text = m_ot.descripcion
    Me.DTVencimiento.value = m_ot.FechaEntrega
    Me.chkMismaFecha.value = CInt(m_ot.MismaFechaEntregaParaDetalles) * -1
    Me.cboMoneda.ListIndex = funciones.PosIndexCbo(m_ot.moneda.id, cboMoneda)
    Me.txtDto.text = m_ot.Descuento
    Me.txtAnticipo.text = m_ot.Anticipo
    Me.txtCantDiasAnticipo.text = m_ot.CantDiasAnticipo
    Me.txtFormaPagoAnticipo.text = m_ot.FormaDePagoAnticipo
    Me.txtFormaPagoSaldo.text = m_ot.FormaDePagoSaldo
    Me.txtCantDiasSaldo.text = m_ot.CantDiasSaldo
    Me.txtFormaPagoAnticipo.text = m_ot.FormaDePagoAnticipo

    Me.dtpInicio.value = m_ot.FechaInicioMarco
    If m_ot.EsMarco Then Me.DTVencimiento.value = m_ot.FechaFinMarco


    Qui.Enabled = Not ActualizacionPrecios
    Command7.Enabled = Not ActualizacionPrecios
    cmdDefinirPrecios.Enabled = Not ActualizacionPrecios
    Command2.Enabled = Not ActualizacionPrecios
    'cmdAgregarPieza.Enabled = Not ActualizacionPrecios
    
    txtReferencia.Enabled = True
    If ActualizacionPrecios Then
        Dim col As JSColumn
        For Each col In Me.grid.Columns
            If col.EditType <> jgexEditNone And col.key <> "precio" Then
                col.EditType = jgexEditNone
            End If
        Next col
    End If

    ConfigurarMismaFecha
    CalcularValorOt
    RecargarDetalles
End Sub
Private Sub RecargarDetalles()
    Me.grid.ItemCount = 0
    Me.grid.ItemCount = m_ot.Detalles.count
End Sub


Private Sub cboCliente_Click()
    If Me.cboCliente.ListIndex <> -1 And Not m_ot Is Nothing Then
        If formLoaded Then
            Set m_ot.cliente = DAOCliente.BuscarPorID(Me.cboCliente.ItemData(Me.cboCliente.ListIndex))
            Set m_ot.Detalles = New Collection
            RecargarDetalles
        End If
    End If

End Sub


Private Sub cboCliente2_Click()
    If Me.cboCliente2.ListIndex <> -1 And Not m_ot Is Nothing Then
        If formLoaded Then
            Set m_ot.ClienteFacturar = DAOCliente.BuscarPorID(Me.cboCliente2.ItemData(Me.cboCliente2.ListIndex))

        End If
    End If

End Sub


Private Sub cboMoneda_Click()
    If Me.cboMoneda.ListIndex <> -1 And Not m_ot Is Nothing Then
        If formLoaded Then
            Set m_ot.moneda = DAOMoneda.GetById(Me.cboMoneda.ItemData(Me.cboMoneda.ListIndex))
        End If
    End If
End Sub

Private Sub chkMismaFecha_Click()
    m_ot.MismaFechaEntregaParaDetalles = Me.chkMismaFecha.value
    ConfigurarMismaFecha
    RecargarDetalles
End Sub

Private Sub ConfigurarMismaFecha()
    If Me.chkMismaFecha.value Then
        grid.Columns("entrega").EditType = jgexEditNone

        For Each tmpDetalle In m_ot.Detalles
            tmpDetalle.FechaEntrega = m_ot.FechaEntrega
        Next tmpDetalle
    Else
        grid.Columns("entrega").EditType = jgexEditCalendarDropDown
    End If

End Sub
Private Sub cmdAgregarPieza_Click()
    If m_ot Is Nothing Then Exit Sub
    If m_ot.cliente Is Nothing Then Exit Sub
    Dim id As Long
    Dim f12 As New frmElegirPieza
    f12.Origen = 2    'desde ot
    f12.OtIdFilter = m_ot.OTMarcoIdPadre
    f12.cliente = m_ot.cliente
    f12.Show 1
End Sub


Private Sub Command10_Click()
    frmMaterializacion.id = m_ot.id
    frmMaterializacion.Ot = True
    frmMaterializacion.Show

End Sub

Private Sub cmdDefinirPrecios_Click()
    If Me.grid.EditMode = jgexEditModeOn Then
        MsgBox "Salga del modo edición de detalles.", vbInformation + vbOKOnly
        Exit Sub
    End If

    If MsgBox("¿Desea asumir los precios de los detalles seleccionados como precios definidos?", vbYesNo + vbQuestion, "Confirmación") = vbYes Then
        Dim va As Boolean
        Dim si As GridEX20.JSSelectedItem
        For Each si In Me.grid.SelectedItems
            If si.RowIndex > 0 And si.RowIndex <= m_ot.Detalles.count Then
                Set tmpDetalle = m_ot.Detalles.item(si.RowIndex)
                va = baseP.definirPrecios(tmpDetalle.Pieza.id, tmpDetalle.Precio, m_ot.moneda.id)
            End If
        Next si

        If Not va Then
            MsgBox "Se produjo algún error, no se guardaron los cambios!", vbCritical, "Error"
        Else
            MsgBox "Definición exitosa de precios!", vbInformation, "Información"
        End If

    End If

End Sub
Private Sub Command2_Click()
    CalcularValorOt
    RecargarDetalles
End Sub
Private Sub CommandGuardar_Click()
    If LenB(Trim$(Me.txtReferencia.text)) = 0 Then
        MsgBox "Falta la referencia", vbInformation + vbOKOnly
        Exit Sub
    End If
    If Me.grid.EditMode = jgexEditModeOn Then
        MsgBox "Salga del modo edición de detalles.", vbInformation + vbOKOnly
        Exit Sub
    End If

    If Me.cboMoneda.ListIndex = 3 Then
        MsgBox "No se puede guardar una OT con Moneda U$A Administrativo. Modifiquelo por favor.", vbInformation + vbOKOnly
        Exit Sub
    End If
    
    If vbYes = MsgBox("¿Confirma la edicion de la orden?", vbYesNo + vbQuestion) Then
        If ActualizacionPrecios Then
            Dim detaOT As DetalleOrdenTrabajo
            Dim result As Boolean
            conectar.BeginTransaction
            For Each detaOT In m_ot.Detalles
                result = DAODetalleOrdenTrabajo.Save(detaOT)
                If Not result Then Exit For
            Next detaOT
            If result Then
                conectar.execute "UPDATE pedidos SET ultima_fecha_actualizacion_precios = NOW() WHERE id = " & m_ot.id
                conectar.execute "UPDATE pedidos SET descripcion = '" & m_ot.descripcion & "' WHERE id = " & m_ot.id
                conectar.CommitTransaction
                 Dim EVENTO As New clsEventoObserver
                Set EVENTO.Elemento = m_ot
                Set EVENTO.Originador = Me
                EVENTO.EVENTO = modificar_
                Channel.Notificar EVENTO, ordenesTrabajo
                MsgBox "Los precios de los detalles se editaron correctamente", vbInformation + vbOKOnly
            Else
                conectar.RollBackTransaction
                MsgBox "Se produjo un error al editar los precios orden", vbCritical, "Error"
            End If
        Else
            If DAOOrdenTrabajo.Save(m_ot) Then
                Dim EVENTO1 As New clsEventoObserver
                Set EVENTO1.Elemento = m_ot
                Set EVENTO1.Originador = Me
                EVENTO1.EVENTO = modificar_
                Channel.Notificar EVENTO1, ordenesTrabajo
                MsgBox "La orden se edito correctamente", vbInformation + vbOKOnly
            Else
                MsgBox "Se produjo un error al editar la orden", vbCritical, "Error"
            End If
        End If



    End If
End Sub

Private Sub Command5_Click()
    If MsgBox("¿Está seguro de salir?", vbYesNo, "Confirmación") = vbYes Then
        Unload Me
    End If
End Sub
Private Sub Command6_Click()
    imprimirOT
End Sub

Private Sub Command7_Click()
    If MsgBox("¿Desea hacer correlativa la numeracion de los ítems?", vbYesNo + vbQuestion, "Confirmación") = vbYes Then
        RenumerarDetalles
        RecargarDetalles
    End If
End Sub

Private Sub RenumerarDetalles()
    Dim x As Long

    For x = 1 To m_ot.Detalles.count
        Set tmpDetalle = m_ot.Detalles(x)
        tmpDetalle.item = Format(x, "000")
    Next x
End Sub
Private Sub Command8_Click()
    Dim A As Boolean
    'a = baseP.informePiezaMateriales(m_ot.Id, 1, True)
    DAOOrdenTrabajo.informePiezaMateriales m_ot.id, 1, True
End Sub
Private Sub Command9_Click()

    Dim dto As DTOPiezaCantidad
    Dim deta As DetalleOrdenTrabajo
    Dim listadtopiezacantidad As New Collection
    For Each deta In m_ot.Detalles
        Set dto = New DTOPiezaCantidad
        Set dto.Pieza = deta.Pieza
        dto.Cantidad = deta.CantidadPedida
        listadtopiezacantidad.Add dto
    Next deta

    Dim frm1 As New frmEstadistiacasEnCurso
    frm1.caption = "Estadisticas de presupuesto activo"
    Set frm1.listadtopiezacantidad = listadtopiezacantidad
    frm1.conjGrabado = False
    frm1.Show
End Sub


Private Sub dtpInicio_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

    If m_ot.EsMarco Then m_ot.FechaInicioMarco = Me.dtpInicio.value

End Sub

Private Sub DTVencimiento_Change()
    m_ot.FechaEntrega = Me.DTVencimiento.value
    If m_ot.EsMarco Then m_ot.FechaFinMarco = m_ot.FechaEntrega

    chkMismaFecha_Click
End Sub

Private Sub verModoEdicion()
    If Me.grid.EditMode = jgexEditModeOn Then
        Me.lblModoEdicion.Visible = True
    Else
        Me.lblModoEdicion.Visible = False
    End If
    
    GroupBoxCotizacionMoneda.Visible = True
End Sub

Private Sub Form_Load()
    FormHelper.Customize Me
    Me.lblModoEdicion.ToolTipText = "Presione <ENTER> para terminar  ó <ESC> para canc"
    formLoaded = False
    Me.grid.ItemCount = 0
    idSuscriber = funciones.CreateGUID()
    Channel.AgregarSuscriptor Me, NuevaOT_

    DAOMoneda.LlenarCombo Me.cboMoneda
    DAOCliente.llenarComboXtremeSuite Me.cboCliente, False, True, False
    DAOCliente.llenarComboXtremeSuite Me.cboCliente2, False, True, False

    GridEXHelper.CustomizeGrid Me.grid, False, True
    CargarOrdenTrabajo
    formLoaded = True
    Set CantArchivos = DAOArchivo.GetCantidadArchivosPorReferencia(OA_Piezas)
    Set CantArchivosDetalle = DAOArchivo.GetCantidadArchivosPorReferencia(OA_OrdenesTrabajoDetalle)
    
    DAOMoneda.llenarComboXtremeSuite Me.ComboBoxValorMoneda, True
    Me.ComboBoxValorMoneda.ListIndex = 3
  
    'Me.caption = caption & " (" & Name & ")"

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Channel.RemoverSuscripcionTotal Me
End Sub



Private Sub grid_BeforeUpdate(ByVal Cancel As GridEX20.JSRetBoolean)
    verModoEdicion
End Sub

Private Sub grid_Click()
    verModoEdicion
End Sub

Private Sub grid_ColumnHeaderClick(ByVal Column As GridEX20.JSColumn)
    GridEXHelper.ColumnHeaderClick Me.grid, Column
End Sub

Private Sub grid_FetchIcon(ByVal RowIndex As Long, ByVal ColIndex As Integer, ByVal RowBookmark As Variant, ByVal IconIndex As GridEX20.JSRetInteger)
    '        On Error Resume Next
    '        Set tmpDetalle = m_ot.Detalles.item(RowIndex)
    '
    '        If CantArchivos.item(tmpDetalle.Pieza.Id) > 0 And ColIndex = 11 Then
    '        IconIndex = 1
    '        End If
End Sub

Private Sub grid_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        grid.EditMode = jgexEditModeOff
        verModoEdicion
    End If

End Sub

Private Sub grid_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next

    Dim idx As Long
    idx = Me.grid.RowIndex(Me.grid.row)

    If Button = 2 And idx > 0 Then
        If m_ot.Detalles(idx).Pieza.EsConjunto Then
            Me.ver.caption = "Ver Conjunto..."
            Me.ver.Tag = 0
        Else
            Me.ver.caption = "Ver Desarrollo..."
            Me.ver.Tag = -1
        End If

        Me.PopupMenu Me.m1
    End If

End Sub

Private Sub grid_SelectionChange()
    Set tmpDetalle = m_ot.Detalles.item(grid.RowIndex(grid.row))
End Sub

Private Sub grid_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex > 0 And RowIndex <= m_ot.Detalles.count Then
        Set tmpDetalle = m_ot.Detalles.item(RowIndex)
        Values(1) = tmpDetalle.item
        Values(2) = tmpDetalle.CantidadPedida
        Values(3) = tmpDetalle.Pieza.nombre
        Values(4) = tmpDetalle.Precio
        Values(5) = tmpDetalle.Precio * tmpDetalle.CantidadPedida
        Values(6) = tmpDetalle.FechaEntrega
        Values(7) = tmpDetalle.Nota
        Values(8) = tmpDetalle.Pieza.CantidadStock
        Values(9) = tmpDetalle.ReservaStock
        Values(10) = tmpDetalle.Pieza.UnidadMedida    '   IIf(tmpDetalle.pieza.EsConjunto, "Conjunto", "Unidad")
        Values(11) = CantArchivos.item(tmpDetalle.Pieza.id)
        Values(12) = CantArchivosDetalle.item(tmpDetalle.id)
    End If
End Sub

Private Sub grid_UnboundUpdate(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    On Error GoTo E:

    Set tmpDetalle = Nothing

    If RowIndex > 0 And RowIndex <= m_ot.Detalles.count Then
        Set tmpDetalle = m_ot.Detalles.item(RowIndex)
        tmpDetalle.item = Values(1)
        tmpDetalle.CantidadPedida = Values(2)
        tmpDetalle.Precio = Values(4)
        tmpDetalle.FechaEntrega = Values(6)
        tmpDetalle.Nota = Values(7)
        tmpDetalle.ReservaStock = Values(9)

        '        CalcularValorOt
    End If

    Exit Sub
E:
    MsgBox Err.Description, vbCritical + vbOKOnly, "Error"
End Sub

Private Property Get ISuscriber_id() As String
    ISuscriber_id = idSuscriber
End Property

Private Function ISuscriber_Notificarse(EVENTO As clsEventoObserver) As Variant
    Dim col As New Collection
    Dim x As Long
    If EVENTO.EVENTO = agregarColeccion_ Then
        Set col = EVENTO.Elemento
        Dim moneda As New clsMoneda
        Dim adm As New classAdministracion

        Dim dto As DTOPiezaDetallePedido


        'For X = 1 To col.count
        For Each dto In col
            Set tmpDetalle = New DetalleOrdenTrabajo
            tmpDetalle.CantidadPedida = 1
            tmpDetalle.FechaEntrega = m_ot.FechaEntrega
            Set tmpDetalle.OrdenTrabajo = m_ot

            Set tmpDetalle.Pieza = dto.Pieza

            If dto.idOt = 0 Then   'ver cuando cree el prox marco
                If tmpDetalle.Pieza.Precio <> 0 Then
                    tmpDetalle.Precio = tmpDetalle.Pieza.Precio
                    If tmpDetalle.Pieza.MonedaPrecio.id <> m_ot.moneda.id Then
                        tmpDetalle.Precio = adm.realizaCambio(tmpDetalle.Pieza.Precio, tmpDetalle.Pieza.MonedaPrecio.id, m_ot.moneda.id)
                    End If
                End If

            Else
                tmpDetalle.Precio = dto.Precio

                If m_ot.EsMarco Then
                    tmpDetalle.idDetalleOtPadre = -1
                Else
                    tmpDetalle.idDetalleOtPadre = dto.idDetalleOt
                End If

                'tmpDetalle.Item = dto.Item
            End If

            Dim d As DetalleOrdenTrabajo
            Dim c As Long: c = 0
            For Each d In m_ot.Detalles
                c = Val(d.item)
            Next d
            tmpDetalle.item = Format(c + 1, "000")

            m_ot.Detalles.Add tmpDetalle

        Next dto

        CalcularValorOt
        RecargarDetalles
        grid.MoveLast
    End If
End Function





Private Sub mnuAdquirirDetalle_Click()
    Dim archi As classArchivos
    Set archi = New classArchivos
    archi.escanearDocumento OrigenArchivos.OA_OrdenesTrabajoDetalle, tmpDetalle.id
End Sub

Private Sub mnuAdquirirPieza_Click()
    Dim archi As classArchivos
    Set archi = New classArchivos
    archi.escanearDocumento OrigenArchivos.OA_Piezas, tmpDetalle.Pieza.id
End Sub

Private Sub mnuArchivoAsociadoDetalle_Click()
    grid_SelectionChange
    Dim F As New frmArchivos2
    F.Origen = OrigenArchivos.OA_OrdenesTrabajoDetalle
    F.ObjetoId = tmpDetalle.id
    F.caption = "OT Nº " & m_ot.IdFormateado & " - Item " & tmpDetalle.item
    F.Show
End Sub

Private Sub mnuArchivoAsociadoPieza_Click()
    grid_SelectionChange
    Dim F As New frmArchivos2
    F.Origen = OrigenArchivos.OA_Piezas
    F.ObjetoId = tmpDetalle.Pieza.id
    F.caption = "Pieza " & tmpDetalle.Pieza.nombre
    F.Show
End Sub

Private Sub mnuIncidenciasDetalle_Click()
    frmVerIncidencias.referencia = tmpDetalle.id
    frmVerIncidencias.Origen = OI_OrdenesTrabajoDetalles
    frmVerIncidencias.Show
End Sub

Private Sub mnuIncidenciasPieza_Click()
    frmVerIncidencias.referencia = tmpDetalle.Pieza.id
    frmVerIncidencias.Origen = OI_Piezas
    frmVerIncidencias.Show
End Sub

Private Sub PushButton1_Click()

    If Me.grid.EditMode = jgexEditModeOn Then
        MsgBox "Salga del modo edición de detalles.", vbInformation + vbOKOnly
        Exit Sub
    End If

    If MsgBox("¿Desea definir los precios de los detalles seleccionados al maximo valor?", vbYesNo + vbQuestion, "Confirmación") = vbYes Then
        Dim va As Boolean
        Dim si As GridEX20.JSSelectedItem
        For Each si In Me.grid.SelectedItems
            If si.RowIndex > 0 And si.RowIndex <= m_ot.Detalles.count Then
                Set tmpDetalle = m_ot.Detalles.item(si.RowIndex)
                tmpDetalle.Precio = DAODetalleOrdenTrabajo.FindBestPriceByPiezaId(tmpDetalle.Pieza.id)
            End If
        Next si
        RecargarDetalles


    End If
End Sub

Private Sub Qui_Click()
    Dim si As GridEX20.JSSelectedItem
    Dim i As Long

    For i = m_ot.Detalles.count To 1 Step -1
        For Each si In Me.grid.SelectedItems
            If si.RowIndex = i Then
                m_ot.Detalles.remove i
                Exit For
            End If
        Next si
    Next i

    If Me.grid.SelectedItems.count > 0 Then
        CalcularValorOt
        RecargarDetalles
    End If
End Sub

Private Sub CalcularValorOt()
    Dim reserva As Double

    Dim tmpPieza As Pieza

    For Each tmpDetalle In m_ot.Detalles

        Set tmpPieza = DAOPieza.FindById(tmpDetalle.Pieza.id, FL_0)

        If tmpPieza.CantidadStock > tmpDetalle.CantidadPedida Then
            reserva = tmpDetalle.CantidadPedida
        Else
            reserva = tmpPieza.CantidadStock
        End If

        'If reserva > 0 Then pinto de rojo la fila si no de negro

        tmpDetalle.ReservaStock = reserva
        Set tmpDetalle.Pieza = tmpPieza
    Next tmpDetalle

    Me.lbltot.caption = funciones.FormatearDecimales(m_ot.Total, 2)

End Sub

Private Sub imprimirOT()
    Dim headercenter As String
    Dim headerLeft As String


    headercenter = "OT NUMERO " & m_ot.id & Chr(10) _
                   & "Cliente: (" & m_ot.cliente.id & ") " & m_ot.cliente.razon & Chr(10) _
                   & "Referencia: " & m_ot.descripcion & Chr(10) _
                   & "Entrega: " & m_ot.FechaEntrega & Chr(10)

    headerLeft = "Total: " & m_ot.moneda.NombreCorto & " " & Format$(m_ot.Total, "0.00") & vbNewLine _
                 & "% Descuento: " & m_ot.Descuento & "%" & Chr(10) _
                 & "% Anticipo: " & m_ot.Anticipo & "%" & Chr(10) _
                 & "Anticipo a " & m_ot.CantDiasAnticipo & " días | FP: " & m_ot.FormaDePagoAnticipo & Chr(10) _
                 & "Saldo a " & m_ot.CantDiasSaldo & " días | FP: " & m_ot.FormaDePagoSaldo & Chr(10) _

With Me.grid.PrinterProperties
        .HeaderDistance = 600
        '.FooterDistance = 1550
        .TopMargin = 2000
        .BottomMargin = 2000

        .FitColumns = True
        .DocumentName = "Orden de Trabajo"

        .RepeatHeaders = True
        .Orientation = jgexPPLandscape
        .HeaderString(jgexHFCenter) = headercenter
        .HeaderString(jgexHFLeft) = headerLeft

    End With

    Load frmPrintPreview
    frmPrintPreview.Move Me.Left, Me.Top, Me.Width, Me.Height
    Me.grid.PrintPreview frmPrintPreview.GEXPreview1
    frmPrintPreview.Show 1

End Sub

Private Sub txtAnticipo_GotFocus()
    foco Me.txtAnticipo
End Sub

Private Sub txtAnticipo_Validate(Cancel As Boolean)
    funciones.ValidarTextBox Me.txtAnticipo, Cancel
    If Not Cancel And IsNumeric(Me.txtAnticipo.text) Then m_ot.Anticipo = CDbl(Me.txtAnticipo.text)
End Sub

Private Sub txtCantDiasAnticipo_GotFocus()
    foco Me.txtCantDiasAnticipo
End Sub

Private Sub txtCantDiasAnticipo_Validate(Cancel As Boolean)
    funciones.ValidarTextBox Me.txtCantDiasAnticipo, Cancel
    If Not Cancel And IsNumeric(Me.txtCantDiasAnticipo.text) Then m_ot.CantDiasAnticipo = CInt(Me.txtCantDiasAnticipo.text)
End Sub

Private Sub txtCantDiasSaldo_GotFocus()
    foco Me.txtCantDiasSaldo
End Sub

Private Sub txtCantDiasSaldo_LostFocus()
    If Not m_ot Is Nothing Then
        m_ot.CantDiasSaldo = Me.txtCantDiasSaldo.text
    End If

End Sub

Private Sub txtCantDiasSaldo_Validate(Cancel As Boolean)
    funciones.ValidarTextBox Me.txtCantDiasSaldo, Cancel
    If Not Cancel And IsNumeric(Me.txtCantDiasSaldo.text) Then m_ot.CantDiasSaldo = CInt(Me.txtCantDiasSaldo.text)
End Sub

Private Sub txtDto_GotFocus()
    foco Me.txtDto
End Sub

Private Sub txtDto_Validate(Cancel As Boolean)
    funciones.ValidarTextBox Me.txtDto, Cancel
    If Not Cancel And IsNumeric(Me.txtDto.text) Then m_ot.Descuento = CDbl(Me.txtDto.text)
End Sub

Private Sub txtFormaPagoAnticipo_LostFocus()
    If Not m_ot Is Nothing Then
        m_ot.FormaDePagoAnticipo = Me.txtFormaPagoAnticipo.text
    End If
End Sub
Private Sub txtFormaPagoSaldo_LostFocus()
    If Not m_ot Is Nothing Then
        m_ot.FormaDePagoSaldo = Me.txtFormaPagoSaldo.text
    End If
End Sub

Private Sub txtReferencia_Change()
    m_ot.descripcion = UCase(Me.txtReferencia)
End Sub

Private Sub ver_Click()
    grid_SelectionChange
    Dim idx As Long
    idx = Me.grid.RowIndex(Me.grid.row)
    If idx > 0 Then
        Dim F As New frmDesarrollo
        Load F
        F.CargarPieza tmpDetalle.Pieza.id   'm_ot.Detalles(idx).Pieza.Id
        F.Show

    End If
End Sub

'Private Sub verIncidencias_Click()
'    Dim idx As Long
'    idx = Me.grid.RowIndex(Me.grid.Row)
'    If idx > 0 Then
'        frmVerIncidencias.referencia = m_ot.Detalles(idx).pieza.id
'        frmVerIncidencias.origen = OrigenIncidencias.OI_Piezas
'        frmVerIncidencias.Show
'    End If
'End Sub
