VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~3.OCX"
Begin VB.Form frmVentasPresupuestoEditar 
   BackColor       =   &H00FF80FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Nuevo presupuesto..."
   ClientHeight    =   10200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14430
   Icon            =   "frmNuevoPresupuesto.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10200
   ScaleWidth      =   14430
   Begin XtremeSuiteControls.GroupBox GroupBox5 
      Height          =   1305
      Left            =   45
      TabIndex        =   71
      Top             =   60
      Width           =   14265
      _Version        =   786432
      _ExtentX        =   25162
      _ExtentY        =   2302
      _StockProps     =   79
      Caption         =   "Datos"
      UseVisualStyle  =   -1  'True
      Begin VB.ComboBox cboCliente 
         Height          =   315
         Left            =   1230
         Style           =   2  'Dropdown List
         TabIndex        =   73
         Top             =   315
         Width           =   12600
      End
      Begin VB.TextBox txtReferencia 
         Height          =   285
         Left            =   1215
         TabIndex        =   72
         Top             =   675
         Width           =   12615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
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
         Left            =   135
         TabIndex        =   75
         Top             =   315
         Width           =   1050
      End
      Begin VB.Label Re 
         Alignment       =   1  'Right Justify
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
         Left            =   135
         TabIndex        =   74
         Top             =   675
         Width           =   1035
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox4 
      Height          =   1875
      Left            =   30
      TabIndex        =   35
      Top             =   1425
      Width           =   14295
      _Version        =   786432
      _ExtentX        =   25215
      _ExtentY        =   3307
      _StockProps     =   79
      Caption         =   "Configuración"
      UseVisualStyle  =   -1  'True
      Begin VB.CommandButton Command9 
         Cancel          =   -1  'True
         Caption         =   "Command9"
         Height          =   375
         Left            =   13395
         TabIndex        =   70
         Top             =   1230
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.TextBox txtMarkupMDO 
         Height          =   285
         Left            =   5205
         TabIndex        =   46
         Top             =   1380
         Width           =   1095
      End
      Begin VB.TextBox men10 
         Height          =   285
         Left            =   5805
         TabIndex        =   45
         Top             =   1020
         Width           =   495
      End
      Begin VB.TextBox men15 
         Height          =   285
         Left            =   5805
         TabIndex        =   44
         Top             =   660
         Width           =   495
      End
      Begin VB.TextBox mas15 
         Height          =   285
         Left            =   5805
         TabIndex        =   43
         Top             =   300
         Width           =   495
      End
      Begin VB.TextBox lblGastos 
         Height          =   285
         Left            =   1380
         TabIndex        =   42
         Top             =   420
         Width           =   975
      End
      Begin VB.TextBox txtDias 
         Height          =   285
         Left            =   9180
         TabIndex        =   41
         Text            =   "1"
         Top             =   300
         Width           =   1215
      End
      Begin VB.TextBox manteOferta 
         Height          =   285
         Left            =   9180
         TabIndex        =   40
         Text            =   "0"
         Top             =   660
         Width           =   1215
      End
      Begin VB.ComboBox cboMoneda 
         Height          =   315
         ItemData        =   "frmNuevoPresupuesto.frx":000C
         Left            =   9180
         List            =   "frmNuevoPresupuesto.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   1380
         Width           =   1215
      End
      Begin VB.TextBox txtMOM 
         Height          =   285
         Left            =   1380
         TabIndex        =   38
         Top             =   780
         Width           =   975
      End
      Begin XtremeSuiteControls.PushButton E 
         Height          =   255
         Left            =   11100
         TabIndex        =   36
         Top             =   1020
         Width           =   1695
         _Version        =   786432
         _ExtentX        =   2990
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Estadisticas"
         BackColor       =   16761024
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton Command2 
         Height          =   375
         Left            =   13005
         TabIndex        =   37
         Top             =   270
         Width           =   1095
         _Version        =   786432
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Recalcular"
         BackColor       =   16761024
         UseVisualStyle  =   -1  'True
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   12885
         Top             =   1155
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
      End
      Begin MSComCtl2.DTPicker DTVencimiento 
         Height          =   255
         Left            =   9180
         TabIndex        =   47
         Top             =   1020
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         Format          =   60162049
         CurrentDate     =   38926
      End
      Begin XtremeSuiteControls.PushButton Command6 
         Height          =   375
         Left            =   13020
         TabIndex        =   48
         Top             =   780
         Width           =   1095
         _Version        =   786432
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Sistema"
         BackColor       =   16761024
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton Command8 
         Height          =   255
         Left            =   11115
         TabIndex        =   49
         Top             =   675
         Width           =   1695
         _Version        =   786432
         _ExtentX        =   2990
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Materiales"
         BackColor       =   16761024
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton Command7 
         Height          =   255
         Left            =   11100
         TabIndex        =   50
         Top             =   300
         Width           =   1695
         _Version        =   786432
         _ExtentX        =   2990
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Imprimir"
         BackColor       =   16761024
         UseVisualStyle  =   -1  'True
      End
      Begin VB.Line Line2 
         X1              =   6780
         X2              =   6780
         Y1              =   300
         Y2              =   1605
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Gastos"
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
         Left            =   300
         TabIndex        =   69
         Top             =   420
         Width           =   735
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "MarkUp Mano de obra"
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
         Left            =   3045
         TabIndex        =   68
         Top             =   1380
         Width           =   1935
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFC0C0&
         Caption         =   "MarkUp Materiales"
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
         Left            =   3405
         TabIndex        =   67
         Top             =   300
         Width           =   1935
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFC0C0&
         Caption         =   "<10Kg"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5205
         TabIndex        =   66
         Top             =   1020
         Width           =   615
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFC0C0&
         Caption         =   "<15Kg"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5205
         TabIndex        =   65
         Top             =   660
         Width           =   615
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFC0C0&
         Caption         =   ">15Kg"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5340
         TabIndex        =   64
         Top             =   300
         Width           =   495
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFC0C0&
         Caption         =   "%"
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
         Left            =   6420
         TabIndex        =   63
         Top             =   285
         Width           =   375
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFC0C0&
         Caption         =   "%"
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
         Left            =   6435
         TabIndex        =   62
         Top             =   675
         Width           =   375
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFC0C0&
         Caption         =   "%"
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
         Left            =   6435
         TabIndex        =   61
         Top             =   1035
         Width           =   255
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Días"
         Height          =   255
         Left            =   10500
         TabIndex        =   60
         Top             =   300
         Width           =   735
      End
      Begin VB.Label Label20 
         BackColor       =   &H00FFC0C0&
         Caption         =   "%"
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
         Left            =   2460
         TabIndex        =   59
         Top             =   420
         Width           =   255
      End
      Begin VB.Label Label22 
         BackColor       =   &H00FFC0C0&
         Caption         =   "%"
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
         Left            =   6435
         TabIndex        =   58
         Top             =   1395
         Width           =   375
      End
      Begin VB.Line Line1 
         X1              =   2775
         X2              =   2775
         Y1              =   300
         Y2              =   1500
      End
      Begin VB.Label Label17 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Mantenimiento de oferta"
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
         Left            =   6900
         TabIndex        =   57
         Top             =   660
         Width           =   2295
      End
      Begin VB.Label Label18 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Días"
         Height          =   255
         Left            =   10500
         TabIndex        =   56
         Top             =   660
         Width           =   375
      End
      Begin VB.Label Label23 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Vencimiento"
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
         Left            =   6900
         TabIndex        =   55
         Top             =   1020
         Width           =   1095
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Tiempo de entrega"
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
         Left            =   6900
         TabIndex        =   54
         Top             =   300
         Width           =   2055
      End
      Begin VB.Label Label24 
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
         Height          =   255
         Left            =   6900
         TabIndex        =   53
         Top             =   1380
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "M.O.Muerta"
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
         Left            =   300
         TabIndex        =   52
         Top             =   780
         Width           =   1095
      End
      Begin VB.Label Label38 
         BackColor       =   &H00FFC0C0&
         Caption         =   "%"
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
         Left            =   2460
         TabIndex        =   51
         Top             =   780
         Width           =   135
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox3 
      Height          =   6780
      Left            =   60
      TabIndex        =   0
      Top             =   3345
      Width           =   14295
      _Version        =   786432
      _ExtentX        =   25215
      _ExtentY        =   11959
      _StockProps     =   79
      Caption         =   "Elementos"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.GroupBox GroupBox2 
         Height          =   1605
         Left            =   105
         TabIndex        =   1
         Top             =   5055
         Width           =   4710
         _Version        =   786432
         _ExtentX        =   8308
         _ExtentY        =   2831
         _StockProps     =   79
         Caption         =   "Forma de Amortizar"
         BackColor       =   16761024
         UseVisualStyle  =   -1  'True
         Begin VB.OptionButton opFijo 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Fijo"
            Height          =   495
            Left            =   3180
            Picture         =   "frmNuevoPresupuesto.frx":0010
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   330
            Width           =   1335
         End
         Begin VB.OptionButton opAutomatica 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Automática"
            Height          =   495
            Left            =   1740
            Picture         =   "frmNuevoPresupuesto.frx":014C
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   330
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton opFabricados 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Por Fabricados"
            Height          =   495
            Left            =   3165
            Picture         =   "frmNuevoPresupuesto.frx":0288
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   930
            Width           =   1335
         End
         Begin VB.OptionButton opCantidad 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Por Cantidad"
            Height          =   495
            Left            =   1740
            Picture         =   "frmNuevoPresupuesto.frx":03C4
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   915
            Width           =   1335
         End
         Begin XtremeSuiteControls.PushButton Command11 
            Height          =   405
            Left            =   150
            TabIndex        =   6
            Top             =   945
            Width           =   1290
            _Version        =   786432
            _ExtentX        =   2275
            _ExtentY        =   714
            _StockProps     =   79
            Caption         =   "Por Fabricados"
            BackColor       =   16761024
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton Command10 
            Height          =   420
            Left            =   150
            TabIndex        =   7
            Top             =   405
            Width           =   1290
            _Version        =   786432
            _ExtentX        =   2275
            _ExtentY        =   741
            _StockProps     =   79
            Caption         =   "Por Cantidad"
            BackColor       =   16761024
            UseVisualStyle  =   -1  'True
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   2025
         Left            =   4875
         TabIndex        =   8
         Top             =   4680
         Width           =   4470
         _Version        =   786432
         _ExtentX        =   7885
         _ExtentY        =   3572
         _StockProps     =   79
         Caption         =   "Condiciones Comerciales"
         BackColor       =   16761024
         UseVisualStyle  =   -1  'True
         Begin VB.TextBox txtDiasPagoSaldo 
            Height          =   285
            Left            =   1410
            TabIndex        =   14
            Text            =   "0"
            Top             =   1305
            Width           =   615
         End
         Begin VB.TextBox txtDiasPagoAnticipo 
            Height          =   285
            Left            =   1410
            TabIndex        =   13
            Text            =   "0"
            Top             =   945
            Width           =   615
         End
         Begin VB.TextBox txtDescuento 
            Height          =   285
            Left            =   3210
            TabIndex        =   12
            Top             =   585
            Width           =   855
         End
         Begin VB.TextBox txtAnticipo 
            Height          =   285
            Left            =   1410
            TabIndex        =   11
            Text            =   "0"
            Top             =   585
            Width           =   615
         End
         Begin VB.TextBox txtFormaPagoAnticipo 
            Height          =   285
            Left            =   3210
            TabIndex        =   10
            Top             =   945
            Width           =   1140
         End
         Begin VB.TextBox txtFormaPagoSaldo 
            Height          =   285
            Left            =   3195
            TabIndex        =   9
            Top             =   1305
            Width           =   1170
         End
         Begin VB.Label lblAFabricar 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Left            =   60
            TabIndex        =   98
            Top             =   1725
            Width           =   5460
         End
         Begin VB.Label Label37 
            BackColor       =   &H00FFC0C0&
            Caption         =   "F.P."
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
            Left            =   2850
            TabIndex        =   24
            Top             =   1305
            Width           =   375
         End
         Begin VB.Label Label36 
            BackColor       =   &H00FFC0C0&
            Caption         =   "F.P."
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
            Left            =   2850
            TabIndex        =   23
            Top             =   945
            Width           =   375
         End
         Begin VB.Label Label35 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Días"
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
            Left            =   2130
            TabIndex        =   22
            Top             =   1305
            Width           =   495
         End
         Begin VB.Label Label34 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Días"
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
            Left            =   2130
            TabIndex        =   21
            Top             =   945
            Width           =   495
         End
         Begin VB.Label Label9 
            BackColor       =   &H00FFC0C0&
            Caption         =   "%"
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
            Left            =   4170
            TabIndex        =   20
            Top             =   585
            Width           =   255
         End
         Begin VB.Label Label14 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Desc"
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
            TabIndex        =   19
            Top             =   405
            Width           =   495
         End
         Begin VB.Label Label33 
            BackColor       =   &H00FFC0C0&
            Caption         =   "%"
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
            Left            =   2130
            TabIndex        =   18
            Top             =   585
            Width           =   735
         End
         Begin VB.Label Label31 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Pago Anticipo"
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
            Left            =   90
            TabIndex        =   17
            Top             =   945
            Width           =   1215
         End
         Begin VB.Label Label30 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Anticipo"
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
            Left            =   570
            TabIndex        =   16
            Top             =   585
            Width           =   735
         End
         Begin VB.Label Label32 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Pago Saldo"
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
            Left            =   330
            TabIndex        =   15
            Top             =   1305
            Width           =   1095
         End
      End
      Begin XtremeSuiteControls.PushButton Qui 
         Height          =   345
         Left            =   135
         TabIndex        =   25
         Top             =   4650
         Width           =   1425
         _Version        =   786432
         _ExtentX        =   2514
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Quitar"
         BackColor       =   16761024
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton Command3 
         Height          =   360
         Left            =   12870
         TabIndex        =   26
         Top             =   210
         Width           =   1245
         _Version        =   786432
         _ExtentX        =   2196
         _ExtentY        =   635
         _StockProps     =   79
         Caption         =   "Guardar"
         BackColor       =   16761024
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton Command1 
         Height          =   315
         Left            =   180
         TabIndex        =   27
         Top             =   195
         Width           =   2625
         _Version        =   786432
         _ExtentX        =   4630
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Elegir Piezas"
         BackColor       =   16761024
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ProgressBar progreso 
         Height          =   270
         Left            =   7350
         TabIndex        =   28
         Top             =   195
         Visible         =   0   'False
         Width           =   2940
         _Version        =   786432
         _ExtentX        =   5186
         _ExtentY        =   476
         _StockProps     =   93
         BackColor       =   16761024
         Appearance      =   6
      End
      Begin GridEX20.GridEX grilla 
         Height          =   3975
         Left            =   165
         TabIndex        =   29
         Top             =   600
         Width           =   13995
         _ExtentX        =   24686
         _ExtentY        =   7011
         Version         =   "2.0"
         AutomaticSort   =   -1  'True
         RecordNavigatorString=   "Presupuesto:|de"
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         GroupFooterStyle=   2
         PreviewColumn   =   "pz"
         PreviewRowLines =   2
         ColumnAutoResize=   -1  'True
         MultiSelect     =   -1  'True
         MethodHoldFields=   -1  'True
         ContScroll      =   -1  'True
         GroupByBoxVisible=   0   'False
         BackColorHeader =   16744576
         ImageCount      =   1
         ImagePicture1   =   "frmNuevoPresupuesto.frx":0500
         RowHeaders      =   -1  'True
         DataMode        =   99
         HeaderFontName  =   "Tahoma"
         FontName        =   "Tahoma"
         GridLines       =   2
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   14
         Column(1)       =   "frmNuevoPresupuesto.frx":081A
         Column(2)       =   "frmNuevoPresupuesto.frx":0A5A
         Column(3)       =   "frmNuevoPresupuesto.frx":0B6E
         Column(4)       =   "frmNuevoPresupuesto.frx":0C7A
         Column(5)       =   "frmNuevoPresupuesto.frx":0D6A
         Column(6)       =   "frmNuevoPresupuesto.frx":0FF2
         Column(7)       =   "frmNuevoPresupuesto.frx":111A
         Column(8)       =   "frmNuevoPresupuesto.frx":129A
         Column(9)       =   "frmNuevoPresupuesto.frx":13D2
         Column(10)      =   "frmNuevoPresupuesto.frx":1542
         Column(11)      =   "frmNuevoPresupuesto.frx":16B2
         Column(12)      =   "frmNuevoPresupuesto.frx":17FA
         Column(13)      =   "frmNuevoPresupuesto.frx":1942
         Column(14)      =   "frmNuevoPresupuesto.frx":1A66
         FmtConditionsCount=   1
         FmtCondition(1) =   "frmNuevoPresupuesto.frx":1B3E
         FormatStylesCount=   16
         FormatStyle(1)  =   "frmNuevoPresupuesto.frx":1C02
         FormatStyle(2)  =   "frmNuevoPresupuesto.frx":1D2A
         FormatStyle(3)  =   "frmNuevoPresupuesto.frx":1DDA
         FormatStyle(4)  =   "frmNuevoPresupuesto.frx":1E8E
         FormatStyle(5)  =   "frmNuevoPresupuesto.frx":1F66
         FormatStyle(6)  =   "frmNuevoPresupuesto.frx":2062
         FormatStyle(7)  =   "frmNuevoPresupuesto.frx":2142
         FormatStyle(8)  =   "frmNuevoPresupuesto.frx":2726
         FormatStyle(9)  =   "frmNuevoPresupuesto.frx":2D0A
         FormatStyle(10) =   "frmNuevoPresupuesto.frx":32FA
         FormatStyle(11) =   "frmNuevoPresupuesto.frx":38E6
         FormatStyle(12) =   "frmNuevoPresupuesto.frx":3972
         FormatStyle(13) =   "frmNuevoPresupuesto.frx":3A4E
         FormatStyle(14) =   "frmNuevoPresupuesto.frx":3B02
         FormatStyle(15) =   "frmNuevoPresupuesto.frx":3BB6
         FormatStyle(16) =   "frmNuevoPresupuesto.frx":3C8E
         ImageCount      =   1
         ImagePicture(1) =   "frmNuevoPresupuesto.frx":3D62
         PrinterProperties=   "frmNuevoPresupuesto.frx":407C
      End
      Begin XtremeSuiteControls.PushButton cmdDefinir 
         Height          =   345
         Left            =   1680
         TabIndex        =   30
         Top             =   4650
         Width           =   1425
         _Version        =   786432
         _ExtentX        =   2514
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Definir"
         BackColor       =   16761024
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton Command13 
         Height          =   345
         Left            =   3225
         TabIndex        =   31
         Top             =   4650
         Width           =   1425
         _Version        =   786432
         _ExtentX        =   2514
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Redondeo"
         BackColor       =   16761024
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnRenumerarItems 
         Height          =   315
         Left            =   2850
         TabIndex        =   32
         Top             =   195
         Width           =   2625
         _Version        =   786432
         _ExtentX        =   4630
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Renumerar Items"
         BackColor       =   16761024
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.TabControl TabControl1 
         Height          =   2025
         Left            =   9450
         TabIndex        =   76
         Top             =   4650
         Width           =   4710
         _Version        =   786432
         _ExtentX        =   8308
         _ExtentY        =   3572
         _StockProps     =   68
         Appearance      =   10
         Color           =   32
         ItemCount       =   4
         Item(0).Caption =   "Precios"
         Item(0).ControlCount=   11
         Item(0).Control(0)=   "Label26"
         Item(0).Control(1)=   "Label25"
         Item(0).Control(2)=   "dtoManual"
         Item(0).Control(3)=   "dtoSistema"
         Item(0).Control(4)=   "Label28"
         Item(0).Control(5)=   "subtotManual"
         Item(0).Control(6)=   "lblTotalManual"
         Item(0).Control(7)=   "Label16"
         Item(0).Control(8)=   "lblTotalSistema"
         Item(0).Control(9)=   "subtotSistema"
         Item(0).Control(10)=   "tota"
         Item(1).Caption =   "Más Datos P.U. Manual"
         Item(1).ControlCount=   10
         Item(1).Control(0)=   "lblTotalMateriales"
         Item(1).Control(1)=   "Label41"
         Item(1).Control(2)=   "lblTotalMdo"
         Item(1).Control(3)=   "Label39"
         Item(1).Control(4)=   "Label29"
         Item(1).Control(5)=   "Label27"
         Item(1).Control(6)=   "lblCosto"
         Item(1).Control(7)=   "lblGg"
         Item(1).Control(8)=   "Label19"
         Item(1).Control(9)=   "lblUtilidad"
         Item(2).Caption =   "Hist Cot"
         Item(2).ControlCount=   1
         Item(2).Control(0)=   "GridHistCot"
         Item(3).Caption =   "Hist Fab"
         Item(3).ControlCount=   1
         Item(3).Control(0)=   "gridHistFab"
         Begin GridEX20.GridEX gridHistFab 
            Height          =   1470
            Left            =   -69895
            TabIndex        =   99
            Top             =   435
            Visible         =   0   'False
            Width           =   4530
            _ExtentX        =   7990
            _ExtentY        =   2593
            Version         =   "2.0"
            BoundColumnIndex=   ""
            ReplaceColumnIndex=   ""
            ColumnAutoResize=   -1  'True
            MethodHoldFields=   -1  'True
            AllowEdit       =   0   'False
            GroupByBoxVisible=   0   'False
            DataMode        =   99
            ColumnHeaderHeight=   285
            IntProp1        =   0
            IntProp2        =   0
            IntProp7        =   0
            ColumnsCount    =   4
            Column(1)       =   "frmNuevoPresupuesto.frx":42A4
            Column(2)       =   "frmNuevoPresupuesto.frx":43C4
            Column(3)       =   "frmNuevoPresupuesto.frx":44B0
            Column(4)       =   "frmNuevoPresupuesto.frx":45A4
            FormatStylesCount=   6
            FormatStyle(1)  =   "frmNuevoPresupuesto.frx":4690
            FormatStyle(2)  =   "frmNuevoPresupuesto.frx":47C8
            FormatStyle(3)  =   "frmNuevoPresupuesto.frx":4878
            FormatStyle(4)  =   "frmNuevoPresupuesto.frx":492C
            FormatStyle(5)  =   "frmNuevoPresupuesto.frx":4A04
            FormatStyle(6)  =   "frmNuevoPresupuesto.frx":4ABC
            ImageCount      =   0
            PrinterProperties=   "frmNuevoPresupuesto.frx":4B9C
         End
         Begin GridEX20.GridEX GridHistCot 
            Height          =   1455
            Left            =   -69895
            TabIndex        =   100
            Top             =   435
            Visible         =   0   'False
            Width           =   4530
            _ExtentX        =   7990
            _ExtentY        =   2566
            Version         =   "2.0"
            BoundColumnIndex=   ""
            ReplaceColumnIndex=   ""
            ColumnAutoResize=   -1  'True
            MethodHoldFields=   -1  'True
            AllowEdit       =   0   'False
            GroupByBoxVisible=   0   'False
            DataMode        =   99
            ColumnHeaderHeight=   285
            IntProp1        =   0
            IntProp2        =   0
            IntProp7        =   0
            ColumnsCount    =   4
            Column(1)       =   "frmNuevoPresupuesto.frx":4D74
            Column(2)       =   "frmNuevoPresupuesto.frx":4E94
            Column(3)       =   "frmNuevoPresupuesto.frx":4F80
            Column(4)       =   "frmNuevoPresupuesto.frx":5070
            FormatStylesCount=   6
            FormatStyle(1)  =   "frmNuevoPresupuesto.frx":515C
            FormatStyle(2)  =   "frmNuevoPresupuesto.frx":5294
            FormatStyle(3)  =   "frmNuevoPresupuesto.frx":5344
            FormatStyle(4)  =   "frmNuevoPresupuesto.frx":53F8
            FormatStyle(5)  =   "frmNuevoPresupuesto.frx":54D0
            FormatStyle(6)  =   "frmNuevoPresupuesto.frx":5588
            ImageCount      =   0
            PrinterProperties=   "frmNuevoPresupuesto.frx":5668
         End
         Begin VB.Label lblUtilidad 
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   -68605
            TabIndex        =   97
            Top             =   1515
            Visible         =   0   'False
            Width           =   2010
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Utilidad"
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
            Left            =   -69580
            TabIndex        =   96
            Top             =   1515
            Visible         =   0   'False
            Width           =   885
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Gastos Grales"
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
            Left            =   -69880
            TabIndex        =   95
            Top             =   1020
            Visible         =   0   'False
            Width           =   1200
         End
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Costo"
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
            Left            =   -69580
            TabIndex        =   94
            Top             =   1275
            Visible         =   0   'False
            Width           =   885
         End
         Begin VB.Label lblGg 
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   -68605
            TabIndex        =   93
            Top             =   990
            UseMnemonic     =   0   'False
            Visible         =   0   'False
            Width           =   2025
         End
         Begin VB.Label lblCosto 
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   -68605
            TabIndex        =   92
            Top             =   1245
            Visible         =   0   'False
            Width           =   2010
         End
         Begin VB.Label Label39 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Total MDO"
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
            Left            =   -69655
            TabIndex        =   91
            Top             =   525
            Visible         =   0   'False
            Width           =   990
         End
         Begin VB.Label lblTotalMdo 
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   -68605
            TabIndex        =   90
            Top             =   495
            Visible         =   0   'False
            Width           =   2040
         End
         Begin VB.Label Label41 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Total MAT"
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
            Left            =   -69670
            TabIndex        =   89
            Top             =   765
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.Label lblTotalMateriales 
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   -68605
            TabIndex        =   88
            Top             =   735
            Visible         =   0   'False
            Width           =   2040
         End
         Begin VB.Label tota 
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Sub Total"
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
            Left            =   285
            TabIndex        =   87
            Top             =   930
            Width           =   855
         End
         Begin VB.Label subtotSistema 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            Height          =   255
            Left            =   1260
            TabIndex        =   86
            Top             =   930
            Width           =   975
         End
         Begin VB.Label lblTotalSistema 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            Height          =   255
            Left            =   1260
            TabIndex        =   85
            Top             =   1410
            Width           =   975
         End
         Begin VB.Label Label16 
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Total"
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
            Left            =   660
            TabIndex        =   84
            Top             =   1410
            Width           =   495
         End
         Begin VB.Label lblTotalManual 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            Height          =   255
            Left            =   2295
            TabIndex        =   83
            Top             =   1410
            Width           =   975
         End
         Begin VB.Label subtotManual 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            Height          =   255
            Left            =   2295
            TabIndex        =   82
            Top             =   930
            Width           =   975
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Descuento"
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
            TabIndex        =   81
            Top             =   1155
            Width           =   930
         End
         Begin VB.Label dtoSistema 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            Height          =   255
            Left            =   1260
            TabIndex        =   80
            Top             =   1170
            Width           =   975
         End
         Begin VB.Label dtoManual 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            Height          =   255
            Left            =   2295
            TabIndex        =   79
            Top             =   1170
            Width           =   975
         End
         Begin VB.Label Label25 
            Alignment       =   2  'Center
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Sistema"
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
            Left            =   1230
            TabIndex        =   78
            Top             =   510
            Width           =   975
         End
         Begin VB.Label Label26 
            Alignment       =   2  'Center
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Venta"
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
            Left            =   2310
            TabIndex        =   77
            Top             =   510
            Width           =   975
         End
      End
      Begin VB.Label lblModoEdicion 
         BackColor       =   &H00FFC0C0&
         Caption         =   "MODO EDICION"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   11265
         TabIndex        =   34
         ToolTipText     =   "Presione <ENTER> para terminar de editar el campo"
         Top             =   225
         Width           =   1335
      End
      Begin VB.Label lblRecalculando 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "Recalculando..."
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
         Left            =   6060
         TabIndex        =   33
         ToolTipText     =   "Presione <ENTER> para terminar de editar el campo"
         Top             =   225
         Visible         =   0   'False
         Width           =   1260
      End
   End
   Begin VB.Menu m1 
      Caption         =   "Menu1"
      Visible         =   0   'False
      Begin VB.Menu Ver 
         Caption         =   "Ver Desarrollo..."
      End
      Begin VB.Menu archivos 
         Caption         =   "Archivos Asociados..."
         Begin VB.Menu mnuArchiPieza 
            Caption         =   "De Pieza..."
         End
         Begin VB.Menu mnuArchiPedido 
            Caption         =   "Del Detalle..."
         End
      End
      Begin VB.Menu scanear 
         Caption         =   "Adquirir..."
         Begin VB.Menu mnuEscaPieza 
            Caption         =   "A Pieza..."
         End
         Begin VB.Menu mnuEscaPedido 
            Caption         =   "Al Detalle..."
         End
      End
      Begin VB.Menu verIncidencias 
         Caption         =   "Ver Incidencias..."
         Begin VB.Menu mnuInciPieza 
            Caption         =   "De Pieza..."
         End
         Begin VB.Menu mnuInciPedido 
            Caption         =   "Del Detalle..."
         End
      End
   End
End
Attribute VB_Name = "frmVentasPresupuestoEditar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Implements ISuscriber
Dim historico As New Collection    'of DTOHistoricoPieza
Dim Inicio As Integer
Dim CantArchivos As Dictionary
Dim A As FormaCotizar
Dim id_suscriber As String
Dim item As String
Dim Detalles As Collection
Dim tmpHist As dtoHistoricoPieza
Dim tmpDetalle As clsPresupuestoDetalle
Dim rows As Long
Dim grabado As Boolean
Dim idpresu As Long
Dim vcosto As Double

Dim claseS As New classStock
Dim tmpPresupuesto As clsPresupuesto
Public Property Let nroPresu(id As Long)
    idpresu = id
End Property
Public Function ProximoItem() As String
    If tmpPresupuesto.DetallePresupuesto.count = 0 Then
        item = 0
    Else
        item = tmpPresupuesto.DetallePresupuesto.count
    End If
    item = item + 1
    ProximoItem = Format(CStr(item), "000")
End Function

Public Function MostrarPresupuesto(Optional BeforSave As Boolean = False)
    Set tmpPresupuesto = DAOPresupuestos.GetById(idpresu)
    Set tmpPresupuesto.DetallePresupuesto = DAOPresupuestosDetalle.GetAllByPresupuesto(tmpPresupuesto)
    Me.txtMOM = tmpPresupuesto.PorcentajeManoObraMuerta
    Me.txtDiasPagoAnticipo = tmpPresupuesto.DiasPagoAnticipo
    Me.txtDiasPagoSaldo = tmpPresupuesto.DiasPagoSaldo
    Me.caption = "Presupuesto en Curso Nro. " & idpresu
    Me.lblGastos = tmpPresupuesto.Gastos
    Me.txtMarkupMDO = tmpPresupuesto.PorcMDO
    Me.men10 = tmpPresupuesto.PorcMen10
    Me.DTVencimiento = Now
    Me.men15 = tmpPresupuesto.PorcMen15
    Me.mas15 = tmpPresupuesto.PorcMas15
    Me.txtAnticipo = tmpPresupuesto.Anticipo
    Me.txtFormaPagoAnticipo = tmpPresupuesto.FormaPagoAnticipo
    Me.txtFormaPagoSaldo = tmpPresupuesto.FormaPagoSaldo
    Me.manteOferta = tmpPresupuesto.manteOferta
    Me.cboMoneda.ListIndex = tmpPresupuesto.moneda.id
    Me.txtReferencia = tmpPresupuesto.detalle
    Me.txtDias = tmpPresupuesto.FechaEntrega
    Me.txtDescuento = tmpPresupuesto.Descuento
    Me.cboCliente.ListIndex = PosIndexCbo(tmpPresupuesto.cliente.id, Me.cboCliente)
    Me.DTVencimiento = tmpPresupuesto.VencimientoPresupuesto

    If Not BeforSave Then llenarLista


End Function

Public Function llenarLista()
    Me.grilla.ItemCount = 0
    Me.grilla.ItemCount = tmpPresupuesto.DetallePresupuesto.count
    VerFabricados
End Function
Private Function Guardar() As Boolean
    Dim tmpParaEstado As New clsPresupuesto
    Set tmpParaEstado = DAOPresupuestos.GetById(tmpPresupuesto.id)

    If tmpParaEstado.EstadoPresupuesto <> ACotizar_ Then
        MsgBox "El presupuesto ya cambio de estado!" & Chr(10) & "No se puede volver a guardar en esta sesión.", vbCritical, "Error"
        Exit Function
    End If
    g = MsgBox("¿Está conforme con los datos ingresados?", vbYesNo, "Confirmación")
    If g = 6 Then
        Set tmpPresupuesto.cliente = DAOCliente.BuscarPorID(CLng(Me.cboCliente.ItemData(Me.cboCliente.ListIndex)))
        tmpPresupuesto.Anticipo = CDbl(Me.txtAnticipo)
        tmpPresupuesto.Descuento = CDbl(Me.txtDescuento)
        tmpPresupuesto.detalle = UCase(Me.txtReferencia)
        tmpPresupuesto.EstadoPresupuesto = ACotizar_
        tmpPresupuesto.FormaPagoAnticipo = UCase(Me.txtFormaPagoAnticipo)
        tmpPresupuesto.FormaPagoSaldo = UCase(Me.txtFormaPagoSaldo)
        tmpPresupuesto.DiasPagoAnticipo = CLng(Me.txtDiasPagoAnticipo)
        tmpPresupuesto.DiasPagoSaldo = CLng(Me.txtDiasPagoSaldo)
        tmpPresupuesto.Gastos = CDbl(Me.lblGastos)
        tmpPresupuesto.id = idpresu
        tmpPresupuesto.manteOferta = Me.manteOferta
        Set tmpPresupuesto.moneda = DAOMoneda.GetById(CLng(Me.cboMoneda.ItemData(Me.cboMoneda.ListIndex)))
        tmpPresupuesto.PorcMas15 = CDbl(mas15)
        tmpPresupuesto.PorcMDO = CDbl(Me.txtMarkupMDO)
        tmpPresupuesto.PorcMen15 = CDbl(men15)
        tmpPresupuesto.PorcMen10 = CDbl(men10)
        tmpPresupuesto.VencimientoPresupuesto = Me.DTVencimiento
        tmpPresupuesto.FechaEntrega = CLng(Me.txtDias)


        If Not DAOPresupuestos.Save(tmpPresupuesto) Then
            MsgBox "Se produjo algun error!", vbCritical, "Error"
        Else
            MsgBox "Actualización exitosa!", vbInformation, "Información"
            Dim EVENTO As New clsEventoObserver
            Set EVENTO.Elemento = tmpPresupuesto
            EVENTO.EVENTO = modificar_
            Set EVENTO.Originador = Me
            Channel.Notificar EVENTO, Presupuestos_
            grabado = True
            MostrarPresupuesto True
        End If
    End If
End Function

Private Sub btnRenumerarItems_Click()
    If MsgBox("¿Esta seguro que desea renumerar los items?", vbYesNo + vbQuestion) = vbYes Then

        Dim count As Long: count = 0
        For Each tmpDetalle In tmpPresupuesto.DetallePresupuesto
            count = count + 1
            tmpDetalle.item = Format(count, "000")
        Next

        Me.grilla.ReBind
        Me.grilla.Refetch
    End If
End Sub

Private Sub cboCliente_Click()
    Inicio = Inicio + 1
    On Error GoTo err1
    If Not tmpPresupuesto Is Nothing Then
        If Me.cboCliente.ItemData(Me.cboCliente.ListIndex) <> CInt(tmpPresupuesto.cliente.id) Then
            h = MsgBox("¿Desea cambiar el cliente seleccionado?", vbYesNo, "Confirmación")
            If h = 6 Then
                Me.grilla.ItemCount = 0
                Set tmpPresupuesto.cliente = DAOCliente.BuscarPorID(Me.cboCliente.ListIndex)
            Else
                Me.cboCliente.ItemData(Me.cboCliente.ListIndex) = tmpPresupuesto.cliente.id
            End If

        End If
    End If
    Exit Sub
err1:

End Sub

Private Sub cboMoneda_Change()
    grabado = False
End Sub

Private Sub cboMoneda_Click()
    On Error Resume Next
    Dim vmo As clsMoneda
    Set vmo = DAOMoneda.GetById(cboMoneda.ItemData(Me.cboMoneda.ListIndex))
    Set tmpPresupuesto.moneda = DAOMoneda.GetById(cboMoneda.ItemData(Me.cboMoneda.ListIndex))
End Sub
Private Sub Command1_Click()
    Dim col As Collection
    Dim id As Long
    Dim f222 As New frmElegirPieza
    f222.Origen = 1    'desde un presupuesto
    f222.cliente = tmpPresupuesto.cliente
    f222.Show 1
End Sub

Private Sub Command10_Click()
    If MsgBox("¿Desea amortizar los items seleccionados por la cantidad?", vbYesNo, "Confirmación") = vbYes Then
        Dim A As JSSelectedItem
        Dim d As clsPresupuestoDetalle

        For Each A In grilla.SelectedItems
            Set d = tmpPresupuesto.DetallePresupuesto(A.RowIndex)
            d.Amortizacion = d.Cantidad
            grilla.RefreshRowIndex A.RowIndex
        Next

        grabado = False
    End If
End Sub

Private Sub Command11_Click()
    Dim idPieza As Long
    Dim Cantidad As Double
    Dim claseP As New classPlaneamiento

    If MsgBox("¿Desea amortizar los items seleccionados por la cantidad fabricada?", vbYesNo, "Confirmación") = vbYes Then
        Dim A As JSSelectedItem
        Dim d As clsPresupuestoDetalle

        For Each A In grilla.SelectedItems
            Set d = tmpPresupuesto.DetallePresupuesto(A.RowIndex)
            d.Amortizacion = claseP.cantidadFabricada(d.Pieza.id, d.Cantidad)
            grilla.RefreshRowIndex A.RowIndex
        Next

        grabado = False
        Set claseP = Nothing

    End If

End Sub



Private Sub Command13_Click()
    Dim A As JSSelectedItem
    Dim b As clsPresupuestoDetalle
    For Each A In Me.grilla.SelectedItems
        Set b = tmpPresupuesto.DetallePresupuesto(A.RowIndex)
        b.ValorManual = funciones.FormatearDecimales(Math.Round(b.ValorManual, 0), 2)

        grilla.RefreshRowIndex A.RowIndex
    Next
End Sub

Private Sub Command2_Click()
    VerFabricados
    recalcule

End Sub

Private Sub Command3_Click()
    If grilla.EditMode = jgexEditModeOn Then
        verModoEdicion
        MsgBox "Salga del modo edicion para continuar!", vbInformation
    Else

        Guardar
    End If
End Sub

Private Sub Command4_Click()
    Dim col As New Collection
    Set col = DAOPieza.FindAll(FL_0)
    Dim Pieza As Pieza
    Dim Cant As Long
    Dim claseP As New classPlaneamiento


    For Each Pieza In col

        Cant = claseP.cantidadFabricada(Pieza.id)

        If Cant > 0 Then
            Pieza.YaFabricada = True
            If Not DAOPieza.Save(Pieza) Then GoTo err1
        End If


    Next Pieza
    Exit Sub
err1:

End Sub

Private Sub Command6_Click()
    Dim Fila As JSSelectedItem
    Dim dato As clsPresupuestoDetalle
    For Each Fila In grilla.SelectedItems
        Set dato = tmpPresupuesto.DetallePresupuesto(Fila.RowIndex)
        dato.ValorManual = dato.ValorSistema
        grilla.RefreshRowIndex Fila.RowIndex
    Next
End Sub

Private Sub Command7_Click()
    imprimirPresu
End Sub

Private Sub imprimirPresu()
    Dim header As String

    headercenter = "PRESUPUESTO NUMERO " & tmpPresupuesto.id & Chr(10) _
                   & "Cliente: " & tmpPresupuesto.cliente.id & " " & tmpPresupuesto.cliente.razon & Chr(10) _
                   & "Referencia: " & tmpPresupuesto.detalle & Chr(10) _
                   & "Entrega: " & tmpPresupuesto.FechaEntrega & " días" & Chr(10)


    headerLeft = "Mark Up: " & Chr(10) _
                 & Space(5) & "Gastos Grales: " & tmpPresupuesto.Gastos & "%, M.O.M: " & tmpPresupuesto.PorcentajeManoObraMuerta & "%" & Chr(10) _
                 & Space(5) & "Mano de Obra: " & tmpPresupuesto.PorcMDO & "%" & Chr(10) _
                 & Space(5) & "Materiales: " & Chr(10) _
                 & Space(10) & "<10Kg: " & tmpPresupuesto.PorcMen10 & "%" & Chr(10) _
                 & Space(10) & "<15Kg: " & tmpPresupuesto.PorcMen15 & "%" & Chr(10) _
                 & Space(10) & ">15Kg: " & tmpPresupuesto.PorcMas15 & "%" & Chr(10)


    footerLeft = "Total Sistema: " & tmpPresupuesto.moneda.NombreCorto & Space(1) & tmpPresupuesto.Total(Sistema) & Chr(10) _
                 & "Total Manual: " & tmpPresupuesto.moneda.NombreCorto & Space(1) & tmpPresupuesto.Total(Manual) & Chr(10) _
                 & "Anticipo " & tmpPresupuesto.Anticipo & " Saldo / FP: " & tmpPresupuesto.FormaPagoAnticipo





    With Me.grilla.PrinterProperties
        .HeaderDistance = 500
        .FooterDistance = 1550
        .TopMargin = 2000
        .BottomMargin = 2000

        .FitColumns = True
        .DocumentName = "Presupuesto"

        .RepeatHeaders = True
        .Orientation = jgexPPLandscape
        .HeaderString(jgexHFCenter) = headercenter
        .HeaderString(jgexHFLeft) = headerLeft

        .FooterString(jgexHFLeft) = footerLeft
        .FooterString(jgexHFCenter) = Now



    End With
    Load frmPrintPreview
    frmPrintPreview.Move Me.Left, Me.Top, Me.Width, Me.Height
    grilla.PrintPreview frmPrintPreview.GEXPreview1
    frmPrintPreview.Show 1
End Sub
Private Sub Command8_Click()
    Dim claseP As New classPlaneamiento
    A = claseP.informePiezaMateriales(idpresu, 2, True)


End Sub

Private Sub Command9_Click()
    If Not grabado Then
        If MsgBox("¿Desea perder los cambios?", vbYesNo, "Confirmación") = vbYes Then
            Unload Me
        End If
    Else
        Unload Me
    End If
End Sub

Private Sub E_Click()
    Dim rs As Recordset
    Dim dto As DTOPiezaCantidad
    Dim dp As clsPresupuestoDetalle
    Dim listadtopiezacantidad As New Collection
    For Each dp In tmpPresupuesto.DetallePresupuesto
        Set dto = New DTOPiezaCantidad
        Set dto.Pieza = dp.Pieza
        dto.Cantidad = dp.Cantidad
        listadtopiezacantidad.Add dto
    Next dp

    Dim frm1 As New frmEstadistiacasEnCurso
    frm1.caption = "Estadisticas de presupuesto activo"
    Set frm1.listadtopiezacantidad = listadtopiezacantidad
    frm1.conjGrabado = False
    frm1.Show
End Sub



Private Sub Form_Deactivate()
    If Not grabado Then
        If MsgBox("¿Desea grabar los cambios?", vbYesNo, "Confirmación") = vbYes Then
            Command3_Click
        End If
    End If
End Sub
Private Sub verModoEdicion()


    Me.lblModoEdicion.Visible = (Me.grilla.EditMode = jgexEditModeOn)


End Sub

Private Sub Form_Load()
    FormHelper.Customize Me
    Inicio = 0
    Me.lblModoEdicion.ToolTipText = "Presione <ENTER> para terminar  ó <ESC> para cancelar"
    verModoEdicion
    Set tmpDetalle = New clsPresupuestoDetalle
    GridEXHelper.CustomizeGrid Me.grilla, False, True
    GridEXHelper.CustomizeGrid Me.GridHistCot, False, False
    GridEXHelper.CustomizeGrid Me.gridHistFab, False, False
    id_suscriber = funciones.CreateGUID
    Channel.AgregarSuscriptor Me, NuevoPresupuesto_
    Channel.AgregarSuscriptor Me, EdicionPieza_
    DAOCliente.LlenarCombo Me.cboCliente, False, False, False
    DAOMoneda.LlenarCombo Me.cboMoneda
    MostrarPresupuesto

    Set CantArchivos = DAOArchivo.GetCantidadArchivosPorReferencia(OA_Piezas)
    grabado = True
End Sub
Private Sub Form_Terminate()
    Channel.RemoverSuscripcionTotal Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Channel.RemoverSuscripcionTotal Me
End Sub


Private Sub gg2_Click()

End Sub

Private Sub GridHistCot_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    On Error Resume Next

    Values(1) = historico(RowIndex).Origen
    Values(2) = historico(RowIndex).Cantidad
    Values(3) = historico(RowIndex).Monto
    Values(4) = historico(RowIndex).FEcha
End Sub

Private Sub gridHistFab_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    On Error Resume Next
    Values(1) = historico(RowIndex).Origen
    Values(2) = historico(RowIndex).Cantidad
    Values(3) = historico(RowIndex).Monto
    Values(4) = historico(RowIndex).FEcha
End Sub

Private Sub grilla_BeforePrintPage(ByVal PageNumber As Long, ByVal nPages As Long)
    grilla.PrinterProperties.FooterString(jgexHFRight) = "Página " & PageNumber & " de " & nPages
End Sub
Private Sub grilla_Click()
    verModoEdicion
End Sub

Private Sub grilla_FetchIcon(ByVal RowIndex As Long, ByVal ColIndex As Integer, ByVal RowBookmark As Variant, ByVal IconIndex As GridEX20.JSRetInteger)
    On Error Resume Next
    Set tmpDetalle = tmpPresupuesto.DetallePresupuesto.item(grilla.RowIndex(RowIndex))

    If CantArchivos.item(tmpDetalle.Pieza.id) > 0 Then
        If ColIndex = 12 Then
            IconIndex = 1
        End If
    End If

End Sub

Private Sub grilla_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 67 And Shift = 2 Then    'CTRL + C
        GridEXHelper.Grid2Clipboard Me.grilla
        DoEvents
        MsgBox "La lista ha sido copiada al portapapeles.", vbInformation + vbOKOnly
    End If

End Sub

Private Sub grilla_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        grilla.EditMode = jgexEditModeOff
        verModoEdicion
    End If
End Sub

    Private Sub grilla_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If tmpPresupuesto.DetallePresupuesto.count > 0 Then
        Set tmpDetalle = tmpPresupuesto.DetallePresupuesto(grilla.RowIndex(grilla.row))
        If Button = 2 Then
            If tmpDetalle.Pieza.EsConjunto Then
                Me.ver.caption = "Ver Conjunto..."
                Me.ver.Tag = 0
            Else
                Me.ver.caption = "Ver Desarrollo..."
                Me.ver.Tag = -1
            End If
            Me.PopupMenu Me.m1
        End If
    End If
End Sub
Private Sub grilla_RowFormat(RowBuffer As GridEX20.JSRowData)
    On Error Resume Next
    Set tmpDetalle = tmpPresupuesto.DetallePresupuesto.item(RowBuffer.RowIndex)
    If tmpDetalle.FormaCotizar = FormaCotizar.automatica_ Then
        RowBuffer.CellStyle(3) = "auto"
    ElseIf tmpDetalle.FormaCotizar = Cantidad_ Then
        RowBuffer.CellStyle(3) = "cantidad"
    ElseIf tmpDetalle.FormaCotizar = fabricados_ Then
        RowBuffer.CellStyle(3) = "fabricados"
    ElseIf tmpDetalle.FormaCotizar = fijo_ Then
        RowBuffer.CellStyle(3) = "fijo"
    End If

    If tmpDetalle.Pieza.YaFabricada Then
        RowBuffer.CellStyle(1) = "ya_fabricado"
    End If


    If tmpDetalle.Pieza.Complejidad = ComplejidadAlta Then
        RowBuffer.CellStyle(14) = "comp_alta"
    ElseIf tmpDetalle.Pieza.Complejidad = Complejidadmedia Then
        RowBuffer.CellStyle(14) = "comp_media"
        ElseIf tmpDetalle.Pieza.Complejidad = ComplejidadBaja Then
        RowBuffer.CellStyle(14) = "comp_baja"
    End If



If tmpDetalle.indiceAjuste < 0 Then
 RowBuffer.CellStyle(13) = "comp_alta"
ElseIf tmpDetalle.indiceAjuste > 0 Then
 RowBuffer.CellStyle(13) = "naranja"
Else
End If

End Sub
Private Sub MostrarFormaCotizar(A As FormaCotizar)
    Select Case A
        Case FormaCotizar.automatica_: Me.opAutomatica.value = True
        Case FormaCotizar.Cantidad_: Me.opCantidad.value = True
        Case FormaCotizar.fabricados_: Me.opFabricados.value = True
        Case FormaCotizar.fijo_: Me.opFijo.value = True
    End Select
End Sub
Private Sub grilla_SelectionChange()
    On Error Resume Next


    MostrarFormaCotizar tmpPresupuesto.DetallePresupuesto(grilla.RowIndex(grilla.row)).FormaCotizar
    rows = grilla.RowIndex(grilla.row)


    If Me.TabControl1(2).Selected Then
        MostrarHistorico 1
    End If

    If Me.TabControl1(3).Selected Then
        MostrarHistorico 2
    End If

End Sub

Private Sub MostrarHistorico(Tipo As Integer)
    Dim deta As DetalleOrdenTrabajo
    Dim deta1 As clsPresupuestoDetalle
    Dim hist As New dtoHistoricoPieza
    Dim col1 As New Collection
    Dim col As New Collection

    Set historico = New Collection
    If Tipo = 2 Then    'fabricado

        col.Add tmpDetalle.Pieza
        Set col1 = DAODetalleOrdenTrabajo.FindAllByPieza(col)
        For Each deta In col1
            Set hist = New dtoHistoricoPieza
            hist.Cantidad = deta.CantidadPedida
            hist.FEcha = deta.FechaEntrega
            hist.Monto = deta.Precio
            hist.Origen = deta.OrdenTrabajo.id & " | " & deta.item
            historico.Add hist
        Next


        Me.gridHistFab.ItemCount = 0
        Me.gridHistFab.ItemCount = historico.count
        Me.gridHistFab.ColumnAutoResize = True
    ElseIf Tipo = 1 Then    'cotizado

        Set col1 = DAOPresupuestosDetalle.GetAllByPieza(tmpDetalle.Pieza.id)
        For Each deta1 In col1
            Set hist = New dtoHistoricoPieza

            hist.Cantidad = deta1.Cantidad
            hist.FEcha = deta1.FechaPresupuesto
            hist.Monto = deta1.ValorManual
            hist.Origen = deta1.idPreuspuesto & " | " & deta1.item
            historico.Add hist
        Next
        Me.GridHistCot.ItemCount = 0
        Me.GridHistCot.ItemCount = historico.count
        Me.GridHistCot.ColumnAutoResize = True

    End If

End Sub
Private Sub grilla_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex > tmpPresupuesto.DetallePresupuesto.count Then Exit Sub
    Set tmpDetalle = tmpPresupuesto.DetallePresupuesto.item(RowIndex)
    With tmpDetalle
        Values(1) = .item
        Values(2) = .Cantidad
        Values(3) = .Amortizacion
        Values(4) = .Detalles
        Values(5) = funciones.FormatearDecimales(.ValorSistema)
        Values(6) = funciones.FormatearDecimales(.ValorManual)
        Values(7) = funciones.FormatearDecimales(.ValorManual * .Cantidad)
        Values(8) = .entrega
        Values(10) = .PorcentajeMAT & "% (" & .TotalMateriales & ")"
        Values(9) = .PorcentajeMDO & "% (" & .TotalMDO & ")"
        Values(11) = .Pieza.nombre
    Values(13) = .indiceAjuste
    Values(14) = enums.EnumTiposComplejidad(.Pieza.Complejidad)
    End With
End Sub

Private Sub grilla_UnboundUpdate(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    On Error GoTo err1
    Set tmpDetalle = tmpPresupuesto.DetallePresupuesto.item(RowIndex)
    tmpDetalle.item = Values(1)
    tmpDetalle.Cantidad = Values(2)
    tmpDetalle.entrega = Values(8)
    tmpDetalle.Amortizacion = Values(3)
    tmpDetalle.ValorManual = Values(6)
    tmpDetalle.Detalles = Values(4)
    tmpDetalle.indiceAjuste = Values(13)
    verModoEdicion
    Exit Sub
err1:
End Sub

Private Property Get ISuscriber_id() As String
    ISuscriber_id = id_suscriber
End Property
Private Function ISuscriber_Notificarse(EVENTO As clsEventoObserver) As Variant
    Dim col As New Collection
    Dim x As Long
    If EVENTO Is Nothing Then
        recalcule
    Else
        If EVENTO.EVENTO = agregarColeccion_ Then
            Set col = EVENTO.Elemento
            Dim dto As DTOPiezaDetallePedido
            'For X = 1 To col.count
            For Each dto In col
                Set tmpDetalle = New clsPresupuestoDetalle
                tmpDetalle.Amortizacion = amortiza(1)
                tmpDetalle.Cantidad = 1
                tmpDetalle.Detalles = ""
                tmpDetalle.entrega = 1
                tmpDetalle.FormaCotizar = automatica_
                tmpDetalle.id = 0
                tmpDetalle.item = ProximoItem
                Set tmpDetalle.presupuesto = tmpPresupuesto
                Set tmpDetalle.Pieza = dto.Pieza    ' col(X)
                tmpDetalle.ValorManual = 0
                tmpDetalle.ValorSistema = 0
                tmpPresupuesto.DetallePresupuesto.Add tmpDetalle
            Next dto
            llenarLista
            grilla.MoveLast
            verModoEdicion
        End If
    End If

    VerFabricados
End Function

Private Sub VerFabricados()
    Dim deta As clsPresupuestoDetalle
    Dim c As Long: c = 0
    Dim fabri As Long: fabri = 0
    Dim tota As Long: tota = 0
    For Each deta In tmpPresupuesto.DetallePresupuesto
        c = c + 1
        If deta.Pieza.YaFabricada Then fabri = fabri + 1
    Next deta

    If tmpPresupuesto.DetallePresupuesto.count > 0 Then
        tota = 100 - funciones.FormatearDecimales((fabri / c) * 100, 1)
        If tota = 0 Then
            Me.lblAFabricar = "** El pedido fue fabricado en algún momento **"
        Else
            Me.lblAFabricar = "** Deberá fabricar por primera vez  el " & tota & "% del pedido **"
        End If

    Else
        Me.lblAFabricar = "** No hay elementos a fabricar. **"
    End If


End Sub

Private Sub lblGastos_Change()
    On Error Resume Next
    grabado = False
    tmpPresupuesto.Gastos = CDbl(Me.lblGastos)
End Sub
Private Sub lblGastos_GotFocus()
    foco Me.lblGastos
End Sub

Private Sub lstPresupuesto_BeforeLabelEdit(Cancel As Integer)
    grabado = False
End Sub

Private Sub lblGastos_Validate(Cancel As Boolean)
    ValidarTextBox Me.lblGastos, Cancel
End Sub



Private Sub manteOferta_Change()
    On Error Resume Next
    grabado = False
    tmpPresupuesto.manteOferta = CLng(manteOferta)
End Sub
Private Sub manteOferta_GotFocus()
    foco Me.manteOferta
End Sub
Private Sub manteOferta_Validate(Cancel As Boolean)
    funciones.ValidarTextBox Me.manteOferta, Cancel
End Sub
Private Sub mas15_GotFocus()
    foco mas15
End Sub
Private Sub mas15_LostFocus()
    On Error Resume Next
    grabado = False
    tmpPresupuesto.PorcMas15 = CDbl(Me.mas15)
End Sub
Private Sub mas15_Validate(Cancel As Boolean)
    ValidarTextBox Me.mas15, Cancel
End Sub
Private Sub men10_GotFocus()
    foco men10
End Sub
Private Sub men10_LostFocus()
    On Error Resume Next
    grabado = False
    tmpPresupuesto.PorcMen10 = CDbl(Me.men10)
End Sub
Private Sub men10_Validate(Cancel As Boolean)
    ValidarTextBox men10, Cancel
End Sub
Private Sub men15_GotFocus()
    foco men15
End Sub
Private Sub men15_LostFocus()
    On Error Resume Next
    grabado = False
    tmpPresupuesto.PorcMen15 = CDbl(Me.men15)
End Sub
Private Sub men15_Validate(Cancel As Boolean)
    ValidarTextBox men15, Cancel
End Sub
Private Sub Option1_Click()
    Me.txtDias.Enabled = True
    tipos = 1
End Sub
Private Sub Option2_Click()
    Me.txtDias.Enabled = False
    tipos = 2
End Sub
Private Sub cambiar(A As FormaCotizar)
    Dim Fila As JSSelectedItem
    Dim row_data As JSRowData
    Dim deta As clsPresupuestoDetalle
    For Each Fila In grilla.SelectedItems
        Set deta = tmpPresupuesto.DetallePresupuesto(Fila.RowIndex)
        deta.FormaCotizar = A
        grilla.RefreshRowIndex Fila.RowIndex
    Next
End Sub
Private Sub mnuArchiPedido_Click()
    Dim ar2 As New frmArchivos2
    ar2.Origen = 11
    ar2.ObjetoId = tmpDetalle.id
    ar2.caption = "Presupuesto Nº  " & tmpDetalle.id
    ar2.Show
End Sub
Private Sub mnuArchiPieza_Click()
    Dim ar1 As New frmArchivos2
    ar1.Origen = 1
    ar1.ObjetoId = tmpDetalle.Pieza.id
    ar1.caption = "Pieza " & tmpDetalle.Pieza.nombre
    ar1.Show
End Sub
Private Sub mnuEscaPedido_Click()
    Dim archivos As New classArchivos
    archivos.escanearDocumento 11, tmpDetalle.id
End Sub
Private Sub mnuEscaPieza_Click()
    Dim archivos As New classArchivos
    archivos.escanearDocumento 1, tmpDetalle.Pieza.id
End Sub
Private Sub mnuInciPedido_Click()
    Dim i1 As New frmVerIncidencias
    i1.referencia = tmpDetalle.id
    i1.Origen = 33
    i1.Show
End Sub
Private Sub mnuInciPieza_Click()
    Dim i2 As New frmVerIncidencias
    i2.referencia = tmpDetalle.Pieza.id
    i2.Origen = 3
    i2.Show
End Sub
Private Sub opAutomatica_Click()
    A = automatica_
    cambiar A
End Sub
Private Sub opCantidad_Click()
    A = Cantidad_
    cambiar A
End Sub
Private Sub opFabricados_Click()
    A = fabricados_
    cambiar A
End Sub
Private Sub opFijo_Click()
    A = fijo_
    cambiar A
End Sub
Private Sub Qui_Click()

    Dim si As GridEX20.JSSelectedItem
    Dim i As Long

    For i = tmpPresupuesto.DetallePresupuesto.count To 1 Step -1
        For Each si In Me.grilla.SelectedItems
            If si.RowIndex = i Then
                tmpPresupuesto.DetallePresupuesto.remove i
                Exit For
            End If
        Next si
    Next i

    Me.llenarLista
End Sub




Private Sub TabControl1_SelectedChanged(ByVal item As Xtremesuitecontrols.ITabControlItem)
    If item.index = 2 Then MostrarHistorico 1
    If item.index = 3 Then MostrarHistorico 2

End Sub

Private Sub txtAnticipo_GotFocus()
    foco Me.txtAnticipo
End Sub
Private Sub txtAnticipo_Validate(Cancel As Boolean)
    On Error Resume Next
    funciones.ValidarTextBox Me.txtAnticipo, Cancel
End Sub
Private Sub txtDescuento_Change()
    On Error Resume Next
    grabado = False
    tmpPresupuesto.Descuento = CDbl(Me.txtDescuento)
End Sub

Private Sub txtDescuento_GotFocus()
    foco Me.txtDescuento
End Sub

Private Sub txtDescuento_Validate(Cancel As Boolean)
    On Error Resume Next
    ValidarTextBox Me.txtDescuento, Cancel
End Sub
Private Sub txtDias_LostFocus()
    On Error Resume Next
    grabado = False
    tmpPresupuesto.FechaEntrega = CLng(Me.txtDias)
End Sub

Private Sub txtDias_Validate(Cancel As Boolean)
    On Error Resume Next
    ValidarTextBox Me.txtDias, Cancel
End Sub

Private Sub txtDiasPagoAnticipo_GotFocus()
    foco Me.txtDiasPagoAnticipo
End Sub
Private Sub txtDiasPagoSaldo_GotFocus()
    foco Me.txtDiasPagoSaldo
End Sub

Private Sub txtFormaPagoAnticipo_GotFocus()
    foco Me.txtFormaPagoAnticipo
End Sub
Private Sub txtFormaPagoSaldo_GotFocus()
    foco Me.txtFormaPagoSaldo
End Sub
Private Sub txtMarkupMDO_GotFocus()
    foco Me.txtMarkupMDO
End Sub
Private Sub txtMarkupMDO_LostFocus()
    On Error Resume Next
    grabado = False
    tmpPresupuesto.PorcMDO = CDbl(Me.txtMarkupMDO)
End Sub
Private Sub txtMarkupMDO_Validate(Cancel As Boolean)
    On Error Resume Next
    ValidarTextBox Me.txtMarkupMDO, Cancel
End Sub
Private Sub txtMOM_Change()
    On Error Resume Next
    grabado = False
    tmpPresupuesto.PorcentajeManoObraMuerta = CDbl(Me.txtMOM)
End Sub
Private Sub txtMOM_GotFocus()
    foco Me.txtMOM
End Sub
Private Sub txtMOM_Validate(Cancel As Boolean)
    funciones.ValidarTextBox Me.txtMOM, Cancel
End Sub
Private Sub txtReferencia_GotFocus()
    foco txtReferencia
End Sub
Public Sub recalcule()
    vcosto = 0
    On Error Resume Next
    Dim mo As Double, ma As Double
    Dim t_mo As Double, t_ma As Double
    Dim Precio As Double, Kg As Double
    Dim rs As Recordset
    Dim FormaCotizar As FormaCotizar
    Dim id As Long
    Dim pp1 As Double
    Dim amorti As Long
    Dim pp2 As Double
    Dim deta As clsPresupuestoDetalle
    Dim Total As Double
    Dim Cantidad As Double
    Dim Cant As Long
    Dim uni As Double
    Dim unitario As Double, totes1 As Double, totes2 As Double
    Dim totConj As Double
    tote = 0
    Me.progreso.max = tmpPresupuesto.DetallePresupuesto.count
    Me.progreso.min = 1
    Me.progreso.Visible = True
    Me.lblRecalculando.Visible = True

    c = 0



    For Each deta In tmpPresupuesto.DetallePresupuesto
        c = c + 1
        Set tmpDetalle = tmpPresupuesto.DetallePresupuesto(nn)
        Me.progreso.value = c
        Dim fl As FetchLevel
        If deta.Pieza.EsConjunto = False Then    'si no es conjunto no hago el fetch es al pedo y gano reindimiento
            fl = FL_0
        Else
            fl = FL_4
        End If
        Set deta.Pieza = DAOPieza.FindById(deta.Pieza.id, fl, True, True)
        deta.CalcularPrecioSistema vcosto, mo, ma
        t_mo = t_mo + mo
        t_ma = t_ma + ma


    Next

            

    Me.progreso.Visible = False
    Me.lblRecalculando.Visible = False

    Me.lblTotalMateriales = tmpPresupuesto.moneda.NombreCorto & " " & tmpPresupuesto.TotalMateriales & " | " & Math.Round(t_ma / tmpPresupuesto.DetallePresupuesto.count, 0) & "%"
    Me.lblTotalMdo = tmpPresupuesto.moneda.NombreCorto & " " & tmpPresupuesto.TotalMDO & " | " & Math.Round(t_mo / tmpPresupuesto.DetallePresupuesto.count, 0) & "%"
    Me.lblCosto = tmpPresupuesto.moneda.NombreCorto & " " & tmpPresupuesto.Total(SMCosto)
    Me.lblGg = tmpPresupuesto.moneda.NombreCorto & " " & tmpPresupuesto.Total(SMGG)
    Me.lblUtilidad = tmpPresupuesto.moneda.NombreCorto & " " & tmpPresupuesto.Total(SMUtilidad)
    calcular
    llenarLista

End Sub

Private Sub calcular()
    Dim subtoM As Double
    Dim subtos As Double
    Dim dtos As Double
    Dim dtoM As Double
    subtoM = funciones.FormatearDecimales(tmpPresupuesto.SubTotal(Manual))
    subtos = funciones.FormatearDecimales(tmpPresupuesto.SubTotal(Sistema))
    dtos = funciones.FormatearDecimales((tmpPresupuesto.Descuento / 100) * subtos)
    dtoSM = funciones.FormatearDecimales((tmpPresupuesto.Descuento / 100) * subtoM)
    Me.subtotManual = funciones.FormatearDecimales(subtoM)
    Me.subtotSistema = funciones.FormatearDecimales(subtos)
    Me.dtoSistema = funciones.FormatearDecimales(dtos)
    Me.dtoManual = funciones.FormatearDecimales(dtos)
    Me.lblTotalSistema = funciones.FormatearDecimales(tmpPresupuesto.Total(Sistema))
    Me.lblTotalManual = funciones.FormatearDecimales(tmpPresupuesto.Total(Manual))
End Sub
Private Sub txtReferencia_LostFocus()
    grabado = False
    tmpPresupuesto.detalle = UCase(Me.txtReferencia)
End Sub
Private Sub ver_Click()
    If tmpPresupuesto.DetallePresupuesto.count > 0 Then
        Set tmpDetalle = tmpPresupuesto.DetallePresupuesto.item(grilla.RowIndex(grilla.row))

        Dim F As New frmDesarrollo
        Load F
        F.CargarPieza tmpDetalle.Pieza.id
        F.Show

    End If
End Sub

