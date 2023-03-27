VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmUsuarios 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Usuarios y grupos del sistema"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   255
   ClientWidth     =   13785
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   13785
   Begin VB.CommandButton Command6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Agregar Seleccionado"
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
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   83
      Top             =   3120
      Width           =   4095
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Quitar seleccionado"
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
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   82
      Top             =   5640
      Width           =   4095
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Marcar por default"
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
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   81
      Top             =   6000
      Width           =   4095
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6285
      Left            =   8565
      TabIndex        =   5
      Top             =   135
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   11086
      _Version        =   393216
      Tabs            =   7
      Tab             =   5
      TabHeight       =   520
      BackColor       =   15786449
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Configuracion"
      TabPicture(0)   =   "frmUsuarios.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "chConf(4)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "chConf(5)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "chConf(7)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "chConf(8)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "chConf(6)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "chConf(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "chConf(1)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "chConf(2)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "chConf(3)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "chConf(63)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "chConf(64)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "chConf(65)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "chConf(71)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "chConf(72)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).ControlCount=   16
      TabCaption(1)   =   "Ventas"
      TabPicture(1)   =   "frmUsuarios.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label3"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "chConf(10)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "chConf(9)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "chConf(17)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "chConf(16)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "chConf(15)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "chConf(13)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "chConf(12)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "chConf(11)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "chConf(18)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "chConf(14)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).ControlCount=   12
      TabCaption(2)   =   "Compras"
      TabPicture(2)   =   "frmUsuarios.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label8"
      Tab(2).Control(1)=   "Label9"
      Tab(2).Control(2)=   "chConf(41)"
      Tab(2).Control(3)=   "chConf(42)"
      Tab(2).Control(4)=   "chConf(56)"
      Tab(2).Control(5)=   "chConf(57)"
      Tab(2).Control(6)=   "chConf(59)"
      Tab(2).Control(7)=   "chConf(60)"
      Tab(2).Control(8)=   "chConf(61)"
      Tab(2).Control(9)=   "chConf(62)"
      Tab(2).Control(10)=   "chConf(68)"
      Tab(2).Control(11)=   "chConf(69)"
      Tab(2).Control(12)=   "chConf(70)"
      Tab(2).Control(13)=   "chConf(73)"
      Tab(2).Control(14)=   "chConf(74)"
      Tab(2).Control(15)=   "chConf(75)"
      Tab(2).Control(16)=   "chConf(76)"
      Tab(2).ControlCount=   17
      TabCaption(3)   =   "Planeamiento"
      TabPicture(3)   =   "frmUsuarios.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label6"
      Tab(3).Control(1)=   "chConf(22)"
      Tab(3).Control(2)=   "chConf(21)"
      Tab(3).Control(3)=   "chConf(20)"
      Tab(3).Control(4)=   "chConf(19)"
      Tab(3).Control(5)=   "chConf(30)"
      Tab(3).Control(6)=   "chConf(29)"
      Tab(3).Control(7)=   "chConf(24)"
      Tab(3).Control(8)=   "chConf(23)"
      Tab(3).Control(9)=   "chConf(26)"
      Tab(3).Control(10)=   "chConf(25)"
      Tab(3).Control(11)=   "chConf(28)"
      Tab(3).Control(12)=   "chConf(27)"
      Tab(3).Control(13)=   "chConf(32)"
      Tab(3).Control(14)=   "chConf(31)"
      Tab(3).Control(15)=   "chConf(33)"
      Tab(3).ControlCount=   16
      TabCaption(4)   =   "Desarrollo"
      TabPicture(4)   =   "frmUsuarios.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "chConf(39)"
      Tab(4).Control(1)=   "chConf(38)"
      Tab(4).Control(2)=   "chConf(36)"
      Tab(4).Control(3)=   "chConf(35)"
      Tab(4).Control(4)=   "chConf(34)"
      Tab(4).Control(5)=   "chConf(37)"
      Tab(4).Control(6)=   "Label5"
      Tab(4).ControlCount=   7
      TabCaption(5)   =   "Administración"
      TabPicture(5)   =   "frmUsuarios.frx":008C
      Tab(5).ControlEnabled=   -1  'True
      Tab(5).Control(0)=   "Label7"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "chConf(40)"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "chConf(43)"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).Control(3)=   "chConf(44)"
      Tab(5).Control(3).Enabled=   0   'False
      Tab(5).Control(4)=   "chConf(45)"
      Tab(5).Control(4).Enabled=   0   'False
      Tab(5).Control(5)=   "chConf(46)"
      Tab(5).Control(5).Enabled=   0   'False
      Tab(5).Control(6)=   "chConf(47)"
      Tab(5).Control(6).Enabled=   0   'False
      Tab(5).Control(7)=   "chConf(48)"
      Tab(5).Control(7).Enabled=   0   'False
      Tab(5).Control(8)=   "chConf(49)"
      Tab(5).Control(8).Enabled=   0   'False
      Tab(5).Control(9)=   "chConf(50)"
      Tab(5).Control(9).Enabled=   0   'False
      Tab(5).Control(10)=   "chConf(51)"
      Tab(5).Control(10).Enabled=   0   'False
      Tab(5).Control(11)=   "chConf(52)"
      Tab(5).Control(11).Enabled=   0   'False
      Tab(5).Control(12)=   "chConf(53)"
      Tab(5).Control(12).Enabled=   0   'False
      Tab(5).Control(13)=   "chConf(54)"
      Tab(5).Control(13).Enabled=   0   'False
      Tab(5).Control(14)=   "chConf(55)"
      Tab(5).Control(14).Enabled=   0   'False
      Tab(5).Control(15)=   "chConf(58)"
      Tab(5).Control(15).Enabled=   0   'False
      Tab(5).Control(16)=   "chConf(77)"
      Tab(5).Control(16).Enabled=   0   'False
      Tab(5).Control(17)=   "chConf(78)"
      Tab(5).Control(17).Enabled=   0   'False
      Tab(5).Control(18)=   "chConf(79)"
      Tab(5).Control(18).Enabled=   0   'False
      Tab(5).Control(19)=   "chConf(80)"
      Tab(5).Control(19).Enabled=   0   'False
      Tab(5).Control(20)=   "chConf(81)"
      Tab(5).Control(20).Enabled=   0   'False
      Tab(5).Control(21)=   "chConf(82)"
      Tab(5).Control(21).Enabled=   0   'False
      Tab(5).Control(22)=   "chConf(83)"
      Tab(5).Control(22).Enabled=   0   'False
      Tab(5).ControlCount=   23
      TabCaption(6)   =   "RRHH"
      TabPicture(6)   =   "frmUsuarios.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "chConf(67)"
      Tab(6).Control(1)=   "chConf(66)"
      Tab(6).Control(2)=   "Label10"
      Tab(6).ControlCount=   3
      Begin VB.CheckBox chConf 
         Caption         =   "FP Ver solo propias"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   83
         Left            =   2865
         TabIndex        =   107
         Tag             =   "521"
         Top             =   3405
         Width           =   1665
      End
      Begin VB.CheckBox chConf 
         Caption         =   "Plan de Cuentas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   82
         Left            =   2880
         TabIndex        =   106
         Tag             =   "519"
         Top             =   3135
         Width           =   1620
      End
      Begin VB.CheckBox chConf 
         Caption         =   "FP Consultas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   81
         Left            =   2880
         TabIndex        =   105
         Tag             =   "518"
         Top             =   2880
         Width           =   1620
      End
      Begin VB.CheckBox chConf 
         Caption         =   "FP Control"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   80
         Left            =   2880
         TabIndex        =   104
         Tag             =   "518"
         Top             =   2625
         Width           =   1620
      End
      Begin VB.CheckBox chConf 
         Caption         =   "OP Consultas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   79
         Left            =   2880
         TabIndex        =   103
         Tag             =   "517"
         Top             =   2370
         Width           =   1620
      End
      Begin VB.CheckBox chConf 
         Caption         =   "OP Control"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   78
         Left            =   2880
         TabIndex        =   102
         Tag             =   "516"
         Top             =   2115
         Width           =   1620
      End
      Begin VB.CheckBox chConf 
         Caption         =   "Caja y Bancos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   77
         Left            =   2880
         TabIndex        =   101
         Tag             =   "515"
         Top             =   1860
         Width           =   1620
      End
      Begin VB.CheckBox chConf 
         Caption         =   "Ver Precios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   76
         Left            =   -72510
         TabIndex        =   100
         Tag             =   "714"
         Top             =   5790
         Width           =   2580
      End
      Begin VB.CheckBox chConf 
         Caption         =   "Administrar Precios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   75
         Left            =   -74280
         TabIndex        =   99
         Tag             =   "713"
         Top             =   5790
         Width           =   2580
      End
      Begin VB.CheckBox chConf 
         Caption         =   "OC Consultas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   74
         Left            =   -74280
         TabIndex        =   98
         Tag             =   "712"
         Top             =   5565
         Width           =   2580
      End
      Begin VB.CheckBox chConf 
         Caption         =   "OC Control"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   73
         Left            =   -74280
         TabIndex        =   97
         Tag             =   "711"
         Top             =   5340
         Width           =   2580
      End
      Begin VB.CheckBox chConf 
         Caption         =   "Ver Eventos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   72
         Left            =   -74280
         TabIndex        =   96
         Tag             =   "116"
         Top             =   5670
         Width           =   2565
      End
      Begin VB.CheckBox chConf 
         Caption         =   "Ver Updates"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   71
         Left            =   -74280
         TabIndex        =   95
         Tag             =   "115"
         Top             =   5400
         Width           =   2565
      End
      Begin VB.CheckBox chConf 
         Caption         =   "PO Crear"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   70
         Left            =   -74280
         TabIndex        =   94
         Tag             =   "709"
         Top             =   4860
         Width           =   2580
      End
      Begin VB.CheckBox chConf 
         Caption         =   "PO Consultar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   69
         Left            =   -74280
         TabIndex        =   93
         Tag             =   "710"
         Top             =   5100
         Width           =   2580
      End
      Begin VB.CheckBox chConf 
         Caption         =   "Reque Anular"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   68
         Left            =   -74280
         TabIndex        =   92
         Tag             =   "708"
         Top             =   4620
         Width           =   2580
      End
      Begin VB.CheckBox chConf 
         Caption         =   "Informe accidente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   67
         Left            =   -74160
         TabIndex        =   90
         Tag             =   "801"
         Top             =   1785
         Width           =   2400
      End
      Begin VB.CheckBox chConf 
         Caption         =   "Siniestros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   66
         Left            =   -74160
         TabIndex        =   89
         Tag             =   "800"
         Top             =   1530
         Width           =   2520
      End
      Begin VB.CheckBox chConf 
         Caption         =   "Archivos de Compras"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   65
         Left            =   -74280
         TabIndex        =   88
         Tag             =   "114"
         Top             =   5130
         Width           =   2565
      End
      Begin VB.CheckBox chConf 
         Caption         =   "Ver Archivos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   64
         Left            =   -74280
         TabIndex        =   87
         Tag             =   "112"
         Top             =   4635
         Width           =   1845
      End
      Begin VB.CheckBox chConf 
         Caption         =   "Adquirir Archivos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   63
         Left            =   -74280
         TabIndex        =   86
         Tag             =   "113"
         Top             =   4875
         Width           =   2565
      End
      Begin VB.CheckBox chConf 
         Caption         =   "Reque Consultas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   62
         Left            =   -74280
         TabIndex        =   80
         Tag             =   "706"
         Top             =   4140
         Width           =   2700
      End
      Begin VB.CheckBox chConf 
         Caption         =   "Reque Procesar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   61
         Left            =   -74280
         TabIndex        =   79
         Tag             =   "704"
         Top             =   4380
         Width           =   2580
      End
      Begin VB.CheckBox chConf 
         Caption         =   "Reque Aprobar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   60
         Left            =   -74280
         TabIndex        =   78
         Tag             =   "707"
         Top             =   3660
         Width           =   2700
      End
      Begin VB.CheckBox chConf 
         Caption         =   "Reque Control"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   59
         Left            =   -74280
         TabIndex        =   77
         Tag             =   "705"
         Top             =   3900
         Width           =   2580
      End
      Begin VB.CheckBox chConf 
         Caption         =   "Centro de Cambio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   58
         Left            =   720
         TabIndex        =   76
         Tag             =   "510"
         Top             =   5100
         Width           =   2805
      End
      Begin VB.CheckBox chConf 
         Caption         =   "Info en pantalla"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   57
         Left            =   -74280
         TabIndex        =   72
         Tag             =   "701"
         Top             =   3420
         Width           =   2580
      End
      Begin VB.CheckBox chConf 
         Caption         =   "Menú completo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   56
         Left            =   -74280
         TabIndex        =   71
         Tag             =   "700"
         Top             =   3180
         Width           =   2580
      End
      Begin VB.CheckBox chConf 
         Caption         =   "Informes Varios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   55
         Left            =   720
         TabIndex        =   70
         Tag             =   "514"
         Top             =   5700
         Width           =   2805
      End
      Begin VB.CheckBox chConf 
         Caption         =   "Informes Cash Flow"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   54
         Left            =   720
         TabIndex        =   69
         Tag             =   "513"
         Top             =   5460
         Width           =   2805
      End
      Begin VB.CheckBox chConf 
         Caption         =   "IIBB Control"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   53
         Left            =   720
         TabIndex        =   68
         Tag             =   "509"
         Top             =   4860
         Width           =   2805
      End
      Begin VB.CheckBox chConf 
         Caption         =   "IIBB Consultas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   52
         Left            =   720
         TabIndex        =   67
         Tag             =   "508"
         Top             =   4620
         Width           =   2805
      End
      Begin VB.CheckBox chConf 
         Caption         =   "Cuentas Corrientes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   51
         Left            =   720
         TabIndex        =   66
         Tag             =   "511"
         Top             =   4380
         Width           =   2805
      End
      Begin VB.CheckBox chConf 
         Caption         =   "Subdiarios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   50
         Left            =   720
         TabIndex        =   65
         Tag             =   "507"
         Top             =   4140
         Width           =   2805
      End
      Begin VB.CheckBox chConf 
         Caption         =   "Cobros Aprobaciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   49
         Left            =   720
         TabIndex        =   64
         Tag             =   "506"
         Top             =   3900
         Width           =   2805
      End
      Begin VB.CheckBox chConf 
         Caption         =   "Cobros Consultas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   48
         Left            =   720
         TabIndex        =   63
         Tag             =   "505"
         Top             =   3660
         Width           =   2805
      End
      Begin VB.CheckBox chConf 
         Caption         =   "Facturas Aprobaciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   47
         Left            =   720
         TabIndex        =   62
         Tag             =   "503"
         Top             =   3060
         Width           =   2805
      End
      Begin VB.CheckBox chConf 
         Caption         =   "Facturas Consultas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   46
         Left            =   720
         TabIndex        =   61
         Tag             =   "502"
         Top             =   2820
         Width           =   2805
      End
      Begin VB.CheckBox chConf 
         Caption         =   "Facturas Control"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   45
         Left            =   720
         TabIndex        =   60
         Tag             =   "501"
         Top             =   2580
         Width           =   2805
      End
      Begin VB.CheckBox chConf 
         Caption         =   "Cobros Control"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   44
         Left            =   720
         TabIndex        =   59
         Tag             =   "504"
         Top             =   3420
         Width           =   2805
      End
      Begin VB.CheckBox chConf 
         Caption         =   "Info Pantalla"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   43
         Left            =   720
         TabIndex        =   58
         Tag             =   "512"
         Top             =   2100
         Width           =   2805
      End
      Begin VB.CheckBox chConf 
         Caption         =   "Proveedores Consultas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   42
         Left            =   -74280
         TabIndex        =   56
         Tag             =   "703"
         Top             =   2100
         Width           =   2580
      End
      Begin VB.CheckBox chConf 
         Caption         =   "Proveedores Control"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   41
         Left            =   -74280
         TabIndex        =   55
         Tag             =   "702"
         Top             =   1860
         Width           =   2700
      End
      Begin VB.CheckBox chConf 
         Caption         =   "Menú Completo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   40
         Left            =   720
         TabIndex        =   53
         Tag             =   "500"
         Top             =   1860
         Width           =   2805
      End
      Begin VB.CheckBox chConf 
         Caption         =   "Manejo Stock"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   39
         Left            =   -74280
         TabIndex        =   52
         Tag             =   "405"
         Top             =   3420
         Width           =   2895
      End
      Begin VB.CheckBox chConf 
         Caption         =   "Consultar Tiempos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   38
         Left            =   -74280
         TabIndex        =   51
         Tag             =   "404"
         Top             =   3180
         Width           =   2895
      End
      Begin VB.CheckBox chConf 
         Caption         =   "Control"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   36
         Left            =   -74280
         TabIndex        =   49
         Tag             =   "402"
         Top             =   2700
         Width           =   2655
      End
      Begin VB.CheckBox chConf 
         Caption         =   "Info en Pantalla"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   35
         Left            =   -74280
         TabIndex        =   48
         Tag             =   "401"
         Top             =   2100
         Width           =   2895
      End
      Begin VB.CheckBox chConf 
         Caption         =   "Menú Completo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   34
         Left            =   -74280
         TabIndex        =   47
         Tag             =   "400"
         Top             =   1860
         Width           =   3015
      End
      Begin VB.CheckBox chConf 
         Caption         =   "Consultas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   37
         Left            =   -74280
         TabIndex        =   46
         Tag             =   "403"
         Top             =   2940
         Width           =   2895
      End
      Begin VB.CheckBox chConf 
         Caption         =   "Remitos Aprobar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   33
         Left            =   -74280
         TabIndex        =   45
         Tag             =   "312"
         Top             =   5580
         Width           =   2610
      End
      Begin VB.CheckBox chConf 
         Caption         =   "Remitos Control"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   31
         Left            =   -74280
         TabIndex        =   44
         Tag             =   "310"
         Top             =   5100
         Width           =   3090
      End
      Begin VB.CheckBox chConf 
         Caption         =   "Remitos Consultas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   32
         Left            =   -74280
         TabIndex        =   43
         Tag             =   "311"
         Top             =   5340
         Width           =   2370
      End
      Begin VB.CheckBox chConf 
         Caption         =   "OE Aprobar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   27
         Left            =   -74280
         TabIndex        =   42
         Tag             =   "304"
         Top             =   3900
         Width           =   2610
      End
      Begin VB.CheckBox chConf 
         Caption         =   "OE Modificar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   28
         Left            =   -74280
         TabIndex        =   41
         Tag             =   "314"
         Top             =   4140
         Width           =   2370
      End
      Begin VB.CheckBox chConf 
         Caption         =   "OE Control"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   25
         Left            =   -74280
         TabIndex        =   40
         Tag             =   "302"
         Top             =   3420
         Width           =   3090
      End
      Begin VB.CheckBox chConf 
         Caption         =   "OE Consultas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   26
         Left            =   -74280
         TabIndex        =   39
         Tag             =   "303"
         Top             =   3660
         Width           =   2370
      End
      Begin VB.CheckBox chConf 
         Caption         =   "OT Aprobar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   23
         Left            =   -74280
         TabIndex        =   37
         Tag             =   "307"
         Top             =   2940
         Width           =   2610
      End
      Begin VB.CheckBox chConf 
         Caption         =   "OT Modificar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   24
         Left            =   -74280
         TabIndex        =   36
         Tag             =   "313"
         Top             =   3180
         Width           =   2370
      End
      Begin VB.CheckBox chConf 
         Caption         =   "Seguimiento Global"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   29
         Left            =   -74280
         TabIndex        =   35
         Tag             =   "308"
         Top             =   4500
         Width           =   3330
      End
      Begin VB.CheckBox chConf 
         Caption         =   "Seguimiento Rutas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   30
         Left            =   -74280
         TabIndex        =   34
         Tag             =   "309"
         Top             =   4740
         Width           =   2250
      End
      Begin VB.CheckBox chConf 
         Caption         =   "Menú Completo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   19
         Left            =   -74280
         TabIndex        =   33
         Tag             =   "300"
         Top             =   1860
         Width           =   2730
      End
      Begin VB.CheckBox chConf 
         Caption         =   "Info en Pantalla"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   20
         Left            =   -74280
         TabIndex        =   32
         Tag             =   "301"
         Top             =   2100
         Width           =   2610
      End
      Begin VB.CheckBox chConf 
         Caption         =   "OT Control"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   21
         Left            =   -74280
         TabIndex        =   31
         Tag             =   "305"
         Top             =   2460
         Width           =   3090
      End
      Begin VB.CheckBox chConf 
         Caption         =   "OT Consultas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   22
         Left            =   -74280
         TabIndex        =   30
         Tag             =   "306"
         Top             =   2700
         Width           =   2370
      End
      Begin VB.CheckBox chConf 
         Caption         =   "Cotizaciones Modificar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   14
         Left            =   -74280
         TabIndex        =   28
         Tag             =   "209"
         Top             =   4260
         Width           =   2400
      End
      Begin VB.CheckBox chConf 
         Caption         =   "Menú completo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   18
         Left            =   -74280
         TabIndex        =   27
         Tag             =   "200"
         Top             =   3060
         Width           =   2400
      End
      Begin VB.CheckBox chConf 
         Caption         =   "Cotizaciones Control"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   11
         Left            =   -74280
         TabIndex        =   24
         Tag             =   "201"
         Top             =   3540
         Width           =   2400
      End
      Begin VB.CheckBox chConf 
         Caption         =   "Cotizaciones consultas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   12
         Left            =   -74280
         TabIndex        =   23
         Tag             =   "202"
         Top             =   3780
         Width           =   2640
      End
      Begin VB.CheckBox chConf 
         Caption         =   "Cotizaciones Aprobar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   13
         Left            =   -74280
         TabIndex        =   22
         Tag             =   "203"
         Top             =   4020
         Width           =   2400
      End
      Begin VB.CheckBox chConf 
         Caption         =   "Pedidos Control"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   15
         Left            =   -74280
         TabIndex        =   21
         Tag             =   "204"
         Top             =   4500
         Width           =   2400
      End
      Begin VB.CheckBox chConf 
         Caption         =   "Pedidos Consultas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   16
         Left            =   -74280
         TabIndex        =   20
         Tag             =   "205"
         Top             =   4740
         Width           =   3480
      End
      Begin VB.CheckBox chConf 
         Caption         =   "Info en Pantalla"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   17
         Left            =   -74280
         TabIndex        =   19
         Tag             =   "208"
         Top             =   3300
         Width           =   2400
      End
      Begin VB.CheckBox chConf 
         Caption         =   "Clientes Control"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   9
         Left            =   -74280
         TabIndex        =   18
         Tag             =   "206"
         Top             =   1860
         Width           =   2520
      End
      Begin VB.CheckBox chConf 
         Caption         =   "Clientes Modificar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   10
         Left            =   -74280
         TabIndex        =   17
         Tag             =   "207"
         Top             =   2100
         Width           =   2400
      End
      Begin VB.CheckBox chConf 
         Caption         =   "Configurar Materiales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   3
         Left            =   -74280
         TabIndex        =   16
         Tag             =   "109"
         Top             =   2355
         Width           =   2325
      End
      Begin VB.CheckBox chConf 
         Caption         =   "Configurar Mano de Obra"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   2
         Left            =   -74280
         TabIndex        =   15
         Tag             =   "108"
         Top             =   2115
         Width           =   3045
      End
      Begin VB.CheckBox chConf 
         Caption         =   "Menú Completo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   -74280
         TabIndex        =   14
         Tag             =   "100"
         Top             =   1875
         Width           =   2565
      End
      Begin VB.CheckBox chConf 
         Caption         =   "Opciones Generales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   -74280
         TabIndex        =   13
         Tag             =   "111"
         Top             =   1635
         Width           =   2685
      End
      Begin VB.CheckBox chConf 
         Caption         =   "Ver Precios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   6
         Left            =   -74280
         TabIndex        =   12
         Tag             =   "110"
         Top             =   3915
         Width           =   1845
      End
      Begin VB.CheckBox chConf 
         Caption         =   "Agenda Modif"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   8
         Left            =   -74280
         TabIndex        =   11
         Tag             =   "104"
         Top             =   4395
         Width           =   1845
      End
      Begin VB.CheckBox chConf 
         Caption         =   "Agenda Ver"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   7
         Left            =   -74280
         TabIndex        =   10
         Tag             =   "103"
         Top             =   4155
         Width           =   1845
      End
      Begin VB.CheckBox chConf 
         Caption         =   "Ver Tablero"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   5
         Left            =   -74280
         TabIndex        =   9
         Tag             =   "102"
         Top             =   3675
         Width           =   1845
      End
      Begin VB.CheckBox chConf 
         Caption         =   "Usuario Activo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   4
         Left            =   -74280
         TabIndex        =   8
         Tag             =   "101"
         Top             =   3435
         Width           =   1845
      End
      Begin VB.Label Label10 
         Caption         =   "Seguridad e Higiene"
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
         Left            =   -74640
         TabIndex        =   91
         Top             =   1185
         Width           =   1815
      End
      Begin VB.Label Label9 
         Caption         =   "Opciones para Compras"
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
         Left            =   -74760
         TabIndex        =   73
         Top             =   2820
         Width           =   2535
      End
      Begin VB.Label Label8 
         Caption         =   "Proveedores"
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
         Left            =   -74760
         TabIndex        =   57
         Top             =   1500
         Width           =   1815
      End
      Begin VB.Label Label7 
         Caption         =   "Opciones para Administracion"
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
         Left            =   240
         TabIndex        =   54
         Top             =   1500
         Width           =   2655
      End
      Begin VB.Label Label5 
         Caption         =   "Opciones para Desarrollo"
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
         Left            =   -74760
         TabIndex        =   50
         Top             =   1500
         Width           =   2655
      End
      Begin VB.Label Label6 
         Caption         =   "Opciones para Planeamiento"
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
         Left            =   -74760
         TabIndex        =   38
         Top             =   1500
         Width           =   2775
      End
      Begin VB.Label Label3 
         Caption         =   "Clientres"
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
         Left            =   -74760
         TabIndex        =   26
         Top             =   1500
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Opciones para Ventas"
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
         Left            =   -74760
         TabIndex        =   25
         Top             =   2820
         Width           =   2535
      End
      Begin VB.Label Label4 
         Caption         =   "Configuración"
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
         Left            =   -74640
         TabIndex        =   7
         Top             =   3075
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Panel de control"
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
         Left            =   -74760
         TabIndex        =   6
         Top             =   1275
         Width           =   1815
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ Opciones ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   6480
      Width           =   13695
      Begin VB.CommandButton Command4 
         BackColor       =   &H00E0E0E0&
         Cancel          =   -1  'True
         Caption         =   "Salir"
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
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   240
         Visible         =   0   'False
         Width           =   1935
      End
      Begin MSComctlLib.ProgressBar pb1 
         Height          =   255
         Left            =   10680
         TabIndex        =   74
         Top             =   240
         Visible         =   0   'False
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Actualizar Grupos"
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
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Actualizar Permisos"
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
         Left            =   8640
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Reiniciar Clave"
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
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1935
      End
   End
   Begin MSComctlLib.ListView lstUsuarios 
      Height          =   6495
      Left            =   120
      TabIndex        =   29
      Top             =   120
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   11456
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Usuario"
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nombre"
         Object.Width           =   2011
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Apellido"
         Object.Width           =   2364
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
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
      Left            =   2760
      TabIndex        =   3
      Top             =   6240
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSComctlLib.ListView lstGruposDisponibles 
      Height          =   2895
      Left            =   4320
      TabIndex        =   84
      Top             =   120
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   5106
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Grupo"
         Object.Width           =   6527
      EndProperty
   End
   Begin MSComctlLib.ListView LstGruposUsuarios 
      Height          =   1935
      Left            =   4320
      TabIndex        =   85
      Top             =   3600
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   3413
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Grupo"
         Object.Width           =   6526
      EndProperty
   End
End
Attribute VB_Name = "frmUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim permi As classPermisos
Dim claseSP As New classSignoplast



Private Sub AgConsultar_Click()
'If Not Me.AgConsultar.value Then Me.AgModificar.value = False
End Sub
Private Function OnOff(Valor)
    If Valor = 0 Then OnOff = "Off" Else OnOff = "On"

End Function

Private Sub Command1_Click()
    If MsgBox("¿Seguro de cambiar password?", vbYesNo, "Confirmación") = vbYes Then
        Dim md As New classMD5
        usu = Me.lstUsuarios.selectedItem.Tag
        pass = md.DigestStrToHexStr(Me.lstUsuarios.selectedItem)
        If claseSP.cambiarPass(usu, pass) Then
            MsgBox "Reinicio exitoso!" & Chr(10) & " Debe reiniciar el sistema para que se efectuen los cambios", vbInformation, "Confirmación"
        Else
            MsgBox "Se produjo un error. No se actualizo password!", vbCritical, "Error"
        End If
        Set md = Nothing
    End If
End Sub

Private Sub Command2_Click()
    MsgBox "Recuerde que para que los cambios surjan efecto, hay que reiniciar el sistema!", vbExclamation, "Información"
    Unload Me
End Sub

Private Sub Command3_Click()
    Dim ids As Long
    If MsgBox("¿Seguro de actualizar?", vbYesNo, "Confirmación") = vbYes Then
        pb1.Visible = True
        it = Me.chConf.count
        pb1.min = 0
        pb1.max = it - 1
        For x = 0 To it - 1
            nro = Me.chConf(x).Tag
            pb1.value = 1
            ids = Me.lstUsuarios.selectedItem.Tag
            Valor = Me.chConf(x).value

            claseSP.verSeleccionado nro, ids, True, Valor
        Next x
        pb1.Visible = False
    End If
End Sub



Private Sub Command4_Click()
    Unload Me
End Sub

Private Sub Command5_Click()

    If MsgBox("¿Desea actualizar los grupos?", vbYesNo, "Confirmación") = vbYes Then
        idUsu = CLng(Me.lstUsuarios.selectedItem.Tag)
        If claseSP.actualizarGrupos(idUsu, Me.LstGruposUsuarios) Then
            MsgBox "Actualizacion exitosa!", vbInformation, "Información"
        Else
            MsgBox "Se produjo un error, no se guardaron lso cambios!", vbCritical, "Error"
        End If
    End If

End Sub

Private Sub Command6_Click()
'veo si está en la lista
    Dim esta As Boolean
    For x = 1 To Me.lstGruposDisponibles.ListItems.count
        If Me.lstGruposDisponibles.ListItems(x).Checked Then
            Id = CLng(Me.lstGruposDisponibles.ListItems(x).Tag)
            grupos = Me.lstGruposDisponibles.ListItems(x)
            esta = False
            For i = 1 To Me.LstGruposUsuarios.ListItems.count
                If Id = Me.LstGruposUsuarios.ListItems(i).Tag Then esta = True
            Next i
            If Not esta Then
                Set P = Me.LstGruposUsuarios.ListItems.Add(, , grupos)
                P.Tag = Id
            End If
        End If
    Next x

End Sub

Private Sub Command7_Click()
    If MsgBox("¿Está seguro de eliminar los items seleecionados?", vbYesNo, "Confirmacion") = vbYes Then
        For i = Me.LstGruposUsuarios.ListItems.count To 1 Step -1
            If Me.LstGruposUsuarios.ListItems(i).Checked = True Then
                Me.LstGruposUsuarios.ListItems.remove (i)
                grabado = False
            End If
        Next i
    End If

End Sub

Private Sub Command8_Click()
    Command5_Click
    If MsgBox("¿Desea marcar este grupo como default?", vbYesNo, "Confirmación") = vbYes Then
        idUsu = CLng(Me.lstUsuarios.selectedItem.Tag)
        gru = CLng(Me.LstGruposUsuarios.selectedItem.Tag)
        If claseSP.ejecutarComando("update sp_permisos.Config set GrupoDefault=" & gru & " where idUsuario=" & idUsu) Then
            MsgBox "Cambio exitoso!", vbInformation, "Información"
            Dim Id As Long
            Id = CLng(Me.lstUsuarios.selectedItem.Tag)
            llenarLSTGruposUsuario Id
        Else
            MsgBox "Se produjo un error, se abortan los cambios!", vbCritical, "Error"
        End If
    End If
End Sub

Private Sub Form_Activate()
    Dim Id As Long

    llenarLST
    llenarLSTGruposDisponibles
    lstUsuarios_ItemClick Me.LstGruposUsuarios.selectedItem
    If Me.lstUsuarios.ListItems.count > 0 Then
        Id = CLng(Me.lstUsuarios.selectedItem.Tag)
        llenarLSTGruposUsuario Id
    End If
    Me.SSTab1.SetFocus
End Sub

Private Function llenarLST()
    Dim rs As Recordset
    Set rs = conectar.RSFactory("select s.id,s.usuario,p.apellido,p.nombre from usuarios s left join personal p on p.id=s.idEmpleado")
    'Set rs = conectar.RSFactory("select s.id,s.usuario,p.apellido,p.nombre from usuarios s inner join personal p on p.id=s.idEmpleado")
    Me.lstUsuarios.ListItems.Clear
    Dim x As ListItem
    While Not rs.EOF
        Set x = Me.lstUsuarios.ListItems.Add(, , rs!usuario)
        If Not IsNull(rs!nombre) Then x.SubItems(1) = rs!nombre
        If Not IsNull(rs!Apellido) Then x.SubItems(2) = rs!Apellido
        x.Tag = rs!Id

        rs.MoveNext
    Wend
    Set rs = Nothing
End Function


Private Function llenarLSTGruposDisponibles()
    Dim rs As Recordset
    Set rs = conectar.RSFactory("select id, grupo from usuariosGrupos")
    Dim x As ListItem
    Me.lstGruposDisponibles.ListItems.Clear
    While Not rs.EOF
        Set x = Me.lstGruposDisponibles.ListItems.Add(, , rs!Grupo)
        x.Tag = rs!Id

        rs.MoveNext
    Wend
    Set rs = Nothing
End Function

Private Function llenarLSTGruposUsuario(idUsuario As Long)
    Dim rs As Recordset

    Set rs = conectar.RSFactory("select GrupoDefault from sp_permisos.Config where idUsuario=" & idUsuario)
    If Not rs.EOF And Not rs.BOF Then
        defa = rs!GrupoDefault
    Else
        defa = -1
    End If

    Set rs = conectar.RSFactory("select g.id, grupo from usuariosGruposDetalle d inner join usuariosGrupos g on d.idGrupo=g.id where d.idUsuario=" & idUsuario)
    Dim x As ListItem

    Me.LstGruposUsuarios.ListItems.Clear
    While Not rs.EOF
        Set x = Me.LstGruposUsuarios.ListItems.Add(, , rs!Grupo)
        x.Tag = rs!Id

        If defa = rs!Id Then
            x.ForeColor = vbRed
        End If
        rs.MoveNext
    Wend
    Set rs = Nothing
End Function





Private Sub Form_Load()
    FormHelper.Customize Me

End Sub

Private Sub lstUsuarios_ItemClick(ByVal item As MSComctlLib.ListItem)
    Dim Id As Long
    If Me.lstUsuarios.ListItems.count > 0 Then
        Id = CLng(Me.lstUsuarios.selectedItem.Tag)
        llenarLSTGruposUsuario Id
        Set permi = Nothing
        verPermisos Id
    End If
End Sub


Public Sub verPermisos(Id As Long)

    it = Me.chConf.count
    For x = 0 To it - 1
        nro = Me.chConf(x).Tag
        Me.chConf(x).value = claseSP.verSeleccionado(nro, Id)
    Next x

End Sub

