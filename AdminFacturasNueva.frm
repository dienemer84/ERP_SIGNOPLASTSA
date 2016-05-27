VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~3.OCX"
Begin VB.Form frmAdminFacturasNueva 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nueva Factura"
   ClientHeight    =   9795
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10275
   Icon            =   "AdminFacturasNueva.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9795
   ScaleWidth      =   10275
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ Datos factura ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3660
      Left            =   60
      TabIndex        =   0
      Top             =   45
      Width           =   10200
      Begin VB.TextBox txtObservar 
         Height          =   285
         Left            =   1560
         TabIndex        =   49
         Top             =   3240
         Width           =   8535
      End
      Begin VB.TextBox txtPercepciones 
         Height          =   285
         Left            =   6600
         TabIndex        =   43
         Text            =   "0"
         Top             =   2040
         Width           =   1695
      End
      Begin VB.ComboBox cboClientes 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   360
         Width           =   4335
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ver"
         Height          =   255
         Left            =   5640
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox txtNroFactura 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   8760
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtOC 
         Height          =   285
         Left            =   1560
         TabIndex        =   3
         Top             =   2520
         Width           =   8535
      End
      Begin VB.TextBox txtFP 
         Height          =   285
         Left            =   1560
         TabIndex        =   2
         ToolTipText     =   "Cantidad de días FF"
         Top             =   2880
         Width           =   2175
      End
      Begin VB.ComboBox cboMoneda 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8760
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1200
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   8760
         TabIndex        =   4
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   63504385
         CurrentDate     =   39176
      End
      Begin XtremeSuiteControls.ComboBox cboPadron 
         Height          =   315
         Left            =   8370
         TabIndex        =   56
         Top             =   1665
         Width           =   1725
         _Version        =   786432
         _ExtentX        =   3043
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Style           =   2
         Appearance      =   6
         UseVisualStyle  =   -1  'True
         Text            =   "cboPadron"
      End
      Begin VB.Label Label14 
         BackColor       =   &H00C0C0C0&
         Caption         =   "días FF"
         Height          =   255
         Left            =   3840
         TabIndex        =   55
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Label lblVencido 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "PADRON VENCIDO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   8400
         TabIndex        =   53
         Top             =   2040
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label lblND 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "N/D"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6600
         TabIndex        =   52
         Top             =   1095
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Condicion "
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
         TabIndex        =   50
         Top             =   3240
         Width           =   1335
      End
      Begin VB.Label lblNC 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "N/C"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6600
         TabIndex        =   48
         Top             =   1080
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lblPadron 
         BackColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   9345
         TabIndex        =   47
         Top             =   1680
         Width           =   615
      End
      Begin VB.Line Line1 
         X1              =   6480
         X2              =   6480
         Y1              =   240
         Y2              =   2280
      End
      Begin VB.Label Label17 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Percep IIBB  según "
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
         Left            =   6600
         TabIndex        =   42
         Top             =   1680
         Width           =   3045
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ciudad"
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
         TabIndex        =   41
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label lblCiudad 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   1200
         TabIndex        =   40
         Top             =   1800
         Width           =   5055
      End
      Begin VB.Label lblCp 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   1200
         TabIndex        =   39
         Top             =   2040
         Width           =   5055
      End
      Begin VB.Label lblLocalidad 
         BackColor       =   &H00E0E0E0&
         Caption         =   " "
         Height          =   255
         Left            =   1200
         TabIndex        =   38
         Top             =   1560
         Width           =   5055
      End
      Begin VB.Label lblDireccion 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   1200
         TabIndex        =   37
         Top             =   1320
         Width           =   5055
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "C.P"
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
         TabIndex        =   36
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Localidad"
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
         TabIndex        =   35
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Dirección"
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
         TabIndex        =   34
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label lblIva 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   1200
         TabIndex        =   32
         Top             =   1080
         Width           =   5055
      End
      Begin VB.Label label 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "C.U.I.T"
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
         TabIndex        =   31
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lblCuit 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   1200
         TabIndex        =   30
         Top             =   840
         Width           =   5055
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
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
         Left            =   120
         TabIndex        =   29
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblTipoFactura2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   30
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   6600
         TabIndex        =   26
         Top             =   240
         Width           =   735
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Número "
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
         Left            =   7680
         TabIndex        =   10
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
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
         Left            =   7680
         TabIndex        =   8
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Moneda "
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
         Left            =   7680
         TabIndex        =   6
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "I.V.A."
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
         TabIndex        =   33
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Referencia "
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
         TabIndex        =   9
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Vencimiento "
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
         TabIndex        =   7
         Top             =   2880
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ Detalle factura ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5850
      Left            =   60
      TabIndex        =   11
      Top             =   3885
      Width           =   10200
      Begin VB.CommandButton btNDIB 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cambiar IIBB"
         Height          =   375
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   3720
         Width           =   1095
      End
      Begin VB.CommandButton btnND 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cambiar IVA"
         Height          =   375
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   3720
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Nueva"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   4440
         Width           =   975
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Salir"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   5400
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Guardar"
         Default         =   -1  'True
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   4920
         Width           =   975
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "[ Totales ]"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   6480
         TabIndex        =   16
         Top             =   3720
         Width           =   3615
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Caption         =   "Subtotal "
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
            Left            =   1440
            TabIndex        =   24
            Top             =   240
            Width           =   855
         End
         Begin VB.Label lblSubTotal 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FF8080&
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
            Left            =   2280
            TabIndex        =   23
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label lbliva2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Caption         =   "I.V.A."
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
            TabIndex        =   22
            Top             =   945
            Width           =   2055
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Caption         =   "Total "
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
            Left            =   1440
            TabIndex        =   21
            Top             =   1680
            Width           =   855
         End
         Begin VB.Label lblTotal 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FF8080&
            BorderStyle     =   1  'Fixed Single
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
            Left            =   2280
            TabIndex        =   20
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Caption         =   "Percepciones "
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
            Left            =   840
            TabIndex        =   19
            Top             =   585
            Width           =   1455
         End
         Begin VB.Label lblVerIva 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FF8080&
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
            Left            =   2280
            TabIndex        =   18
            Top             =   945
            Width           =   1215
         End
         Begin VB.Label lblTotalPercep 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FF8080&
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
            Left            =   2280
            TabIndex        =   17
            Top             =   585
            Width           =   1215
         End
      End
      Begin VB.CommandButton bElegirRto 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Remito"
         Height          =   375
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   5400
         Width           =   975
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Concepto"
         Height          =   375
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   4920
         Width           =   975
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Quitar"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   3720
         Width           =   1095
      End
      Begin VB.CommandButton Command10 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Recalcular"
         Height          =   375
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   4440
         Width           =   975
      End
      Begin MSComctlLib.ListView lstFactura 
         Height          =   3375
         Left            =   120
         TabIndex        =   25
         Top             =   255
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   5953
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Item"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Cant"
            Object.Width           =   1323
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Detalle"
            Object.Width           =   6967
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Unitario"
            Object.Width           =   1623
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Importe"
            Object.Width           =   1623
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Origen"
            Object.Width           =   1623
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Remito"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "idEntrega"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "% Dto"
            Object.Width           =   1411
         EndProperty
      End
   End
End
Attribute VB_Name = "frmAdminFacturasNueva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim claseP As New classPlaneamiento
Dim claseC As New classConfigurar
Dim vIdFactura As Long
Dim vCuit As Double
'origen facturado
'1- desde remito.
'2- por conceptos.
Dim Discrimina As Integer
Dim Alicuota As Double
Dim tipoNC As Integer
Dim idsentrega As Collection
Dim errores As Boolean
Dim idCliente As Long
Dim claseS As New classStock
Dim claseA As New classAdministracion
Dim rs As recordset
Dim vTipoFactura As Integer
Dim vorigen As Integer
Dim grabado As Boolean

Public Property Let NotaDeCredito(ntipoNC As Integer)
    tipoNC = ntipoNC
End Property

Public Property Let idFactura(nIdFactura As Long)
    vIdFactura = nIdFactura
End Property
Public Property Get idFactura() As Long
    idFactura = vIdFactura
End Property


Public Property Let OrigenFactura(nOrigen As Integer)
    vorigen = nOrigen
End Property

Public Property Get OrigenFactura() As Integer
    OrigenFactura = vorigen
End Property

Private Sub bElegirRto_Click()
    If idCliente = Empty Or idCliente = 0 Then
        MsgBox "No hay un cliente definido!", vbCritical
        Exit Sub
    Else
        If Not errores Then
            Dim idEntrega As Long
            frmPlaneamientoRemitosListaProceso.idCliMostrar = idCliente
            frmPlaneamientoRemitosListaProceso.Mostrar = 2    ' idcliente 'CLng(Me.cboClientes.ItemData(Me.cboClientes.ListIndex))
            frmPlaneamientoRemitosListaProceso.Show 1
            If funciones.queRemitoElegido <> -1 Then
                frmPlaneamientoRemitosDetalle.usable = True
                frmPlaneamientoRemitosDetalle.rtoNro = funciones.queRemitoElegido
                frmPlaneamientoRemitosDetalle.Show 1
                'traigo la coleccion de elementos a facturar
                'desde el remito elegido
                Set idsentrega = funciones.idEntrega
                agregarAFactura idsentrega
                Set idsentrega = Nothing
                'Me.txtDto = verDescuento
                Me.txtOC = verOC
            End If
        End If
    End If

End Sub
Private Sub totalizarFactura(Optional ByRef Total As Double)
'creo una factura temporaria con detalles para poder utilizar el calculador de totales de la clase

Dim fac As New Factura
Set fac.Cliente = DAOCliente.BuscarPorID(Me.cboClientes.ItemData(Me.cboClientes.ListIndex))
Set fac.Moneda = DAOMoneda.GetById(Me.cboMoneda.ItemData(Me.cboMoneda.ListIndex))
fac.AlicuotaPercepcionesIIBB = 1 + (CDbl(txtPercepciones.text) / 100)
fac.AlicuotaAplicada = fac.Cliente.tipoIva.Alicuota
fac.Detalles = New Collection
fac.EstaDiscriminada = Discrimina 'fac.Cliente.tipoIva.TipoFactura.Discrimina

Dim deta As FacturaDetalle
Dim li As ListItem

For Each li In Me.lstFactura.ListItems
    Set deta = New FacturaDetalle
    Set deta.Factura = fac
    
    deta.cantidad = Val(li.SubItems(1))
    deta.Bruto = Val(li.SubItems(3))
    
    deta.IvaAplicado = li.ListSubItems(1).Tag
    deta.IBAplicado = li.ListSubItems(2).Tag

    deta.PorcentajeDescuento = Val(li.SubItems(8))
    fac.Detalles.Add deta
Next li


Me.lbliva2.caption = "IVA " & fac.AlicuotaAplicada & "%"
Me.lblVerIva.caption = funciones.FormatearDecimales(fac.TotalIVA)
Me.lblSubTotal.caption = funciones.FormatearDecimales(fac.TotalSubTotal)
'ver que pasa aca con las percecpciones
Me.lblTotalPercep.caption = funciones.FormatearDecimales(fac.totalPercepciones)

Me.lbltotal.caption = funciones.FormatearDecimales(fac.Total)

''Me.txtOC = verOC
'    Dim subto_ As Double
'    Dim subto1_ As Double
'    Dim dto_ As Double
'    Dim porDtos_ As Double
'    Dim ali_ As Double
'    Dim porAli_ As Double
'    Dim total_ As Double
'    Dim percep_ As Double
'    Dim tot_solo As Double
'    Dim tot_sin_ib As Double
'    Dim tot_sin As Double
'    Dim per As Double
'    Dim tot As Double
'    'If Trim(Me.txtDto) = Empty Then Exit Sub
'    tot = 0
'    tot_solo = 0
'    'If Not IsNumeric(Me.txtDto) Then Exit Sub
'    'dto = CDbl(Me.txtDto)
'    'dto = 1 - (dto / 100)
'
'    'dto1 = CDbl(Me.txtDto) / 100
'    'sumo los parciales de lo que lleva iva
'    For X = 1 To Me.lstFactura.ListItems.count
'        If Me.lstFactura.ListItems(X).ListSubItems(1).Tag = 1 Then
'            tot = tot + CDbl(Me.lstFactura.ListItems(X).ListSubItems(4))
'            tot_solo = tot_solo + funciones.formatearDecimales(CDbl(Me.lstFactura.ListItems(X).ListSubItems(3).Tag), 2) * funciones.formatearDecimales(CDbl(Me.lstFactura.ListItems(X).ListSubItems(1)), 2)
'
'        End If
'    Next X
'
'    'sumo los parciales de lo que no lleva iva
'    For X = 1 To Me.lstFactura.ListItems.count
'        If Me.lstFactura.ListItems(X).ListSubItems(1).Tag = 0 Then
'            tot_sin = tot_sin + CDbl(Me.lstFactura.ListItems(X).ListSubItems(4))
'        End If
'    Next X
'
'    'sumo los parciales de lo que no lleva IIBB
'    For X = 1 To Me.lstFactura.ListItems.count
'        If Me.lstFactura.ListItems(X).ListSubItems(2).Tag = 0 Then
'            tot_sin_ib = tot_sin_ib + funciones.formatearDecimales(CDbl(Me.lstFactura.ListItems(X).ListSubItems(4)), 2)
'        End If
'    Next X
'
'
'
'
'    'claseA.totalizarFactura vIdFactura, subto_, subto1_, dto_, porDtos_, ali_, porAli_, total_, percep_
'
'    If Discrimina = 0 Then    'no discrimina el IVA
'        suba = Format(Math.Round(tot, 2), "0.00")
'        Me.lblSubTotal = suba
'        subto = suba * dto
'        sub_completo = subto
'        tot_sin_ib = 0
'        ali = 1
'        Me.lblTotal = subto
'        Me.lblIva = Empty
'        Me.lblVerIva = funciones.formatearDecimales(0, 2)
'    ElseIf Discrimina = 1 Then
'        subto = CDbl(Format(Math.Round(tot, 2), "0.00"))
'        sub_sin = CDbl(Format(Math.Round(tot_sin, 2), "0.00"))
'        sub_completo = subto + sub_sin
'        Me.lblSubTotal = funciones.formatearDecimales(subto + sub_sin, 2)
'        Me.lbliva2 = "IVA " & funciones.formatearDecimales(Alicuota, 2) & "% "
'        ali = 1 + (Alicuota / 100)
'        ali2 = ali - 1
'        Me.lblVerIva = funciones.formatearDecimales(ali2 * subto * dto, 2)
'
'    End If
'
'    indPer = CDbl(Me.txtPercepciones)
'    ADA = CDbl(funciones.formatearDecimales(tot_sin, 2))
'    AD = funciones.formatearDecimales(tot_solo, 2) + ADA - funciones.formatearDecimales(tot_sin_ib, 2)
'
'
'    'a = (() * dto) * (indPer / 100)
'
'    per = funciones.formatearDecimales(((funciones.formatearDecimales(tot_solo, 2) + CDbl(funciones.formatearDecimales(tot_sin, 2)) - funciones.formatearDecimales(tot_sin_ib, 2)) * dto) * (indPer / 100), 2)
'    'per = claseA.calcularPercepciones(vIdFactura, (sub_completo - tot_sin_ib) * dto, CDbl(Me.txtPercepciones))
'    Me.lblTotalPercep = funciones.formatearDecimales(per, 2)
'    'Me.lblVerDto = funciones.formatearDecimales(sub_completo * dto1, 2)
'    tota = Format(Math.Round((((subto * ali) + sub_sin) * dto) + per, 2), "0.00")
'    Me.lblTotal = tota
'    Total = tota
End Sub
Private Sub LimpiarFactura()
    Me.lstFactura.ListItems.Clear
    'Me.txtDto = 0
    Me.lbliva2 = "I.V.A."
    Me.lblTotalPercep = Empty
    Me.lblVerIva = Empty
    Me.lblLocalidad = Empty
    Me.lbltotal = Empty
    Me.lblSubTotal = Empty
    idCliente = Empty
    Discrimina = Empty
    Alicuota = Empty
    Me.lblCiudad = Empty
    Me.lblCp = Empty
    Me.lblDireccion = Empty
    Me.lblCp = Empty
    Me.lblCuit = Empty
    Me.lblIva = Empty
    Me.txtFP = Empty
    'Me.lblTipoFactura = Empty
    Me.txtNroFactura = Empty
    Me.lblTipoFactura2 = Empty

    If vIdFactura <= 0 Then
        'Me.txtNroFactura = claseA.proximaFactura
        vIdFactura = -1
    End If
End Sub
Private Function agregarAFactura(Optional coleccion As Collection = Nothing, Optional concepto As Boolean = False)
    Dim x As ListItem
    'If coleccion Is Nothing Then Exit Function
    Dim valor As Double
    Dim valor_solo As Double
    If idCliente = Empty Or idCliente = 0 Then
        MsgBox "No hay un cliente definido!"
        Exit Function
    End If
    If Me.lstFactura.ListItems.count < funciones.itemsPorFactura Then
        If Not coleccion Is Nothing Then
            For i = 1 To coleccion.count
                id = coleccion(i)
                'recorro la coleccion
                'y chequeo que el id a facturar no este
                'ya agregado en la factura
                esta = False
                For l = 1 To Me.lstFactura.ListItems.count
                    If IsNumeric(Me.lstFactura.ListItems(l).Tag) Then    'creo q con esto filtro si es un concepto o no
                        idEnFactura = CLng(Me.lstFactura.ListItems(l).Tag)
                        If idEnFactura = id Then
                            esta = True
                        End If
                    End If
                Next

                If Not esta Then

                    Set rs = conectar.RSFactory("select e.origen from entregas e where e.id=" & id)
                    If Not rs.EOF And Not rs.BOF Then
                        origen = rs!origen
                    End If
                    If origen = 3 Or origen = 4 Then 'ot
                        Set rs = conectar.RSFactory("select e.id,e.valor,e.remito,e.idPedido,e.cantidad,e.concepto as detalle,e.origen from entregas e where e.id =" & id)
                    ElseIf origen = 1 Then  'conc
                        Set rs = conectar.RSFactory("select e.id,e.valor,e.remito,e.idPedido,e.cantidad,s.detalle,e.origen from entregas e,detalles_pedidos dp,stock s where e.idDetallePedido=dp.id and dp.idpieza=s.id and e.id =" & id)
                    ElseIf origen = 2 Then    'oe
                        Set rs = conectar.RSFactory("select e.id,e.valor,e.remito,e.idPedido,e.cantidad,s.detalle,e.origen from entregas e,detallesPedidosEntregas dp,stock s where e.idDetallePedido=dp.id and dp.idpieza=s.id and e.id =" & id)
                    End If

                    it = Me.lstFactura.ListItems.count + 1
                    If Not rs.EOF And Not rs.BOF Then
                        Set x = Me.lstFactura.ListItems.Add(, , Format(it, "000"))
                        x.SubItems(1) = rs!cantidad
                        x.SubItems(2) = UCase(rs!detalle)
                        If rs!origen = 1 Then
                            ori = "O/T"
                        ElseIf rs!origen = 2 Then
                            ori = "O/E"
                        ElseIf rs!origen = 3 Then
                            ori = "Concepto"
                        End If
                        If Discrimina = 1 Then    '
                            valor = rs!valor
                            valor_solo = rs!valor
                        ElseIf Discrimina = 0 Then
                            ali = 1 + (Alicuota / 100)
                            valor = rs!valor * ali
                            valor_solo = rs!valor
                        End If
                        x.SubItems(3) = funciones.FormatearDecimales(valor, 2)
                        x.ListSubItems(3).Tag = funciones.FormatearDecimales(valor_solo, 2)
                        x.SubItems(4) = funciones.FormatearDecimales(valor * rs!cantidad, 2)
                        If rs!origen = 3 Then
                            x.SubItems(5) = ori
                        Else
                            x.SubItems(5) = ori & " " & Format(rs!idpedido, "0000")
                        End If
                        x.SubItems(6) = rs!remito
                        'x.SubItems(7) = rs!id
                        x.ListSubItems(1).Tag = 1
                        x.ListSubItems(2).Tag = 1
                        x.SubItems(8) = funciones.FormatearDecimales(0)
                        x.Tag = rs!id
                        If Me.lstFactura.ListItems.count = funciones.itemsPorFactura Then
                            MsgBox "La factura se completo. Por favor utilice otra.", vbExclamation, "Información"
                            Exit Function
                        End If
                    End If


                Else
                    MsgBox "algunos items ya estan en la factura"
                End If

            Next i

        Else
            If concepto Then
                it = Me.lstFactura.ListItems.count + 1
                Set x = Me.lstFactura.ListItems.Add(, , Format(it, "000"))
                x.SubItems(1) = funciones.FormatearDecimales(funciones.ConcCantidad, 2)
                x.SubItems(2) = UCase(funciones.ConcConc)
                If Discrimina = 1 Then  '
                    valor = funciones.ConcValor
                    valor_solo = funciones.ConcValor
                ElseIf Discrimina = 0 Then
                    ali = 1 + (Alicuota / 100)
                    valor = ConcValor * ali
                    valor_solo = ConcValor
                End If
                x.SubItems(3) = funciones.FormatearDecimales(valor, 2)
                x.ListSubItems(3).Tag = funciones.FormatearDecimales(valor_solo, 2)

                x.SubItems(4) = funciones.FormatearDecimales(valor * funciones.ConcCantidad, 2)
                x.SubItems(5) = "Concepto"
                x.SubItems(6) = ""
                'x.SubItems(7) = -1
                x.SubItems(8) = funciones.FormatearDecimales(funciones.DescuentoDetalleFactura, 2)
                x.Tag = -1
                x.ListSubItems(1).Tag = 1
                x.ListSubItems(2).Tag = 1
                funciones.ConcValor = 0
                funciones.ConcCantidad = 0
                funciones.ConcConc = Empty
            End If
        End If

    Else
        MsgBox "Factura completa.Por Favor utilice otra.", vbExclamation, "Información"
    End If
    grabado = False
    Set coleccion = Nothing
    funciones.idEntrega = coleccion
    totalizarFactura
End Function
Private Sub Command1_Click()
    If Not grabado Then
        If MsgBox("¿Está seguro de abandonar la factura?", vbYesNo, "Confirmación") = vbYes Then
            Unload Me
        End If
    Else
        Unload Me
    End If
End Sub

Private Sub btNDIB_Click()
    If Me.lstFactura.ListItems.count > 0 Then
        If Me.lstFactura.selectedItem.ListSubItems(2).Tag = 1 Then
            If MsgBox("¿Desea marcarlo sin percepciones?", vbYesNo, "Consulta") = vbYes Then
                Me.lstFactura.selectedItem.ListSubItems(2).Tag = 0
            End If
        Else
            If MsgBox("¿Desea marcarlo con percepciones ?", vbYesNo, "Consulta") = vbYes Then
                Me.lstFactura.selectedItem.ListSubItems(2).Tag = 1
            End If


        End If

    End If
    totalizarFactura
End Sub

Private Sub btnND_Click()
    If Me.lstFactura.ListItems.count > 0 Then
        If Me.lstFactura.selectedItem.ListSubItems(1).Tag = 1 Then
            If MsgBox("¿Desea marcarlo sin iva?", vbYesNo, "Consulta") = vbYes Then
                Me.lstFactura.selectedItem.ListSubItems(1).Tag = 0
            End If
        Else
            If MsgBox("¿Desea marcarlo con iva ?", vbYesNo, "Consulta") = vbYes Then
                Me.lstFactura.selectedItem.ListSubItems(1).Tag = 1
            End If


        End If

    End If
    totalizarFactura
End Sub

Private Sub cboClientes_Click()
    nuevoCli = Me.cboClientes.ItemData(Me.cboClientes.ListIndex)

    If nuevoCli <> idCliente And idCliente > 0 Then
        Command7_Click    'limpio la factura

    End If
End Sub

Private Sub cmdCopiarFC_Click()

End Sub



Private Sub Command10_Click()
    totalizarFactura
End Sub

Private Sub Command2_Click()
    Dim IDtipo As Long
    On Error GoTo err100

    'limpiarFactura

    'seteo el cliente
    '-----------------------------------------------------------------------'
    If vIdFactura <= 0 Then
        'si la factura es nueva, tomo el idCliene del combo
        idCliente = Me.cboClientes.ItemData(Me.cboClientes.ListIndex)
    Else
        'si la fc es vieja lo traigo de la bbdd
        Set rs = conectar.RSFactory("select idcliente from AdminFacturas where id=" & vIdFactura)
        If Not rs.EOF And Not rs.BOF Then
            idCliente = rs!idCliente
        Else
            MsgBox "ERROR!!"
            Exit Sub
        End If
    End If

    Set rs = conectar.RSFactory("select * from clientes where id=" & idCliente)
    If Not rs.EOF And Not rs.BOF Then
        'valido el CUIT
        If Not IsNumeric(rs!Cuit) Or rs!Cuit < 0 Or Len(rs!Cuit) <> 11 Then
            MsgBox "Hay un error con el CUIT!, no podrá continuar!", vbCritical, "Error"
            Me.bElegirRto.Enabled = False
            Me.Command10.Enabled = False
            Me.Command3.Enabled = False
            Me.Command7.Enabled = False
            Me.Command8.Enabled = False
            Me.Command6.Enabled = False
            Me.Command7.Enabled = False
            Exit Sub
        Else
            vCuit = rs!Cuit
            Me.lblCuit = rs!Cuit
            Me.lblIva = IVA(rs!IVA)
            Me.lblDireccion = rs!Domicilio
            Me.lblCp = rs!CP
            Me.lblCiudad = rs!Ciudad
            Me.lblLocalidad = rs!localidad
            Me.txtFP = rs!FP

            'Command9_Click
        End If
    End If


    '-----------------------------------------------------------------------'


    If vIdFactura <= 0 Then    'esto cuando la factura no se grabo
        Set rs = conectar.RSFactory("select aci.alicuota as alicuotaValor, acf.DiscriminaIVA,acf.id as idTipoFactura,acf.TipoFactura,concat(aci.detalle,' ',aci.alicuota,'%') as alicuota from clientes c inner join AdminConfigFacturas acf on acf.idIVA=c.iva inner join AdminConfigIVA aci on aci.idIVA=acf.idIVA  where c.id=" & idCliente)

        Set rs = conectar.RSFactory("select aci.alicuota as alicuotaValor, acf.DiscriminaIVA,acf.id as idTipoFactura,ft.TipoFactura,concat(aci.detalle,' ',aci.alicuota,'%') as alicuota from clientes c inner join AdminConfigFacturas acf on acf.idIVA=c.iva inner join AdminConfigIVA aci on aci.idIVA=acf.idIVA  inner join AdminConfigFacturasTipos ft on acf.tipoFactura=ft.id where c.id=" & idCliente)


        If Not rs.EOF Or Not rs.BOF Then
            IDtipo = rs!idtipoFactura
            Tipo = rs!TipoFactura
            Discrimina = rs!discriminaIVA
            ' idCliente = rs!idCliente
            Alicuota = rs!alicuotavalor
            errores = False
            Me.lblTipoFactura2 = Tipo
            Me.txtNroFactura = Format(DAOFactura.proximaFactura(IDtipo), "0000")
            verPercepciones
        Else
            LimpiarFactura
            Alicuota = -1
            Discrimina = -1
            vCuit = -1
            errores = True
        End If

    Else    'esto cuando esta en modo edicion

        'anulo la sig. linea, dejando q cada ves q editen la factura lea los datos del cliente con respecto al iva y facturas
        'Set rs = claseS.CrearRS("select f.alicuotaaplicada as alicuotaValor, f.Discriminada ,acf.TipoFactura from AdminFacturas f inner join AdminConfigFacturas acf on acf.id=f.tipoFactura  where f.id=" & vidFactura)
        'Set rs = claseS.CrearRS("select aci.alicuota as alicuotaValor, acf.DiscriminaIVA,acf.TipoFactura,concat(aci.detalle,' ',aci.alicuota,'%') as alicuota from clientes c inner join AdminConfigFacturas acf on acf.idIVA=c.iva inner join AdminConfigIVA aci on aci.idIVA=acf.idIVA  where c.id=" & idCliente)

        Set rs = conectar.RSFactory("select aci.alicuota as alicuotaValor, acf.DiscriminaIVA,ft.TipoFactura,concat(aci.detalle,' ',aci.alicuota,'%') as alicuota from clientes c inner join AdminConfigFacturas acf on acf.idIVA=c.iva inner join AdminConfigIVA aci on aci.idIVA=acf.idIVA inner join AdminConfigFacturasTipos ft on acf.TipoFactura=ft.id  where c.id=" & idCliente)
        If Not rs.EOF Or Not rs.BOF Then
            'Me.lblTipoFactura = rs!tipoFactura & " - " & rs!alicuotavalor

            Discrimina = rs!discriminaIVA
            Alicuota = rs!alicuotavalor
            errores = False

            Tipo = rs!TipoFactura
            ' tipoReal = claseA.tipoFactura(idCliente)
            ' If tipo <> tipoReal Then MsgBox "Hay una diferencia entre el tipo de FC original y el tipo de FC real, se ajusta", vbCritical, "Error"


            Me.lblTipoFactura2 = Tipo    'claseA.tipoFactura(idCliente)
            verPercepciones
        Else
            'Me.lblTipoFactura = "ERROR - DATO NO DEFINIDO."
            LimpiarFactura
            Alicuota = -1
            Discrimina = -1
            errores = True
        End If

    End If

    grabado = False
    Exit Sub
err100:
    MsgBox Err.Description
    errores = True

End Sub

Private Sub verPercepciones()
    If Me.cboPadron.ListIndex = 0 Then
        tabla = "IIBB"
    Else
        tabla = "IIBB_ant"
    End If

    
    Set rs = conectar.RSFactory("select * from sp_permisos." & tabla & " where cuit=" & vCuit)
    If Not rs.EOF And Not rs.BOF Then

        FechaDesde = rs!FechaHasta
        f_desde_anio = Right(FechaDesde, 4)
        f_desde_mes = Mid(FechaDesde, 3, 2)
        f_desde_dia = Mid(FechaDesde, 1, 2)
        Fhasta = f_desde_dia & "/" & f_desde_mes & "/" & f_desde_anio

        If Now() > Fhasta Then
            Me.lblVencido.Visible = True
        Else
            Me.lblVencido.Visible = False
        End If

        Me.txtPercepciones = rs!Percepcion

    End If
End Sub

Private Sub GrabarFactura()
    Dim pib As Double
    Dim errorFact As Boolean
    Dim nroFactura As Long
    Dim IdMoneda As Integer
    Dim dto As Double
    Dim FP As String, oc As String, origen As Integer
    Dim TipoFactura As Long
    Dim obs As String

    If Trim(Me.txtNroFactura) = Empty Or Trim(Me.txtOC) = Empty Or Trim(Me.txtFP) = Empty Then errorFact = True
    If errorFact Then
        MsgBox "Debe completar todos los datos requeridos por la factura", vbCritical, "Error"
        Exit Sub
    Else
        If MsgBox("¿Está seguro de continuar?", vbYesNo, "Confirmación") = vbYes Then
            'If vidFactura = -1 Then
            'si estamos acá, es porque la factura no fue nunca grabada, por ende.. y por silogismo disyuntivo
            'hay que crearla de 0. Una vez creada, se camba ese -1 por el id de la factura a fin
            'de la proxima grabada sea solo modif.
            nroFactura = CLng(Me.txtNroFactura)
            IdMoneda = CLng(Me.cboMoneda.ItemData(Me.cboMoneda.ListIndex))
            FP = UCase(Me.txtFP)
            oc = UCase(Me.txtOC)
            obs = UCase(Me.txtObservar)
            TipoFactura = claseA.TipoFactura(idCliente)
            'dto = CDbl(Me.txtDto)    'aca, calculo que voy a poenr si hay algun descuento por anticipo o algo asi.
            pib = CDbl(Me.txtPercepciones)
            If Not IsNumeric(dto) Then dto = 0
            'If obs = Empty Then obs = "!"
            If dto = Empty Then dto = 0
            Dim Tipo As Integer

            If Not claseA.crearFactura(idCliente, IdMoneda, dto, FP, oc, vorigen, Me.DTPicker1, TipoFactura, Me.lstFactura, Alicuota, Discrimina, tipoNC, , , nroFactura, vIdFactura, pib, obs) Then
                MsgBox "Se produjo un error al crear la factura.", vbCritical, "Error"
            Else
                MsgBox "La factura se guardó con éxito!", vbInformation, "Información"
                vIdFactura = nroFactura
            End If
        End If
        grabado = True
    End If

End Sub
Private Sub Command3_Click()
    GrabarFactura
End Sub
Private Sub Command5_Click()
    Command1_Click
End Sub
Private Sub Command6_Click()
    If idCliente = Empty Or idCliente = 0 Then
        MsgBox "No hay un cliente definido!", vbCritical
        Exit Sub
    Else
        FrmAdminFacturarConcepto.Show 1
        If funciones.ConcCantidad > 0 Or funciones.ConcValor > 0 Or Not funciones.ConcConc = Empty Then
            agregarAFactura , True     'trae los datos de funciones.conc.....
        End If
    End If
End Sub

Private Sub Command7_Click()
    If MsgBox("¿Está seguro de limpiar la factura?", vbYesNo, "Confirmación") = vbYes Then
        LimpiarFactura
        grabado = False
        vIdFactura = -1
    End If

End Sub

Private Sub Command8_Click()
    For i = Me.lstFactura.ListItems.count To 1 Step -1
        If Me.lstFactura.ListItems(i).Checked = True Then
            Me.lstFactura.ListItems.Remove (i)
        End If
    Next i
totalizarFactura
End Sub

Public Function verPadron() As Double
    If Not IsNumeric(vCuit) Then
        MsgBox "El cliente no tiene cuit aplicado, no se podrá facturar", vbCritical, "Error"
    Else
        If idCliente = Empty Or idCliente = 0 Then
            MsgBox "No hay un cliente definido!", vbCritical
            Exit Function
        Else
            Dim r As recordset
            Set r = conectar.RSFactory("select percepcion from sp_permisos.IIBB where cuit=" & vCuit)
            If Not r.EOF And Not r.BOF Then
                verPadron = r!Percepcion
            End If
        End If
    End If
End Function

Private Sub Command9_Click()

End Sub

Private Sub DTPicker1_Click()
    grabado = False
End Sub
Private Sub Form_Load()

FormHelper.Customize Me
    Me.cboPadron.Clear
    cboPadron.AddItem "Padrón Actual"
        Me.cboPadron.ItemData(Me.cboPadron.NewIndex) = 0
    cboPadron.AddItem "Padrón Anterior"
        Me.cboPadron.ItemData(Me.cboPadron.NewIndex) = 1
        
        Me.cboPadron.ListIndex = 0
    
    LimpiarFactura
    claseS.llenar_combo_clientes Me.cboClientes, 9999
    claseC.llenarCboMonedas Me.cboMoneda
    grabado = False
    Me.DTPicker1 = Now
    errores = False
    Me.btnND.Visible = False
    If vIdFactura > 0 Then    'es una factura ya grabada
        cargarDatos
    Else
        If tipoNC = 0 Then
            Me.lblND.Visible = False
            Me.lblNC.Visible = False
            Me.bElegirRto.Enabled = True
            Me.btnND.Visible = False
            Me.btNDIB.Visible = False
        ElseIf tipoNC = 1 Then
            Me.lblND.Visible = False
            Me.lblNC.Visible = True
            Me.bElegirRto.Enabled = False
            Me.btnND.Visible = True
            Me.btNDIB.Visible = False
        ElseIf tipoNC = 2 Then
            Me.lblND.Visible = True
            Me.lblNC.Visible = False
            Me.bElegirRto.Enabled = False
            Me.btnND.Visible = True
            Me.btNDIB.Visible = True
        End If
    End If
End Sub
Private Sub cargarDatos()
    Dim rs3 As recordset
    Dim strsql As String
    Dim x As ListItem
    Dim valecant As Double
    Dim vale As Double
    Dim vale_solo As Double
    'Set rs = claseA.CrearRS("select formaPago, nroFactura,ordencompra,idcliente,discriminada from AdminFacturas where id=" & vidFactura)
    'command2_click
    Command2_Click
    Set rs = conectar.RSFactory("select f.idMoneda, f.descuento,f.id as id_factura,f.observaciones,f.tipo ,f.aliPercIB as valor,f.formaPago, f.nroFactura,f.ordencompra,f.idcliente,f.discriminada from AdminFacturas f where f.id=" & vIdFactura)

    If Not rs.EOF And Not rs.BOF Then
        idCliente = rs!idCliente


        desculo = rs!Descuento

        nroFactura = rs!nroFactura
        obs = rs!Observaciones
        TipoFactura = claseA.TipoFactura(idCliente)
        Discrimina = rs!Discriminada
        tipoNC = rs!Tipo
'
'            factura = 0
'    notaCredito = 1
'    NotaDebito = 2
        If tipoNC = 0 Then
            Me.lblNC.Visible = False
            Me.lblND.Visible = False
            Me.bElegirRto.Enabled = True
            Me.btnND.Visible = False
            Me.btNDIB.Visible = False
        ElseIf tipoNC = 1 Then
            Me.btnND.Visible = False
            Me.lblNC.Visible = True
            Me.bElegirRto.Enabled = False
            Me.btnND.Visible = False
            Me.btNDIB.Visible = False
        ElseIf tipoNC = 2 Then
            Me.lblNC.Visible = False
            Me.lblND.Visible = True
            Me.btnND.Visible = True
            Me.bElegirRto.Enabled = False
            Me.btnND.Visible = True
            Me.btNDIB.Visible = True
        End If
        Me.cboClientes.ListIndex = funciones.PosIndexCbo(idCliente, Me.cboClientes)
        Me.cboMoneda.ListIndex = funciones.PosIndexCbo(rs!IdMoneda, Me.cboMoneda)
        
        Me.txtFP = rs!FormaPago
        Me.txtOC = rs!OrdenCompra
        percep = rs!valor
        Me.txtObservar = obs

        'Me.txtDto = desculo

        'muestro los iibb del padron
        Me.lblPadron.caption = Me.verPadron & " %"
        Me.txtPercepciones = Round((percep - 1) * 100, 2)
        Me.txtNroFactura = Format(nroFactura, "0000")
        If Discrimina = 1 Then    'no discrimina el IVA
            ali = 1    'no discrimina
        ElseIf Discrimina = 0 Then
            ali = 1 + (Alicuota / 100)
        End If

        strsql = "select * from AdminFacturasDetalleNueva where idFactura=" & vIdFactura

        Set rs3 = conectar.RSFactory(strsql)
        it = 0
        While Not rs3.EOF
            vale = rs3!valor * ali
            vale_solo = rs3!valor
            valecant = funciones.FormatearDecimales(vale, 2) * funciones.FormatearDecimales(rs3!cantidad, 2)
            it = it + 1
            Set x = Me.lstFactura.ListItems.Add(, , Format(it, "000"))
            x.SubItems(1) = Format(rs3!cantidad, "0.00")
            x.SubItems(2) = rs3!detalle
            x.SubItems(3) = funciones.FormatearDecimales(vale, 2)
            x.ListSubItems(3).Tag = funciones.FormatearDecimales(vale_solo, 2)
            x.SubItems(4) = funciones.FormatearDecimales(valecant, 2)


            claseP.datosEntrega rs3!idEntrega, origen, remito
            x.SubItems(5) = origen
            x.SubItems(6) = remito
            '    x.SubItems(7) = rs3!idEntrega
            x.Tag = rs3!idEntrega
            x.ListSubItems(1).Tag = rs3!IVA
            x.ListSubItems(2).Tag = rs3!IB
            
            'el 7 es el id
            x.SubItems(8) = funciones.FormatearDecimales(rs3!porcentaje_descuento, 2)
            rs3.MoveNext
        Wend
        totalizarFactura

    Else
        MsgBox "Error!", vbCritical, "Error"
        Exit Sub
    End If
    Set rs3 = Nothing
End Sub


Private Sub txtDto_Change()
    totalizarFactura
End Sub



Private Sub lstFactura_DblClick()
    If IsSomething(Me.lstFactura.selectedItem) Then
        Dim tmp As String
        tmp = InputBox("Especifice el porcentaje de descuento del item", "Descuento", Me.lstFactura.selectedItem.SubItems(8))
        If LenB(tmp) > 0 Then
            Me.lstFactura.selectedItem.SubItems(8) = funciones.FormatearDecimales(Val(tmp))
            totalizarFactura
        End If
    End If
End Sub

Private Sub txtFP_Change()
    grabado = False
End Sub

Private Sub txtFP_Validate(Cancel As Boolean)
    If Not IsNumeric(txtFP) Then Cancel = True Else Cancel = False
End Sub

Private Sub txtNroFactura_Change()
    grabado = False
End Sub
Private Sub txtOC_Change()
    grabado = False
End Sub
Function llenarFacturaDesdeRemito(IdRemito As Long) As Boolean
    llenarFacturaDesdeRemito = True
    Set rs = conectar.RSFactory("select e.valor,e.idPedido,e.cantidad,s.detalle,e.origen from entregas e,detalles_pedidos dp,stock s where e.idDetallePedido=dp.id and dp.idpieza=s.id and e.remito =" & IdRemito & " and e.origen=1 union all select e.valor,e.idPedido,e.cantidad,s.detalle,e.origen from entregas e,detallesPedidosEntregas dp,stock s where e.idDetallePedido=dp.id and dp.idpieza=s.id and e.remito=" & IdRemito & " and e.origen=2")
    While Not rs.EOF
        it = it + 1
        Set x = Me.lstFactura.ListItems.Add(, , Format(it, "000"))
        x.SubItems(1) = rs!cantidad
        x.SubItems(2) = rs!detalle
        If rs!origen = 1 Then ori = "O/T" Else ori = "O/E"
        x.SubItems(3) = ori & " " & Format(rs!idpedido, "0000")
        x.SubItems(4) = rs!valor
        x.ListSubItems(4).Tag = rs!valor
        rs.MoveNext
    Wend
End Function
Private Function verDescuento() As Double
    Dim strsql As String
    Dim rs As recordset
    'recorro la lista y veo los items agregados
    'si son todos de la misma OT verifico que tenga descuento aplicado y lifting
    igual = True
    For x = 1 To Me.lstFactura.ListItems.count
        'ide = Me.lstFactura.ListItems(x).ListSubItems(7)
        ide = Me.lstFactura.ListItems(x).Tag
        strsql = "select dto from pedidos p inner join detalles_pedidos dp on p.id=dp.idpedido inner join entregas e on e.idDetallePedido=dp.id  where e.id=" & ide
        Set rs = conectar.RSFactory(strsql)
        If Not rs.EOF And Not rs.BOF Then
            dto = rs!dto
        End If
        If x = 1 Then dtoAnterior = dto
        If dto = dtoAnterior Then
            Descuento = dto
        Else
            igual = False
        End If
    Next x
    If Not igual Then Descuento = 0
    verDescuento = Descuento
End Function
Private Function verOC() As String
'muestra variables que dicebn descuento, pero es para ver si es todo de la misma OT/OC
    Dim strsql As String
    Dim rs As recordset
    Dim s As New classStock
    'recorro la lista y veo los items agregados
    'si son todos de la misma OT verifico que tenga descuento aplicado y lifting
    origen = 1
    igual = True
    For x = 1 To Me.lstFactura.ListItems.count
        '    ide = Me.lstFactura.ListItems(x).ListSubItems(7)
        ide = Me.lstFactura.ListItems(x).Tag
        If ide = -1 Then
            verOC = "VARIOS"
            Exit Function
        End If
        Dim rs1 As recordset
        Set rs1 = conectar.RSFactory("select origen from entregas where id=" & ide)
        If Not rs1.EOF And Not rs1.BOF Then
            origen = rs1!origen
        End If

        If origen = 1 Or origen = 4 Then
            strsql = "select p.id from pedidos p inner join detalles_pedidos dp on p.id=dp.idpedido inner join entregas e on e.idDetallePedido=dp.id  where e.id=" & ide
        ElseIf origen = 2 Then
            strsql = "select p.id from PedidosEntregas p inner join detallesPedidosEntregas dp on p.id=dp.idpedidoEntrega inner join entregas e on e.idDetallePedido=dp.id  where e.id=" & ide
        ElseIf origen = 3 Then

        End If


        If origen = 3 Then    ' si es concepto no jodas
            Exit Function
        End If
        Set rs = conectar.RSFactory(strsql)
        If Not rs.EOF And Not rs.BOF Then
            idOC = rs!id
        End If
        If x = 1 Then IDOCAnterior = idOC
        If idOC <> IDOCAnterior Then
            igual = False
        End If


    Next x

    If Not igual Then detalle = "INGRESE NRO DE OC"
    If Me.lstFactura.ListItems.count > 0 Then
        If origen = 1 Or origen = 4 Then
            Set rs = conectar.RSFactory("select descripcion from pedidos where id=" & idOC)
        ElseIf origen = 2 Then
            Set rs = conectar.RSFactory("select referencia as descripcion from PedidosEntregas where id=" & idOC)
        End If
        If Not rs.EOF And Not rs.BOF Then
            detalle = rs!Descripcion
        End If

    End If

    verOC = detalle
End Function


