VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAdminNuevaFactura 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Nueva Factura..."
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14265
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   14265
   ShowInTaskbar   =   0   'False
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
      Height          =   6615
      Left            =   4680
      TabIndex        =   21
      Top             =   0
      Width           =   9495
      Begin VB.CommandButton Command10 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Recalcular"
         Height          =   255
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   4920
         Width           =   1575
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Quitar"
         Height          =   255
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   4560
         Width           =   1575
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Facturar concepto"
         Height          =   255
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   6240
         Width           =   1575
      End
      Begin VB.CommandButton bElegirRto 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Facturar remito"
         Height          =   255
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   5880
         Width           =   1575
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
         Left            =   5760
         TabIndex        =   22
         Top             =   4560
         Width           =   3615
         Begin VB.CommandButton Command4 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Ver"
            CausesValidation=   0   'False
            Height          =   195
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   50
            Top             =   960
            Width           =   735
         End
         Begin VB.TextBox txtDto 
            Height          =   285
            Left            =   1320
            TabIndex        =   45
            Text            =   "0"
            Top             =   600
            Width           =   855
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
            TabIndex        =   53
            Top             =   960
            Width           =   1215
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
            TabIndex        =   52
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label lblVerDto 
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
            TabIndex        =   51
            Top             =   600
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
            TabIndex        =   49
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Caption         =   "Descuento "
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
            TabIndex        =   44
            Top             =   600
            Width           =   1215
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
            TabIndex        =   39
            Top             =   1680
            Width           =   1215
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
            TabIndex        =   38
            Top             =   1680
            Width           =   855
         End
         Begin VB.Label lbliva2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Caption         =   "I.V.A. "
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
            TabIndex        =   37
            Top             =   1320
            Width           =   2055
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
            TabIndex        =   36
            Top             =   240
            Width           =   1215
         End
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
            TabIndex        =   35
            Top             =   240
            Width           =   855
         End
      End
      Begin MSComctlLib.ListView lstFactura 
         Height          =   4215
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   7435
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
         NumItems        =   8
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
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Unitario"
            Object.Width           =   1623
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
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
            SubItemIndex    =   6
            Text            =   "Remito"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "idEntrega"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Shape Shape1 
         Height          =   1575
         Left            =   3960
         Top             =   4920
         Width           =   1695
      End
      Begin VB.Label lblTipoFactura2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   99.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1815
         Left            =   3960
         TabIndex        =   40
         Top             =   4560
         Width           =   1695
      End
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Command1"
      Height          =   375
      Left            =   1320
      TabIndex        =   20
      Top             =   8280
      Width           =   1215
   End
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
      Height          =   3975
      Left            =   0
      TabIndex        =   19
      Top             =   2640
      Width           =   4575
      Begin VB.CommandButton Command9 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Buscar"
         Height          =   255
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox txtPercepciones 
         Height          =   285
         Left            =   1440
         TabIndex        =   54
         Text            =   "0"
         Top             =   2160
         Width           =   1935
      End
      Begin VB.ComboBox cboPercepcionesAplicadas 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   48
         Top             =   3000
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Nueva"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   3000
         Width           =   975
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Salir"
         Height          =   375
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   3480
         Width           =   975
      End
      Begin VB.ComboBox cboMoneda 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   600
         Width           =   1575
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Guardar"
         Default         =   -1  'True
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3480
         Width           =   975
      End
      Begin VB.TextBox txtFP 
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Top             =   1440
         Width           =   3015
      End
      Begin VB.TextBox txtOC 
         Height          =   285
         Left            =   1440
         TabIndex        =   5
         Top             =   1800
         Width           =   3015
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   255
         Left            =   1200
         TabIndex        =   3
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         _Version        =   393216
         Format          =   63635457
         CurrentDate     =   39176
      End
      Begin VB.TextBox txtNroFactura 
         Height          =   285
         Left            =   3600
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Percep  IIBB"
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
         TabIndex        =   47
         Top             =   2160
         Width           =   1215
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
         Left            =   120
         TabIndex        =   32
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblTipoFactura 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   1440
         TabIndex        =   30
         Top             =   1080
         Width           =   3015
      End
      Begin VB.Label Label10 
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
         Left            =   360
         TabIndex        =   27
         Top             =   1440
         Width           =   1095
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
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   975
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
         Left            =   360
         TabIndex        =   25
         Top             =   1800
         Width           =   1095
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
         Left            =   2640
         TabIndex        =   24
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo "
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
         Left            =   360
         TabIndex        =   23
         Top             =   1080
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ Datos Cliente ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   4575
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ver"
         Height          =   255
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   360
         Width           =   615
      End
      Begin VB.ComboBox cboClientes 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   2655
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
         Left            =   120
         TabIndex        =   29
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label lblCiudad 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   1080
         TabIndex        =   28
         Top             =   1320
         Width           =   3375
      End
      Begin VB.Label lblIva 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   1080
         TabIndex        =   18
         Top             =   2040
         Width           =   3375
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
         Left            =   120
         TabIndex        =   17
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label lblCuit 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   1080
         TabIndex        =   16
         Top             =   1080
         Width           =   3375
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
         Left            =   120
         TabIndex        =   15
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lblCp 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   1080
         TabIndex        =   14
         Top             =   1800
         Width           =   3375
      End
      Begin VB.Label lblLocalidad 
         BackColor       =   &H00E0E0E0&
         Caption         =   " "
         Height          =   255
         Left            =   1080
         TabIndex        =   13
         Top             =   1560
         Width           =   3375
      End
      Begin VB.Label lblDireccion 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   1080
         TabIndex        =   12
         Top             =   840
         Width           =   3375
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "C.P."
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
         TabIndex        =   11
         Top             =   1800
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
         Left            =   120
         TabIndex        =   10
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
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   855
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
         TabIndex        =   8
         Top             =   360
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmAdminNuevaFactura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim claseP As New classPlaneamiento
Dim claseC As New classConfigurar
Dim vidFactura As Long
Dim vcuit As Double
'origen facturado
'1- desde remito.
'2- por conceptos.
Dim discrimina As Integer
Dim alicuota As Double
Dim idsentrega As Collection
Dim errores As Boolean
Dim idCliente As Long
Dim claseS As New classStock
Dim claseA As New classAdministracion
Dim rs As Recordset
Dim vorigen As Integer
Dim grabado As Boolean
Public Property Let idFactura(nidFactura As Long)
vidFactura = nidFactura
End Property
Public Property Get idFactura() As Long
idFactura = vidFactura
End Property

Public Property Let OrigenFactura(norigen As Integer)
  vorigen = norigen
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
    frmPlaneamientoRemitosListaProceso.mostrar = CLng(Me.cboClientes.ItemData(Me.cboClientes.ListIndex))
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
        Me.txtDto = verDescuento
        Me.txtOC = verOC
    End If
End If
End If

End Sub
Private Sub totalizarFactura(Optional ByRef total As Double)
Dim per As Double
Dim tot As Double
If Trim(Me.txtDto) = Empty Then Exit Sub
tot = 0
If Not IsNumeric(Me.txtDto) Then Exit Sub
dto = CDbl(Me.txtDto)
dto = 1 - (dto / 100)

dto1 = CDbl(Me.txtDto) / 100
'sumo los parciales
For X = 1 To Me.lstFactura.ListItems.count
 tot = tot + CDbl(Me.lstFactura.ListItems(X).ListSubItems(4))
Next X

'me fijo las retenciones
'per = claseA.calcularPercepciones(tot, idCliente)


If discrimina = 0 Then 'no discrimina el IVA
 suba = Format(Math.Round(tot, 2), "0.00")
 Me.lblSubTotal = suba
 subto = suba * dto
 Me.lblTotal = subto
 Me.lblIva = Empty
 Me.lblVerIva = 0
ElseIf discrimina = 1 Then
 subto = Format(Math.Round(tot, 2), "0.00")
 Me.lblSubTotal = subto
 
 Me.lbliva2 = "IVA " & funciones.formatearDecimales(alicuota, 2) & "% "
 ali = 1 + (alicuota / 100)
 ali2 = ali - 1
 Me.lblVerIva = funciones.formatearDecimales(ali2 * subto, 2)
 'veriva
  End If
'verdto
per = claseA.calcularPercepciones(tot * dto, idCliente, , , CDbl(Me.txtPercepciones))
Me.lblTotalPercep = funciones.formatearDecimales(per, 2)
Me.lblVerDto = funciones.formatearDecimales(subto * dto1, 2)
tota = Format(Math.Round((tot * ali * dto) + per, 2), "0.00")
Me.lblTotal = tota
total = tota
End Sub
Private Sub limpiarFactura()
Me.lstFactura.ListItems.Clear
Me.txtDto = 0
Me.lbliva2 = "I.V.A."
Me.lblTotalPercep = Empty
Me.lblVerIva = Empty
Me.lblLocalidad = Empty
Me.lblTotal = Empty
Me.lblSubTotal = Empty
idCliente = Empty
discrimina = Empty
alicuota = Empty
Me.lblCiudad = Empty
Me.lblCp = Empty
Me.lblDireccion = Empty
Me.lblCp = Empty
Me.lblCuit = Empty
Me.lblIva = Empty
Me.txtFP = Empty
Me.lblTipoFactura = Empty
Me.txtNroFactura = Empty
Me.lblTipoFactura2 = Empty

If vidFactura <= 0 Then
'Me.txtNroFactura = claseA.proximaFactura
vidFactura = -1
End If
End Sub
Private Function agregarAFactura(Optional coleccion As Collection)
Dim X As ListItem
'If coleccion Is Nothing Then Exit Function
Dim valor As Double
If idCliente = Empty Or idCliente = 0 Then
 MsgBox "No hay un cliente definido!"
 Exit Function
End If
If Me.lstFactura.ListItems.count < funciones.itemsPorFactura Then
     If Not coleccion Is Nothing Then
         For I = 1 To coleccion.count
            id = coleccion(I)
            'recorro la coleccion
            'y chequeo que el id a facturar no este
            'ya agregado en la factura
            esta = False
            For l = 1 To Me.lstFactura.ListItems.count
                If IsNumeric(Me.lstFactura.ListItems(l).ListSubItems(6)) Then 'creo q con esto filtro si es un concepto o no
                    idEnFactura = CLng(Me.lstFactura.ListItems(l).ListSubItems(6))
                    If idEnFactura = id Then
                        esta = True
                    End If
                End If
            Next
        
    If Not esta Then
      
          Set rs = claseA.CrearRS("select e.origen from entregas e where e.id=" & id)
            If Not rs.EOF And Not rs.BOF Then
                origen = rs!origen
            End If
        If origen = 3 Then
           Set rs = claseA.CrearRS("select e.id,e.valor,e.remito,e.idPedido,e.cantidad,e.concepto as detalle,e.origen from entregas e where e.id =" & id)
        Else
            Set rs = claseA.CrearRS("select e.id,e.valor,e.remito,e.idPedido,e.cantidad,s.detalle,e.origen from entregas e,detalles_pedidos dp,stock s where e.idDetallePedido=dp.id and dp.idpieza=s.id and e.id =" & id)
        End If
    End If
    it = Me.lstFactura.ListItems.count + 1
    If Not rs.EOF And Not rs.BOF Then
        Set X = Me.lstFactura.ListItems.Add(, , Format(it, "000"))
            X.SubItems(1) = rs!cantidad
            X.SubItems(2) = rs!detalle
            If rs!origen = 1 Then
                ori = "O/T"
            ElseIf rs!origen = 2 Then
                ori = "O/E"
            ElseIf rs!origen = 3 Then
                ori = "CONC"
            End If
          If discrimina = 1 Then '
          valor = rs!valor
        ElseIf discrimina = 0 Then
         ali = 1 + (alicuota / 100)
         valor = rs!valor * ali
        End If
        X.SubItems(3) = funciones.formatearDecimales(valor, 2)
        X.SubItems(4) = funciones.formatearDecimales(valor * rs!cantidad, 2)
        If rs!origen = 3 Then
            X.SubItems(5) = ori
        Else
            X.SubItems(5) = ori & " " & Format(rs!idpedido, "0000")
        End If
        X.SubItems(6) = rs!remito
        X.SubItems(7) = rs!id
        
        If Me.lstFactura.ListItems.count = funciones.itemsPorFactura Then
              MsgBox "La factura se completo. Por favor utilice otra.", vbExclamation, "Información"
            Exit Function
        End If
  'End If


Else
MsgBox "algunos items ya estan en la factura"
 End If
 Next I

Else
it = Me.lstFactura.ListItems.count + 1
Set X = Me.lstFactura.ListItems.Add(, , Format(it, "000"))
        X.SubItems(1) = funciones.ConcCantidad
        X.SubItems(2) = funciones.ConcConc
        If discrimina = 1 Then  '
          valor = ConcValor
        ElseIf discrimina = 0 Then
         ali = 1 + (alicuota / 100)
         valor = ConcValor * ali
        End If
        X.SubItems(3) = funciones.formatearDecimales(valor, 2)
        X.SubItems(4) = funciones.formatearDecimales(valor * funciones.ConcCantidad, 2)
        X.SubItems(5) = "Concepto"
        X.SubItems(6) = "N/D"
        X.SubItems(7) = -1
        funciones.ConcValor = 0
        funciones.ConcCantidad = 0
        funciones.ConcConc = Empty
End If

Else
MsgBox "Factura completa.Por Favor utilice otra.", vbExclamation, "Información"
End If
grabado = False
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

Private Sub Command10_Click()
totalizarFactura
End Sub

Private Sub Command2_Click()
 On Error GoTo err100
limpiarFactura
If vidfactua <= 0 Then
        idCliente = Me.cboClientes.ItemData(Me.cboClientes.ListIndex)
End If


Set rs = claseS.CrearRS("select * from clientes where id=" & idCliente)
If Not rs.EOF And Not rs.BOF Then
claseA.llenarComboPercepcionesAplicadas Me.cboPercepcionesAplicadas, idCliente
Me.lblDireccion = rs!domicilio
Me.lblCp = rs!cp
Me.lblCiudad = rs!ciudad
Me.lblCuit = rs!cuit
If IsNumeric(rs!cuit) Then vcuit = CDbl(rs!cuit) Else vcuit = -1
'VALIDAR CUIT PARA VER PERCEPCION
Me.lblIva = Iva(rs!Iva)
Me.lblLocalidad = rs!localidad
Me.txtFP = rs!fp
Command9_Click
End If


If vidFactura <= 0 Then 'esto cuando la factura no se grabo
Set rs = claseS.CrearRS("select aci.alicuota as alicuotaValor, acf.DiscriminaIVA,acf.TipoFactura,concat(aci.detalle,' ',aci.alicuota,'%') as alicuota from sp.clientes c inner join sp.AdminConfigFacturas acf on acf.idIVA=c.iva inner join sp.AdminConfigIVA aci on aci.idIVA=acf.idIVA  where c.id=" & idCliente)
If Not rs.EOF Or Not rs.BOF Then
 Me.lblTipoFactura = rs!tipoFactura & " - " & rs!alicuota
   
 discrimina = rs!discriminaIVA
 alicuota = rs!alicuotavalor
 errores = False
 Me.lblTipoFactura2 = rs!tipoFactura
 verPercepciones
 Else
 Me.lblTipoFactura = "ERROR - DATO NO DEFINIDO."
 limpiarFactura
 alicuota = -1
 discrimina = -1
 errores = True
End If

Else 'esto cuando esta en modo edicion

'anulo la sig. linea, dejando q cada ves q editen la factura lea los datos del cliente con respecto al iva y facturas
'Set rs = claseS.CrearRS("select f.alicuotaaplicada as alicuotaValor, f.Discriminada ,acf.TipoFactura from AdminFacturas f inner join AdminConfigFacturas acf on acf.id=f.tipoFactura  where f.id=" & vidFactura)
Set rs = claseS.CrearRS("select aci.alicuota as alicuotaValor, acf.DiscriminaIVA,acf.TipoFactura,concat(aci.detalle,' ',aci.alicuota,'%') as alicuota from sp.clientes c inner join sp.AdminConfigFacturas acf on acf.idIVA=c.iva inner join sp.AdminConfigIVA aci on aci.idIVA=acf.idIVA  where c.id=" & idCliente)
If Not rs.EOF Or Not rs.BOF Then
 Me.lblTipoFactura = rs!tipoFactura & " - " & rs!alicuotavalor
   
 discrimina = rs!discriminaIVA
 alicuota = rs!alicuotavalor
 errores = False
 Me.lblTipoFactura2 = rs!tipoFactura
 verPercepciones
 Else
 Me.lblTipoFactura = "ERROR - DATO NO DEFINIDO."
 limpiarFactura
 alicuota = -1
 discrimina = -1
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


End Sub

Private Sub Command3_Click()
Dim errorFact As Boolean
Dim nroFactura As Long, idMoneda As Integer, dto As Double, fp As String, oc As String, origen As Integer
Dim tipoFactura As Long
If Trim(Me.txtNroFactura) = Empty Or Trim(Me.txtOC) = Empty Or Trim(Me.txtFP) = Empty Then errorFact = True
If errorFact Then
 MsgBox "Debe completar todos los datos requeridos por la factura", vbCritical, "Error"
 Exit Sub
Else
 If MsgBox("¿Está seguro de continuar?", vbYesNo, "Confirmación") = vbYes Then
If vidFactura = -1 Then
 'si estamos acá, es porque la factura no fue nunca grabada, por ende.. y por silogismo disyuntivo
 'hay que crearla de 0. Una vez creada, se camba ese -1 por el id de la factura a fin
 'de la proxima grabada sea solo modif.
If Not claseA.ExisteFactura(CLng(Me.txtNroFactura)) Then
'perfecto.
'creamos la factura y traemos el número creado.
 nroFactura = CLng(Me.txtNroFactura)
 'idCliente = CLng(Me.cboClientes.ItemData(Me.cboClientes.ListIndex))
 idMoneda = CLng(Me.cboMoneda.ItemData(Me.cboMoneda.ListIndex))
 fp = normaliza(Me.txtFP)
 oc = normaliza(Me.txtOC)
 tipoFactura = claseA.tipoFactura(idCliente)
 dto = CDbl(Me.txtDto) 'aca, calculo que voy a poenr si hay algun descuento por anticipo o algo asi.
If Not IsNumeric(dto) Then dto = 0
If dto = Empty Then dto = 0
Dim tipo As Integer



 If Not claseA.crearFactura(nroFactura, idCliente, idMoneda, dto, fp, oc, vorigen, Me.DTPicker1, tipoFactura, Me.lstFactura, alicuota, discrimina, 0) Then
    MsgBox "Se produjo un error al crear la factura.", vbCritical, "Error"
 Else
    MsgBox "La factura se creo con éxito!", vbInformation, "Información"
    vidFactura = nroFactura
 End If
Else
MsgBox "El número de factura ya existe.", vbCritical, "Error"
End If
Else
'si estamos acá es porque la factura ya se creo, esta en modo edicion y hay que actualizar los datos

End If
grabado = True
'imprimirFactura
 End If
End If
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
        
        agregarAFactura 'trae los datos de funciones.conc.....
    End If
End If
End Sub

Private Sub Command7_Click()
If MsgBox("¿Está seguro de limpiar la factura?", vbYesNo, "Confirmación") = vbYes Then
limpiarFactura
grabado = False
Me.txtNroFactura = claseA.proximaFactura
End If

End Sub

Private Sub Command8_Click()
For I = Me.lstFactura.ListItems.count To 1 Step -1
If Me.lstFactura.ListItems(I).Checked = True Then
 Me.lstFactura.ListItems.Remove (I)
End If
Next I

End Sub

Private Sub Command9_Click()
If Not IsNumeric(vcuit) Then
  MsgBox "El cliente no tiene cuit aplicado, no se podrá facturar", vbCritical, "Error"
Else
If idCliente = Empty Or idCliente = 0 Then
MsgBox "No hay un cliente definido!", vbCritical
Exit Sub
Else
Dim r As Recordset
Set r = claseC.CrearRS("select percepcion from sp_permisos.IIBB where cuit=" & vcuit)
If Not r.EOF And Not r.BOF Then
Me.txtPercepciones = r!percepcion
End If
End If
End If
End Sub

Private Sub DTPicker1_Click()
grabado = False
End Sub
Private Sub Form_Load()
limpiarFactura
claseS.llenar_combo_clientes Me.cboClientes, 9999
claseC.llenarCboMonedas Me.cboMoneda

grabado = False
Me.DTPicker1 = Now
errores = False
If vidFactura <= 0 Then
    'Me.txtNroFactura = claseA.proximaFactura
Else
    
    cargarDatos
End If
End Sub

Private Sub cargarDatos()
Set rs = claseA.CrearRS("select formaPago, ordencompra,idcliente from AdminFacturas where id=" & vidFactura)
        If Not rs.EOF And Not rs.BOF Then
            idCliente = rs!idCliente
            Me.cboClientes.ListIndex = funciones.PosIndexCbo(idCliente, Me.cboClientes)
            Me.txtFP = rs!formaPago
            Me.txtOC = rs!ordenCompra
            Command2_Click
            Me.txtNroFactura = vidFactura
            totalizarFactura
        Else
            MsgBox "Error!", vbCritical, "Error"
            Exit Sub
    
        End If
                
End Sub

Private Sub txtDto_Change()
totalizarFactura
End Sub

Private Sub txtFP_Change()
grabado = False
End Sub

Private Sub txtNroFactura_Change()
grabado = False
End Sub

Private Sub txtOC_Change()
grabado = False
End Sub
Function llenarFacturaDesdeRemito(idremito As Long) As Boolean
llenarFacturaDesdeRemito = True
Set rs = claseA.CrearRS("select e.valor,e.idPedido,e.cantidad,s.detalle,e.origen from entregas e,detalles_pedidos dp,stock s where e.idDetallePedido=dp.id and dp.idpieza=s.id and e.remito =" & idremito & " and e.origen=1 union all select e.valor,e.idPedido,e.cantidad,s.detalle,e.origen from entregas e,detallesPedidosEntregas dp,stock s where e.idDetallePedido=dp.id and dp.idpieza=s.id and e.remito=" & idremito & " and e.origen=2")
While Not rs.EOF
it = it + 1
Set X = Me.lstFactura.ListItems.Add(, , Format(it, "000"))
    X.SubItems(1) = rs!cantidad
    X.SubItems(2) = rs!detalle
    If rs!origen = 1 Then ori = "O/T" Else ori = "O/E"
    X.SubItems(3) = ori & " " & Format(rs!idpedido, "0000")
    X.SubItems(4) = rs!valor
rs.MoveNext
Wend
End Function
Private Function verDescuento() As Double
Dim strsql As String
Dim rs As Recordset
'recorro la lista y veo los items agregados
'si son todos de la misma OT verifico que tenga descuento aplicado y lifting
igual = True
For X = 1 To Me.lstFactura.ListItems.count
idE = Me.lstFactura.ListItems(X).ListSubItems(7)
 strsql = "select dto from pedidos p inner join detalles_pedidos dp on p.id=dp.idpedido inner join entregas e on e.idDetallePedido=dp.id  where e.id=" & idE
Set rs = claseP.CrearRS(strsql)
If Not rs.EOF And Not rs.BOF Then
dto = rs!dto
End If
If X = 1 Then dtoAnterior = dto
If dto = dtoAnterior Then
    descuento = dto
Else
igual = False
End If
Next X
If Not igual Then descuento = 0
verDescuento = descuento
End Function
Private Function verOC() As String
'muestra variables que dicebn descuento, pero es para ver si es todo de la misma OT/OC
Dim strsql As String
Dim rs As Recordset
'recorro la lista y veo los items agregados
'si son todos de la misma OT verifico que tenga descuento aplicado y lifting
igual = True
For X = 1 To Me.lstFactura.ListItems.count
idE = Me.lstFactura.ListItems(X).ListSubItems(7)
If idE = -1 Then Exit Function
 strsql = "select p.id from pedidos p inner join detalles_pedidos dp on p.id=dp.idpedido inner join entregas e on e.idDetallePedido=dp.id  where e.id=" & idE
Set rs = claseP.CrearRS(strsql)
If Not rs.EOF And Not rs.BOF Then
IDOC = rs!id
End If
If X = 1 Then IDOCAnterior = dto

If IDOC <> IDOCAnterior Then
    igual = False
End If

Next X
If Not igual Then detalle = "INGRESE NRO DE OC"
Set rs = claseP.CrearRS("select descripcion from pedidos where id=" & IDOC)
If Not rs.EOF And Not rs.BOF Then
detalle = rs!descripcion
End If
verOC = detalle
End Function




