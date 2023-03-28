VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmNuevoElemento 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Nuevo elemento..."
   ClientHeight    =   11025
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   13125
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   11025
   ScaleWidth      =   13125
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Pieza"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10935
      Left            =   0
      TabIndex        =   11
      Top             =   15
      Width           =   13095
      Begin VB.TextBox txtNombreElemento 
         Height          =   285
         Left            =   960
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   240
         Width           =   8895
      End
      Begin XtremeSuiteControls.ComboBox cboClientes 
         Height          =   315
         Left            =   2880
         TabIndex        =   2
         Top             =   600
         Width           =   10080
         _Version        =   786432
         _ExtentX        =   17780
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Text            =   "ComboBox1"
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0C0C0&
         Caption         =   "[ Mano de obra ]"
         Height          =   4815
         Left            =   120
         TabIndex        =   32
         Top             =   6000
         Width           =   12855
         Begin VB.CommandButton Command2 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Invertir"
            Height          =   255
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   81
            Top             =   4440
            Width           =   735
         End
         Begin VB.CommandButton Command5 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Agregar"
            Default         =   -1  'True
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   11655
            Style           =   1  'Graphical
            TabIndex        =   59
            Top             =   1650
            Width           =   1095
         End
         Begin VB.CommandButton Command4 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Quitar"
            Height          =   255
            Left            =   960
            Style           =   1  'Graphical
            TabIndex        =   48
            Top             =   4440
            Width           =   735
         End
         Begin VB.TextBox txtTiempo 
            Height          =   285
            Left            =   960
            TabIndex        =   36
            Text            =   "Text1"
            ToolTipText     =   "Tiempo en minutos"
            Top             =   1200
            Width           =   3255
         End
         Begin VB.TextBox txtCantOp 
            Height          =   285
            Left            =   1440
            TabIndex        =   35
            Text            =   "Text1"
            Top             =   840
            Width           =   2775
         End
         Begin VB.CommandButton btnAgregarMDO 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Agregar"
            Height          =   375
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   1680
            Width           =   975
         End
         Begin VB.Frame Frame7 
            BackColor       =   &H00C0C0C0&
            Caption         =   "[ Detalle ]"
            ForeColor       =   &H00000000&
            Height          =   1095
            Left            =   4320
            TabIndex        =   37
            Top             =   360
            Width           =   8415
            Begin VB.Label lblValor 
               BackColor       =   &H00C0C0C0&
               Height          =   255
               Left            =   6720
               TabIndex        =   76
               Top             =   720
               Width           =   1575
            End
            Begin VB.Label lblSector 
               BackColor       =   &H00C0C0C0&
               Height          =   255
               Left            =   1200
               TabIndex        =   51
               Top             =   720
               Width           =   3975
            End
            Begin VB.Label Label25 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0C0&
               Caption         =   "Sector "
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
               Top             =   720
               Width           =   1095
            End
            Begin VB.Label lblCPP 
               BackColor       =   &H00C0C0C0&
               Height          =   255
               Left            =   6720
               TabIndex        =   44
               Top             =   240
               Width           =   1575
            End
            Begin VB.Label lblDescripcion 
               BackColor       =   &H00C0C0C0&
               Height          =   255
               Left            =   1200
               TabIndex        =   43
               Top             =   480
               Width           =   7095
            End
            Begin VB.Label lblTarea 
               BackColor       =   &H00C0C0C0&
               Height          =   255
               Left            =   1200
               TabIndex        =   42
               Top             =   240
               Width           =   3975
            End
            Begin VB.Label Label22 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0C0&
               Caption         =   "Cant x Proceso "
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
               Left            =   5160
               TabIndex        =   41
               Top             =   240
               Width           =   1575
            End
            Begin VB.Label Label20 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0C0&
               Caption         =   "Descripcion "
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
               TabIndex        =   40
               Top             =   480
               Width           =   1095
            End
            Begin VB.Label T 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0C0&
               Caption         =   "Tarea "
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
               TabIndex        =   39
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label va 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0C0&
               Caption         =   "Valor "
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
               Left            =   5160
               TabIndex        =   75
               Top             =   720
               Width           =   1575
            End
         End
         Begin VB.TextBox txtCodigoMDO 
            Height          =   285
            Left            =   840
            TabIndex        =   34
            Text            =   "Text1"
            Top             =   480
            Width           =   3375
         End
         Begin MSComctlLib.ListView ListView2 
            Height          =   2175
            Left            =   120
            TabIndex        =   47
            Top             =   2160
            Width           =   12615
            _ExtentX        =   22251
            _ExtentY        =   3836
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   16777215
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   11
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   529
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "Código"
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Text            =   "Cant OP"
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Text            =   "Tiempo"
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   4
               Text            =   "Sector"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   5
               Text            =   "CPP"
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   6
               Text            =   "Tarea"
               Object.Width           =   4762
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "idcpp"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   8
               Text            =   "Descripcion"
               Object.Width           =   4057
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   9
               Text            =   "T.Total"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   10
               Text            =   "Costo $"
               Object.Width           =   1587
            EndProperty
         End
         Begin VB.CommandButton btnModificar 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Modificar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   11640
            Style           =   1  'Graphical
            TabIndex        =   79
            Top             =   1680
            Width           =   1095
         End
         Begin VB.Label lblidStock 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Label27"
            Height          =   255
            Left            =   7680
            TabIndex        =   80
            Top             =   1800
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label lblCtoMDO 
            BackColor       =   &H00C0C0C0&
            Height          =   255
            Left            =   12120
            TabIndex        =   78
            Top             =   4440
            Width           =   615
         End
         Begin VB.Label Label21 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Costo $"
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
            Left            =   11280
            TabIndex        =   77
            Top             =   4440
            Width           =   735
         End
         Begin VB.Label Label31 
            BackColor       =   &H00C0C0C0&
            Caption         =   "M.D.O"
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
            Left            =   7080
            TabIndex        =   58
            Top             =   4440
            Width           =   615
         End
         Begin VB.Label Label30 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Cambio"
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
            Left            =   8520
            TabIndex        =   57
            Top             =   4440
            Width           =   735
         End
         Begin VB.Label Label29 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Fijos"
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
            Left            =   9960
            TabIndex        =   56
            Top             =   4440
            Width           =   495
         End
         Begin VB.Label lblcambio 
            BackColor       =   &H00C0C0C0&
            Height          =   255
            Left            =   9255
            TabIndex        =   55
            Top             =   4440
            Width           =   615
         End
         Begin VB.Label lblmdo 
            BackColor       =   &H00C0C0C0&
            Height          =   255
            Left            =   7710
            TabIndex        =   54
            Top             =   4440
            Width           =   615
         End
         Begin VB.Label lblfijos 
            BackColor       =   &H00C0C0C0&
            Height          =   255
            Left            =   10440
            TabIndex        =   53
            Top             =   4440
            Width           =   615
         End
         Begin VB.Label lblidCPP 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Label26"
            Height          =   255
            Left            =   6960
            TabIndex        =   52
            Top             =   1800
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label lblidMDO 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Label25"
            Height          =   255
            Left            =   6240
            TabIndex        =   49
            Top             =   1800
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label Label24 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Tiempo"
            Height          =   255
            Left            =   240
            TabIndex        =   46
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label Label23 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Cant Operarios"
            Height          =   255
            Left            =   240
            TabIndex        =   45
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label Label19 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Código"
            Height          =   255
            Left            =   240
            TabIndex        =   33
            Top             =   480
            Width           =   855
         End
      End
      Begin VB.TextBox txtIdCliente 
         Height          =   285
         Left            =   960
         TabIndex        =   1
         Text            =   "Text2"
         Top             =   600
         Width           =   1815
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "[ Materiales ]"
         Height          =   4935
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   12855
         Begin VB.TextBox txtCodigoMaterial 
            Height          =   285
            Left            =   855
            TabIndex        =   82
            Text            =   "Text1"
            Top             =   375
            Width           =   1005
         End
         Begin VB.CommandButton Command3 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Calcular "
            Height          =   255
            Left            =   3120
            Style           =   1  'Graphical
            TabIndex        =   74
            Top             =   4560
            Width           =   1095
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00C0C0C0&
            Caption         =   "[ Medida Pieza ]"
            Enabled         =   0   'False
            Height          =   1215
            Left            =   4440
            TabIndex        =   67
            Top             =   240
            Width           =   2055
            Begin VB.TextBox txtLargoPieza 
               Height          =   285
               Left            =   720
               TabIndex        =   8
               Top             =   360
               Width           =   1215
            End
            Begin VB.TextBox txtAnchoPieza 
               Height          =   285
               Left            =   735
               TabIndex        =   9
               Top             =   720
               Width           =   1215
            End
            Begin VB.Label Label14 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Largo"
               Height          =   255
               Left            =   120
               TabIndex        =   69
               Top             =   360
               Width           =   615
            End
            Begin VB.Label Label9 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Ancho"
               Height          =   255
               Left            =   120
               TabIndex        =   68
               Top             =   720
               Width           =   615
            End
         End
         Begin VB.CommandButton Command10 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Ok"
            Height          =   255
            Left            =   2520
            Style           =   1  'Graphical
            TabIndex        =   66
            Top             =   4560
            Width           =   495
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "frmNuevoElemento.frx":0000
            Left            =   960
            List            =   "frmNuevoElemento.frx":0010
            Style           =   2  'Dropdown List
            TabIndex        =   65
            Top             =   4560
            Width           =   1455
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   2655
            Left            =   135
            TabIndex        =   27
            Top             =   1800
            Width           =   12615
            _ExtentX        =   22251
            _ExtentY        =   4683
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   15
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Código"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "id"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Text            =   "Descripción"
               Object.Width           =   9172
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Text            =   "Esp mm"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   4
               Text            =   "Pieza"
               Object.Width           =   2716
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "x"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "y"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "XT"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "YT"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   9
               Text            =   "Scrap %"
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   10
               Text            =   "Kg"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   11
               Text            =   "M2/Ml"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   12
               Text            =   "Costo $"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   13
               Text            =   "term"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   14
               Text            =   "cantidad"
               Object.Width           =   0
            EndProperty
         End
         Begin VB.TextBox txtScrap 
            Height          =   285
            Left            =   840
            TabIndex        =   5
            Text            =   "Text2"
            ToolTipText     =   "Valor porcentual"
            Top             =   1080
            Width           =   1335
         End
         Begin VB.TextBox txtCantidad 
            Height          =   285
            Left            =   840
            TabIndex        =   4
            Text            =   "Text1"
            Top             =   720
            Width           =   1335
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Agregar"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   1440
            Width           =   735
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00C0C0C0&
            Caption         =   "[ Detalles ]"
            Height          =   1215
            Left            =   6600
            TabIndex        =   19
            Top             =   240
            Width           =   6135
            Begin VB.Label lblXHoja 
               BackColor       =   &H00C0C0C0&
               Height          =   255
               Left            =   3960
               TabIndex        =   71
               Top             =   600
               Width           =   2055
            End
            Begin VB.Label lblm2 
               BackColor       =   &H00C0C0C0&
               Height          =   255
               Left            =   3960
               TabIndex        =   61
               Top             =   840
               Width           =   2055
            End
            Begin VB.Label lblkg 
               BackColor       =   &H00C0C0C0&
               Height          =   255
               Left            =   960
               TabIndex        =   60
               Top             =   840
               Width           =   1455
            End
            Begin VB.Label lblEspesor 
               BackColor       =   &H00C0C0C0&
               Height          =   255
               Left            =   960
               TabIndex        =   25
               Top             =   600
               Width           =   1455
            End
            Begin VB.Label Label13 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0C0&
               Caption         =   "Espesor "
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
               TabIndex        =   24
               Top             =   600
               Width           =   855
            End
            Begin VB.Label lblMaterial 
               BackColor       =   &H00C0C0C0&
               Height          =   255
               Left            =   960
               TabIndex        =   21
               Top             =   360
               Width           =   5055
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0C0&
               Caption         =   "Material "
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
               TabIndex        =   20
               Top             =   360
               Width           =   855
            End
            Begin VB.Label Label18 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0C0&
               Caption         =   "Cant x Hoja "
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
               Left            =   2400
               TabIndex        =   70
               Top             =   600
               Width           =   1575
            End
            Begin VB.Label Label7 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0C0&
               Caption         =   "Total KG "
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
               TabIndex        =   63
               Top             =   840
               Width           =   855
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0C0&
               Caption         =   "Total M2/Ml "
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
               Left            =   2400
               TabIndex        =   62
               Top             =   840
               Width           =   1575
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00C0C0C0&
            Caption         =   "[ Medida Hoja ]"
            Enabled         =   0   'False
            Height          =   1215
            Left            =   2280
            TabIndex        =   16
            Top             =   240
            Width           =   2055
            Begin VB.TextBox txtAnchoTerm 
               Height          =   285
               Left            =   720
               TabIndex        =   7
               Top             =   720
               Width           =   1215
            End
            Begin VB.TextBox txtLargoTerm 
               Height          =   285
               Left            =   720
               TabIndex        =   6
               Top             =   360
               Width           =   1215
            End
            Begin VB.Label Label5 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Ancho"
               Height          =   255
               Left            =   120
               TabIndex        =   18
               Top             =   720
               Width           =   615
            End
            Begin VB.Label Label4 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Largo"
               Height          =   255
               Left            =   120
               TabIndex        =   17
               Top             =   360
               Width           =   615
            End
         End
         Begin XtremeSuiteControls.PushButton Command6 
            Height          =   285
            Left            =   1920
            TabIndex        =   3
            Top             =   360
            Width           =   300
            _Version        =   786432
            _ExtentX        =   529
            _ExtentY        =   494
            _StockProps     =   79
            Caption         =   "..."
            BackColor       =   16744576
            UseVisualStyle  =   -1  'True
         End
         Begin VB.Label Label26 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Costo $"
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
            Left            =   11280
            TabIndex        =   73
            Top             =   4560
            Width           =   735
         End
         Begin VB.Label lblCosto 
            BackColor       =   &H00C0C0C0&
            Height          =   255
            Left            =   12120
            TabIndex        =   72
            Top             =   4560
            Width           =   615
         End
         Begin VB.Label Label8 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Marcados"
            Height          =   255
            Left            =   120
            TabIndex        =   64
            Top             =   4560
            Width           =   735
         End
         Begin VB.Label lblTotalKg 
            BackColor       =   &H00C0C0C0&
            Height          =   255
            Left            =   8520
            TabIndex        =   31
            Top             =   4560
            Width           =   615
         End
         Begin VB.Label lblTotalM2 
            BackColor       =   &H00C0C0C0&
            Height          =   255
            Left            =   10440
            TabIndex        =   30
            Top             =   4560
            Width           =   615
         End
         Begin VB.Label Label17 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Total M2/Ml"
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
            Left            =   9240
            TabIndex        =   29
            Top             =   4560
            Width           =   1215
         End
         Begin VB.Label Label16 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Total KG"
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
            TabIndex        =   28
            Top             =   4560
            Width           =   855
         End
         Begin VB.Label lblidMaterial 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Label16"
            Height          =   255
            Left            =   960
            TabIndex        =   26
            Top             =   1440
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label Label12 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Scrap"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label Label11 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Cantidad"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   735
            Width           =   735
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Código"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   360
            Width           =   615
         End
      End
      Begin XtremeSuiteControls.ComboBox cboComplejidad 
         Height          =   315
         Left            =   11115
         TabIndex        =   83
         Top             =   210
         Width           =   1875
         _Version        =   786432
         _ExtentX        =   3307
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin VB.Label lblComplejidad 
         Caption         =   "Complejidad"
         Height          =   225
         Left            =   9975
         TabIndex        =   84
         Top             =   285
         Width           =   1110
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nombre"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmNuevoElemento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim grabado As Boolean
Dim rss As Recordset
Dim baseM As New classConfigurar
Dim base As New classNuevoElemento
Dim baseS As New classStock
'Dim baseSP As New classSignoplast
Dim Kg
Dim m2



Public Sub calcularTotalMateriales(ByVal lst As ListView, ByRef Kg, ByRef m2, ByRef costo)
    Dim K As Double, m As Double, c As Double
    Dim i As Integer

    For i = 1 To lst.ListItems.count
        K = K + CDbl(lst.ListItems(i).ListSubItems(10))
        m = m + CDbl(lst.ListItems(i).ListSubItems(11))
        c = c + CDbl(lst.ListItems(i).ListSubItems(12))
    Next
    Kg = K
    m2 = m
    costo = c
    Me.lblTotalKg = Kg
    Me.lblTotalM2 = m2
    Me.lblCosto = costo
End Sub
Public Sub calcular_totales_mdo()
    canti = Me.ListView2.ListItems.count
    TotalMDO = 0
    totalCAMBIO = 0
    totalFIJO = 0
    cto = 0
    For i = 1 To canti
        If (CInt(Me.ListView2.ListItems(i).ListSubItems(7))) = -1 Then totalCAMBIO = totalCAMBIO + CDbl(Me.ListView2.ListItems(i).ListSubItems(9))
        If (CInt(Me.ListView2.ListItems(i).ListSubItems(7))) = 0 Then totalFIJO = totalFIJO + CDbl(Me.ListView2.ListItems(i).ListSubItems(9))
        If (CInt(Me.ListView2.ListItems(i).ListSubItems(7))) > 0 Then TotalMDO = TotalMDO + CDbl(Me.ListView2.ListItems(i).ListSubItems(9))
        cto = cto + CDbl(Me.ListView2.ListItems(i).ListSubItems(10))
    Next i
    Me.lblfijos = Math.Round(totalFIJO, 2)
    Me.lblmdo = Math.Round(TotalMDO, 2)
    Me.lblcambio = Math.Round(totalCAMBIO, 2)
    Me.lblCtoMDO = Math.Round(cto, 2)
End Sub
Private Function verDetalleMateriales(Id)
    Dim Kg As Double, m2ml As Double
    Dim descripcion As String
    Dim costo As Double

    x = Me.txtAnchoPieza
    y = Me.txtLargoPieza
    x1 = Me.txtAnchoTerm
    y1 = Me.txtLargoTerm
    If Trim(x) = Empty Then x = 0
    If Trim(x1) = Empty Then x1 = 0
    If Trim(y) = Empty Then y = 0
    If Trim(y1) = Empty Then y1 = 0

    Cant = CDbl(Me.txtCantidad)

    baseS.calcularM2MLKGMaterial x1, y1, Id, Scrap, x, y, Cant, Kg, m2ml, Pieza, costo, 0
    cxh = funciones.cantxhoja(x, y, x1, y1)

    baseS.ejecutar "select m.espesor,m.descripcion, g.grupo, r.rubro from materiales m,grupos g, rubros r where m.id_grupo=g.id and m.id_rubro=r.id and  m.id=" & Id
    descripcion = baseS.descripcion
    Espesor = baseS.Espesor
    Grupo = baseS.Grupo
    rubro = baseS.rubro
    Scrap = Val(Me.txtScrap)
    verDetalle cxh, descripcion, Espesor, Kg, m2ml, Grupo, rubro
End Function

Private Sub btnAgregarMDO_Click()
    If Not IsNumeric(Me.txtCantOp) Or Not IsNumeric(Me.txtTiempo) Then
        MsgBox "Ingrese datos válidos por favor", vbCritical, "Error"
    Else

        Set x = Me.ListView2.ListItems.Add(, , Id)
        x.SubItems(1) = Me.lblidMDO
        x.SubItems(2) = Me.txtCantOp
        x.SubItems(3) = Me.txtTiempo
        x.SubItems(4) = Me.lblSector
        x.SubItems(5) = Me.lblCPP
        x.SubItems(7) = Me.lblidCPP
        x.SubItems(6) = Me.lblTarea
        x.SubItems(8) = Me.lblDescripcion

        Valor = CDbl(Me.lblValor)
        Tiempo = CDbl(Me.txtTiempo)
        cpp = CInt(Me.lblidCPP)
        cantop = CDbl(Me.txtCantOp)
        If cpp > 0 Then    '(cpp variable)
            totmin = cantop * Tiempo / cpp
            totplata = totmin * Valor
        Else
            totmin = cantop * Tiempo
            totplata = totmin * Valor

        End If
        x.SubItems(9) = Math.Round(totmin, 2)
        x.SubItems(10) = Math.Round(totplata, 2)

        Me.txtCodigoMDO.SetFocus
    End If
    grabado = False
    Me.calcular_totales_mdo
End Sub


Private Sub btnModificar_Click()
    On Error GoTo errb
    Dim idPieza As Long

    If IsNumeric(Trim(Me.lblidStock)) Then    'si est definido el id de la pieza a modificar
        ErrorCode = 0
        idPieza = CLng(Me.lblidStock)
        If Trim(Me.txtNombreElemento) = Empty Then
            MsgBox "Error, debe completar todos los campos.", vbCritical, "Error"
        Else
            Dim h As VbMsgBoxResult



            h = MsgBox("¿Está conforme con los datos ingresados?", vbYesNo, "Confirmación")
            If h = 6 Then
                base.modificar Me.ListView1, Me.ListView2, Me.txtNombreElemento, CInt(Me.txtIdCliente), idPieza, CInt(Me.cboComplejidad.ItemData(Me.cboComplejidad.ListIndex))
                grabado = True

                Channel.Notificar Nothing, EdicionPieza_
            End If
        End If
    End If
    Exit Sub
errb:
    MsgBox Err.Description

End Sub

Private Sub cboClientes_Click()
    On Error Resume Next
    Me.txtIdCliente = Me.cboClientes.ItemData(Me.cboClientes.ListIndex)
End Sub

Private Sub cboComplejidad_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next

End Sub

Private Sub Command1_Click()
    On Error GoTo erra
    Dim Kg As Double, m2ml As Double
    Dim descripcion As String
    If Trim(Me.txtCodigoMaterial) <> Empty Then
        codigo = UCase(Me.txtCodigoMaterial)
        Id = baseM.QueIdMaterial(Trim(Me.txtCodigoMaterial))
        x = Me.txtAnchoPieza
        y = Me.txtLargoPieza
        x1 = Me.txtAnchoTerm
        y1 = Me.txtLargoTerm

        If Trim(x) = Empty Then x = 0
        If Trim(x1) = Empty Then x1 = 0
        If Trim(y) = Empty Then y = 0
        If Trim(y1) = Empty Then y1 = 0

        Cant = CDbl(Me.txtCantidad)
        'si existe el código
        If Id <> -1 Then
            cxh = funciones.cantxhoja(x, y, x1, y1)
            'si cxh = 0 then exit
            If cxh <= 0 Then
                MsgBox "No puede procesar con estas dimensiones!", vbInformation, "Error"
                Exit Sub
            End If
            baseS.ejecutar "select m.id_unidad,m.espesor,m.descripcion, g.grupo, r.rubro from materiales m,grupos g,rubros r where m.id_grupo=g.id and m.id_rubro=r.id and  m.id=" & Id
            descripcion = baseS.descripcion
            Espesor = baseS.Espesor
            Grupo = baseS.Grupo
            rubro = baseS.rubro
            If Trim(Me.txtScrap) = Empty Or Not IsNumeric(Me.txtScrap) Then
                Scrap = 0
            Else
                Scrap = CDbl(Me.txtScrap)
            End If
            uni = baseS.idUnidad

            'If uni = 3 Then cxh = 1
            'si son ml cantxhoja debe quedar en 1

            If uni = 4 Or uni = 1 Then
                x = 0
                x1 = 0
                y = 0
                y1 = 0
            End If
            Dim costo As Double

            baseS.calcularM2MLKGMaterial x, y, Id, Scrap, x1, y1, Cant, Kg, m2ml, Pieza, costo, 0
            'agrego datos a la lista
            Dim h As ListItem
            Set h = Me.ListView1.ListItems.Add(, , codigo)
            h.SubItems(1) = Id
            h.SubItems(2) = rubro & " " & Grupo & " " & descripcion
            h.SubItems(3) = Espesor
            h.SubItems(4) = Pieza
            h.SubItems(5) = y1
            h.SubItems(6) = x1
            h.SubItems(7) = y
            h.SubItems(8) = x
            h.SubItems(9) = Scrap
            h.SubItems(10) = Kg
            h.SubItems(11) = m2ml
            h.SubItems(12) = funciones.FormatearDecimales(costo, 2)
            h.SubItems(13) = ""
            h.SubItems(14) = Cant


            verDetalle cxh, descripcion, Espesor, Kg, m2ml, Grupo, rubro

        End If
    End If
    calcularTotalMateriales Me.ListView1, Kg, m2, costo

    grabado = False
    Exit Sub
erra:
    MsgBox Err.Description

End Sub

Private Function verDetalle(cantxhoja, MAT As String, esp, Kg, m2, Grupo, rubro)
    Me.lblXHoja = cantxhoja
    Me.lblEspesor = esp
    Me.lblm2 = m2
    Me.lblkg = Kg
    Me.lblMaterial = truncar(rubro, 40) & " " & truncar(Grupo, 40) & " " & truncar(MAT, 40)

End Function

Private Sub Command10_Click()
    If Me.Combo1.ListIndex = 0 Then
        If MsgBox("¿Está seguro de eliminar los items seleecionados?", vbYesNo, "Confirmacion") = vbYes Then
            quitar

        End If
    Else
        term (Me.Combo1.ListIndex)
    End If

End Sub

Private Sub Command2_Click()
'base.ver_detalle_elemento Trim(Me.txtCodigoMaterial), Me, 0
    For i = 1 To Me.ListView2.ListItems.count
        If Me.ListView2.ListItems(i).Checked Then
            Me.ListView2.ListItems(i).Checked = False
        Else
            Me.ListView2.ListItems(i).Checked = True
        End If

    Next i

End Sub

Private Sub quitar()
    For i = Me.ListView1.ListItems.count To 1 Step -1
        If Me.ListView1.ListItems(i).Checked = True Then
            Me.ListView1.ListItems.remove (i)
            grabado = False
        End If
        Me.calcularTotalMateriales Me.ListView1, Kg, m2, costo
    Next i
End Sub

Private Sub Command3_Click()
    If Me.Combo1.ListIndex = 1 Then
        For i = 1 To Me.ListView1.ListItems.count
            If Me.ListView1.ListItems(i).Checked = True Then
                codigo = Me.ListView1.ListItems(i)
                descripcion = Me.ListView1.ListItems(i).ListSubItems(2)
                Kg = Me.ListView1.ListItems(i).ListSubItems(10)
                m2 = Me.ListView1.ListItems(i).ListSubItems(11)
                Largo = Me.ListView1.ListItems(i).ListSubItems(7)
                Ancho = Me.ListView1.ListItems(i).ListSubItems(8)
                Cantidad = Me.ListView1.ListItems(i).ListSubItems(14)
                medida = Me.ListView1.ListItems(i).ListSubItems(4)
                Set x = frmConfigurarTerminacion.lstPiezas.ListItems.Add(, , codigo)
                x.SubItems(1) = descripcion
                x.SubItems(2) = Kg
                x.SubItems(3) = m2
                x.SubItems(4) = medida
                x.SubItems(5) = 2
                x.SubItems(6) = 1
                x.SubItems(7) = Cantidad
                x.SubItems(8) = Largo
                x.SubItems(9) = Ancho

            End If
        Next i

        frmConfigurarTerminacion.nuevo_form = Me
        frmConfigurarTerminacion.Show 1
        calcularTotalMateriales Me.ListView1, Kg, m2, costo

    End If
    grabado = False

End Sub

Private Sub Command4_Click()
    If MsgBox("¿Está seguro de eliminar los items seleecionados?", vbYesNo, "Confirmacion") = vbYes Then
        For i = Me.ListView2.ListItems.count To 1 Step -1
            If Me.ListView2.ListItems(i).Checked = True Then
                Me.ListView2.ListItems.remove (i)
                grabado = False
            End If
        Next i
    End If
    Me.calcular_totales_mdo
End Sub





Private Sub Command5_Click()
'agregar
    On Error GoTo errb
    ErrorCode = 0
    If baseS.buscar_pieza(Trim(Me.txtNombreElemento)) > 0 Then
        MsgBox "El nombre asignado ya existe en la base de datos", vbCritical, "Error"
    Else

        If Me.txtIdCliente = vbNullString Or Me.txtIdCliente = 0 Or Trim(Me.txtNombreElemento) = Empty Then
            MsgBox "Error, debe completar todos los campos.", vbCritical, "Error"
        Else
            Dim h As VbMsgBoxResult
            h = MsgBox("¿Está conforme con los datos ingresados?", vbYesNo, "Confirmación")
            If h = 6 Then
                base.agregar Me.ListView1, Me.ListView2, Me.txtNombreElemento, CInt(Me.txtIdCliente), CInt(Me.cboComplejidad.ItemData(Me.cboComplejidad.ListIndex))
                grabado = False
            End If
        End If
    End If
    Exit Sub
errb:
    MsgBox Err.Description
End Sub

Private Sub term(nu)
    If nu = 1 Then
        color1 = vbRed
    Else
        color1 = vbBlue
    End If

    For i = 1 To Me.ListView1.ListItems.count

        If Me.ListView1.ListItems(i).Checked = True Then
            If Me.ListView1.ListItems(i).ListSubItems(13).ForeColor = color1 Then
                Me.ListView1.ListItems(i).ListSubItems(13).ForeColor = vbBlack
            Else
                Me.ListView1.ListItems(i).ListSubItems(13).ForeColor = color1
            End If
        End If

    Next i

End Sub


Private Sub Command9_Click()

    Dim x As ListItem
    For i = 1 To Me.ListView1.ListItems.count
        If Me.ListView1.ListItems(i).ListSubItems(11).text = "X1" Then
            codigo = Me.ListView1.ListItems(i).ListSubItems(1)
            descripcion = Me.ListView1.ListItems(i).ListSubItems(2)
            Kg = Me.ListView1.ListItems(i).ListSubItems(7)
            m2 = Me.ListView1.ListItems(i).ListSubItems(8)
            Largo = Me.ListView1.ListItems(i).ListSubItems(4)
            Ancho = Me.ListView1.ListItems(i).ListSubItems(5)

            medida = Me.ListView1.ListItems(i).ListSubItems(3)
            Set x = frmConfigurarTerminacion.lstPiezas.ListItems.Add(, , codigo)
            x.SubItems(1) = descripcion
            x.SubItems(2) = Kg
            x.SubItems(3) = m2
            x.SubItems(4) = medida
            x.SubItems(5) = 2
            x.SubItems(6) = 1
            x.SubItems(7) = Largo
            x.SubItems(8) = Ancho


        End If

    Next i
    frmConfigurarTerminacion.Show 1
End Sub

Private Sub Command6_Click()
    Dim frm As New frmMaterialesLista2_modal
    frm.Usable = True
    frm.Show 1

    If IsSomething(Selecciones.Material) Then
        Me.txtCodigoMaterial = Selecciones.Material.codigo
        Set Selecciones.Material = Nothing
    End If

End Sub

Private Sub Form_Activate()

    Me.calcular_totales_mdo
    Me.calcularTotalMateriales Me.ListView1, Kg, m2, costo

End Sub


Private Sub Form_Load()
    FormHelper.Customize Me
    Me.Combo1.ListIndex = 0
    Me.limpiar_txt
    Me.calcular_totales_mdo
    'Me.calcularTotalMateriales Me.ListView1, kg, m2, costo
    Me.lblCosto = costo
    Me.lblTotalKg = Kg
    Me.lblTotalM2 = m2
    grabado = False
    DAOCliente.llenarComboXtremeSuite Me.cboClientes

    cboComplejidad.AddItem "Baja"
    cboComplejidad.ItemData(cboComplejidad.NewIndex) = 1
    cboComplejidad.AddItem "Media"
    cboComplejidad.ItemData(cboComplejidad.NewIndex) = 2
    cboComplejidad.AddItem "Alta"
    cboComplejidad.ItemData(cboComplejidad.NewIndex) = 3
    Me.cboComplejidad.ListIndex = 0

    ''Me.caption = caption & " (" & Name & ")"


End Sub
Function limpiar_txt()
    Me.txtAnchoPieza = Empty
    Me.txtCodigoMaterial = Empty
    Me.txtIdCliente = Empty
    Me.txtLargoPieza = Empty
    Me.txtNombreElemento = Empty
    Me.lblMaterial = Empty
    Me.txtCantidad = 1
    Me.txtScrap = 25
    Me.lblEspesor = Empty
    Me.txtCodigoMDO = Empty
    Me.txtTiempo = Empty
    Me.lblkg = Empty
    Me.lblm2 = Empty
    Me.txtCantOp = 1

End Function

Private Sub Form_Unload(Cancel As Integer)
    If Not grabado Then
        If MsgBox("¿Desea descartar los cambios?", vbYesNo, "Confirmación") = vbYes Then
            Unload Me
        Else
            Cancel = 1
        End If
    End If
End Sub

Private Sub lblCliente_Click()

End Sub

Private Sub ListView1_DblClick()
    If Me.ListView1.ListItems.count > 0 Then

        Scrap = Me.ListView1.selectedItem.ListSubItems(9)
        xhoja = Me.ListView1.selectedItem.ListSubItems(5)
        yhoja = Me.ListView1.selectedItem.ListSubItems(6)
        xpieza = Me.ListView1.selectedItem.ListSubItems(7)
        ypieza = Me.ListView1.selectedItem.ListSubItems(8)
        Cantidad = Me.ListView1.selectedItem.ListSubItems(14)
        codigo = Me.ListView1.selectedItem
        idcodigo = Me.ListView1.selectedItem.ListSubItems(1)

        frmDesarrolloModificarMaterial.nuevo_form = Me
        frmDesarrolloModificarMaterial.txtScrap = Scrap
        frmDesarrolloModificarMaterial.txtCantidad = Cantidad
        frmDesarrolloModificarMaterial.txtCodigo = codigo
        frmDesarrolloModificarMaterial.txtXHoja = xhoja
        frmDesarrolloModificarMaterial.txtXPieza = xpieza
        frmDesarrolloModificarMaterial.txtYHoja = yhoja
        frmDesarrolloModificarMaterial.txtYPieza = ypieza
        frmDesarrolloModificarMaterial.txtDetalle = Me.ListView1.selectedItem.Tag
        frmDesarrolloModificarMaterial.Show 1
    End If
End Sub

Private Sub ListView1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim li As ListItem
    Set li = Me.ListView1.HitTest(x, y)
    If li Is Nothing Then
        Me.ListView1.ToolTipText = ""
    Else
        Me.ListView1.ToolTipText = li.Tag
    End If

End Sub

Private Sub ListView2_DblClick()
    frmModificarMDO.nuevo_form = Me
    frmModificarMDO.lblCPP = Me.ListView2.selectedItem.ListSubItems(7)

    frmModificarMDO.lblSector = Me.ListView2.selectedItem.ListSubItems(4)
    frmModificarMDO.lblTarea = Me.ListView2.selectedItem.ListSubItems(1) & " - " & Me.ListView2.selectedItem.ListSubItems(6)
    frmModificarMDO.lblDescripcion = Me.ListView2.selectedItem.ListSubItems(8)
    frmModificarMDO.idDesMDO = Me.ListView2.selectedItem.ListSubItems(1)
    frmModificarMDO.txtCantOp = Me.ListView2.selectedItem.ListSubItems(2)
    frmModificarMDO.txtTiempo = Me.ListView2.selectedItem.ListSubItems(3)
    frmModificarMDO.txtDetalle = Me.ListView2.selectedItem.Tag
    frmModificarMDO.Show 1
    Me.calcular_totales_mdo

End Sub


Private Sub ListView2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim li As ListItem
    Set li = Me.ListView2.HitTest(x, y)
    If li Is Nothing Then
        Me.ListView2.ToolTipText = ""
    Else
        Me.ListView2.ToolTipText = li.Tag
    End If
End Sub

Private Sub txtAnchoPieza_Change()
    grabado = False
End Sub

Private Sub txtAnchoPieza_GotFocus()
    foco Me.txtAnchoPieza
End Sub


Private Sub txtAnchoPieza_LostFocus()
    Me.Command1.SetFocus
    Id = baseM.QueIdMaterial(Trim(Me.txtCodigoMaterial))
    If Id <> -1 Then
        verDetalleMateriales (Id)
    End If
End Sub

Private Sub txtAnchoPieza_Validate(Cancel As Boolean)
    If Not IsNumeric(Me.txtAnchoPieza) Then Cancel = True
End Sub

Private Sub txtAnchoTerm_Change()
    grabado = False
End Sub

Private Sub txtAnchoTerm_GotFocus()
    foco Me.txtAnchoTerm
End Sub

Private Sub txtAnchoTerm_LostFocus()
    Id = baseM.QueIdMaterial(Trim(Me.txtCodigoMaterial))
    If Id <> -1 Then
        verDetalleMateriales (Id)
    End If
End Sub

Private Sub txtAnchoTerm_Validate(Cancel As Boolean)
    If Not IsNumeric(Me.txtAnchoTerm) Then Cancel = True
End Sub

Private Sub txtCantidad_Change()
    grabado = False
End Sub

Private Sub txtCantidad_GotFocus()
    foco Me.txtCantidad
End Sub

Private Sub txtCantidad_Validate(Cancel As Boolean)
    If Not IsNumeric(Me.txtCantidad) Then Cancel = True
End Sub

Private Sub txtCantOp_Change()
    If Trim(Me.txtCodigoMDO) = Empty Or Trim(Me.txtCantOp) = Empty Or Trim(Me.txtTiempo) = Empty Then

        Me.btnAgregarMDO.Enabled = False
    Else
        Me.btnAgregarMDO.Enabled = True

    End If
    grabado = False
End Sub

Private Sub txtCantOp_GotFocus()
    foco txtCantOp
End Sub

Private Sub txtCantOp_Validate(Cancel As Boolean)
    If Not IsNumeric(Me.txtCantOp) Then Cancel = True
End Sub

Private Sub txtCodigoMaterial_Change()
    On Error GoTo err1
    Dim r As Recordset
    Dim descripcion As String
    Dim Kg As Double
    Dim m2ml As Double
    Dim rubro As String, Grupo As String

    If Trim(Me.txtCodigoMaterial) = Empty Then Exit Sub


    Id = baseM.QueIdMaterial(Trim(Me.txtCodigoMaterial))

    Set rss = conectar.RSFactory("select valor_unitario from materiales where id=" & Id)
    If Id <> -1 Then
        estado = rss!valor_unitario
        If estado > 0 Then
            If Id <> -1 Then
                Me.Command1.Enabled = True
                Me.Frame3.Enabled = True
                Me.Frame4.Enabled = True
                verDetalleMateriales (Id)
                Set r = RSFactory("select largo,ancho, id_unidad from materiales where id=" & Id)
                If Not r.EOF And Not r.BOF Then
                    idUnidad = r!id_Unidad
                    Largo = r!Largo
                    Ancho = r!Ancho
                Else
                    Exit Sub
                End If
                Set r = Nothing
                If idUnidad = 2 Then
                    'habilito los campos necesarios para M2
                    Me.txtAnchoTerm = Ancho
                    Me.txtLargoTerm = Largo
                    Me.txtAnchoPieza.Enabled = True
                    Me.txtAnchoTerm.Enabled = True
                    Me.txtLargoPieza.Enabled = True
                    Me.txtLargoTerm.Enabled = True
                    Me.txtScrap.Enabled = True
                ElseIf idUnidad = 4 Or idUnidad = 1 Then
                    'si es un elemento unitario o por Kg deshabilito todo
                    Me.txtAnchoPieza = 0
                    Me.txtLargoPieza = 0
                    Me.txtAnchoTerm = 0
                    Me.txtLargoTerm = 0
                    Me.txtAnchoPieza.Enabled = False
                    Me.txtAnchoTerm.Enabled = False
                    Me.txtLargoPieza.Enabled = False
                    Me.txtLargoTerm.Enabled = False
                    Me.txtScrap = 0
                    Me.txtScrap.Enabled = False
                ElseIf idUnidad = 3 Then
                    'habilito lo  necesitan los ml
                    Me.txtAnchoPieza.Enabled = False
                    Me.txtAnchoTerm.Enabled = False
                    Me.txtLargoPieza.Enabled = True
                    Me.txtLargoTerm.Enabled = True
                    Me.txtAnchoTerm = 0
                    Me.txtAnchoPieza = 0
                    Me.txtLargoTerm = Largo

                    Me.txtScrap.Enabled = True
                End If
            Else
                Me.Command1.Enabled = False
                Me.Frame3.Enabled = False
                Me.Frame4.Enabled = False
            End If
        Else
            MsgBox "El material está en estado inactivo" & Chr(10) & "Imposible utilizar en un desarrollo", vbCritical, "Error"
            Me.txtCodigoMaterial = Empty
        End If
    End If
    grabado = False
    Exit Sub
err1:
End Sub


Private Sub txtCodigoMaterial_GotFocus()
    foco Me.txtCodigoMaterial
    'id = baseS.buscar_pieza(Trim(Me.txtCodigoMaterial))

End Sub


Private Sub txtCodigoMDO_Change()
    If Trim(Me.txtCodigoMDO) = Empty Or Trim(Me.txtCantOp) = Empty Or Trim(Me.txtTiempo) = Empty Then

        Me.btnAgregarMDO.Enabled = False
    Else
        Me.btnAgregarMDO.Enabled = True

    End If
    grabado = False
End Sub

Private Sub txtCodigoMDO_GotFocus()
    foco Me.txtCodigoMDO
End Sub

Private Sub txtCodigoMDO_KeyPress(KeyAscii As Integer)
    Set base = New classNuevoElemento
    If KeyAscii = 13 Then
        base.ver_detalle_mdo CInt(Me.txtCodigoMDO), idcpp, cantxproc, mdoDescrip, Tarea, Sector, Valor
        Me.lblidMDO = Me.txtCodigoMDO
        Me.lblCPP = cantxproc
        Me.lblidCPP = idcpp
        lblTarea = Tarea
        lblDescripcion = mdoDescrip
        lblSector = Sector
        Me.lblValor = Valor


    End If
End Sub

Private Sub txtCodigoMDO_LostFocus()
    Set base = New classNuevoElemento
    If Not Trim(Me.txtCodigoMDO) = Empty Then
        base.ver_detalle_mdo CInt(Me.txtCodigoMDO), idcpp, cantxproc, mdoDescrip, Tarea, Sector, Valor
        Me.lblidMDO = Me.txtCodigoMDO
        Me.lblCPP = cantxproc
        Me.lblidCPP = idcpp
        lblTarea = Tarea
        lblDescripcion = mdoDescrip
        lblSector = Sector
        Me.lblValor = Valor

    End If
End Sub

Private Sub txtCodigoMDO_Validate(Cancel As Boolean)
    If Not IsNumeric(Me.txtCodigoMDO) Then Cancel = True
End Sub

Private Sub txtIdCliente_Change()
    On Error Resume Next
    grabado = False
    If Not Trim(Me.txtIdCliente) = Empty And IsNumeric(Me.txtIdCliente) Then
        'Me.lblCliente = DAOCliente.BuscarPorID(CLng(Me.txtIdCliente)).Razon

        Me.cboClientes.ListIndex = funciones.PosIndexCbo(CLng(Me.txtIdCliente), Me.cboClientes)
    End If
End Sub

Private Sub txtIdCliente_GotFocus()
    foco Me.txtIdCliente
End Sub

Private Sub txtLargoPieza_Change()
    grabado = False
End Sub

Private Sub txtLargoPieza_GotFocus()
    foco Me.txtLargoPieza
End Sub

Private Sub txtLargoPieza_LostFocus()
    Id = baseM.QueIdMaterial(Trim(Me.txtCodigoMaterial))
    If Id <> -1 Then
        verDetalleMateriales (Id)
    End If
End Sub

Private Sub txtLargoPieza_Validate(Cancel As Boolean)
    If Not IsNumeric(Me.txtLargoPieza) Then Cancel = True
End Sub

Private Sub txtLargoTerm_Change()
    grabado = False
End Sub
Private Sub txtLargoTerm_GotFocus()
    foco Me.txtLargoTerm
End Sub

Private Sub txtLargoTerm_LostFocus()
    Id = baseM.QueIdMaterial(Trim(Me.txtCodigoMaterial))
    If Id <> -1 Then
        verDetalleMateriales (Id)
    End If
End Sub

Private Sub txtLargoTerm_Validate(Cancel As Boolean)
    If Not IsNumeric(Me.txtLargoTerm) Then Cancel = True
End Sub

Private Sub txtNombreElemento_Change()
    grabado = False
End Sub

Private Sub txtNombreElemento_GotFocus()
    foco Me.txtNombreElemento
End Sub

Private Sub txtScrap_Change()
    grabado = False
End Sub

Private Sub txtScrap_GotFocus()
    foco Me.txtScrap
End Sub

Private Sub txtTiempo_Change()
    If Trim(Me.txtCodigoMDO) = Empty Or Trim(Me.txtCantOp) = Empty Or Trim(Me.txtTiempo) = Empty Then

        Me.btnAgregarMDO.Enabled = False
    Else
        Me.btnAgregarMDO.Enabled = True

    End If
    grabado = False
End Sub

Private Sub txtTiempo_GotFocus()
    foco txtTiempo

End Sub

Private Sub txtTiempo_Validate(Cancel As Boolean)
    If Not IsNumeric(Me.txtTiempo) Then Cancel = True
End Sub
