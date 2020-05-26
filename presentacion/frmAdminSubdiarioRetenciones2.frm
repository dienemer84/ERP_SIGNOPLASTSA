VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~3.OCX"
Begin VB.Form frmAdminSubdiarioRetenciones2 
   Caption         =   "Subdiario de retenciones"
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7500
   ScaleWidth      =   12240
   Begin GridEX20.GridEX gridSubdiario 
      Height          =   5145
      Left            =   60
      TabIndex        =   6
      Top             =   1935
      Width           =   12090
      _ExtentX        =   21325
      _ExtentY        =   9075
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      MethodHoldFields=   -1  'True
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   2
      Column(1)       =   "frmAdminSubdiarioRetenciones2.frx":0000
      Column(2)       =   "frmAdminSubdiarioRetenciones2.frx":00C8
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminSubdiarioRetenciones2.frx":016C
      FormatStyle(2)  =   "frmAdminSubdiarioRetenciones2.frx":02A4
      FormatStyle(3)  =   "frmAdminSubdiarioRetenciones2.frx":0354
      FormatStyle(4)  =   "frmAdminSubdiarioRetenciones2.frx":0408
      FormatStyle(5)  =   "frmAdminSubdiarioRetenciones2.frx":04E0
      FormatStyle(6)  =   "frmAdminSubdiarioRetenciones2.frx":0598
      ImageCount      =   0
      PrinterProperties=   "frmAdminSubdiarioRetenciones2.frx":0678
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1890
      Left            =   45
      TabIndex        =   0
      Top             =   15
      Width           =   12165
      _Version        =   786432
      _ExtentX        =   21458
      _ExtentY        =   3334
      _StockProps     =   79
      Caption         =   "Parámetros de búsqueda"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.PushButton cmdGenerar 
         Height          =   360
         Left            =   2940
         TabIndex        =   5
         Top             =   660
         Width           =   1125
         _Version        =   786432
         _ExtentX        =   1984
         _ExtentY        =   635
         _StockProps     =   79
         Caption         =   "Generar"
         UseVisualStyle  =   -1  'True
      End
      Begin MSComCtl2.DTPicker DTHasta 
         Height          =   255
         Left            =   1065
         TabIndex        =   1
         Top             =   900
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         _Version        =   393216
         Format          =   16646145
         CurrentDate     =   39660
      End
      Begin MSComCtl2.DTPicker DTDesde 
         Height          =   255
         Left            =   1065
         TabIndex        =   2
         Top             =   540
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         _Version        =   393216
         Format          =   16646145
         CurrentDate     =   39660
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Desde"
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
         Left            =   345
         TabIndex        =   4
         Top             =   540
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Hasta"
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
         Left            =   345
         TabIndex        =   3
         Top             =   900
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmAdminSubdiarioRetenciones2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public colret As Collection

Private Sub Form_Load()
    Customize Me
    ListaRetenciones

    GridEXHelper.CustomizeGrid Me.gridSubdiario, False, False
    ArmarCol
End Sub

Private Sub Form_Resize()
    Me.GroupBox1.Width = Me.ScaleWidth - 100
    Me.gridSubdiario.Width = Me.GroupBox1.Width

End Sub

Private Sub ListaRetenciones()
    Set colret = DAORetenciones.FindAll
End Sub

Private Sub ArmarCol()
    Dim ret As Retencion
    Me.gridSubdiario.Columns.Clear
    Me.gridSubdiario.Columns.Add "Fecha", jgexText, jgexEditNone, "fecha"
    Me.gridSubdiario.Columns.Add "Razón Social", jgexText, jgexEditNone, "razon"
    Me.gridSubdiario.Columns.Add "CUIT", jgexText, jgexEditNone, "cuit"
    Me.gridSubdiario.Columns.Add "Nro. Retencion", jgexText, jgexEditNone, "nroRet"
    For Each ret In colret
        Me.gridSubdiario.Columns.Add ret.codigo, jgexText, jgexEditNone, ret.id
    Next


    Me.gridSubdiario.ColumnAutoResize = True
End Sub

