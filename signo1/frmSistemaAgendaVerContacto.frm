VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmSistemaAgendaVerContacto 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Contácto"
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   10635
   ShowInTaskbar   =   0   'False
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3135
      Left            =   0
      TabIndex        =   3
      Top             =   2760
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   5530
      _Version        =   393216
      BackColor       =   16777215
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Lista de contáctos"
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "detalle"
         Caption         =   "Detalle"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "tel1"
         Caption         =   "Tel 1"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "tel2"
         Caption         =   "Tel  2"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "mail"
         Caption         =   "E-Mail"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "mas"
         Caption         =   "Más..."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         AllowSizing     =   0   'False
         RecordSelectors =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   2805.166
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1800
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1800
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1800
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2099.906
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ Datos de registro ]"
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
      TabIndex        =   2
      Top             =   0
      Width           =   10575
   End
   Begin MSComctlLib.ListView lstContactos 
      Height          =   3375
      Left            =   0
      TabIndex        =   1
      Top             =   6720
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   5953
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Detalle"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Teléfono 1"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Teléfono 2"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "E-Mail"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Más datos..."
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   9600
      TabIndex        =   0
      Top             =   6000
      Width           =   975
   End
End
Attribute VB_Name = "frmSistemaAgendaVerContacto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strsql As String
Dim rs As Recordset
Dim claseSP As New classAdministracion

Dim vIdContacto As Long

Public Property Let idContacto(nIdContacto As Long)
    vIdContacto = nIdContacto
End Property




Private Sub Command1_Click()
    Unload Me
End Sub

Public Sub llenarDatosContactos()
    Dim x As ListItem
    strsql = "select * from datos_agenda where id_agenda=" & vIdContacto
    Set rs = conectar.RSFactoryCliente(strsql)
    While Not rs.EOF
        Set x = Me.lstContactos.ListItems.Add(, , rs!detalle)
        x.SubItems(1) = rs!tel1
        x.SubItems(2) = rs!tel2
        x.SubItems(3) = rs!mail
        x.SubItems(4) = rs!mas
        rs.MoveNext
    Wend

    Set Me.DataGrid1.DataSource = rs
End Sub


Private Sub Form_Load()
    FormHelper.Customize Me
    llenarDatosContactos
End Sub

