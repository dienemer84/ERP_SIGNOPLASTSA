VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmComprasOrdenesHistorial 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Historial"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   255
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2880
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ Historial ] "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin MSComctlLib.ListView lstHisotial 
         Height          =   2415
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   4260
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Fecha"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nota"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Usuario"
            Object.Width           =   2117
         EndProperty
      End
   End
End
Attribute VB_Name = "frmComprasOrdenesHistorial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim claseC As New classCompras
Dim vidOrden As Long
Dim rs As Recordset
Public Property Let idOrden(nidOrden As Long)
    vidOrden = nidOrden
End Property
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    FormHelper.Customize Me
    Dim x As ListItem
    Set rs = conectar.RSFactory("select f.nota,f.fecha,u.usuario from ComprasOrdenesHistorial f inner join usuarios u on f.idUsuario=u.id where idOrden=" & vidOrden)
    While Not rs.EOF
        Set x = Me.lstHisotial.ListItems.Add(, , Format(rs!FEcha, "dd-mm-yyyy"))
        x.SubItems(1) = rs!Nota
        x.SubItems(2) = rs!usuario
        rs.MoveNext
    Wend
End Sub



