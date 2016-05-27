VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAdminFacturasHistorial 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Historial Factura"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   4005
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   4005
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   255
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2520
      Width           =   975
   End
   Begin MSComctlLib.ListView lstHisotial 
      Height          =   2415
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3975
      _ExtentX        =   7011
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
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Usuario"
         Object.Width           =   2117
      EndProperty
   End
End
Attribute VB_Name = "frmAdminFacturasHistorial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim claseA As New classAdministracion
Dim vIdFactura As Long
Public Sub idFactura(nIdFactura As Long)
    vIdFactura = nIdFactura
End Sub
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    FormHelper.Customize Me
    Dim x As ListItem
    Set rs = conectar.RSFactory("select f.nota,f.fecha,u.usuario from AdminFacturasHistorial f inner join usuarios u on f.idUsuario=u.id where idFactura=" & vIdFactura)
    While Not rs.EOF
        Set x = Me.lstHisotial.ListItems.Add(, , Format(rs!FEcha, "dd-mm-yyyy"))
        x.SubItems(1) = rs!Nota
        x.SubItems(2) = rs!usuario
        rs.MoveNext
    Wend

End Sub
