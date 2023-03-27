VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmRemitosEntregados 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remitos aplicados..."
   ClientHeight    =   4365
   ClientLeft      =   7680
   ClientTop       =   8310
   ClientWidth     =   5085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   5085
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Command1"
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   5160
      Width           =   615
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   4920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemitosEntregados.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lstRemitos 
      Height          =   4335
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   7646
      Arrange         =   2
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label idPedidoEntrega 
      Caption         =   "Label1"
      Height          =   255
      Left            =   1920
      TabIndex        =   0
      Top             =   5160
      Width           =   1095
   End
End
Attribute VB_Name = "frmRemitosEntregados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim vorigen As Integer
Public Property Let Origen(Origen)
    vorigen = Origen
End Property

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    llenarLstRemitos
End Sub
Private Sub llenarLstRemitos()
    Me.lstRemitos.ListItems.Clear
    Dim rs As Recordset
    idp = CLng(Me.idPedidoEntrega)
    Set rs = conectar.RSFactory("select r.id,r.numero from entregas e inner join remitos r on e.remito=r.id where idPedido=" & idp & " and origen=" & vorigen & " group by remito")





    While Not rs.EOF
        'si el remito es -1, significa que no sale y queda en stock
        If rs!numero > 0 Then
            Set x = Me.lstRemitos.ListItems.Add(, , rs!numero, 1)
            x.Tag = rs!Id
        End If
        rs.MoveNext
    Wend



End Sub

Private Sub Form_Load()
    FormHelper.Customize Me
End Sub

Private Sub lstRemitos_DblClick()



    Dim frm As frmPlaneamientoRemitoVer
    Set frm = New frmPlaneamientoRemitoVer
    Set frm.Remito = DAORemitoS.FindById(Me.lstRemitos.selectedItem.Tag)
    frm.MostrarInfoAdministracion = True
    frm.Show


End Sub
