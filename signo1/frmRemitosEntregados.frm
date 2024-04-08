VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmRemitosEntregados 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remitos aplicados..."
   ClientHeight    =   5415
   ClientLeft      =   7680
   ClientTop       =   8310
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   6045
   Begin XtremeSuiteControls.PushButton btnFormato 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   4920
      Width           =   5775
      _Version        =   786432
      _ExtentX        =   10186
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Ver Listado"
      UseVisualStyle  =   -1  'True
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Command1"
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   7920
      Width           =   615
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   7680
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
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   7646
      Arrange         =   2
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
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
   Begin XtremeSuiteControls.Label lblContador 
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   4440
      Width           =   2535
      _Version        =   786432
      _ExtentX        =   4471
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Cantidad Remitos: 0"
      Alignment       =   1
   End
   Begin VB.Label idPedidoEntrega 
      Caption         =   "Label1"
      Height          =   255
      Left            =   2040
      TabIndex        =   0
      Top             =   7920
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

Private Sub btnFormato_Click()
    If Me.lstRemitos.View = lvwIcon Then
        Me.lstRemitos.View = lvwSmallIcon

        Me.btnFormato.caption = "Ver Iconos"
    ElseIf Me.lstRemitos.View = lvwSmallIcon Then
        Me.lstRemitos.View = lvwIcon
        Me.btnFormato.caption = "Ver Listado"
    End If
    
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    llenarLstRemitos
    Me.lblContador.caption = "Cantidad Remitos : " & Me.lstRemitos.ListItems.count
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
