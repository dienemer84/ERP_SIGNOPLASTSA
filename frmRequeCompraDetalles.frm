VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmComprasRequeDetalles 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Detalle"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   1995
   ClientWidth     =   15765
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   15765
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15735
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Cancel          =   -1  'True
         Caption         =   "Salir"
         Height          =   375
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   3720
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Imprimir"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   3720
         Width           =   1095
      End
      Begin MSComctlLib.ListView lstReque 
         Height          =   3015
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   15495
         _ExtentX        =   27331
         _ExtentY        =   5318
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   11
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Rubro"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Grupo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Descripción"
            Object.Width           =   4586
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Espesor"
            Object.Width           =   1323
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Detalle"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Un"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Largo"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Ancho"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "cant"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Entrega"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Orden de trabajo"
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
         Left            =   12240
         TabIndex        =   8
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha"
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
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sector"
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
         Left            =   3480
         TabIndex        =   6
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblSector 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Sector"
         Height          =   255
         Left            =   4200
         TabIndex        =   4
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label lblFecha 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fecha"
         Height          =   255
         Left            =   840
         TabIndex        =   3
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label lblOt 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Orden de trabajo"
         Height          =   255
         Left            =   13800
         TabIndex        =   2
         Top             =   240
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmComprasRequeDetalles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As recordset
Dim classCompras As New classCompras
Dim idr As Long
Public Property Let idReque(nidr As Long)
idr = nidr
End Property

Private Sub Command2_Click()
If MsgBox("¿Está seguro de salir?", vbYesNo, "Confirmación") = vbYes Then
    Unload Me
End If

End Sub

Private Sub Form_Activate()
Me.lstReque.Refresh
End Sub

Private Sub Form_Load()
llenarLST
Me.Refresh
Me.lstReque.Refresh
End Sub

Private Sub llenarLST()
Dim X As ListItem
Set rs = conectar.RSFactory("select r.fechaCreado,r.idPedido,s.sector from sp.ComprasRequerimientos r inner join sp.sectores s on s.id=r.idSector where r.id=" & idr)
Me.lblFecha = Format(rs!FechaCreado, "DD/MM/YYYY")
Me.lblSector = rs!sector
ot = Format(rs!idpedido, "0000")
If ot = -1 Then
 Me.lblOt = "Stock"
Else
 Me.lblOt = id
End If
Set rs = conectar.RSFactory("select rd.cantidad,rd.fechaEntrega,rd.detalle,rd.y,rd.x,rd.idMaterial,m.espesor,m.codigo,m.descripcion,m.pesoxunidad, if( m.id_unidad=1,'Kg',if(m.id_unidad=2,'M2',if(m.id_unidad=3,'Ml','Un'))) as unidad,g.grupo,r.rubro from ComprasRequerimientosDetalles rd  inner join materiales m on m.id = rd.idMaterial inner join grupos g on g.id=m.id_grupo inner join rubros r on m.id_rubro=r.id where idReque=" & idr)
While Not rs.EOF
Set X = Me.lstReque.ListItems.Add(, , rs!codigo)
      
                X.SubItems(1) = rs!rubro
                X.SubItems(2) = rs!grupo
                X.SubItems(3) = rs!descripcion
                X.SubItems(4) = rs!espesor

                X.SubItems(6) = rs!unidad
                X.SubItems(5) = rs!detalle
                X.SubItems(7) = rs!cantidad
                X.SubItems(8) = rs!X
                X.SubItems(9) = rs!Y
                
                X.SubItems(10) = rs!fechaEntrega
                rs.MoveNext
Wend

End Sub

