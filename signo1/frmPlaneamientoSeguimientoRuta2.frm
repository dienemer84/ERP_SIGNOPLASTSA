VERSION 5.00
Begin VB.Form frmPlaneamientoSeguimientoRuta2 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Seguimiento especial..."
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6945
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   2640
      TabIndex        =   15
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   1320
      TabIndex        =   14
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox txtLegajo 
      Height          =   285
      Left            =   1800
      TabIndex        =   12
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox txt2 
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Legajo Operativo"
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
      TabIndex        =   13
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label lblCantidad 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   1080
      TabIndex        =   11
      Top             =   1440
      Width           =   5655
   End
   Begin VB.Label lblDeta 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   1080
      TabIndex        =   10
      Top             =   1200
      Width           =   5655
   End
   Begin VB.Label lblOt 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   1080
      TabIndex        =   9
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tarea"
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
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cantidad"
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
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Detalle"
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
      TabIndex        =   6
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Item"
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
      TabIndex        =   5
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "O/T"
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
      TabIndex        =   4
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Código de tarea"
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
      Left            =   0
      TabIndex        =   3
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label lblTarea 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   1800
      Width           =   5655
   End
   Begin VB.Label lblItem 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   960
      Width           =   5655
   End
End
Attribute VB_Name = "frmPlaneamientoSeguimientoRuta2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim plan As New classPlaneamiento

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
FormHelper.Customize Me
End Sub

Private Sub txt2_GotFocus()
    foco Me.txt2

End Sub

Private Sub txt2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        llenarDatos Trim(Me.txt2)
        Me.txt2 = Empty
        Me.txt2.SetFocus
    End If
End Sub



Private Sub llenarDatos(Leido As String)
On Error GoTo err1
    Dim haa() As String
    'haa = Split(Leido, ".")
    ' 'area = CLng(haa(0))
''
  '  idDetallePedido = CLng(haa(1))
'
   ' idDEtallePedidoConj = CLng(haa(2))

    Dim rs As recordset
    Dim q As String
    q = " SELECT *,dp.item FROM stock s" _
        & " LEFT JOIN detalles_pedidos dp ON dp.idPieza = s.id " _
        & "LEFT JOIN PlaneamientoTiemposProcesos ptp ON ptp.idDetallePedido = dp.id " _
        & "LEFT JOIN tareas t ON t.id=ptp.codigoTarea " _
        & "LEFT JOIN sectores sec ON t.id_sector=sec.id " _
        & "Where ptp.id =  " & Leido
    
    'Set rs = conectar.RSFactory("select dp.item,dp.idPedido,dp.nota,dp.cantidad,s.detalle from detalles_pedidos dp inner join stock s on dp.idPieza=s.id where dp.id=" & idDetallePedido)
    Set rs = conectar.RSFactory(q)
    
    
    If Not rs.EOF And Not rs.BOF Then
        Me.lblDeta = rs!detalle & " (" & rs!Nota & ")"
        Me.lblItem = rs!Item
        Me.lblOt = Format(rs!idpedido, "0000")
        Me.lblCantidad = formatearDecimales(rs!Cantidad, 2)
        Me.lblTarea = rs!codigoTarea & "-" & rs!tarea & " (" & rs!Sector & ")"
    Else

        Exit Sub
    End If

    'Set rs = conectar.RSFactory("select concat(s.sector,' - ',t.tarea) as tarea from tareas t inner join sectores s on t.id_sector=s.id where t.id=" & tarea)
    'If Not rs.EOF And Not rs.BOF Then
        
    'Else

        Exit Sub
    
    Set rs = Nothing
Exit Sub

err1:

End Sub

