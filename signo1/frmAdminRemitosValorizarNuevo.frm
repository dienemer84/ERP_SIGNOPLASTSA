VERSION 5.00
Begin VB.Form frmAdminRemitosValorizarNuevo 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ingrese nuevo valor..."
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ Remito ]"
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
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5535
      Begin VB.CommandButton Command2 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   2880
         TabIndex        =   10
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Aceptar"
         Default         =   -1  'True
         Height          =   255
         Left            =   1440
         TabIndex        =   9
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox txtNuevoValor 
         Height          =   285
         Left            =   1440
         TabIndex        =   5
         Top             =   1200
         Width           =   3975
      End
      Begin VB.Label lblValorActual 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   1440
         TabIndex        =   8
         Top             =   840
         Width           =   3975
      End
      Begin VB.Label lblCantidad 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   1440
         TabIndex        =   7
         Top             =   600
         Width           =   3975
      End
      Begin VB.Label lblDetalle 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   1440
         TabIndex        =   6
         Top             =   360
         Width           =   3975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nuevo Valor "
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
         TabIndex        =   4
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Valor Actual "
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
         TabIndex        =   3
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cantidad "
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
         TabIndex        =   2
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Detalle "
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
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmAdminRemitosValorizarNuevo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs2 As Recordset
Dim rs As Recordset
Dim claseP As New classPlaneamiento
Dim vRemito As Long
Dim vIdEntrega As Long
Public Property Let remito(nRemito As Long)
    vRemito = nRemito
End Property
Public Property Let idEntrega(nIdEntrega As Long)
    vIdEntrega = nIdEntrega
End Property
Private Sub Command1_Click()
    frmPlaneamientoRemitosDetalle.lstRemito.selectedItem.ListSubItems(4).text = funciones.FormatearDecimales(CDbl(Me.txtNuevoValor), 2)
    Unload Me
End Sub
Private Sub Command2_Click()
    Unload Me
End Sub
Private Sub Form_Load()
FormHelper.Customize Me
    Me.Frame1.caption = "[ Remito " & vRemito & " ]"
    strsql = "select id,idPedido,idDetallePedido,cantidad,remito,origen,valor,concepto from entregas where id=" & vIdEntrega
    Set rs = conectar.RSFactory(strsql)
    If Not rs.EOF And Not rs.BOF Then
        ide = rs!id
        idpedido = rs!idpedido
        cantidad = rs!cantidad
        remo = rs!remito
        origen = rs!origen
        valor = funciones.FormatearDecimales(rs!valor, 2)
        concepto = rs!concepto
        Me.txtNuevoValor = valor
        idDetallePedido = rs!idDetallePedido
        If idpedido = -1 Then
            detalle = concepto
        Else
            Set rs2 = conectar.RSFactory("select s.detalle from detalles_pedidos dp inner join stock s on dp.idPieza=s.id where dp.id=" & idDetallePedido)
            If Not rs2.EOF And Not rs2.BOF Then
                detalle = rs2!detalle
            Else
                Exit Sub
            End If
        End If
    Else
        Exit Sub
    End If
    Me.lblDetalle = detalle
    Me.lblValorActual = rs!valor
    Me.lblCantidad = rs!cantidad
End Sub
