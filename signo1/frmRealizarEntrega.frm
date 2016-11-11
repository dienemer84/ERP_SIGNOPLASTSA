VERSION 5.00
Begin VB.Form frmPlaneamientoRealizarEntrega 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Realizar entrega"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5625
   ClipControls    =   0   'False
   Icon            =   "frmRealizarEntrega.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   5625
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Envíar a Stock.."
      Height          =   450
      Left            =   120
      TabIndex        =   18
      Top             =   2640
      Width           =   1635
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1200
      TabIndex        =   5
      Text            =   "0"
      Top             =   1920
      Width           =   4335
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   2280
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Realizar"
      Default         =   -1  'True
      Height          =   465
      Left            =   3930
      TabIndex        =   3
      Top             =   2640
      Width           =   1605
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
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
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Fabricados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   16
      Top             =   840
      Width           =   945
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Entregados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   15
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "A entregar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   1920
      Width           =   900
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Remito"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   2310
      Width           =   600
   End
   Begin VB.Label lblPieza 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Label6"
      Height          =   195
      Left            =   1200
      TabIndex        =   12
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblPedidos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Label7"
      Height          =   195
      Left            =   1200
      TabIndex        =   11
      Top             =   480
      Width           =   480
   End
   Begin VB.Label lblEntregados 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Label8"
      Height          =   195
      Left            =   1200
      TabIndex        =   10
      Top             =   1200
      Width           =   480
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Pedidos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Width           =   690
   End
   Begin VB.Label lblFabricados 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Label7"
      Height          =   195
      Left            =   1200
      TabIndex        =   8
      Top             =   840
      Width           =   480
   End
   Begin VB.Label lblDeStock 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Label8"
      Height          =   195
      Left            =   1200
      TabIndex        =   7
      Top             =   1560
      Width           =   480
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "De stock"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   780
   End
   Begin VB.Label lblIdPieza 
      Caption         =   "Label7"
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label lblItem 
      Caption         =   "Label1"
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label lblOT 
      Caption         =   "Label1"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   3840
      Width           =   1095
   End
End
Attribute VB_Name = "frmPlaneamientoRealizarEntrega"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public deta As DetalleOrdenTrabajo
Dim baseP As New classPlaneamiento
Dim strsql As String
Private remitoId As Long
Public TipoOrden As TipoOt


Private Sub Command1_Click()
    Dim aentregar As Double
    Dim idPieza As Long
    saldo = CLng(Me.lblPedidos) - CLng(Me.lblEntregados)
    Entregados = CLng(Me.lblEntregados)
    Fabricados = CLng(Me.lblFabricados)
    deStock = CLng(Me.lblDeStock)
    Dim Remito As Long
    Remito = remitoId
    Dim rs As Recordset
    Set rs = conectar.RSFactory("select count(id) as cr from entregas where remito=" & CLng(Remito))
    If rs!cr >= funciones.itemsPorRemito Then
        MsgBox "Imposible agregar al remito, pués el remito esta completo. " & Chr(10) & "Cree o utlice otro remito"
        Exit Sub
    End If

    idpedido = Me.lblOT
    pedidos = CInt(Me.lblPedidos)
    saldo_fabricados = CDbl(Me.lblPedidos) - CDbl(Me.lblEntregados)
    aentregar = CDbl(Me.Text1)
    idPieza = CLng(Me.lblIdPieza)    'id en detalles_pedidos
    'si entrego menos o igual que lo q tengio fabricado, proceso.
    If aentregar <= saldo_fabricados + deStock Then
        'si entrego menos o igual q lo que resta entregar, proceso.
        'este caso puiede suceder si tengo fabricado mas que lo q piden
        'If aEntregar <= saldo Then
        'rutina de entrega
        If Not IsNumeric(Me.Text2) Then
            MsgBox "Ingrese datos de Remito válidos", vbCritical, "Error"
        Else

            Dim mostrar_obs As Boolean
            mostrar_obs = (MsgBox("Incluír observaciones del detalle de la orden trabajo?", vbYesNo, "Consulta") = vbYes)

            Dim mostrar_header As String

            If MsgBox("Incluír información de la orden trabajo?", vbYesNo, "Consulta") = vbYes Then

                Dim rss As Recordset
                Set rss = conectar.RSFactory("select descripcion from pedidos where id=" & deta.OrdenTrabajo.id)


                mostrar_header = rss!descripcion
            End If

            If baseP.RealizarEntrega(1, Remito, aentregar, idPieza, CLng(idpedido), , , , deta, mostrar_obs, mostrar_header) Then
                Unload Me
            Else
                MsgBox "Se produjo algún error", vbCritical, "Error"
            End If
        End If

        'verifico si está completo el item
        Entregados = aentregar + Entregados
        If Entregados = pedidos Then
            'aumento stock
            saldo_stock = Fabricados - pedidos
        End If

        '   Else
        '       MsgBox "Está intentando entregar más de lo que debe. ", vbCritical, "Error"
        '   End If
    Else
        MsgBox "Está intentando entregar más de lo que tiene fabricado", vbCritical, "Error"
    End If
End Sub

Private Sub Command2_Click()
On Error GoTo err1
Dim Cant As Double

    Dim Disponibles As Double
    Disponibles = deta.CantidadFabricados - deta.CantidadEnviadasAStock
  
    Dim res As String

    Dim envio As Double: envio = Val(Me.Text1)
    
    
    
    
  If Disponibles - envio < 0 Then
        MsgBox "Cantidad Insuficiente Para enviar a Stock"
        Exit Sub
  End If
    
     DAODetalleOrdenTrabajo.EnviarAStock deta, envio
    




Exit Sub
err1:
MsgBox Err.Description, vbCritical, Err.Source

End Sub

Private Sub Form_Activate()
    Me.caption = "Realizar entrega [ OT Nº " & Me.lblOT & " - Item " & Me.lblItem & " ]"
End Sub

Private Sub Form_Load()
    FormHelper.Customize Me
    
    
    Me.Command1.Enabled = (TipoOrden = OT_Entrega Or Me.TipoOrden = OT_TRADICIONAL)
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
    If Not IsNumeric(Me.Text1) Then Cancel = True
End Sub

Private Sub Text2_DblClick()
    frmPlaneamientoRemitosListaProceso.mostrar = 0
    frmPlaneamientoRemitosListaProceso.Show 1
    If IsSomething(Selecciones.RemitoElegido) Then
        Me.Text2 = Selecciones.RemitoElegido.numero
        remitoId = Selecciones.RemitoElegido.id
    Else
        Me.Text2 = Empty
    End If
End Sub
