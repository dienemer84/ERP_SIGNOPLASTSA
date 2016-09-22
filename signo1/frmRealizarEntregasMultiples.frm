VERSION 5.00
Begin VB.Form frmPlaneamientoRealizarEntregaMultiple 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Realizar entrega"
   ClientHeight    =   1035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4335
   ClipControls    =   0   'False
   Icon            =   "frmRealizarEntregasMultiples.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1035
   ScaleWidth      =   4335
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "A Stock.."
      Height          =   330
      Left            =   3060
      TabIndex        =   5
      Top             =   540
      Width           =   1155
   End
   Begin VB.TextBox txtNroRto 
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   4
      ToolTipText     =   "Doble click para seleccionar remito"
      Top             =   120
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   330
      Left            =   1620
      TabIndex        =   3
      Top             =   540
      Width           =   1275
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Remito Nro"
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
      Left            =   210
      TabIndex        =   6
      Top             =   135
      Width           =   1095
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
Attribute VB_Name = "frmPlaneamientoRealizarEntregaMultiple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public idP As Long
Dim claseP As New classPlaneamiento
Dim vec()
Private remitoId As Long
Public TipoOrden() As TipoOt

Public Function vector(nvec() As Long)
    'Erase vec
    ReDim vec(UBound(nvec))
    For i = 0 To UBound(nvec)
        vec(i) = nvec(i)
    Next i
End Function

Private Sub Command1_Click()
    error2 = False

    'idp = CLng(frmEntregas.lblIdOT)
    Dim rs As Recordset

    Dim mostrar_obs As Boolean
    mostrar_obs = (MsgBox("Incluír observaciones del detalle de la orden trabajo?", vbYesNo, "Consulta") = vbYes)

    Dim mostrar_header As String

    If MsgBox("Incluír información de la orden trabajo?", vbYesNo, "Consulta") = vbYes Then
        Dim Ot As OrdenTrabajo
        Set Ot = DAOOrdenTrabajo.FindById(idP)
        If IsSomething(Ot) Then
            mostrar_header = Ot.descripcion
        End If
    End If


    If Not claseP.RealizarEntrega(3, remitoId, , , idP, 1, vec, , , mostrar_obs, mostrar_header) Then    'modo 3 es entrega multiple
        MsgBox "Se produjo algun error", vbCritical, "Error"
    Else
        MsgBox "Entregado correctamente!", vbInformation, "Información"
        Unload Me
    End If


End Sub

Private Sub Command2_Click()
    error2 = False

    rto = -1

    Dim Ot As OrdenTrabajo
    Set Ot = DAOOrdenTrabajo.FindById(idP)
    Set Ot.Detalles = DAODetalleOrdenTrabajo.FindAllByOrdenTrabajo(Ot.id, True, True, True)

    If DAOOrdenTrabajo.Cerrar(Ot, True) Then
        MsgBox "El pedido " & Pedido.id & " se cerro correctamente.", vbInformation, "Información"
        Unload Me
        Unload frmEntrega

    End If
End Sub

Private Sub Command3_Click()
    If MsgBox("¿Está seguro de salir?", vbYesNo, "Confirmación") = vbYes Then
        Unload Me
    End If

End Sub




Private Sub Form_Load()

    FormHelper.Customize Me
    ver
End Sub

Private Sub txtNroRto_Change()
    ver
End Sub
Public Sub ver()
    
        Command1.Enabled = Trim(txtNroRto) <> Empty And TipoOrden = OT_Entrega
    
    
End Sub

Private Sub txtNroRto_DblClick()
    frmPlaneamientoRemitosListaProceso.mostrar = 0
    frmPlaneamientoRemitosListaProceso.Show 1
    If IsSomething(Selecciones.RemitoElegido) Then
        Me.txtNroRto = Selecciones.RemitoElegido.numero
        remitoId = Selecciones.RemitoElegido.id
    Else
        Me.txtNroRto = Empty
    End If
End Sub

