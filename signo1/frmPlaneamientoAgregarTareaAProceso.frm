VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmPlaneamientoAgregarTareaAProceso 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Agregar Tarea a Proceso"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.PushButton btnGuardar 
      Height          =   435
      Left            =   1755
      TabIndex        =   6
      Top             =   2895
      Width           =   1605
      _Version        =   786432
      _ExtentX        =   2831
      _ExtentY        =   767
      _StockProps     =   79
      Caption         =   "Guardar"
      UseVisualStyle  =   -1  'True
   End
   Begin VB.ComboBox cboTareas 
      Height          =   315
      Left            =   1755
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1335
      Width           =   3450
   End
   Begin VB.TextBox txtObservaciones 
      Height          =   885
      Left            =   1755
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1830
      Width           =   3450
   End
   Begin VB.ComboBox cboSectores 
      Height          =   315
      Left            =   1755
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   870
      Width           =   3450
   End
   Begin XtremeSuiteControls.PushButton btnCancelar 
      Height          =   435
      Left            =   3600
      TabIndex        =   7
      Top             =   2895
      Width           =   1605
      _Version        =   786432
      _ExtentX        =   2831
      _ExtentY        =   767
      _StockProps     =   79
      Caption         =   "Cancelar"
      UseVisualStyle  =   -1  'True
   End
   Begin VB.Label lblPieza 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sector"
      Height          =   195
      Left            =   1770
      TabIndex        =   11
      Top             =   510
      Width           =   465
   End
   Begin VB.Label lblPiezaaaa 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   1200
      TabIndex        =   10
      Top             =   495
      Width           =   480
   End
   Begin VB.Label lblOrdenTrabajo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sector"
      Height          =   195
      Left            =   1770
      TabIndex        =   9
      Top             =   150
      Width           =   465
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Orden de Trabajo"
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
      Left            =   180
      TabIndex        =   8
      Top             =   150
      Width           =   1500
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Height          =   195
      Left            =   1155
      TabIndex        =   5
      Top             =   1365
      Width           =   510
   End
   Begin VB.Label lblObservaciones 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Observaciones"
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
      Left            =   405
      TabIndex        =   2
      Top             =   1845
      Width           =   1275
   End
   Begin VB.Label lblSector 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Height          =   195
      Left            =   1095
      TabIndex        =   1
      Top             =   900
      Width           =   570
   End
End
Attribute VB_Name = "frmPlaneamientoAgregarTareaAProceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private vIdDetallePedido As Long
Private sectores As Collection
Private tmpSector As clsSector
Public PIEZA_ID As Long
Public pedido_id As Long
Private Pieza As Pieza
Public idDetallePedidoConjunto As Long

Public Property Let idDetallePedido(nvalue As Long)
    vIdDetallePedido = nvalue
End Property

Private Sub btnCancelar_Click()
    Unload Me
End Sub

Private Sub btnGuardar_Click()
    If Me.cboTareas.ListIndex = -1 Then Exit Sub

    Dim q As String
    '    Dim conjunto As Long
    '
    '    q = "SELECT conjunto FROM stock WHERE id = " & PIEZA_ID
    '    Set rs = conectar.RSFactory(q)
    '    If Not rs.EOF Then
    '        conjunto = rs!conjunto
    '        conjunto = conjunto + 1   'lo cambie de *-1
    '    End If


    q = "insert into PlaneamientoTiemposProcesos" _
      & " (idPedido, idPieza, idDetallePedido, idDetallePedidoConj, codigoTarea, estado, conjunto, duracion, agregado, observacion_agregado)" _
      & " Values" _
      & " ({idPedido}, {idPieza}, {idDetallePedido}, {idDetallePedidoConj}, {codigoTarea}, 0, {conjunto}, 0, 1, {observacion_agregado})"

    q = Replace$(q, "{idPieza}", PIEZA_ID)
    q = Replace$(q, "{idPedido}", Me.pedido_id)

    '    If pieza.EsConjunto Then
    '        Dim dtoConjunto As DetalleOTConjuntoDTO
    '        Set dtoConjunto = DAODetalleOrdenTrabajo.FindConjuntoById(vIdDetallePedido)
    '        If IsSomething(dtoConjunto) Then
    '            q = Replace$(q, "{idDetallePedido}", dtoConjunto.idDetallePedido)
    '            q = Replace$(q, "{idDetallePedidoConj}", dtoConjunto.Id)
    '        Else
    '            q = Replace$(q, "{idDetallePedido}", 0)
    '            q = Replace$(q, "{idDetallePedidoConj}", vIdDetallePedido)
    '        End If
    '    Else
    '        q = Replace$(q, "{idDetallePedido}", vIdDetallePedido)
    '        q = Replace$(q, "{idDetallePedidoConj}", 0)
    '    End If

    q = Replace$(q, "{idDetallePedido}", vIdDetallePedido)
    q = Replace$(q, "{idDetallePedidoConj}", idDetallePedidoConjunto)

    q = Replace$(q, "{codigoTarea}", Me.cboTareas.ItemData(Me.cboTareas.ListIndex))
    q = Replace$(q, "{conjunto}", CInt(Pieza.EsConjunto) * -1)
    q = Replace$(q, "{observacion_agregado}", conectar.Escape(Me.txtObservaciones.text))

    TareaAgregada = conectar.execute(q)
    If TareaAgregada Then
        Unload Me
    Else
        MsgBox "Ocurrió un error al agregar la tarea.", vbCritical
    End If
End Sub

Private Sub cboSectores_Click()
    If Me.cboSectores.ListIndex = -1 Then Exit Sub
    FillComboBox Me.cboTareas, DAOTareas.FindAll(DAOTareas.TABLA_TAREA & "." & DAOTareas.CAMPO_ID_SECTOR & " = " & Me.cboSectores.ItemData(Me.cboSectores.ListIndex)), "Tarea", "Id", True
End Sub


Private Sub Form_Load()
    TareaAgregada = False

    Me.lblOrdenTrabajo.caption = pedido_id
    Set Pieza = DAOPieza.FindById(PIEZA_ID, FL_0, False, False)
    Me.lblPieza = Pieza.nombre

    FormHelper.Customize Me
    If DAOUsuarios.GetById(funciones.getUser).Empleado Is Nothing Then
        Set sectores = DAOSectores.GetAll()
    Else
        Set sectores = DAOSectores.GetByIdEmpleado(DAOUsuarios.GetById(funciones.getUser).Empleado.Id)
    End If

    'FillComboBox Me.cboSectores, sectores, "Sector", "id", True
    DAOSectores.LlenarCombo Me.cboSectores, sectores

End Sub

