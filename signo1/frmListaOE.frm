VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~3.OCX"
Begin VB.Form frmPlaneamientoOELista 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ordenes de entrega..."
   ClientHeight    =   6255
   ClientLeft      =   1800
   ClientTop       =   2985
   ClientWidth     =   11565
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   11565
   ShowInTaskbar   =   0   'False
   Begin GridEX20.GridEX gridEntregas 
      Height          =   4455
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   7858
      Version         =   "2.0"
      AllowRowSizing  =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      MethodHoldFields=   -1  'True
      AllowEdit       =   0   'False
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   7
      Column(1)       =   "frmListaOE.frx":0000
      Column(2)       =   "frmListaOE.frx":0110
      Column(3)       =   "frmListaOE.frx":01FC
      Column(4)       =   "frmListaOE.frx":02E8
      Column(5)       =   "frmListaOE.frx":03DC
      Column(6)       =   "frmListaOE.frx":04D0
      Column(7)       =   "frmListaOE.frx":05C0
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmListaOE.frx":06B4
      FormatStyle(2)  =   "frmListaOE.frx":07EC
      FormatStyle(3)  =   "frmListaOE.frx":089C
      FormatStyle(4)  =   "frmListaOE.frx":0950
      FormatStyle(5)  =   "frmListaOE.frx":0A28
      FormatStyle(6)  =   "frmListaOE.frx":0AE0
      ImageCount      =   0
      PrinterProperties=   "frmListaOE.frx":0BC0
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "Volver"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4680
      Width           =   1095
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1545
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11445
      _Version        =   786432
      _ExtentX        =   20188
      _ExtentY        =   2725
      _StockProps     =   79
      Caption         =   "Parámetros de búsqueda"
      UseVisualStyle  =   -1  'True
      Begin VB.TextBox txtNroRemito 
         Height          =   285
         Left            =   1425
         TabIndex        =   4
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   285
         Left            =   1425
         TabIndex        =   3
         Top             =   645
         Width           =   6855
      End
      Begin XtremeSuiteControls.ComboBox cboClientes 
         Height          =   315
         Left            =   4530
         TabIndex        =   5
         Top             =   240
         Width           =   6015
         _Version        =   786432
         _ExtentX        =   10610
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   285
         Left            =   10620
         TabIndex        =   6
         Top             =   225
         Width           =   495
         _Version        =   786432
         _ExtentX        =   873
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "X"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton cmdBuscar 
         Default         =   -1  'True
         Height          =   375
         Left            =   8445
         TabIndex        =   7
         Top             =   600
         Width           =   1095
         _Version        =   786432
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Buscar"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.DateTimePicker dtpDesde 
         Height          =   315
         Left            =   5730
         TabIndex        =   8
         Top             =   1065
         Width           =   1470
         _Version        =   786432
         _ExtentX        =   2593
         _ExtentY        =   556
         _StockProps     =   68
         CheckBox        =   -1  'True
         Format          =   1
      End
      Begin XtremeSuiteControls.DateTimePicker dtpHasta 
         Height          =   315
         Left            =   7905
         TabIndex        =   9
         Top             =   1065
         Width           =   1470
         _Version        =   786432
         _ExtentX        =   2593
         _ExtentY        =   556
         _StockProps     =   68
         CheckBox        =   -1  'True
         Format          =   1
      End
      Begin XtremeSuiteControls.ComboBox cboRangos 
         Height          =   315
         Left            =   1380
         TabIndex        =   10
         Top             =   1050
         Width           =   3645
         _Version        =   786432
         _ExtentX        =   6429
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Número"
         Height          =   255
         Left            =   270
         TabIndex        =   16
         Top             =   270
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   645
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         Height          =   255
         Left            =   3315
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
      Begin XtremeSuiteControls.Label Label6 
         Height          =   195
         Left            =   7335
         TabIndex        =   13
         Top             =   1125
         Width           =   420
         _Version        =   786432
         _ExtentX        =   741
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Hasta"
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label5 
         Height          =   195
         Left            =   5160
         TabIndex        =   12
         Top             =   1110
         Width           =   465
         _Version        =   786432
         _ExtentX        =   820
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Desde"
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label7 
         Height          =   195
         Left            =   795
         TabIndex        =   11
         Top             =   1110
         Width           =   480
         _Version        =   786432
         _ExtentX        =   847
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Rango"
         AutoSize        =   -1  'True
      End
   End
   Begin VB.Menu entrega 
      Caption         =   "entregas"
      Visible         =   0   'False
      Begin VB.Menu OENumero 
         Caption         =   "Numero"
         Enabled         =   0   'False
      End
      Begin VB.Menu vereditar 
         Caption         =   "vereditar"
      End
      Begin VB.Menu as234 
         Caption         =   "-"
      End
      Begin VB.Menu AprobarOE 
         Caption         =   "Aprobar"
      End
      Begin VB.Menu remitar 
         Caption         =   "Remitar..."
      End
      Begin VB.Menu cerrarOE 
         Caption         =   "Cerrar..."
      End
      Begin VB.Menu verHistorialOE 
         Caption         =   "Ver historial..."
      End
      Begin VB.Menu nada 
         Caption         =   "-"
      End
      Begin VB.Menu RtosEntregados 
         Caption         =   "Remitos entregados...."
      End
      Begin VB.Menu printOrder 
         Caption         =   "Imprimir..."
      End
   End
End
Attribute VB_Name = "frmPlaneamientoOELista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim tmpOe As OrdenDeEntrega
Dim ordenes As New Collection



Public Sub LlenarListaOE()

    Set ordenes = DAOOrdenDeEntrega.GetAll()
    Me.gridEntregas.ItemCount = ordenes.count
End Sub


Private Sub AprobarOE_Click()
    'Dim vidOe As Long
    'vidOe = CLng(Me.lstOE.selectedItem)
'    If MsgBox("¿Está seguro de aprobar la O/E?", vbYesNo, "Confirmación") = vbYes Then
'        If claseP.AprobarOrdenEntrega(vidOe) Then
'            MsgBox "Orden de Entrega aprobada con éxito!", vbInformation, "Información"
'        Else
'            MsgBox "Se produjo un error al aprobar la OE!", vbCritical, "Error"
'        End If
'    End If

End Sub

Private Sub cmdBuscar_Click()
LlenarListaOE
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    FormHelper.Customize Me
        GridEXHelper.CustomizeGrid Me.gridEntregas, True
        GridEXHelper.AutoSizeColumns Me.gridEntregas, True
        Me.gridEntregas.ItemCount = 0
End Sub



Private Sub lstOE_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    'If Me.lstOE.ListItems.count > 0 Then
        If Button = 2 Then
            'IDOE = Me.lstOE.selectedItem
            Me.OENumero.caption = "[ Nro. " & IDOE & " ]"
            Set rs = conectar.RSFactory("select estado from PedidosEntregas where id=" & IDOE)

            If rs!Estado = 1 Then    'pendiente
                Me.vereditar.caption = "Editar..."
                Me.remitar.Enabled = False
                Me.cerrarOE.Enabled = False
                Me.RtosEntregados.Enabled = False
                vereditarOE = 1
                Me.printOrder.Enabled = False
                Me.AprobarOE.Enabled = True
            ElseIf rs!Estado = 2 Then    'aprobado
                Me.remitar.Enabled = True
                Me.vereditar.Enabled = True
                Me.cerrarOE.Enabled = False
                Me.RtosEntregados.Enabled = False
                Me.vereditar.caption = "Ver..."
                vereditarOE = 3
                Me.printOrder.Enabled = True
                Me.AprobarOE.Enabled = False
            ElseIf rs!Estado = 4 Then    'entregada
                Me.remitar.Enabled = False
                Me.cerrarOE.Enabled = True
                Me.RtosEntregados.Enabled = False
                Me.vereditar.caption = "Ver..."
                vereditarOE = 3
                Me.printOrder.Enabled = False
                Me.AprobarOE.Enabled = False
            ElseIf rs!Estado = 3 Then    'finalizada
                Me.vereditar.caption = "Ver..."
                vereditarOE = 3
                Me.remitar.Enabled = False
                Me.cerrarOE.Enabled = False
                Me.RtosEntregados.Enabled = True
                Me.printOrder.Enabled = False
                Me.AprobarOE.Enabled = False
            End If

            If Not Permisos.planOEaprobaciones Then Me.AprobarOE.Enabled = False
            Me.PopupMenu entrega
        End If
    'End If

End Sub

Private Sub gridEntregas_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error Resume Next
    Set tmpOe = ordenes.item(RowIndex)
    With Values
        .value(1) = tmpOe.id
        .value(2) = tmpOe.FEcha
        .value(3) = tmpOe.cliente.razon
        .value(4) = tmpOe.referencia
        .value(5) = tmpOe.usuarioCreador.usuario
        .value(6) = tmpOe.usuarioAprobador.usuario
        .value(7) = enumEstadoOrdenEntrega(tmpOe.Estado)
    End With

End Sub

Private Sub printOrder_Click()
   ' If Me.lstOE.ListItems.count > 0 Then
        'claseP.imprimirOrdenEntrega (CLng(Me.lstOE.selectedItem))
  '  End If


End Sub

Private Sub remitar_Click()
   'If Me.lstOE.ListItems.count > 0 Then
        'frmRemitar.idPedidoEntrega = CLng(Me.lstOE.selectedItem)
'        frmRemitar.idPe = CLng(Me.lstOE.selectedItem)
'        frmRemitar.Frame1.caption = "[ Nro." & Me.lstOE.selectedItem & " ]"
'        frmRemitar.Show
   ' End If
End Sub


Private Sub RtosEntregados_Click()
  '  If Me.lstOE.ListItems.count > 0 Then
        frmRemitosEntregados.Origen = 2
        'frmRemitosEntregados.idPedidoEntrega = Me.lstOE.selectedItem
        'frmRemitosEntregados.caption = "Nro." & Me.lstOE.selectedItem
'        frmRemitosEntregados.Show
   ' End If
End Sub

Private Sub vereditar_Click()
    If vereditarOE = 3 Then    'ver (porque la oe esta finalziada)
     '   frmPlaneamientoOEVer.IDOE = CLng(Me.lstOE.selectedItem)
        frmPlaneamientoOEVer.Show
    ElseIf vereditarOE = 1 Then    'editar (porq la oe esta en proceso)
    '    frmPlaneamientoOEEditar.IDOE = CLng(Me.lstOE.selectedItem)
        frmPlaneamientoOEEditar.Show
    End If
End Sub
