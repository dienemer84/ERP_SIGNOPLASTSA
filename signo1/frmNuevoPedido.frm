VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmVentasPedidoNuevo 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Nuevo Pedido..."
   ClientHeight    =   7200
   ClientLeft      =   2805
   ClientTop       =   2880
   ClientWidth     =   11760
   ClipControls    =   0   'False
   Icon            =   "frmNuevoPedido.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   11760
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.ComboBox cboClientes 
      Height          =   315
      Left            =   1275
      TabIndex        =   23
      Top             =   1005
      Width           =   7290
      _Version        =   786432
      _ExtentX        =   12859
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Appearance      =   5
      Text            =   "ComboBox1"
   End
   Begin VB.CommandButton btnGenerar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Generar"
      Height          =   375
      Left            =   7275
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   6570
      Width           =   1785
   End
   Begin VB.Frame frameTotal 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1335
      Left            =   9480
      TabIndex        =   12
      Top             =   5760
      Width           =   2175
      Begin VB.TextBox txtDto 
         Height          =   285
         Left            =   1200
         TabIndex        =   13
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblMoneda 
         BackColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   1200
         TabIndex        =   20
         Top             =   0
         Width           =   855
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         Caption         =   "Moneda"
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
         TabIndex        =   19
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         Caption         =   "Total"
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
         TabIndex        =   18
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblTotalTotal 
         BackColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   1200
         TabIndex        =   17
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label SubTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         Caption         =   "Sub Total"
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
         TabIndex        =   16
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblSubTotal 
         BackColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   1200
         TabIndex        =   15
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         Caption         =   "Descuento"
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
         TabIndex        =   14
         Top             =   720
         Width           =   975
      End
   End
   Begin VB.TextBox TxtNroPresupuesto 
      Height          =   285
      Left            =   1665
      TabIndex        =   5
      Top             =   150
      Width           =   1095
   End
   Begin VB.TextBox txtDescripcion 
      Height          =   285
      Left            =   1320
      TabIndex        =   4
      Top             =   600
      Width           =   7215
   End
   Begin VB.TextBox txtEntregaDias 
      Height          =   285
      Left            =   9600
      TabIndex        =   3
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Buscar"
      Default         =   -1  'True
      Height          =   375
      Left            =   2850
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      Caption         =   "Usar sólo esta fecha"
      Height          =   330
      Left            =   8925
      TabIndex        =   0
      Top             =   1290
      Width           =   1980
   End
   Begin MSComCtl2.DTPicker txtEntrega 
      Height          =   300
      Left            =   9600
      TabIndex        =   1
      Top             =   960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   393216
      Format          =   62324737
      CurrentDate     =   38861
   End
   Begin GridEX20.GridEX grilla 
      Height          =   3975
      Left            =   120
      TabIndex        =   11
      Top             =   1680
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   7011
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      RecordNavigatorString=   "Presupuesto:|de"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      GroupFooterStyle=   2
      PreviewColumn   =   "pz"
      PreviewRowLines =   2
      ColumnAutoResize=   -1  'True
      MultiSelect     =   -1  'True
      MethodHoldFields=   -1  'True
      ContScroll      =   -1  'True
      GroupByBoxVisible=   0   'False
      BackColorHeader =   16744576
      RowHeaders      =   -1  'True
      DataMode        =   99
      HeaderFontName  =   "Tahoma"
      FontName        =   "Tahoma"
      GridLines       =   2
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   7
      Column(1)       =   "frmNuevoPedido.frx":000C
      Column(2)       =   "frmNuevoPedido.frx":024C
      Column(3)       =   "frmNuevoPedido.frx":0340
      Column(4)       =   "frmNuevoPedido.frx":0430
      Column(5)       =   "frmNuevoPedido.frx":054C
      Column(6)       =   "frmNuevoPedido.frx":06AC
      Column(7)       =   "frmNuevoPedido.frx":07C4
      FormatStylesCount=   11
      FormatStyle(1)  =   "frmNuevoPedido.frx":08C8
      FormatStyle(2)  =   "frmNuevoPedido.frx":09F0
      FormatStyle(3)  =   "frmNuevoPedido.frx":0AA0
      FormatStyle(4)  =   "frmNuevoPedido.frx":0B54
      FormatStyle(5)  =   "frmNuevoPedido.frx":0C2C
      FormatStyle(6)  =   "frmNuevoPedido.frx":0D28
      FormatStyle(7)  =   "frmNuevoPedido.frx":0E08
      FormatStyle(8)  =   "frmNuevoPedido.frx":13EC
      FormatStyle(9)  =   "frmNuevoPedido.frx":19D0
      FormatStyle(10) =   "frmNuevoPedido.frx":1FC0
      FormatStyle(11) =   "frmNuevoPedido.frx":25AC
      ImageCount      =   0
      PrinterProperties=   "frmNuevoPedido.frx":2638
   End
   Begin XtremeSuiteControls.ComboBox cboListaOt 
      Height          =   315
      Left            =   1335
      TabIndex        =   24
      Top             =   6615
      Width           =   3765
      _Version        =   786432
      _ExtentX        =   6641
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Style           =   2
      Text            =   "ComboBox1"
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "Agregar a OT"
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
      Left            =   135
      TabIndex        =   22
      Top             =   6690
      Width           =   1155
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Nro. Presupuesto"
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
      TabIndex        =   10
      Top             =   180
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Cliente"
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
      Left            =   540
      TabIndex        =   9
      Top             =   1035
      Width           =   735
   End
   Begin VB.Label F 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Entrega"
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
      Left            =   8790
      TabIndex        =   8
      Top             =   630
      Width           =   855
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Descripción"
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
      Left            =   180
      TabIndex        =   7
      Top             =   630
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
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
      Left            =   8925
      TabIndex        =   6
      Top             =   990
      Width           =   615
   End
End
Attribute VB_Name = "frmVentasPedidoNuevo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim formLoading As Boolean
Dim presupuesto As clsPresupuesto
Dim detalle As clsPresupuestoDetalle
Dim tmpCliente As clsCliente
Dim Ot As OrdenTrabajo
Dim detalle_ot As DetalleOrdenTrabajo
Private Sub btnGenerar_Click()
    Dim solo As Boolean
    solo = Me.Check1.value
    Dim idpedido As Long
    If IsNumeric(Me.TxtNroPresupuesto) Then
        presu = Me.TxtNroPresupuesto
        If MsgBox("¿Está seguro de generar el pedido?", vbYesNo, "Confirmación") = vbYes Then
            idpedido = Me.cboListaOt.ItemData(Me.cboListaOt.ListIndex)
            Set Ot = DAOPresupuestos.CrearOT(presupuesto, idpedido, txtDescripcion.text)
            If IsSomething(Ot) Then
                MsgBox "OT creada con exito. Número " & Ot.IdFormateado
                Unload Me
            End If
        End If
    End If
End Sub
Private Sub cboClientes_Click()
    On Error GoTo err1
    If formLoading Then Exit Sub
    Set presupuesto.cliente = DAOCliente.BuscarPorID(Me.cboClientes.ItemData(Me.cboClientes.ListIndex))
    MostrarOTPendientes Me.cboClientes.ItemData(Me.cboClientes.ListIndex)
    Exit Sub
err1:
End Sub
Private Sub Command2_Click()
    If MsgBox("¿Está seguro de salir?", vbYesNo, "Confirmación") = vbYes Then
        Unload Me
    End If
End Sub
Private Sub BuscarPresu()
    If Trim(Me.TxtNroPresupuesto) <> Empy Then
        Set presupuesto = DAOPresupuestos.GetById(CLng(Me.TxtNroPresupuesto))
        If Not presupuesto Is Nothing Then
            If presupuesto.EstadoPresupuesto = Enviado_ Then
                Set presupuesto.DetallePresupuesto = DAOPresupuestosDetalle.GetAllByPresupuesto(presupuesto)
                llenarLista
                Me.lblMoneda = presupuesto.moneda.NombreCorto
                Me.cboClientes.ListIndex = funciones.PosIndexCbo(presupuesto.cliente.Id, cboClientes)
                Me.txtDescripcion = presupuesto.detalle
                Me.txtEntregaDias = presupuesto.FechaEntrega & " días"
                Me.txtEntrega = Now + presupuesto.FechaEntrega
                Me.verTotal
                Me.txtDto = presupuesto.Descuento
            Else
                Select Case presupuesto.EstadoPresupuesto
                Case 1: MsgBox "Primero debería mandar el presupuesto al cliente.", vbCritical, "Error"
                Case 3: MsgBox "No puede procesar un prespuesto ya procesado por planeamiento.", vbCritical, "Error"
                Case 5: MsgBox "Error 5"
                Case 6: MsgBox "No puede procesar un prespuesto no terminado.", vbCritical, "Error"
                Case 8: MsgBox "No puede procesar un prespuesto desactivado.", vbCritical, "Error"
                End Select
                Exit Sub
            End If
        End If
    End If
End Sub
Private Sub llenarLista()
    Me.grilla.ItemCount = 0
    Me.grilla.ItemCount = presupuesto.DetallePresupuesto.count
End Sub



Private Sub Command4_Click()
    BuscarPresu
End Sub
Private Sub Form_Load()
    FormHelper.Customize Me
    GridEXHelper.CustomizeGrid Me.grilla, False, True
    Me.grilla.AllowDelete = True
    Me.grilla.ItemCount = 0    'pongo la grilla en 0
    formLoading = True
    DAOCliente.llenarComboXtremeSuite Me.cboClientes, False, False, False
    Me.txtEntrega = Now
    formLoading = False
    MostrarOTPendientes Me.cboClientes.ItemData(Me.cboClientes.ListIndex)
    Me.cboListaOt.ListIndex = 0

    ''Me.caption = caption & " (" & Name & ")"


End Sub
Private Sub MostrarOTPendientes(idCliente)

    Dim col As Collection
    Dim Ot As OrdenTrabajo

    Set col = DAOOrdenTrabajo.FindAll(DAOOrdenTrabajo.TABLA_PEDIDO & "." & DAOOrdenTrabajo.CAMPO_ESTADO & "=" & EstadoOrdenTrabajo.EstadoOT_Pendiente & " AND " & DAOOrdenTrabajo.CAMPO_CLIENTE_ID & " = " & idCliente)
    Me.cboListaOt.Clear

    Me.cboListaOt.AddItem "Nueva OT"
    Me.cboListaOt.ItemData(Me.cboListaOt.NewIndex) = -1
    For Each Ot In col
        Me.cboListaOt.AddItem Ot.Id & " - " & Ot.descripcion
        Me.cboListaOt.ItemData(Me.cboListaOt.NewIndex) = Ot.Id
    Next
    Me.cboListaOt.ListIndex = 0
End Sub

Private Sub grilla_UnboundDelete(ByVal RowIndex As Long, ByVal Bookmark As Variant)
    If RowIndex > 0 And presupuesto.DetallePresupuesto.count > 0 Then
        presupuesto.DetallePresupuesto.remove RowIndex
        verTotal
    End If
End Sub

Private Sub grilla_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If presupuesto Is Nothing Then Exit Sub
    If RowIndex > presupuesto.DetallePresupuesto.count Then Exit Sub
    Set detalle = presupuesto.DetallePresupuesto.item(RowIndex)
    With detalle
        Values(1) = .item
        Values(2) = .Cantidad
        Values(3) = .Detalles
        Values(4) = funciones.FormatearDecimales(.ValorManual)
        Values(5) = funciones.FormatearDecimales(.ValorManual * .Cantidad)
        Values(6) = .entrega
        Values(7) = .Pieza.nombre
    End With
End Sub
Public Function verTotal()
    Me.txtDto = presupuesto.Descuento
    Me.lblSubTotal = funciones.FormatearDecimales(presupuesto.Total(Manual))
    Me.lblTotalTotal = funciones.FormatearDecimales(presupuesto.TotalConDescuento(Manual))
End Function

Private Sub grilla_UnboundUpdate(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    presupuesto.DetallePresupuesto(RowIndex).Cantidad = Values(2)
    presupuesto.DetallePresupuesto(RowIndex).ValorManual = Values(4)
    presupuesto.DetallePresupuesto(RowIndex).Detalles = Values(3)
    presupuesto.DetallePresupuesto(RowIndex).entrega = Values(6)
    verTotal
End Sub

Private Sub txtDescripcion_Change()
'    presupuesto.detalle = UCase(Me.txtDescripcion)

End Sub
Private Sub txtDto_Change()
    On Error GoTo err1:
    presupuesto.Descuento = CDbl(Me.txtDto)
    verTotal
err1:

End Sub
Private Sub txtDto_Validate(Cancel As Boolean)
    funciones.ValidarTextBox Me.txtDto, Cancel
End Sub



