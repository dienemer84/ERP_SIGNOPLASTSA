VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmComprasPeticionesLista 
   BackColor       =   &H00F0E1D1&
   Caption         =   "Peticiones de Oferta"
   ClientHeight    =   6195
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   9990
   Icon            =   "frmComprasPeticionesLista.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6195
   ScaleWidth      =   9990
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1230
      Left            =   45
      TabIndex        =   1
      Top             =   30
      Width           =   9840
      _Version        =   786432
      _ExtentX        =   17357
      _ExtentY        =   2170
      _StockProps     =   79
      Caption         =   "Filtros"
      UseVisualStyle  =   -1  'True
      Begin VB.ComboBox cboEstado 
         Height          =   315
         Left            =   2370
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   255
         Width           =   2220
      End
      Begin VB.TextBox txtReque 
         Height          =   285
         Left            =   1020
         TabIndex        =   3
         Top             =   285
         Width           =   525
      End
      Begin XtremeSuiteControls.PushButton cmdBuscar 
         Height          =   465
         Left            =   5070
         TabIndex        =   2
         Top             =   585
         Width           =   1260
         _Version        =   786432
         _ExtentX        =   2222
         _ExtentY        =   820
         _StockProps     =   79
         Caption         =   "Buscar"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton cmdLimpiarProveedor 
         Height          =   330
         Left            =   4635
         TabIndex        =   5
         Top             =   705
         Width           =   255
         _Version        =   786432
         _ExtentX        =   450
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "X"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cboProveedor 
         Height          =   315
         Left            =   1020
         TabIndex        =   9
         Top             =   690
         Width           =   3570
         _Version        =   786432
         _ExtentX        =   6297
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Text            =   "ComboBox1"
      End
      Begin MSComDlg.CommonDialog cmd 
         Left            =   7125
         Top             =   450
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00F0E1D1&
         Caption         =   "Estado"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1770
         TabIndex        =   8
         Top             =   315
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00F0E1D1&
         Caption         =   "Proveedor"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   180
         TabIndex        =   7
         Top             =   690
         Width           =   735
      End
      Begin VB.Label lblReque 
         AutoSize        =   -1  'True
         BackColor       =   &H00F0E1D1&
         Caption         =   "Nº reque"
         ForeColor       =   &H008B4215&
         Height          =   195
         Left            =   285
         TabIndex        =   6
         Top             =   315
         Width           =   630
      End
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   4845
      Left            =   15
      TabIndex        =   0
      Top             =   1320
      Width           =   9900
      _ExtentX        =   17463
      _ExtentY        =   8546
      Version         =   "2.0"
      PreviewRowIndent=   300
      AutomaticSort   =   -1  'True
      HoldSortSettings=   -1  'True
      DefaultGroupMode=   1
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      PreviewColumn   =   "materiales"
      PreviewRowLines =   1
      ColumnAutoResize=   -1  'True
      MethodHoldFields=   -1  'True
      BackColorInfoText=   16777215
      ForeColorHeader =   65280
      GroupByBoxInfoText=   "Arrastre una columna para agrupar"
      AllowEdit       =   0   'False
      BackColorGBBox  =   14068360
      BackColorHeader =   14068304
      DataMode        =   99
      HeaderFontName  =   "Tahoma"
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   10
      Column(1)       =   "frmComprasPeticionesLista.frx":000C
      Column(2)       =   "frmComprasPeticionesLista.frx":0154
      Column(3)       =   "frmComprasPeticionesLista.frx":0248
      Column(4)       =   "frmComprasPeticionesLista.frx":0370
      Column(5)       =   "frmComprasPeticionesLista.frx":044C
      Column(6)       =   "frmComprasPeticionesLista.frx":052C
      Column(7)       =   "frmComprasPeticionesLista.frx":0600
      Column(8)       =   "frmComprasPeticionesLista.frx":06E0
      Column(9)       =   "frmComprasPeticionesLista.frx":07B0
      Column(10)      =   "frmComprasPeticionesLista.frx":0880
      GroupCount      =   1
      Group(1)        =   "frmComprasPeticionesLista.frx":09A0
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmComprasPeticionesLista.frx":0A08
      FormatStyle(2)  =   "frmComprasPeticionesLista.frx":0B40
      FormatStyle(3)  =   "frmComprasPeticionesLista.frx":0BF0
      FormatStyle(4)  =   "frmComprasPeticionesLista.frx":0CA4
      FormatStyle(5)  =   "frmComprasPeticionesLista.frx":0D7C
      FormatStyle(6)  =   "frmComprasPeticionesLista.frx":0E34
      ImageCount      =   0
      PrinterProperties=   "frmComprasPeticionesLista.frx":0F14
   End
   Begin VB.Menu menu 
      Caption         =   "menu"
      Visible         =   0   'False
      Begin VB.Menu editar 
         Caption         =   "Editar..."
      End
      Begin VB.Menu mnuFinalizar 
         Caption         =   "Finalizar"
      End
      Begin VB.Menu mnuExportar 
         Caption         =   "Exportar a Excel"
      End
      Begin VB.Menu mnuCrearOrdenCompra 
         Caption         =   "Crear Orden de Compra"
      End
   End
End
Attribute VB_Name = "frmComprasPeticionesLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Implements ISuscriber

Private susc_id As String
Dim tmp As clsPeticionOferta
Dim peticiones As Collection
Dim po As clsPeticionOferta

Dim tmpColNombreMateriales As Collection
Dim tmpDetPetOferta As clsPeticionOfertaDetalle

Private Sub cmdBuscar_Click()
    CargarPeticiones
End Sub

Private Sub cmdLimpiarProveedor_Click()
    Me.cboProveedor.ListIndex = -1
End Sub

Private Sub editar_Click()
    Dim frm As frmComprasPeticionesOfertaNueva
    Set frm = New frmComprasPeticionesOfertaNueva
    A = Me.GridEX1.RowIndex(Me.GridEX1.row)
    frm.peticion = peticiones.item(A)
    frm.Show
End Sub
Private Sub Form_Load()
    Customize Me
    GridEXHelper.CustomizeGrid Me.GridEX1, True

    Me.GridEX1.ItemCount = 0

    susc_id = funciones.CreateGUID()

    'Channel.AgregarSuscriptor Me, EdicionDetallePeticionOferta
    'no es necesario

    Me.cboEstado.Clear

    Dim i As Long

    For i = 0 To UBound(estado_po, 1)
        Me.cboEstado.AddItem enumEstadoPO(i)
        Me.cboEstado.ItemData(Me.cboEstado.NewIndex) = i
    Next i
    Me.cboEstado.ListIndex = 0


    DAOProveedor.llenarComboXtremeSuite Me.cboProveedor
    Me.cboProveedor.ListIndex = -1



    CargarPeticiones
End Sub
Private Sub CargarPeticiones()
    Dim F As String

    If IsNumeric(Me.txtReque.text) Then
        F = F & " AND id_reque = " & Me.txtReque.text
    End If

    If Me.cboProveedor.ListIndex <> -1 Then
        F = F & " AND id_proveedor = " & Me.cboProveedor.ItemData(Me.cboProveedor.ListIndex)
    End If

    F = F & " AND estado = " & Me.cboEstado.ItemData(Me.cboEstado.ListIndex)


    Set peticiones = DAOPeticionOferta.GetAll(F)
    Me.GridEX1.ItemCount = peticiones.count
    Me.GridEX1.RefreshGroups True

    GridEXHelper.AutoSizeColumns Me.GridEX1
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Me.GridEX1.Width = Me.ScaleWidth - 100
    Me.GridEX1.Height = Me.ScaleHeight - 1350
    GroupBox1.Width = Me.ScaleWidth - 100
End Sub



Private Sub Form_Unload(Cancel As Integer)
    Channel.RemoverSuscripcionTotal Me
End Sub




Private Sub GridEX1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim rectmp As clsPeticionOferta
    Dim gr
    If Button = 2 Then
        Dim r As Recordset
        Dim est As EstadoPO
        gr = Me.GridEX1.RowIndex(Me.GridEX1.row)
        If gr = 0 Then Exit Sub
        Set rectmp = peticiones.item(gr)
        Me.editar.Enabled = (rectmp.estado = EstadoPO.Pendiente_)
        Me.mnuFinalizar.Enabled = (rectmp.estado = EstadoPO.Pendiente_)
        Me.mnuCrearOrdenCompra.Enabled = (rectmp.estado = EstadoPO.Finalizado_)
        Me.PopupMenu menu
    End If
End Sub

Private Sub GridEX1_SelectionChange()
    If Me.GridEX1.RowIndex(Me.GridEX1.row) > 0 Then
        Set po = peticiones(Me.GridEX1.RowIndex(Me.GridEX1.row))
    Else
        Set po = Nothing
    End If
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Set tmp = peticiones.item(RowIndex)
    With tmp
        Values(1) = .idReque
        Values(2) = .numero
        Values(3) = .idReque
        Values(4) = .FechaEmision
        Values(5) = .FechaSolicitada
        Values(6) = .Proveedor.RazonSocial
        Values(7) = .usuarioCreador.usuario
        Values(8) = enums.enumEstadoPO(.estado)
        Values(9) = .moneda.NombreCorto

        Set tmpColNombreMateriales = New Collection
        For Each tmpDetPetOferta In .detalle
            tmpColNombreMateriales.Add tmpDetPetOferta.DetalleReque.Material.descripcion
        Next tmpDetPetOferta

        Values(10) = funciones.JoinCollectionValues(tmpColNombreMateriales, ", ")
    End With

End Sub

Private Property Get ISuscriber_id() As String
    ISuscriber_id = susc_id
End Property

Private Function ISuscriber_Notificarse(EVENTO As clsEventoObserver) As Variant
    CargarPeticiones
End Function

Private Sub mnuCrearOrdenCompra_Click()
    If Not po Is Nothing Then
        If MsgBox("¿Desea crear la orden de compra?", vbQuestion + vbYesNo) = vbYes Then
            If CrearOrdenCompra(po) Then
                Me.GridEX1.RefreshRowIndex Me.GridEX1.RowIndex(Me.GridEX1.row)
            Else
                MsgBox "Hubo un error", vbCritical
            End If
        End If
    End If
End Sub

Private Sub mnuExportar_Click()
    DAOPeticionOferta.ExportarExcel po.numero, Me.cmd
End Sub

Private Sub mnuFinalizar_Click()
    If Not po Is Nothing Then
        po.detalle = DAOPeticionOfertaDetalle.FindAll(po.numero)
        If po.IsValid() Then
            If MsgBox("¿Desea establecer la PO como finalizada?", vbQuestion + vbYesNo) = vbYes Then
                If DAOPeticionOferta.CambiarEstado(po, EstadoPO.Finalizado_) Then
                    Me.GridEX1.RefreshRowIndex Me.GridEX1.RowIndex(Me.GridEX1.row)
                Else
                    MsgBox "Hubo un error", vbCritical
                End If
            End If
        Else
            MsgBox "Los detalles de la PO tienen cantidad o precio igual a cero. Verifique.", vbOKOnly + vbInformation, "Error"
        End If
    End If
End Sub

Private Sub PushButton1_Click()

End Sub
