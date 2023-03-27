VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmComprasPOComprar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Comprar"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   14565
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.PushButton PushButton1 
      Height          =   555
      Left            =   105
      TabIndex        =   3
      Top             =   6675
      Width           =   1560
      _Version        =   786432
      _ExtentX        =   2752
      _ExtentY        =   979
      _StockProps     =   79
      Caption         =   "Agregar a Compra"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.ComboBox cboPOs 
      Height          =   315
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      _Version        =   786432
      _ExtentX        =   11668
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Appearance      =   3
      Text            =   "ComboBox1"
   End
   Begin GridEX20.GridEX gridReq 
      Height          =   2955
      Left            =   105
      TabIndex        =   1
      Top             =   495
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   5212
      Version         =   "2.0"
      PreviewRowIndent=   200
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      PreviewColumn   =   6
      PreviewRowLines =   1
      ColumnAutoResize=   -1  'True
      MultiSelect     =   -1  'True
      MethodHoldFields=   -1  'True
      BackColorInfoText=   16777215
      AllowDelete     =   -1  'True
      BorderStyle     =   3
      GroupByBoxVisible=   0   'False
      BackColorHeader =   16761024
      RowHeaders      =   -1  'True
      DataMode        =   99
      BackColorBkg    =   16777215
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   7
      Column(1)       =   "frmComprasPOComprar.frx":0000
      Column(2)       =   "frmComprasPOComprar.frx":0120
      Column(3)       =   "frmComprasPOComprar.frx":0214
      Column(4)       =   "frmComprasPOComprar.frx":030C
      Column(5)       =   "frmComprasPOComprar.frx":0400
      Column(6)       =   "frmComprasPOComprar.frx":04F4
      Column(7)       =   "frmComprasPOComprar.frx":0610
      FormatStylesCount=   9
      FormatStyle(1)  =   "frmComprasPOComprar.frx":0724
      FormatStyle(2)  =   "frmComprasPOComprar.frx":085C
      FormatStyle(3)  =   "frmComprasPOComprar.frx":090C
      FormatStyle(4)  =   "frmComprasPOComprar.frx":09C0
      FormatStyle(5)  =   "frmComprasPOComprar.frx":0A98
      FormatStyle(6)  =   "frmComprasPOComprar.frx":0B50
      FormatStyle(7)  =   "frmComprasPOComprar.frx":0C30
      FormatStyle(8)  =   "frmComprasPOComprar.frx":0CC0
      FormatStyle(9)  =   "frmComprasPOComprar.frx":0D54
      ImageCount      =   0
      PrinterProperties=   "frmComprasPOComprar.frx":0DE4
   End
   Begin GridEX20.GridEX gridPO 
      Height          =   2955
      Left            =   120
      TabIndex        =   2
      Top             =   3660
      Width           =   7440
      _ExtentX        =   13123
      _ExtentY        =   5212
      Version         =   "2.0"
      PreviewRowIndent=   200
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      PreviewRowLines =   1
      ColumnAutoResize=   -1  'True
      MultiSelect     =   -1  'True
      MethodHoldFields=   -1  'True
      BackColorInfoText=   16777215
      AllowDelete     =   -1  'True
      BorderStyle     =   3
      GroupByBoxVisible=   0   'False
      BackColorHeader =   16761024
      RowHeaders      =   -1  'True
      DataMode        =   99
      BackColorBkg    =   16777215
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   5
      Column(1)       =   "frmComprasPOComprar.frx":0FBC
      Column(2)       =   "frmComprasPOComprar.frx":10A8
      Column(3)       =   "frmComprasPOComprar.frx":117C
      Column(4)       =   "frmComprasPOComprar.frx":1250
      Column(5)       =   "frmComprasPOComprar.frx":131C
      FormatStylesCount=   9
      FormatStyle(1)  =   "frmComprasPOComprar.frx":13EC
      FormatStyle(2)  =   "frmComprasPOComprar.frx":1524
      FormatStyle(3)  =   "frmComprasPOComprar.frx":15D4
      FormatStyle(4)  =   "frmComprasPOComprar.frx":1688
      FormatStyle(5)  =   "frmComprasPOComprar.frx":1760
      FormatStyle(6)  =   "frmComprasPOComprar.frx":1818
      FormatStyle(7)  =   "frmComprasPOComprar.frx":18F8
      FormatStyle(8)  =   "frmComprasPOComprar.frx":1988
      FormatStyle(9)  =   "frmComprasPOComprar.frx":1A1C
      ImageCount      =   0
      PrinterProperties=   "frmComprasPOComprar.frx":1AAC
   End
   Begin GridEX20.GridEX gridCompra 
      Height          =   6060
      Left            =   7695
      TabIndex        =   4
      Top             =   510
      Width           =   6690
      _ExtentX        =   11800
      _ExtentY        =   10689
      Version         =   "2.0"
      PreviewRowIndent=   200
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      PreviewRowLines =   1
      ColumnAutoResize=   -1  'True
      MultiSelect     =   -1  'True
      MethodHoldFields=   -1  'True
      BackColorInfoText=   16777215
      AllowDelete     =   -1  'True
      BorderStyle     =   3
      GroupByBoxVisible=   0   'False
      BackColorHeader =   16761024
      RowHeaders      =   -1  'True
      DataMode        =   99
      BackColorBkg    =   16777215
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   4
      Column(1)       =   "frmComprasPOComprar.frx":1C84
      Column(2)       =   "frmComprasPOComprar.frx":1D70
      Column(3)       =   "frmComprasPOComprar.frx":1E44
      Column(4)       =   "frmComprasPOComprar.frx":1F18
      FormatStylesCount=   9
      FormatStyle(1)  =   "frmComprasPOComprar.frx":1FE4
      FormatStyle(2)  =   "frmComprasPOComprar.frx":211C
      FormatStyle(3)  =   "frmComprasPOComprar.frx":21CC
      FormatStyle(4)  =   "frmComprasPOComprar.frx":2280
      FormatStyle(5)  =   "frmComprasPOComprar.frx":2358
      FormatStyle(6)  =   "frmComprasPOComprar.frx":2410
      FormatStyle(7)  =   "frmComprasPOComprar.frx":24F0
      FormatStyle(8)  =   "frmComprasPOComprar.frx":2580
      FormatStyle(9)  =   "frmComprasPOComprar.frx":2614
      ImageCount      =   0
      PrinterProperties=   "frmComprasPOComprar.frx":26A4
   End
End
Attribute VB_Name = "frmComprasPOComprar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim compra As New Collection
Dim reqs As Collection
Dim req As clsRequerimiento
Dim vDetalle As clsRequeMateriales
Dim pos As Collection
Dim pod As clsPeticionOfertaDetalle
Dim item As clsPeticionOfertaDetalle
Dim prov As clsProveedor


Private Sub cboPOs_Change()
    Set req = DAORequerimiento.FindById(cboPOs.ItemData(Me.cboPOs.ListIndex), True, , , True)
    Me.gridReq.ItemCount = req.Materiales.count
End Sub

Private Sub Form_Load()
    Customize Me
    GridEXHelper.CustomizeGrid Me.gridReq, False, False
    GridEXHelper.CustomizeGrid Me.gridPO, False, False
    GridEXHelper.CustomizeGrid Me.gridCompra, False, False

    gridReq.ItemCount = 0
    gridCompra.ItemCount = 0
    'Traigo Reques que tienen PO's Con items pendientes
    Set reqs = DAORequerimiento.FindAll(, True, , , True)

    For Each req In reqs
        Me.cboPOs.AddItem req.Id & " | " & req.Sector.Sector & " | " & req.StringDestino
        Me.cboPOs.ItemData(Me.cboPOs.NewIndex) = req.Id
    Next req






    If reqs.count > 0 Then
        Me.cboPOs.ListIndex = 0
        cboPOs_Change
    End If

End Sub



Private Sub gridCompra_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Set item = compra.item(RowIndex)

    Values(1) = item.POid
    Values(2) = DAOProveedor.FindById(pod.ProveedorId).RazonSocial
    Values(3) = item.Total
    Values(4) = item.Valor
End Sub

Private Sub gridPO_SelectionChange()
    Set pod = pos(Me.gridPO.RowIndex(Me.gridPO.row))

End Sub

Private Sub gridPO_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Set pod = pos.item(RowIndex)

    Values(1) = pod.POid
    Values(2) = DAOProveedor.FindById(pod.ProveedorId).RazonSocial
    Values(3) = pod.Total
    Values(4) = pod.Valor
    If Not BuscarEnColeccion(compra, CStr(pod.Id)) Then
        Values(5) = enums.EstadosPeticionOfertaDetalle.item(CStr(pod.estado))
    Else
        Values(5) = enums.EstadosPeticionOfertaDetalle.item(CStr(EstadoPeticionOfertaDetalle.EPOD_comprado))
    End If



End Sub

Private Sub gridReq_SelectionChange()
    On Error Resume Next
    Set vDetalle = req.Materiales.item(Me.gridReq.RowIndex(Me.gridReq.row))




    Set pos = DAOPeticionOfertaDetalle.FindAll(, "id_detalle_reque=" & vDetalle.Id, False)
    Me.gridPO.ItemCount = 0
    Me.gridPO.ItemCount = pos.count
    '
    '    'armo las columnas nuevas
    '    gridPO.Columns.Clear
    '    For Each pod In pos
    '        Set prov = DAOProveedor.FindById(pod.ProveedorId)
    '        gridPO.Columns.Add "PO " & pod.POid & " | " & prov.RazonSocial, jgexText
    '     Next pod
    '    GridEXHelper.AutoSizeColumns Me.gridPO, True
End Sub

Private Sub gridReq_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If req.Materiales.count > 0 Then
        Set vDetalle = req.Materiales.item(RowIndex)
        With vDetalle
            Values(1) = .Cantidad
            Values(2) = .observaciones
            Values(3) = .Material.codigo
            Values(4) = .Material.Grupo.rubros.rubro
            Values(5) = .Material.Grupo.Grupo
            Values(6) = "Material: " & .Material.descripcion & " | Medidas: " & funciones.JoinCollectionValues(.Material.Atributos, ", ")
            Values(7) = enums.enumEstadoRequeCompra(.estado)
        End With
    End If
End Sub

Private Sub PushButton1_Click()

    If Not funciones.BuscarEnColeccion(compra, CStr(pod.Id)) Then compra.Add pod, CStr(pod.Id)
    Me.gridCompra.ItemCount = compra.count
End Sub
