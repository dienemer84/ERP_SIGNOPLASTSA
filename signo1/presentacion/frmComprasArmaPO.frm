VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmComprasArmaPO 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Creacion de PO"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmComprasArmaPO.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   10680
   Begin XtremeSuiteControls.PushButton btnSeleccionar 
      Height          =   300
      Left            =   5970
      TabIndex        =   4
      Top             =   105
      Width           =   1140
      _Version        =   786432
      _ExtentX        =   2011
      _ExtentY        =   529
      _StockProps     =   79
      Caption         =   "Seleccionar"
      UseVisualStyle  =   -1  'True
   End
   Begin GridEX20.GridEX grid 
      Height          =   6945
      Left            =   60
      TabIndex        =   1
      Top             =   495
      Width           =   10545
      _ExtentX        =   18600
      _ExtentY        =   12250
      Version         =   "2.0"
      PreviewRowIndent=   300
      HoldSortSettings=   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      PreviewColumn   =   "detalle"
      PreviewRowLines =   1
      MethodHoldFields=   -1  'True
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   5
      Column(1)       =   "frmComprasArmaPO.frx":000C
      Column(2)       =   "frmComprasArmaPO.frx":014C
      Column(3)       =   "frmComprasArmaPO.frx":028C
      Column(4)       =   "frmComprasArmaPO.frx":03A8
      Column(5)       =   "frmComprasArmaPO.frx":04D8
      GroupCount      =   1
      Group(1)        =   "frmComprasArmaPO.frx":0610
      SortKeysCount   =   2
      SortKey(1)      =   "frmComprasArmaPO.frx":0678
      SortKey(2)      =   "frmComprasArmaPO.frx":06E0
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmComprasArmaPO.frx":0748
      FormatStyle(2)  =   "frmComprasArmaPO.frx":0870
      FormatStyle(3)  =   "frmComprasArmaPO.frx":0920
      FormatStyle(4)  =   "frmComprasArmaPO.frx":09D4
      FormatStyle(5)  =   "frmComprasArmaPO.frx":0AAC
      FormatStyle(6)  =   "frmComprasArmaPO.frx":0B64
      ImageCount      =   0
      PrinterProperties=   "frmComprasArmaPO.frx":0C44
   End
   Begin XtremeSuiteControls.PushButton btnCrearPO 
      Height          =   405
      Left            =   8700
      TabIndex        =   0
      Top             =   7545
      Width           =   1905
      _Version        =   786432
      _ExtentX        =   3360
      _ExtentY        =   714
      _StockProps     =   79
      Caption         =   "Crear Peticiones Of."
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.ComboBox cboProveedores 
      Height          =   315
      Left            =   2520
      TabIndex        =   3
      Top             =   90
      Width           =   3360
      _Version        =   786432
      _ExtentX        =   5927
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Sorted          =   -1  'True
      Style           =   2
      Text            =   "ComboBox1"
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Seleccionar items del proveedor"
      Height          =   195
      Left            =   150
      TabIndex        =   2
      Top             =   150
      Width           =   2280
   End
End
Attribute VB_Name = "frmComprasArmaPO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private detalles As Collection
Private det As clsRequeMateriales
Private potenciales() As PotencialPetOf
Dim pot As PotencialPetOf

Private Type PotencialPetOf
    Proveedor As clsProveedor
    detalle As clsRequeMateriales
    GenerarPetOf As Boolean
End Type

Private proveedores As New Collection


Private Sub btnCrearPO_Click()
    Dim i As Long
    Dim col As New Collection
    Dim tmpCol As Collection

    If Me.grid.ItemCount = 0 Then Exit Sub

    If Me.grid.EditMode = jgexEditModeOn Then
        MsgBox "Todavia esta editando algun item, seleccionelo y presione [ENTER].", vbOKOnly + vbExclamation
        Me.grid.SetFocus
        Exit Sub
    End If

    For i = 1 To UBound(potenciales, 1)
        If potenciales(i).GenerarPetOf Then
            Set tmpCol = New Collection
            tmpCol.Add potenciales(i).detalle
            tmpCol.Add potenciales(i).Proveedor.Id
            col.Add tmpCol
        End If
    Next i

    If col.count > 0 Then
        If DAOPeticionOferta.crearPO(col) Then
            FillGrid
            Me.grid.Refresh
            Me.grid.RefreshGroups

            Dim eve As New clsEventoObserver
            Set eve.Elemento = Nothing
            Set eve.Originador = Me
            eve.EVENTO = modificar_
            eve.Tipo = RequerimientosCompra_
            Channel.Notificar eve, RequerimientosCompra_

            MsgBox "PO creadas.", vbOKOnly + vbInformation
        Else
            MsgBox "Hubo alguno error al crear las PO.", vbOKOnly + vbCritical
        End If
    Else
        MsgBox "Debe seleccionar algun item para generar las PO.", vbExclamation
    End If


End Sub

Private Sub btnSeleccionar_Click()
    If Me.cboProveedores.ListIndex >= 0 Then
        Dim provid As Long
        provid = Me.cboProveedores.ItemData(Me.cboProveedores.ListIndex)
        Dim i As Long
        For i = 1 To UBound(potenciales, 1)
            potenciales(i).GenerarPetOf = (potenciales(i).Proveedor.Id = provid)
        Next i

        Me.grid.ItemCount = 0
        Me.grid.ItemCount = UBound(potenciales, 1)

        GridEXHelper.AutoSizeColumns Me.grid

    End If
End Sub

Private Sub Form_Load()
    Customize Me
    GridEXHelper.CustomizeGrid Me.grid, True, True
    FillGrid
End Sub

Private Sub FillGrid()
    Me.grid.ItemCount = 0

    Set proveedores = New Collection

    Set detalles = DAORequeMateriales.FindListosParaPetOf()
    Erase potenciales

    Dim prov As clsProveedor

    Dim dimension As Long: dimension = 0

    'recorrer con indce de atras para adlenate
    Dim i As Long
    For Each det In detalles
        For i = det.ListaProveedores.count To 1 Step -1
            If DAOPeticionOfertaDetalle.FindAll(, "pod.id_detalle_reque = " & det.Id & " and po.id_proveedor = " & det.ListaProveedores(i).Id).count > 0 Then
                det.ListaProveedores.remove i
            End If
        Next i
    Next det


    For Each det In detalles
        dimension = dimension + det.ListaProveedores.count
    Next
    If dimension = 0 Then Exit Sub
    ReDim potenciales(1 To dimension)

    dimension = 0
    For Each det In detalles
        For Each prov In det.ListaProveedores
            dimension = dimension + 1

            Set pot.detalle = det
            Set pot.Proveedor = prov

            If Not funciones.BuscarEnColeccion(proveedores, CStr(prov.Id)) Then
                proveedores.Add prov, CStr(prov.Id)
            End If

            pot.GenerarPetOf = False

            potenciales(dimension) = pot
        Next prov
    Next det


    Me.grid.ItemCount = UBound(potenciales, 1)

    GridEXHelper.AutoSizeColumns Me.grid


    Me.cboProveedores.Clear
    For Each prov In proveedores
        Me.cboProveedores.AddItem prov.RazonSocial
        Me.cboProveedores.ItemData(Me.cboProveedores.NewIndex) = prov.Id
    Next prov
    Me.cboProveedores.ListIndex = -1

End Sub


Private Sub grid_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If rowIndex > 0 And UBound(potenciales, 1) > 0 Then
        pot = potenciales(rowIndex)
        Values(1) = pot.Proveedor.RazonSocial
        Values(2) = pot.detalle.RequeId
        Values(3) = pot.detalle.Material.descripcion
        Values(4) = pot.detalle.observaciones
        Values(5) = pot.GenerarPetOf
    End If
End Sub

Private Sub grid_UnboundUpdate(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If rowIndex > 0 And UBound(potenciales, 1) > 0 Then
        potenciales(rowIndex).GenerarPetOf = CBool(Values(5))
    End If
End Sub
