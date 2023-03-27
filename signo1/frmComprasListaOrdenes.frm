VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmComprasOrdenesLista 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Ordenes de compra"
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   11355
   ClipControls    =   0   'False
   Icon            =   "frmComprasListaOrdenes.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6975
   ScaleWidth      =   11355
   Begin VB.CommandButton btnBuscar 
      Caption         =   "Buscar"
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   5265
      Width           =   1785
   End
   Begin GridEX20.GridEX grid 
      Height          =   5145
      Left            =   -30
      TabIndex        =   0
      Top             =   15
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   9075
      Version         =   "2.0"
      PreviewRowIndent=   500
      DefaultGroupMode=   1
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      PreviewRowLines =   2
      CalendarTodayText=   "Hoy"
      CalendarNoneText=   "Vacio"
      ColumnAutoResize=   -1  'True
      MultiSelect     =   -1  'True
      MethodHoldFields=   -1  'True
      ContScroll      =   -1  'True
      ForeColorInfoText=   16777215
      BackColorInfoText=   8421504
      GroupByBoxInfoText=   "Arrastre una columna aqui para ordenar por dicha columna."
      AllowEdit       =   0   'False
      BackColorGBBox  =   8421504
      BackColorHeader =   16761024
      RowHeaders      =   -1  'True
      ItemCount       =   1
      DataMode        =   99
      HeaderFontName  =   "Tahoma"
      FontName        =   "Tahoma"
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   4
      Column(1)       =   "frmComprasListaOrdenes.frx":000C
      Column(2)       =   "frmComprasListaOrdenes.frx":0160
      Column(3)       =   "frmComprasListaOrdenes.frx":0254
      Column(4)       =   "frmComprasListaOrdenes.frx":0398
      FormatStylesCount=   10
      FormatStyle(1)  =   "frmComprasListaOrdenes.frx":0488
      FormatStyle(2)  =   "frmComprasListaOrdenes.frx":05B0
      FormatStyle(3)  =   "frmComprasListaOrdenes.frx":0660
      FormatStyle(4)  =   "frmComprasListaOrdenes.frx":0714
      FormatStyle(5)  =   "frmComprasListaOrdenes.frx":07C8
      FormatStyle(6)  =   "frmComprasListaOrdenes.frx":08A0
      FormatStyle(7)  =   "frmComprasListaOrdenes.frx":0980
      FormatStyle(8)  =   "frmComprasListaOrdenes.frx":0A4C
      FormatStyle(9)  =   "frmComprasListaOrdenes.frx":0B18
      FormatStyle(10) =   "frmComprasListaOrdenes.frx":0BEC
      ImageCount      =   0
      PrinterProperties=   "frmComprasListaOrdenes.frx":0CB0
   End
   Begin VB.Menu mnuc 
      Caption         =   "mnu"
      Visible         =   0   'False
      Begin VB.Menu nro 
         Caption         =   "[nro]"
         Enabled         =   0   'False
      End
      Begin VB.Menu editar 
         Caption         =   "Editar..."
      End
      Begin VB.Menu aprobar 
         Caption         =   "Aprobar..."
      End
      Begin VB.Menu sadf 
         Caption         =   "-"
      End
      Begin VB.Menu verDetalle 
         Caption         =   "Ver Detalle..."
      End
      Begin VB.Menu verHistorial 
         Caption         =   "Ver Historial"
      End
   End
End
Attribute VB_Name = "frmComprasOrdenesLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private ordenes As Collection
Private Orden As OrdenCompra

Private Sub btnBuscar_Click()
    Buscar
End Sub

Private Sub Buscar()
    Me.grid.ItemCount = 0
    Set ordenes = DAOOrdenCompra.FindAll()
    Me.grid.ItemCount = ordenes.count
End Sub

Private Sub Form_Load()
    Customize Me
    GridEXHelper.CustomizeGrid Me.grid, True
    Me.grid.ItemCount = 0
End Sub

Private Sub grid_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex > 0 Then
        Set Orden = ordenes.item(RowIndex)
        Values(1) = Orden.Id
        If Orden.Proveedor Is Nothing Then
            Values(2) = Empty
        Else
            Values(2) = Orden.Proveedor.RazonSocial
        End If
        Values(3) = Orden.FechaCreacion
        Values(4) = Orden.estado

    End If
End Sub
