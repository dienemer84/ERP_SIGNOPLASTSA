VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmAdminVentasListaRECSegunCbte 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Recibos de cobro vinculados"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6720
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   6720
   ShowInTaskbar   =   0   'False
   Begin GridEX20.GridEX gridREC 
      Height          =   2655
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   4683
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      ColumnAutoResize=   -1  'True
      ReadOnly        =   -1  'True
      MethodHoldFields=   -1  'True
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   3
      Column(1)       =   "frmAdminVentasListaRECSegunCbte.frx":0000
      Column(2)       =   "frmAdminVentasListaRECSegunCbte.frx":016C
      Column(3)       =   "frmAdminVentasListaRECSegunCbte.frx":02AC
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminVentasListaRECSegunCbte.frx":03F4
      FormatStyle(2)  =   "frmAdminVentasListaRECSegunCbte.frx":052C
      FormatStyle(3)  =   "frmAdminVentasListaRECSegunCbte.frx":05DC
      FormatStyle(4)  =   "frmAdminVentasListaRECSegunCbte.frx":0690
      FormatStyle(5)  =   "frmAdminVentasListaRECSegunCbte.frx":0768
      FormatStyle(6)  =   "frmAdminVentasListaRECSegunCbte.frx":0820
      ImageCount      =   0
      PrinterProperties=   "frmAdminVentasListaRECSegunCbte.frx":0900
   End
   Begin XtremeSuiteControls.PushButton btnCerrar 
      Height          =   495
      Left            =   2640
      TabIndex        =   2
      Top             =   4080
      Width           =   1455
      _Version        =   786432
      _ExtentX        =   2566
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Cerrar"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.Label lblInstrucciones 
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   6495
      _Version        =   786432
      _ExtentX        =   11456
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "* Doble click para abrir el Recibo seleccionado."
   End
   Begin XtremeSuiteControls.Label lblNumeroCbte 
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      _Version        =   786432
      _ExtentX        =   11033
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Comprobante N°"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmAdminVentasListaRECSegunCbte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public vFactura As Factura

Dim ordenes As New Collection
Private Orden As OrdenPago

Dim recibos As New Collection
Private Recibo As Recibo


Public Property Let Factura(nFactura As Factura)
    If IsSomething(nFactura) Then
        Set vFactura = DAOFactura.FindById(nFactura.Id)
    End If
End Property


Private Sub btnCerrar_Click()
    Unload Me
    
End Sub

Private Sub Form_Load()

    Me.Height = 5265
    Me.Width = 6870
    
    Me.Left = frmPrincipal.ScaleWidth / 3
    Me.Top = frmPrincipal.ScaleHeight / 4
    
    FormHelper.Customize Me
    
    Me.lblNumeroCbte(0).caption = "Comprobante: " & vFactura.NumeroFormateado & " | " & UCase(vFactura.cliente.razon)
    
    MostrarRecibos
    
End Sub

Public Sub MostrarRecibos()

   Dim filter As String
    filter = "1 = 1"

    filter = filter & " AND  adminrec.idFactura = " & vFactura.Id

    gridREC.ItemCount = 0
    
    Set recibos = DAORecibo.FindAllByCbte(filter)
    
   Me.gridREC.ItemCount = recibos.count

End Sub

Private Sub gridREC_ColumnHeaderClick(ByVal Column As GridEX20.JSColumn)
    GridEXHelper.ColumnHeaderClick Me.gridREC, Column
End Sub

Private Sub gridREC_DblClick()
    gridREC_SelectionChange
    verREC
End Sub

Private Sub gridREC_RowFormat(RowBuffer As GridEX20.JSRowData)
    If RowBuffer.rowIndex > 0 And recibos.count > 0 Then
        Set Recibo = recibos.item(RowBuffer.rowIndex)
        If Recibo.estado = EstadoRecibo.Aprobado Then
            RowBuffer.CellStyle(3) = "aprobada"
        ElseIf Recibo.estado = EstadoRecibo.Reciboanulado Then
            RowBuffer.RowStyle = "anulada2"

            RowBuffer.CellStyle(3) = "anulada"
        ElseIf Recibo.estado = EstadoRecibo.Pendiente Then
            RowBuffer.CellStyle(3) = "pendiente"
        End If
    End If
End Sub

Private Sub gridREC_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If rowIndex > 0 And recibos.count > 0 Then
    
        Set Recibo = recibos.item(rowIndex)
        
        Values(1) = Recibo.Id
        Values(2) = Recibo.FEcha
        Values(3) = enums.EnumEstadoRecibo(Recibo.estado)
        
    End If
End Sub

Private Sub verREC()
    
    Dim F As New frmAdminCobranzasNuevoRecibo
    F.editar = False
    F.reciboId = Recibo.Id
    F.Show
    
End Sub

Private Sub gridREC_SelectionChange()
    On Error Resume Next
    Set Recibo = recibos.item(gridREC.rowIndex(gridREC.row))
    
End Sub

