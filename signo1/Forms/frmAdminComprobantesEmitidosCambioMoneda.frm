VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmAdminComprobantesEmitidosCambioMoneda 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Hacer conversión"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6825
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   6825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.PushButton PushButtonCancelar 
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   6360
      Width           =   1815
      _Version        =   786432
      _ExtentX        =   3201
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Cancelar"
      UseVisualStyle  =   -1  'True
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   4320
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   3201
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      ColumnAutoResize=   -1  'True
      MethodHoldFields=   -1  'True
      Options         =   -1
      RecordsetType   =   1
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   4
      Column(1)       =   "frmAdminComprobantesEmitidosCambioMoneda.frx":0000
      Column(2)       =   "frmAdminComprobantesEmitidosCambioMoneda.frx":016C
      Column(3)       =   "frmAdminComprobantesEmitidosCambioMoneda.frx":0298
      Column(4)       =   "frmAdminComprobantesEmitidosCambioMoneda.frx":03EC
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminComprobantesEmitidosCambioMoneda.frx":04AC
      FormatStyle(2)  =   "frmAdminComprobantesEmitidosCambioMoneda.frx":05E4
      FormatStyle(3)  =   "frmAdminComprobantesEmitidosCambioMoneda.frx":0694
      FormatStyle(4)  =   "frmAdminComprobantesEmitidosCambioMoneda.frx":0748
      FormatStyle(5)  =   "frmAdminComprobantesEmitidosCambioMoneda.frx":0820
      FormatStyle(6)  =   "frmAdminComprobantesEmitidosCambioMoneda.frx":08D8
      ImageCount      =   0
      PrinterProperties=   "frmAdminComprobantesEmitidosCambioMoneda.frx":09B8
   End
   Begin XtremeSuiteControls.PushButton PushButtonAceptar 
      Height          =   495
      Left            =   4920
      TabIndex        =   0
      Top             =   6360
      Width           =   1815
      _Version        =   786432
      _ExtentX        =   3201
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Aceptar"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   735
      Left            =   240
      TabIndex        =   5
      Top             =   3000
      Width           =   4215
      _Version        =   786432
      _ExtentX        =   7435
      _ExtentY        =   1296
      _StockProps     =   79
      Caption         =   "Label3"
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   735
      Left            =   240
      TabIndex        =   4
      Top             =   2040
      Width           =   4095
      _Version        =   786432
      _ExtentX        =   7223
      _ExtentY        =   1296
      _StockProps     =   79
      Caption         =   "Label2"
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6615
      _Version        =   786432
      _ExtentX        =   11668
      _ExtentY        =   2566
      _StockProps     =   79
      Caption         =   "Label1"
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmAdminComprobantesEmitidosCambioMoneda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Mn As clsMoneda
Private mon As New Collection
Public Ot As OrdenTrabajo
Public Factura As Factura

Dim q As String

Private Sub Form_Load()
    Customize Me
    Me.GridEX1.ItemCount = 0
    Me.caption = "Conversión de Moneda (" & Name & ")"

    Me.Label1.caption = "La OT origen a facturar tiene como Moneda asignada: " & vbCrLf & "" & Ot.moneda.NombreCorto & " -" & Ot.moneda.NombreLargo & ". " & vbCrLf & " " & vbCrLf & "" _
                      & "¿Desea realizar la conversión de acuerdo al valor de otra moneda? " & vbCrLf & "" & vbCrLf & "" _
                      & " Si desea modificarlo, seleccione el valor por el cual convertir y luego Acepte esta ventana. " & vbCrLf & "" _
                      & " En el caso de que desea mantener el mismo valor, seleccione Cancelar esta ventana."

    Me.Label2.caption = "Moneda de OT: " & Ot.moneda.NombreCorto & "- " & Ot.moneda.NombreLargo & "."

    Me.Label3.caption = "Moneda de Comprobante: " & Factura.moneda.NombreCorto & "- " & Factura.moneda.NombreLargo & "."


    llenarLista

End Sub

Private Sub llenarLista()
    Set mon = DAOMoneda.GetAll(q)

    Me.GridEX1.ItemCount = 0
    Me.GridEX1.ItemCount = mon.count

End Sub

Private Sub GridEX1_ColumnHeaderClick(ByVal Column As GridEX20.JSColumn)
    GridEXHelper.ColumnHeaderClick Me.GridEX1, Column

End Sub


Private Sub GridEX1_SelectionChange()
    Set Mn = mon.item(Me.GridEX1.rowIndex(Me.GridEX1.row))

End Sub

Private Sub GridEX1_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Set Mn = mon.item(rowIndex)

    Values(1) = Mn.NombreCorto
    Values(2) = Mn.NombreLargo
    Values(3) = Mn.Cambio
    Values(4) = Mn.FechaActual

End Sub

Private Sub PushButtonAceptar_Click()
    GridEX1_SelectionChange
    'Set Monedas.MonedaConvertibles = Mn
    Unload Me

End Sub

Private Sub PushButtonCancelar_Click()
'Set Monedas.MonedaConvertibles = Factura.moneda.MonedaCambio
    MsgBox ("Se calculará por el valor de la moneda de la Factura")
    Unload Me

End Sub
