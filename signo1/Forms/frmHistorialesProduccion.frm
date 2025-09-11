VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmHistorialesProduccion 
   Caption         =   "Form1"
   ClientHeight    =   3780
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13065
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3780
   ScaleWidth      =   13065
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      Default         =   -1  'True
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3240
      Width           =   1095
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   5318
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      PreviewColumn   =   "Mensaje"
      PreviewRowLines =   1
      ColumnAutoResize=   -1  'True
      MethodHoldFields=   -1  'True
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      BackColorHeader =   16761024
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      ColumnsCount    =   3
      Column(1)       =   "frmHistorialesProduccion.frx":0000
      Column(2)       =   "frmHistorialesProduccion.frx":0164
      Column(3)       =   "frmHistorialesProduccion.frx":0258
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmHistorialesProduccion.frx":033C
      FormatStyle(2)  =   "frmHistorialesProduccion.frx":0474
      FormatStyle(3)  =   "frmHistorialesProduccion.frx":0524
      FormatStyle(4)  =   "frmHistorialesProduccion.frx":05D8
      FormatStyle(5)  =   "frmHistorialesProduccion.frx":06B0
      FormatStyle(6)  =   "frmHistorialesProduccion.frx":0804
      ImageCount      =   0
      PrinterProperties=   "frmHistorialesProduccion.frx":08E4
   End
   Begin XtremeSuiteControls.Label lblPieza 
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   3240
      Width           =   3015
      _Version        =   786432
      _ExtentX        =   5318
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Label1"
   End
End
Attribute VB_Name = "frmHistorialesProduccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim tmp As clsHistorialProduccion
Private vLista As Collection

' === Propiedad lista (objeto)
Public Property Set lista(ByVal nValue As Collection)
    Set vLista = nValue
End Property

Public Property Get lista() As Collection
    Set lista = vLista
End Property

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub llenarGrilla()
    If vLista Is Nothing Then
        Me.GridEX1.ItemCount = 0
    Else
        Me.GridEX1.ItemCount = vLista.count
    End If
End Sub

Private Sub Form_Load()
    FormHelper.Customize Me
    
    If vLista Is Nothing Or vLista.count = 0 Then
        Me.lblPieza.caption = "Sin historial"
    Else
        Me.lblPieza.caption = CStr(vLista.count) & " eventos"
    End If
    
    llenarGrilla
    GridEXHelper.CustomizeGrid Me.GridEX1
End Sub

Private Sub Form_Resize()
    Me.GridEX1.Width = Me.ScaleWidth
    Me.GridEX1.Height = Me.ScaleHeight - 650
    Me.Command1.Top = Me.GridEX1.Height + 150
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, _
    ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)

    If vLista Is Nothing Then Exit Sub
    If RowIndex < 1 Or RowIndex > vLista.count Then Exit Sub  ' Janus 1-based

    Set tmp = vLista.item(RowIndex)

    ' OJO: usa los nombres reales de tu clase (Fecha/mensaje/Usuario)
    Values(1) = tmp.FEcha          ' si tu clase la expuso como FEcha, mantené FEcha
    Values(2) = tmp.CantFabricada
    Values(3) = tmp.Accion
End Sub

