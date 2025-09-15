VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmHistorialesProduccion 
   Caption         =   "Historial de producción"
   ClientHeight    =   3780
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13065
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3780
   ScaleWidth      =   13065
   WindowState     =   2  'Maximized
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
      PreviewRowLines =   1
      ColumnAutoResize=   -1  'True
      MethodHoldFields=   -1  'True
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      BackColorHeader =   16761024
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      ColumnsCount    =   17
      Column(1)       =   "frmHistorialesProduccion.frx":0000
      Column(2)       =   "frmHistorialesProduccion.frx":0194
      Column(3)       =   "frmHistorialesProduccion.frx":02F4
      Column(4)       =   "frmHistorialesProduccion.frx":0478
      Column(5)       =   "frmHistorialesProduccion.frx":05F8
      Column(6)       =   "frmHistorialesProduccion.frx":0784
      Column(7)       =   "frmHistorialesProduccion.frx":090C
      Column(8)       =   "frmHistorialesProduccion.frx":0A88
      Column(9)       =   "frmHistorialesProduccion.frx":0C00
      Column(10)      =   "frmHistorialesProduccion.frx":0D74
      Column(11)      =   "frmHistorialesProduccion.frx":0EE4
      Column(12)      =   "frmHistorialesProduccion.frx":1048
      Column(13)      =   "frmHistorialesProduccion.frx":11A8
      Column(14)      =   "frmHistorialesProduccion.frx":130C
      Column(15)      =   "frmHistorialesProduccion.frx":146C
      Column(16)      =   "frmHistorialesProduccion.frx":15B4
      Column(17)      =   "frmHistorialesProduccion.frx":170C
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmHistorialesProduccion.frx":1860
      FormatStyle(2)  =   "frmHistorialesProduccion.frx":1998
      FormatStyle(3)  =   "frmHistorialesProduccion.frx":1A48
      FormatStyle(4)  =   "frmHistorialesProduccion.frx":1AFC
      FormatStyle(5)  =   "frmHistorialesProduccion.frx":1BD4
      FormatStyle(6)  =   "frmHistorialesProduccion.frx":1D28
      ImageCount      =   0
      PrinterProperties=   "frmHistorialesProduccion.frx":1E08
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   3240
      Width           =   3375
      _Version        =   786432
      _ExtentX        =   5953
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

Public Enum ColsHistorial
    cUsuarioOperacion = 1
    cUsuarioRecibio
    cCantRecibidaOld
    cCantRecibidaNew
    cCantFabricadaOld
    cCantFabricadaNew
    cCantScrapOld
    cCantScrapNew
    cFechaInicioOld
    cFechaInicioNew
    cFechaFinOld
    cFechaFinNew
    cProcesoOld
    cProcesoNew
    cAccion
    cObservacion
    cFecha
End Enum

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
        Me.Label1.caption = "Sin historial"
    Else
        Me.Label1.caption = CStr(vLista.count) & " eventos"
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
                                    ByVal Bookmark As Variant, _
                                    ByVal Values As GridEX20.JSRowData)

    If vLista Is Nothing Then Exit Sub
    If RowIndex < 1 Or RowIndex > vLista.count Then Exit Sub  ' Janus 1-based

    Set tmp = vLista.item(RowIndex)


    '--- Usuarios
    If Not tmp.UsuarioOperacion Is Nothing Then
        Values(cUsuarioOperacion) = tmp.UsuarioOperacion.usuario
    End If
    If Not tmp.UsuarioRecibio Is Nothing Then
        Values(cUsuarioRecibio) = tmp.UsuarioRecibio.usuario
    End If
    
    '--- Cantidades
    Values(cCantRecibidaOld) = tmp.CantRecibidaOld
    Values(cCantRecibidaNew) = tmp.CantRecibidaNew
    Values(cCantFabricadaOld) = tmp.CantFabricadaOld
    Values(cCantFabricadaNew) = tmp.CantFabricadaNew
    Values(cCantScrapOld) = tmp.CantScrapOld
    Values(cCantScrapNew) = tmp.CantScrapNew

    '--- Fechas
    Values(cFechaInicioOld) = tmp.FechaInicioOld
    Values(cFechaInicioNew) = tmp.FechaInicioNew
    Values(cFechaFinOld) = tmp.FechaFinOld
    Values(cFechaFinNew) = tmp.FechaFinNew

    '--- Procesos
    If Not tmp.ProcesoOld Is Nothing Then
        Values(cProcesoOld) = tmp.ProcesoOld.Modulo  ' o .Id si querés el numérico
    End If
    
    If Not tmp.ProcesoNew Is Nothing Then
        Values(cProcesoNew) = tmp.ProcesoNew.Modulo   ' o .Id
    End If
    
    '--- Acción / Nota / Fecha
    Values(cAccion) = tmp.Accion
    Values(cObservacion) = tmp.Observacion
    Values(cFecha) = tmp.FEcha
End Sub


