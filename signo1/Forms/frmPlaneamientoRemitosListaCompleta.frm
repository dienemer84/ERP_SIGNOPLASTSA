VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmPlaneamientoRemitosListaCompleta 
   Caption         =   "Form1"
   ClientHeight    =   9165
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16530
   LinkTopic       =   "Form1"
   ScaleHeight     =   9165
   ScaleWidth      =   16530
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.GroupBox GroupBox 
      Height          =   2175
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   16215
      _Version        =   786432
      _ExtentX        =   28601
      _ExtentY        =   3836
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.PushButton btnBuscar 
         Height          =   615
         Left            =   13800
         TabIndex        =   2
         Top             =   1440
         Width           =   2295
         _Version        =   786432
         _ExtentX        =   4048
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "Buscar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.GroupBox GroupBox 
         Height          =   1335
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   4695
         _Version        =   786432
         _ExtentX        =   8281
         _ExtentY        =   2355
         _StockProps     =   79
         Caption         =   "Fecha"
         BackColor       =   12632256
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.DateTimePicker dtpDesde 
            Height          =   315
            Left            =   855
            TabIndex        =   4
            Top             =   735
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
            Left            =   3030
            TabIndex        =   5
            Top             =   735
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
            Left            =   825
            TabIndex        =   6
            Top             =   240
            Width           =   2190
            _Version        =   786432
            _ExtentX        =   3863
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Style           =   2
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.Label Label7 
            Height          =   195
            Left            =   240
            TabIndex        =   9
            Top             =   300
            Width           =   480
            _Version        =   786432
            _ExtentX        =   847
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Rango"
            BackColor       =   12632256
            AutoSize        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label5 
            Height          =   195
            Left            =   285
            TabIndex        =   8
            Top             =   780
            Width           =   465
            _Version        =   786432
            _ExtentX        =   820
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Desde"
            BackColor       =   12632256
            AutoSize        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label6 
            Height          =   195
            Left            =   2460
            TabIndex        =   7
            Top             =   795
            Width           =   420
            _Version        =   786432
            _ExtentX        =   741
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Hasta"
            BackColor       =   12632256
            AutoSize        =   -1  'True
         End
      End
   End
   Begin GridEX20.GridEX gridRemitosCompletos 
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   2400
      Width           =   16215
      _ExtentX        =   28601
      _ExtentY        =   11456
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      MethodHoldFields=   -1  'True
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   13
      Column(1)       =   "frmPlaneamientoRemitosListaCompleta.frx":0000
      Column(2)       =   "frmPlaneamientoRemitosListaCompleta.frx":013C
      Column(3)       =   "frmPlaneamientoRemitosListaCompleta.frx":0254
      Column(4)       =   "frmPlaneamientoRemitosListaCompleta.frx":036C
      Column(5)       =   "frmPlaneamientoRemitosListaCompleta.frx":047C
      Column(6)       =   "frmPlaneamientoRemitosListaCompleta.frx":0594
      Column(7)       =   "frmPlaneamientoRemitosListaCompleta.frx":06B4
      Column(8)       =   "frmPlaneamientoRemitosListaCompleta.frx":07D4
      Column(9)       =   "frmPlaneamientoRemitosListaCompleta.frx":08E4
      Column(10)      =   "frmPlaneamientoRemitosListaCompleta.frx":0A10
      Column(11)      =   "frmPlaneamientoRemitosListaCompleta.frx":0B30
      Column(12)      =   "frmPlaneamientoRemitosListaCompleta.frx":0C50
      Column(13)      =   "frmPlaneamientoRemitosListaCompleta.frx":0D60
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmPlaneamientoRemitosListaCompleta.frx":0E94
      FormatStyle(2)  =   "frmPlaneamientoRemitosListaCompleta.frx":0FCC
      FormatStyle(3)  =   "frmPlaneamientoRemitosListaCompleta.frx":107C
      FormatStyle(4)  =   "frmPlaneamientoRemitosListaCompleta.frx":1130
      FormatStyle(5)  =   "frmPlaneamientoRemitosListaCompleta.frx":1208
      FormatStyle(6)  =   "frmPlaneamientoRemitosListaCompleta.frx":12C0
      ImageCount      =   0
      PrinterProperties=   "frmPlaneamientoRemitosListaCompleta.frx":13A0
   End
End
Attribute VB_Name = "frmPlaneamientoRemitosListaCompleta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Remito As Remito
'Dim tmp As remitoDetalle
Private remitos As New Collection
Private col As New Collection
Dim tmp As remitoDetalle
Public colremitosdetalles As New Collection
Dim filtro As String

Public VerInfoAdministracion As Boolean
Public VerInfoPlaneamiento As Boolean

Dim claseP As New classPlaneamiento
Dim facturasRemitos As Dictionary
Dim m_Archivos As Dictionary


Private Sub btnBuscar_Click()
    LlenarGridDetalles
End Sub

Public Sub LlenarGrid()

    Set m_Archivos = DAOArchivo.GetCantidadArchivosPorReferencia(OA_Remitos)

    filtro = "1 = 1"
    
    If Not IsNull(Me.dtpDesde.value) Then
        filtro = filtro & " and  " & DAORemitoS.TABLA_REMITO & "." & DAORemitoS.CAMPO_FECHA & " >= " & conectar.Escape(Format(Me.dtpDesde.value, "yyyy-mm-dd 00:00:00"))
    End If

    If Not IsNull(Me.dtpHasta.value) Then
        filtro = filtro & " and  " & DAORemitoS.TABLA_REMITO & "." & DAORemitoS.CAMPO_FECHA & " <= " & conectar.Escape(Format(Me.dtpHasta.value, "yyyy-mm-dd 23:59:59"))
    End If

    Set remitos = DAORemitoS.FindAll("and " & filtro)

    Dim remi As Remito
    Dim remitosId As New Collection
    
    For Each remi In remitos
        remitosId.Add remi.Id
    Next
    
    Set facturasRemitos = New Dictionary

    If remitosId.count > 0 Then
        Set facturasRemitos = DAOFactura.FindAllByRemitos(remitosId)
    End If

    Me.gridRemitosCompletos.ItemCount = 0
    Me.gridRemitosCompletos.ItemCount = remitos.count
    Me.gridRemitosCompletos.Update

End Sub

Public Sub LlenarGridDetalles()
    Dim col As New Collection
    
    If Not IsNull(Me.dtpDesde.value) Then
        filtro = filtro & " and  r." & DAORemitoS.CAMPO_FECHA & " >= " & conectar.Escape(Format(Me.dtpDesde.value, "yyyy-mm-dd 00:00:00"))
    End If

    If Not IsNull(Me.dtpHasta.value) Then
        filtro = filtro & " and  r." & DAORemitoS.CAMPO_FECHA & " <= " & conectar.Escape(Format(Me.dtpHasta.value, "yyyy-mm-dd 23:59:59"))
    End If

'   Set Remito.Detalles = DAORemitoSDetalle.FindAllByRemito(Remito.Id, False, True)
'    Dim remitos As Collection
'    Set remitos = DAORemitoSDetalle.FindAll(filtro)
'    Set Remito.Detalles = DAORemitoSDetalle.FindAllByRemito(Remito.Id, False, True)


    Set col = DAORemitoSDetalle.FindAll(filtro)
    Me.gridRemitosCompletos.ItemCount = 0
    Me.gridRemitosCompletos.ItemCount = col.count
    
'    llenarLista
    
End Sub


'Private Sub llenarLista()
'    Me.gridRemitosCompletos.ItemCount = 0
'    Me.gridRemitosCompletos.ItemCount = colremitosdetalles.count
'
'End Sub

Private Sub gridRemitosCompletos_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
   
    '<--- Datos del detalle del Remito --->
        If rowIndex > 0 And col.count > 0 Then
        Set tmp = colremitosdetalles(rowIndex)
'            Values(7) = tmp.RemitoAlQuePertenece.EstadoFacturado
            Values(8) = rowIndex
            Values(9) = tmp.FEcha
            Values(10) = ""
            Values(11) = ""
            Values(12) = ""
            Values(13) = ""
            
        End If
'            If Not IsSomething(tmp.DetallePedido) Then
'                .value(7) = tmp.VerOrigen & Chr(10) & tmp.observaciones
'            Else
'                .value(7) = tmp.VerOrigen & " | " & tmp.DetallePedido.item & Chr(10) & tmp.observaciones
'            End If
'            .value(4) = funciones.FormatearDecimales(tmp.Cantidad, 2)
'            .value(5) = funciones.FormatearDecimales(tmp.Valor, 2)
'            .value(6) = tmp.VerFacturado

  
    '<--- Datos del detalle del Remito --->


End Sub

Private Sub cboRangos_Click()
    funciones.CalculateDateRange Me.cboRangos, Me.dtpDesde, Me.dtpHasta

End Sub


Private Sub Form_Load()
    LlenarGridDetalles

    Dim i As Integer
    funciones.FillComboBoxDateRanges Me.cboRangos
    For i = 0 To Me.cboRangos.ListCount - 1
        If Me.cboRangos.ItemData(i) = DateRangeValue.DRV_YearCurrent Then Exit For
    Next i
    Me.cboRangos.ListIndex = i
        
End Sub
