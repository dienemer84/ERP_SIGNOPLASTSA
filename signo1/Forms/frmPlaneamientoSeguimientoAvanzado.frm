VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmPlaneamientoSeguimientoAvanzado 
   Caption         =   "Seguimiento de Producci�n"
   ClientHeight    =   13725
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   14415
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   13725
   ScaleWidth      =   14415
   WindowState     =   2  'Maximized
   Begin GridEX20.GridEX gridAlmacenes 
      Height          =   3495
      Left            =   7200
      TabIndex        =   16
      Top             =   9240
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   6165
      Version         =   "2.0"
      BoundColumnIndex=   "id"
      ReplaceColumnIndex=   "nombre"
      ActAsDropDown   =   -1  'True
      ColumnAutoResize=   -1  'True
      HideSelection   =   2
      MethodHoldFields=   -1  'True
      ContScroll      =   -1  'True
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      ColumnHeaders   =   0   'False
      RowHeaders      =   -1  'True
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   2
      Column(1)       =   "frmPlaneamientoSeguimientoAvanzado.frx":0000
      Column(2)       =   "frmPlaneamientoSeguimientoAvanzado.frx":015C
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmPlaneamientoSeguimientoAvanzado.frx":0250
      FormatStyle(2)  =   "frmPlaneamientoSeguimientoAvanzado.frx":0388
      FormatStyle(3)  =   "frmPlaneamientoSeguimientoAvanzado.frx":0438
      FormatStyle(4)  =   "frmPlaneamientoSeguimientoAvanzado.frx":04EC
      FormatStyle(5)  =   "frmPlaneamientoSeguimientoAvanzado.frx":05C4
      FormatStyle(6)  =   "frmPlaneamientoSeguimientoAvanzado.frx":067C
      ImageCount      =   0
      PrinterProperties=   "frmPlaneamientoSeguimientoAvanzado.frx":075C
   End
   Begin GridEX20.GridEX gridSectores 
      Height          =   3495
      Left            =   3720
      TabIndex        =   15
      Top             =   9240
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   6165
      Version         =   "2.0"
      BoundColumnIndex=   "id"
      ReplaceColumnIndex=   "sector"
      ActAsDropDown   =   -1  'True
      ColumnAutoResize=   -1  'True
      HideSelection   =   2
      MethodHoldFields=   -1  'True
      ContScroll      =   -1  'True
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      ColumnHeaders   =   0   'False
      RowHeaders      =   -1  'True
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   2
      Column(1)       =   "frmPlaneamientoSeguimientoAvanzado.frx":0934
      Column(2)       =   "frmPlaneamientoSeguimientoAvanzado.frx":0A58
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmPlaneamientoSeguimientoAvanzado.frx":0B4C
      FormatStyle(2)  =   "frmPlaneamientoSeguimientoAvanzado.frx":0C84
      FormatStyle(3)  =   "frmPlaneamientoSeguimientoAvanzado.frx":0D34
      FormatStyle(4)  =   "frmPlaneamientoSeguimientoAvanzado.frx":0DE8
      FormatStyle(5)  =   "frmPlaneamientoSeguimientoAvanzado.frx":0EC0
      FormatStyle(6)  =   "frmPlaneamientoSeguimientoAvanzado.frx":0F78
      ImageCount      =   0
      PrinterProperties=   "frmPlaneamientoSeguimientoAvanzado.frx":1058
   End
   Begin GridEX20.GridEX gridUsuarios 
      Height          =   3495
      Left            =   240
      TabIndex        =   14
      Top             =   9240
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   6165
      Version         =   "2.0"
      BoundColumnIndex=   "id"
      ReplaceColumnIndex=   "legajo"
      ActAsDropDown   =   -1  'True
      ColumnAutoResize=   -1  'True
      HideSelection   =   2
      MethodHoldFields=   -1  'True
      ContScroll      =   -1  'True
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      ColumnHeaders   =   0   'False
      NewRowPos       =   1
      RowHeaders      =   -1  'True
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   3
      Column(1)       =   "frmPlaneamientoSeguimientoAvanzado.frx":1230
      Column(2)       =   "frmPlaneamientoSeguimientoAvanzado.frx":1330
      Column(3)       =   "frmPlaneamientoSeguimientoAvanzado.frx":1424
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmPlaneamientoSeguimientoAvanzado.frx":1518
      FormatStyle(2)  =   "frmPlaneamientoSeguimientoAvanzado.frx":1650
      FormatStyle(3)  =   "frmPlaneamientoSeguimientoAvanzado.frx":1700
      FormatStyle(4)  =   "frmPlaneamientoSeguimientoAvanzado.frx":17B4
      FormatStyle(5)  =   "frmPlaneamientoSeguimientoAvanzado.frx":188C
      FormatStyle(6)  =   "frmPlaneamientoSeguimientoAvanzado.frx":1944
      ImageCount      =   0
      PrinterProperties=   "frmPlaneamientoSeguimientoAvanzado.frx":1A24
   End
   Begin XtremeSuiteControls.GroupBox GroupBox2 
      Height          =   6975
      Left            =   120
      TabIndex        =   11
      Top             =   1920
      Width           =   15975
      _Version        =   786432
      _ExtentX        =   28178
      _ExtentY        =   12303
      _StockProps     =   79
      Caption         =   "Detalle de OT"
      UseVisualStyle  =   -1  'True
      Begin GridEX20.GridEX gridDetalles 
         Height          =   5895
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   15705
         _ExtentX        =   27702
         _ExtentY        =   10398
         Version         =   "2.0"
         PreviewRowIndent=   100
         AutomaticSort   =   -1  'True
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         PreviewRowLines =   1
         ColumnAutoResize=   -1  'True
         MethodHoldFields=   -1  'True
         ContScroll      =   -1  'True
         GroupByBoxVisible=   0   'False
         RowHeaders      =   -1  'True
         DataMode        =   99
         FontSize        =   9.75
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   18
         Column(1)       =   "frmPlaneamientoSeguimientoAvanzado.frx":1BFC
         Column(2)       =   "frmPlaneamientoSeguimientoAvanzado.frx":1D80
         Column(3)       =   "frmPlaneamientoSeguimientoAvanzado.frx":1EE0
         Column(4)       =   "frmPlaneamientoSeguimientoAvanzado.frx":2040
         Column(5)       =   "frmPlaneamientoSeguimientoAvanzado.frx":21C8
         Column(6)       =   "frmPlaneamientoSeguimientoAvanzado.frx":2308
         Column(7)       =   "frmPlaneamientoSeguimientoAvanzado.frx":2488
         Column(8)       =   "frmPlaneamientoSeguimientoAvanzado.frx":25F4
         Column(9)       =   "frmPlaneamientoSeguimientoAvanzado.frx":2760
         Column(10)      =   "frmPlaneamientoSeguimientoAvanzado.frx":28B4
         Column(11)      =   "frmPlaneamientoSeguimientoAvanzado.frx":2A34
         Column(12)      =   "frmPlaneamientoSeguimientoAvanzado.frx":2BAC
         Column(13)      =   "frmPlaneamientoSeguimientoAvanzado.frx":2D3C
         Column(14)      =   "frmPlaneamientoSeguimientoAvanzado.frx":2EAC
         Column(15)      =   "frmPlaneamientoSeguimientoAvanzado.frx":3024
         Column(16)      =   "frmPlaneamientoSeguimientoAvanzado.frx":318C
         Column(17)      =   "frmPlaneamientoSeguimientoAvanzado.frx":331C
         Column(18)      =   "frmPlaneamientoSeguimientoAvanzado.frx":3440
         FormatStylesCount=   10
         FormatStyle(1)  =   "frmPlaneamientoSeguimientoAvanzado.frx":3578
         FormatStyle(2)  =   "frmPlaneamientoSeguimientoAvanzado.frx":36B0
         FormatStyle(3)  =   "frmPlaneamientoSeguimientoAvanzado.frx":3760
         FormatStyle(4)  =   "frmPlaneamientoSeguimientoAvanzado.frx":3814
         FormatStyle(5)  =   "frmPlaneamientoSeguimientoAvanzado.frx":38EC
         FormatStyle(6)  =   "frmPlaneamientoSeguimientoAvanzado.frx":39A4
         FormatStyle(7)  =   "frmPlaneamientoSeguimientoAvanzado.frx":3A84
         FormatStyle(8)  =   "frmPlaneamientoSeguimientoAvanzado.frx":3B38
         FormatStyle(9)  =   "frmPlaneamientoSeguimientoAvanzado.frx":3BCC
         FormatStyle(10) =   "frmPlaneamientoSeguimientoAvanzado.frx":3C58
         ImageCount      =   0
         PrinterProperties=   "frmPlaneamientoSeguimientoAvanzado.frx":3D08
      End
      Begin XtremeSuiteControls.ProgressBar pgbAvanceOT 
         Height          =   255
         Left            =   1320
         TabIndex        =   23
         Top             =   0
         Width           =   6735
         _Version        =   786432
         _ExtentX        =   11880
         _ExtentY        =   450
         _StockProps     =   93
         Appearance      =   6
      End
      Begin XtremeSuiteControls.Label lblAvancePct 
         Height          =   195
         Left            =   8280
         TabIndex        =   22
         Top             =   50
         Width           =   855
         _Version        =   786432
         _ExtentX        =   1508
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Label1"
      End
      Begin XtremeSuiteControls.Label lblSectorGrande 
         Height          =   495
         Left            =   1320
         TabIndex        =   13
         Top             =   6360
         Width           =   5415
         _Version        =   786432
         _ExtentX        =   9551
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "SECTOR: "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "@Arial Unicode MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   16455
      _Version        =   786432
      _ExtentX        =   29025
      _ExtentY        =   2143
      _StockProps     =   79
      Caption         =   "P�rametros de b�squeda"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.GroupBox grpAvanceOT 
         Height          =   855
         Left            =   11160
         TabIndex        =   17
         Top             =   240
         Width           =   5175
         _Version        =   786432
         _ExtentX        =   9128
         _ExtentY        =   1508
         _StockProps     =   79
         Caption         =   "Totalizador"
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.Label lblCantScrap 
            Height          =   195
            Left            =   2640
            TabIndex        =   21
            Top             =   525
            Width           =   2415
            _Version        =   786432
            _ExtentX        =   4260
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Cant. Scrap Total:"
         End
         Begin XtremeSuiteControls.Label lblCantFabricada 
            Height          =   195
            Left            =   2640
            TabIndex        =   20
            Top             =   240
            Width           =   2415
            _Version        =   786432
            _ExtentX        =   4260
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Cant. Fabricada Total:"
         End
         Begin XtremeSuiteControls.Label lblCantRecibida 
            Height          =   195
            Left            =   240
            TabIndex        =   19
            Top             =   525
            Width           =   2055
            _Version        =   786432
            _ExtentX        =   3625
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Cant. Recibida Total:"
         End
         Begin XtremeSuiteControls.Label lblCantPedida 
            Height          =   195
            Left            =   240
            TabIndex        =   18
            Top             =   240
            Width           =   2175
            _Version        =   786432
            _ExtentX        =   3836
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Cant. Pedida Total:"
         End
      End
      Begin XtremeSuiteControls.ComboBox cboSectores 
         Height          =   315
         Left            =   1560
         TabIndex        =   9
         Top             =   720
         Width           =   3375
         _Version        =   786432
         _ExtentX        =   5953
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.GroupBox fraDatosOT 
         Height          =   855
         Left            =   5160
         TabIndex        =   1
         Tag             =   "Datos de la Orden de Trabajo N� "
         Top             =   240
         Width           =   5955
         _Version        =   786432
         _ExtentX        =   10504
         _ExtentY        =   1508
         _StockProps     =   79
         Caption         =   "Datos de la Orden de Trabajo N� "
         UseVisualStyle  =   -1  'True
         Begin VB.Label lblEstado 
            AutoSize        =   -1  'True
            Caption         =   "Estado:"
            Height          =   195
            Left            =   3960
            TabIndex        =   5
            Top             =   240
            Width           =   540
         End
         Begin VB.Label lblFechaEntrega 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Entrega:"
            Height          =   195
            Left            =   3360
            TabIndex        =   4
            Top             =   525
            Width           =   1095
         End
         Begin VB.Label lblFechaCreado 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Creada:"
            Height          =   195
            Left            =   165
            TabIndex        =   3
            Top             =   525
            Width           =   1050
         End
         Begin VB.Label lblCliente 
            AutoSize        =   -1  'True
            Caption         =   "Cliente:"
            Height          =   195
            Left            =   165
            TabIndex        =   2
            Top             =   240
            Width           =   525
         End
      End
      Begin XtremeSuiteControls.PushButton cmdBuscar 
         Height          =   345
         Left            =   3600
         TabIndex        =   6
         Top             =   240
         Width           =   1380
         _Version        =   786432
         _ExtentX        =   2434
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Buscar"
         ForeColor       =   9126421
         Appearance      =   6
      End
      Begin XtremeSuiteControls.FlatEdit txtOTNro 
         Height          =   300
         Left            =   1560
         TabIndex        =   7
         Top             =   262
         Width           =   1935
         _Version        =   786432
         _ExtentX        =   3413
         _ExtentY        =   529
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   1335
         _Version        =   786432
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Sector:"
         Alignment       =   1
      End
      Begin VB.Label lblOT 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Orden de Trabajo:"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   315
         Width           =   1290
      End
   End
   Begin VB.Menu menu 
      Caption         =   "menuSeguimiento"
      Begin VB.Menu menu_Observaciones 
         Caption         =   "Ver Observaciones"
      End
      Begin VB.Menu menu_historial 
         Caption         =   "Ver Historial"
      End
      Begin VB.Menu menu_desarrollo 
         Caption         =   "Ver Desarrollo"
      End
      Begin VB.Menu menu_archivos_asociados 
         Caption         =   "Ver Archivos Asociados"
         Begin VB.Menu AADLPieza 
            Caption         =   "De la pieza..."
         End
         Begin VB.Menu AADelDetalle 
            Caption         =   "Del Detalle..."
         End
      End
      Begin VB.Menu verIncidencias 
         Caption         =   "Ver Incidencias"
         Begin VB.Menu IdePieza 
            Caption         =   "Incidencias de Pieza..."
         End
         Begin VB.Menu IdelDetalle 
            Caption         =   "Incidencias del Detalle..."
         End
      End
   End
End
Attribute VB_Name = "frmPlaneamientoSeguimientoAvanzado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim idPieza As Long
Private tmpDetalle As DetalleOrdenTrabajo
Private m_ot As OrdenTrabajo
Private Empleado As clsEmpleado
Private empleados As New Collection
Private Sector As clsSector
Private sectores As New Collection
Private Almacen As clsAlmacen
Private almacenes As New Collection

Dim col As Collection
Private detallesPlanos As Collection
Dim dto As clsFilaPlanoRow
Private m_rows As New Collection    'colecci�n de clsFilaPlanoRow

Public Enum cols
  cID = 1
  cItem = 2
  cTipo = 3
  cUM = 4
  cNombre = 5
  cCantPedida = 6
  cCantRecibida = 7
  cCantFabricada = 8
  cCantScrap = 9
  cFechaIni = 10
  cHoraIni = 11
  cFechaFin = 12
  cHoraFin = 13
  cUsuarioRecibio = 14
  cAlmacen = 15
  cProcesoSig = 16
  cEsConjunto = 17  ' oculto
  cObservaciones = 18
End Enum

Private Enum CampoSumaPlano
   csCantPedida = 1
   csCantRecibida
   csCantFabricada
   csCantScrap
End Enum


Private Function SumarPlano(ByVal campo As CampoSumaPlano, Optional ByVal soloPiezas As Boolean = True) As Double
    Dim r As clsFilaPlanoRow
    Dim acc As Double
    If detallesPlanos Is Nothing Then Exit Function
    For Each r In detallesPlanos
        If (Not soloPiezas) Or (soloPiezas And Not r.EsConjunto) Then
            Select Case campo
                Case csCantPedida:    acc = acc + NzDbl(r.cantpedida)
                Case csCantRecibida:  acc = acc + NzDbl(r.CantRecibida)
                Case csCantFabricada:  acc = acc + NzDbl(r.CantFabricada)
                Case csCantScrap:      acc = acc + NzDbl(r.CantScrap)
            End Select
        End If
    Next
    SumarPlano = acc
End Function


Private Sub btnCargarSector_Click(): llenarDataGrid: End Sub



Private Sub cboSectores_Click()

    llenarDataGrid
    
    Me.lblSectorGrande.caption = "SECTOR: " & Me.cboSectores.Text
    
    refrescarAvanceOT
    

End Sub


Private Sub cmdBuscar_Click()

    If Not IsNumeric(Me.txtOTNro.Text) Then Exit Sub
    Set m_ot = DAOOrdenTrabajo.FindById(Me.txtOTNro.Text)
    If m_ot Is Nothing Then
        MsgBox "La Orden de Trabajo N� " & Me.txtOTNro.Text & " no existe.", vbInformation + vbOKOnly
        Exit Sub
    End If
    Set m_ot.detalles = DAODetalleOrdenTrabajo.FindAllByOrdenTrabajo(m_ot.Id)
    
'''    Me.lblSectorGrande.caption = "SECTOR: CORTE"
    
    cargarDetallesOT
    
    llenarDataGrid              ' <<< volver a cargar
    
    refrescarAvanceOT
    
        
End Sub


Private Sub totalizarAvance()

    Dim totPedida As Double
    totPedida = SumarPlano(csCantPedida, True)  ' solo piezas

    Me.lblCantPedida.caption = "Cant. Pedida Total: " & Format(totPedida, "0.##")
    
    ' Si quer�s mostrar m�s totales (opcional):
    Me.lblCantRecibida.caption = "Cant. Recibida Total: " & Format(SumarPlano(csCantRecibida, True), "0.##")
    Me.lblCantFabricada.caption = "Cant. Fabricada Total: " & Format(SumarPlano(csCantFabricada, True), "0.##")
    Me.lblCantScrap.caption = "Cant. Scrap Total: " & Format(SumarPlano(csCantScrap, True), "0.##")


End Sub



Private Sub cargarDetallesOT()
        Me.fraDatosOT.caption = Me.fraDatosOT.Tag & m_ot.Id
        Me.lblCliente.caption = "Cliente: " & m_ot.Cliente.razon
        Me.lblFechaCreado.caption = "Fecha Creada: " & m_ot.fechaCreado
        Me.lblFechaEntrega.caption = "Fecha Entrega: " & m_ot.FechaEntrega
        Me.lblEstado.caption = "Estado: " & funciones.estado_pedido(m_ot.estado)

End Sub


Private Sub Form_Load()
    FormHelper.Customize Me
    GridEXHelper.CustomizeGrid Me.gridDetalles, , True
    GridEXHelper.CustomizeGrid Me.gridSectores, False, False
    GridEXHelper.CustomizeGrid Me.gridUsuarios, False, False

    EnsureConjuntoStyle          '<< agrega el estilo
    
    Me.txtOTNro = "6002"
    
    Set sectores = DAOSectores.GetAllModulos()
    Me.gridSectores.ItemCount = sectores.count
    Set Me.gridDetalles.Columns("procesosiguiente").DropDownControl = Me.gridSectores

    Set empleados = DAOEmpleados.GetAll()
    Me.gridUsuarios.ItemCount = empleados.count
    Set Me.gridDetalles.Columns("recibio").DropDownControl = Me.gridUsuarios
    
    Set almacenes = DAOAlmacenes.GetAll()
    Me.gridAlmacenes.ItemCount = almacenes.count
    Set Me.gridDetalles.Columns("almacen").DropDownControl = Me.gridAlmacenes

   DAOSectores.LlenarComboXtremeModulos Me.cboSectores
   Me.cboSectores.ListIndex = 0
    
    Set detallesPlanos = New Collection
    
    refrescarAvanceOT
    
End Sub


Private Sub Form_Resize()
   On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    
    With fraDatosOT
        .Left = 5160
        .Top = 240
        .Width = 5955
    End With
    
    With Me.grpAvanceOT
        .Left = 11160
        .Top = 240
        .Width = Me.GroupBox2.Width
    End With
    
    ' GroupBox ocupa todo el ancho, arriba
    With GroupBox1
        .Left = 100
        .Top = 100
        .Width = Me.ScaleWidth - 200
        .Height = 1215
    End With

    With GroupBox2
        .Left = 100
        .Top = 1315
        .Width = Me.ScaleWidth - 200
        .Height = Me.Height - 2500
    End With
    
    ' Grid ajustado al resto de la ventana
    With gridDetalles
        .Top = gridDetalles.Top
        
        .Height = Me.GroupBox2.Height - 800
        .Width = Me.GroupBox2.Width - 800
        
        On Error Resume Next
        If .Columns.count >= cObservaciones Then
            .Columns(cID).Width = 500
            .Columns(cItem).Width = 500
            .Columns(cUM).Width = 500
            .Columns(cNombre).Width = 4000
            .Columns(cCantPedida).Width = 800
            .Columns(cCantFabricada).Width = 800
            .Columns(cCantRecibida).Width = 800
            .Columns(cCantScrap).Width = 800
            
            .Columns(cFechaIni).Width = 800
            .Columns(cHoraIni).Width = 800
            .Columns(cFechaFin).Width = 800
            .Columns(cHoraFin).Width = 800

            .Columns(cFechaIni).Format = "dd/MM/yyyy"
            .Columns(cHoraIni).Format = "HH:mm"
            .Columns(cFechaFin).Format = "dd/MM/yyyy"
            .Columns(cHoraFin).Format = "HH:mm"


       End If
        
        On Error GoTo 0
    
    End With
    
        With Me.lblSectorGrande
        .Left = 100
        .Top = gridDetalles.Top + Me.gridDetalles.Height
    End With
End Sub


Private Sub EnsureConjuntoStyle()
    Dim fs As GridEX20.JSFormatStyle
    On Error Resume Next
    Set fs = gridDetalles.FormatStyles.item("ConjuntoBold")
    On Error GoTo 0
    If fs Is Nothing Then
        Set fs = gridDetalles.FormatStyles.Add("ConjuntoBold")
        ' Seg�n versi�n, uno de estos compila:
        On Error Resume Next
        fs.FontBold = True
        If Err.Number <> 0 Then
            Err.Clear
         End If
        On Error GoTo 0
    End If
End Sub


Private Sub llenarDataGrid()
    If m_ot Is Nothing Then
        gridDetalles.ItemCount = 0
        Exit Sub
    End If
    If m_ot.detalles Is Nothing Or m_ot.detalles.count = 0 Then
        gridDetalles.ItemCount = 0
        Exit Sub
    End If

    Set detallesPlanos = New Collection
    
    frmAviso.mostrar "Cargando datos..."
    
    construirPlano
       
    gridDetalles.ItemCount = detallesPlanos.count

    On Error Resume Next
    
    gridDetalles.ReBind
    
    gridDetalles.Refresh
    
    On Error GoTo 0
    
    totalizarAvance   ' <-- refresca los totales luego de bind
    
    refrescarAvanceOT
    
End Sub


Private Sub construirPlano()
    Dim d As DetalleOrdenTrabajo
    
   Set detallesPlanos = New Collection

    For Each d In m_ot.detalles
        agregarFilaDetalle d, 0
        If Not d.Pieza Is Nothing Then
            If d.Pieza.EsConjunto Then
                agregarHijos d.Id, d.Pieza.Id, 1, CLng(d.CantidadPedida) 'factor ra�z
            End If
        End If
    Next
End Sub


Private Sub agregarFilaDetalle(ByVal d As DetalleOrdenTrabajo, ByVal Nivel As Integer)
    
    Dim r As clsFilaPlanoRow: Set r = New clsFilaPlanoRow
    
    r.item = CStr(d.item)
    r.IdTabla = d.Id
    r.cantpedida = d.CantidadPedida
    r.Nivel = Nivel

    If Not d.Pieza Is Nothing Then
        r.idPiezaPedido = d.Pieza.Id
        r.nombre = d.Pieza.nombre
        r.UnidadMedida = d.Pieza.UnidadMedida
        r.EsConjunto = d.Pieza.EsConjunto
    Else
        On Error Resume Next
        r.idPiezaPedido = NzLng(d.Id) ' si no existe la prop, quedar� 0
        On Error GoTo 0
        r.nombre = IIf(LenB(NzStr(d.NombrePiezaHistorico)) > 0, d.NombrePiezaHistorico, "Pieza sin cat�logo")
        r.UnidadMedida = "-"
        r.EsConjunto = False
    End If

    ' <<< NUEVO: si NO es conjunto, traer �ltimo avance simple
    If Not r.EsConjunto And r.idPiezaPedido > 0 Then
        Dim sid As Long: sid = NzLng(Me.cboSectores.ItemData(Me.cboSectores.ListIndex))
        
        Dim av As AvanceSimpleDTO
        av = DAOProduccion.FindAvanceSimple(m_ot.Id, r.IdTabla, sid, False) ' True para fallback

        r.CantRecibida = av.CantRecibida
        r.CantFabricada = av.CantFabricada
        r.CantScrap = av.CantScrap
        r.FechaInicio = av.FechaInicio
        r.HoraInicio = av.HoraInicio
        r.FechaFin = av.FechaFin
        r.HoraFin = av.HoraFin
        r.UsuarioRecibio = av.Recibio
        r.Almacen = av.Almacen
        r.ProcesoSiguiente = av.SiguienteProceso
        r.Observaciones = av.Observaciones
        
    End If

    detallesPlanos.Add r
        
End Sub


Private Sub agregarFilaDTO(ByVal dto As DetalleOTConjuntoDTO, _
                           ByVal Nivel As Integer, _
                           ByVal factor As Long)
                          
    Dim r As clsFilaPlanoRow
    Set r = New clsFilaPlanoRow
    
    r.item = CStr(dto.IdentificadorPosicion)
    
    ' id del registro en dpc (detalle del conjunto)
    r.IdTabla = dto.Id
    
    If Not dto.Pieza Is Nothing Then
        r.idPiezaPedido = dto.Pieza.Id
        r.nombre = dto.Pieza.nombre
        r.UnidadMedida = dto.Pieza.UnidadMedida
        r.EsConjunto = dto.Pieza.EsConjunto
    End If
    
    r.cantpedida = dto.CantidadTotalStatic
    r.Nivel = Nivel
    
    r.CantRecibida = dto.CantidadRecibida
    r.CantFabricada = dto.CantidadFabricada
    r.CantScrap = dto.CantidadScrap
    
    r.FechaInicio = dto.FechaInicio
    r.HoraInicio = dto.HoraInicio
    r.FechaFin = dto.FechaFin
    r.HoraFin = dto.HoraFin
    
    r.Almacen = dto.Almacen
    r.UsuarioRecibio = dto.Recibio
    r.ProcesoSiguiente = dto.SiguienteProceso
    
    r.Observaciones = dto.Observaciones
    
    detallesPlanos.Add r
  
End Sub


Private Sub agregarHijos(ByVal idDetallePedido As Long, _
                         ByVal idPiezaPadre As Long, _
                         ByVal Nivel As Integer, _
                         ByVal factor As Long)
                         
    Dim hijos As Collection
    Dim dto As DetalleOTConjuntoDTO
    
    Dim SectorID As Long: SectorID = NzLng(Me.cboSectores.ItemData(Me.cboSectores.ListIndex))
    
    Set hijos = DAOProduccion.FindAllConjuntoProduccion(idDetallePedido, idPiezaPadre, vbNullString, False, 0, SectorID)
    
    If hijos Is Nothing Then Exit Sub

    For Each dto In hijos
        agregarFilaDTO dto, Nivel, factor
        If Not dto.Pieza Is Nothing Then
            If dto.Pieza.EsConjunto Then
                agregarHijos idDetallePedido, dto.Pieza.Id, Nivel + 1, dto.CantidadTotalStatic
            End If
        End If
    Next
End Sub


Private Sub gridDetalles_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If detallesPlanos.count > 0 Then
    
        seleccionarDetalle
        
        If Button = 2 Then
   
            Me.PopupMenu Me.menu

        End If
    End If
End Sub


Private Sub seleccionarDetalle()
    On Error Resume Next
    Set dto = detallesPlanos.item(Me.gridDetalles.RowIndex(Me.gridDetalles.row))

End Sub


Private Sub gridDetalles_RowFormat(RowBuffer As GridEX20.JSRowData)
    On Error Resume Next
    
    Dim v As Variant
    v = RowBuffer.DisplayValue(cEsConjunto)
    
    If Not IsNull(v) Then
        If v = 1 Then
            RowBuffer.RowStyle = "ConjuntoBold"
        Else
            RowBuffer.RowStyle = ""
        End If
    Else
        RowBuffer.RowStyle = ""
    End If
    
    On Error GoTo 0
End Sub

Private Sub gridDetalles_UnboundReadData(ByVal RowIndex As Long, _
    ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)

    If detallesPlanos Is Nothing Then Exit Sub
    Dim i As Long: i = ToCollIndex(RowIndex, detallesPlanos.count)
    If i < 1 Then Exit Sub

    Dim r As clsFilaPlanoRow
    Set r = detallesPlanos.item(i)

    Values(cID) = r.IdTabla
    Values(cItem) = r.item
    Values(cTipo) = IIf(r.EsConjunto, "Conjunto", "Pieza")
    Values(cUM) = NzStr(r.UnidadMedida)
    Values(cNombre) = NzStr(String$(r.Nivel * 3, " ") & NzStr(r.nombre))
    Values(cCantPedida) = r.cantpedida
    Values(cEsConjunto) = IIf(r.EsConjunto, 1, 0)

    If r.EsConjunto Then
        Values(cCantRecibida) = Null
        Values(cCantFabricada) = Null
        Values(cCantScrap) = Null
        Values(cFechaIni) = Null
        Values(cHoraIni) = Null
        Values(cFechaFin) = Null
        Values(cHoraFin) = Null
        Values(cUsuarioRecibio) = Null
        Values(cAlmacen) = Null
        Values(cProcesoSig) = Null
        Values(cObservaciones) = Null
    Else
        Values(cCantRecibida) = r.CantRecibida
        Values(cCantFabricada) = r.CantFabricada
        Values(cCantScrap) = r.CantScrap
        Values(cFechaIni) = IIf(r.FechaInicio = 0, Null, r.FechaInicio)
        Values(cHoraIni) = IIf(r.HoraInicio = 0, Null, r.HoraInicio)
        Values(cFechaFin) = IIf(r.FechaFin = 0, Null, r.FechaFin)
        Values(cHoraFin) = IIf(r.FechaFin = 0, Null, r.HoraFin)
        Values(cUsuarioRecibio) = r.UsuarioRecibio
        Values(cAlmacen) = NzStr(r.Almacen)
        Values(cProcesoSig) = NzStr(r.ProcesoSiguiente)
        Values(cObservaciones) = NzStr(r.Observaciones)
    End If
End Sub


Private Sub gridDetalles_UnboundUpdate(ByVal RowIndex As Long, _
    ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)

    If detallesPlanos Is Nothing Then Exit Sub
    Dim i As Long: i = ToCollIndex(RowIndex, detallesPlanos.count)
    If i < 1 Then Exit Sub

    Dim r As clsFilaPlanoRow
    Set r = detallesPlanos.item(i)

    If r.EsConjunto Then Exit Sub  ' opcional: no guardar conjuntos

    With r
        .IdPedido = m_ot.Id
        .IdSector = NzLng(Me.cboSectores.ItemData(Me.cboSectores.ListIndex))
        .IdTabla = Values(cID)
        .CantRecibida = NzDbl(Values(cCantRecibida))
        .CantFabricada = NzDbl(Values(cCantFabricada))
        .CantScrap = NzDbl(Values(cCantScrap))
        .FechaInicio = NzDate(Values(cFechaIni))
        .HoraInicio = NzDate(Values(cHoraIni))
        .FechaFin = NzDate(Values(cFechaFin))
        .HoraFin = NzDate(Values(cHoraFin))
        .UsuarioRecibio = NzLng(Values(cUsuarioRecibio))   ' <-- FIX
        .Almacen = NzLng(Values(cAlmacen))
        .ProcesoSiguiente = NzStr(Values(cProcesoSig))
        .Observaciones = NzStr(Values(cObservaciones))

    End With
    
    Dim prev As AvanceSimpleDTO
    prev = DAOProduccion.FindAvanceSimple(m_ot.Id, r.IdTabla, r.IdSector, True) ' True=fallback

    If Not DAOProduccion.Save(r) Then
        MsgBox "Hubo un error al guardar"
    Else
        frmAviso.mostrar "Guardando..."
        Call DAOProduccionHistorial.agregar(r, "CARGA DE DATOS", prev)
        
        totalizarAvance   ' <-- vuelve a calcular
        
        refrescarAvanceOT
            
    End If
    
End Sub



' Helpers
Private Function NzLng(v As Variant) As Long
    If IsNull(v) Or v = "" Then NzLng = 0 Else NzLng = CLng(v)
End Function

Private Function NzDbl(v As Variant) As Double
    If IsNull(v) Or v = "" Then NzDbl = 0 Else NzDbl = CDbl(v)
End Function

Private Function NzStr(v As Variant) As String
    If IsNull(v) Then NzStr = "" Else NzStr = CStr(v)
End Function

Private Function NzDate(v As Variant) As Variant
    If IsDate(v) Then NzDate = CDate(v) Else NzDate = Null
End Function

Private Function ToCollIndex(ByVal rowIdx As Long, ByVal n As Long) As Long
    ' Convierte 0/1-based de Janus a 1-based de Collection
    If n <= 0 Then ToCollIndex = 0: Exit Function
    If rowIdx <= 0 Then
        ToCollIndex = 1
    ElseIf rowIdx > n Then
        ToCollIndex = n
    Else
        ToCollIndex = rowIdx
    End If
End Function


Private Sub gridSectores_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex <= sectores.count Then
        Set Sector = sectores.item(RowIndex)
        Values(1) = Sector.Sectorizacion
        Values(2) = Sector.Modulo
    End If
End Sub


Private Sub gridUsuarios_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex <= empleados.count Then
        Set Empleado = empleados.item(RowIndex)
        Values(1) = Empleado.Id
        Values(2) = Empleado.legajo
        Values(3) = UCase(Empleado.NombreCompleto)
    End If
End Sub


Private Sub gridAlmacenes_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex <= almacenes.count Then
        Set Almacen = almacenes.item(RowIndex)
        Values(1) = Almacen.Id
        Values(2) = Almacen.Almacen
    End If
End Sub


Private Sub AADelDetalle_Click()
    seleccionarDetalle
    Dim F As New frmArchivos2
    F.Origen = OrigenArchivos.OA_OrdenesTrabajoDetalle
    F.ObjetoId = dto.IdTabla
    F.caption = "OT N� " & m_ot.IdFormateado & " - Item " & dto.item
    F.Show
End Sub


Private Sub AADLPieza_Click()
    seleccionarDetalle
    Dim F As New frmArchivos2
    F.Origen = OrigenArchivos.OA_Piezas
    F.ObjetoId = dto.idPiezaPedido
    F.caption = "Pieza " & dto.nombre
    F.Show
End Sub


Private Sub IDePieza_Click()
    frmVerIncidencias.referencia = dto.IdTabla
    frmVerIncidencias.Origen = OI_Piezas
    frmVerIncidencias.Show
End Sub


Private Sub IdelDetalle_Click()
    frmVerIncidencias.referencia = dto.idPiezaPedido
    frmVerIncidencias.Origen = OI_OrdenesTrabajoDetalles
    frmVerIncidencias.Show
End Sub


Private Sub menu_desarrollo_Click()
    seleccionarDetalle
    Dim idx As Long
    idx = Me.gridDetalles.RowIndex(Me.gridDetalles.row)
    If idx > 0 Then
        Dim F As New frmDesarrollo
        Load F
        F.CargarPieza dto.idPiezaPedido   'm_ot.Detalles(idx).Pieza.Id

        F.Show
    End If

End Sub


Private Sub menu_historial_Click()
    If dto Is Nothing Then Exit Sub

    Dim SectorID As Long
    SectorID = NzLng(Me.cboSectores.ItemData(Me.cboSectores.ListIndex))

    Dim hist As Collection
    Set hist = DAOProduccionHistorial.GetAllByPieza(dto.IdTabla)

    If hist Is Nothing Or hist.count = 0 Then
        MsgBox "Sin historial para este �tem.", vbInformation
        Exit Sub
    End If

    Set frmHistorialesProduccion.lista = hist
    frmHistorialesProduccion.Show
    
End Sub


Private Sub menu_Observaciones_Click()

    seleccionarDetalle
    Dim idx As Long
    
    idx = Me.gridDetalles.RowIndex(Me.gridDetalles.row)
    
    If idx > 0 Then
        Dim F As New frmPlaneamientoSeguimientoObservaciones
        Load F
        F.CargarObservacion dto.Observaciones

        F.Show
    End If
End Sub


Private Function PorcentajeAvanceOT(Optional ByRef totPed As Double, Optional ByRef totFab As Double) As Double
    totPed = SumarPlano(csCantPedida, True)       ' solo piezas
    totFab = SumarPlano(csCantFabricada, True)    ' solo piezas

    Dim pct As Double
    If totPed <= 0 Then
        pct = 0
    Else
        pct = (totFab / totPed) * 100#
    End If

    ' Clamp para evitar pasar de 0..100 visualmente
    If pct < 0 Then pct = 0
    If pct > 100 Then pct = 100

    PorcentajeAvanceOT = pct
End Function

Private Sub refrescarAvanceOT()
    Dim totPed As Double, totFab As Double
    Dim pct As Double: pct = PorcentajeAvanceOT(totPed, totFab)

    ' Mejor resoluci�n en la barra: 0..1000 (0.1%)
    With pgbAvanceOT
        .min = 0
        .max = 1000
        .value = CLng(pct * 10)   ' 57.3% -> 573
    End With

    Me.lblAvancePct.caption = Format(pct, "0.0") & " %"

    ' Si ya mostr�s estos totales en otras etiquetas, pod�s comentar estas l�neas:
    Me.lblCantPedida.caption = "Cant. Pedida Total: " & Format(totPed, "0.##")
    ' (Opcional) si quer�s un label de fabricada total, crea lblCantFabricadaTotal:
    'Me.lblCantFabricadaTotal.Caption = "Cant. Fabricada Total: " & Format(totFab, "0.##")
End Sub


