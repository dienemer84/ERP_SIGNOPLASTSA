VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmPlaneamientoSeguimientoAvanzado 
   Caption         =   "Seguimiento de Producción"
   ClientHeight    =   13725
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   16245
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   13725
   ScaleWidth      =   16245
   WindowState     =   2  'Maximized
   Begin GridEX20.GridEX gridSectores 
      Height          =   3495
      Left            =   3720
      TabIndex        =   15
      Top             =   9240
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
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
      Column(1)       =   "frmPlaneamientoSeguimientoAvanzado.frx":0000
      Column(2)       =   "frmPlaneamientoSeguimientoAvanzado.frx":0124
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmPlaneamientoSeguimientoAvanzado.frx":0218
      FormatStyle(2)  =   "frmPlaneamientoSeguimientoAvanzado.frx":0350
      FormatStyle(3)  =   "frmPlaneamientoSeguimientoAvanzado.frx":0400
      FormatStyle(4)  =   "frmPlaneamientoSeguimientoAvanzado.frx":04B4
      FormatStyle(5)  =   "frmPlaneamientoSeguimientoAvanzado.frx":058C
      FormatStyle(6)  =   "frmPlaneamientoSeguimientoAvanzado.frx":0644
      ImageCount      =   0
      PrinterProperties=   "frmPlaneamientoSeguimientoAvanzado.frx":0724
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
      ReplaceColumnIndex=   "usuario"
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
      ColumnsCount    =   2
      Column(1)       =   "frmPlaneamientoSeguimientoAvanzado.frx":08FC
      Column(2)       =   "frmPlaneamientoSeguimientoAvanzado.frx":09FC
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmPlaneamientoSeguimientoAvanzado.frx":0AF0
      FormatStyle(2)  =   "frmPlaneamientoSeguimientoAvanzado.frx":0C28
      FormatStyle(3)  =   "frmPlaneamientoSeguimientoAvanzado.frx":0CD8
      FormatStyle(4)  =   "frmPlaneamientoSeguimientoAvanzado.frx":0D8C
      FormatStyle(5)  =   "frmPlaneamientoSeguimientoAvanzado.frx":0E64
      FormatStyle(6)  =   "frmPlaneamientoSeguimientoAvanzado.frx":0F1C
      ImageCount      =   0
      PrinterProperties=   "frmPlaneamientoSeguimientoAvanzado.frx":0FFC
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
      Caption         =   "GroupBox2"
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
         PreviewColumn   =   "id"
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
         ColumnsCount    =   14
         Column(1)       =   "frmPlaneamientoSeguimientoAvanzado.frx":11D4
         Column(2)       =   "frmPlaneamientoSeguimientoAvanzado.frx":1348
         Column(3)       =   "frmPlaneamientoSeguimientoAvanzado.frx":14A8
         Column(4)       =   "frmPlaneamientoSeguimientoAvanzado.frx":1608
         Column(5)       =   "frmPlaneamientoSeguimientoAvanzado.frx":1790
         Column(6)       =   "frmPlaneamientoSeguimientoAvanzado.frx":18D0
         Column(7)       =   "frmPlaneamientoSeguimientoAvanzado.frx":1A58
         Column(8)       =   "frmPlaneamientoSeguimientoAvanzado.frx":1BB4
         Column(9)       =   "frmPlaneamientoSeguimientoAvanzado.frx":1D10
         Column(10)      =   "frmPlaneamientoSeguimientoAvanzado.frx":1E5C
         Column(11)      =   "frmPlaneamientoSeguimientoAvanzado.frx":1FD0
         Column(12)      =   "frmPlaneamientoSeguimientoAvanzado.frx":2134
         Column(13)      =   "frmPlaneamientoSeguimientoAvanzado.frx":22AC
         Column(14)      =   "frmPlaneamientoSeguimientoAvanzado.frx":243C
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmPlaneamientoSeguimientoAvanzado.frx":2560
         FormatStyle(2)  =   "frmPlaneamientoSeguimientoAvanzado.frx":2698
         FormatStyle(3)  =   "frmPlaneamientoSeguimientoAvanzado.frx":2748
         FormatStyle(4)  =   "frmPlaneamientoSeguimientoAvanzado.frx":27FC
         FormatStyle(5)  =   "frmPlaneamientoSeguimientoAvanzado.frx":28D4
         FormatStyle(6)  =   "frmPlaneamientoSeguimientoAvanzado.frx":298C
         ImageCount      =   0
         PrinterProperties=   "frmPlaneamientoSeguimientoAvanzado.frx":2A6C
      End
      Begin XtremeSuiteControls.Label lblSectorGrande 
         Height          =   495
         Left            =   120
         TabIndex        =   13
         Top             =   6360
         Width           =   5415
         _Version        =   786432
         _ExtentX        =   9551
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "..."
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
      Caption         =   "Párametros de búsqueda"
      UseVisualStyle  =   -1  'True
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
         Left            =   6720
         TabIndex        =   1
         Tag             =   "Datos de la Orden de Trabajo Nº "
         Top             =   240
         Width           =   9075
         _Version        =   786432
         _ExtentX        =   16007
         _ExtentY        =   1508
         _StockProps     =   79
         Caption         =   "Datos de la Orden de Trabajo Nº "
         UseVisualStyle  =   -1  'True
         Begin VB.Label lblEstado 
            AutoSize        =   -1  'True
            Caption         =   "Estado:"
            Height          =   195
            Left            =   4830
            TabIndex        =   5
            Top             =   240
            Width           =   540
         End
         Begin VB.Label lblFechaEntrega 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Entrega:"
            Height          =   195
            Left            =   4305
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
      Begin VB.Menu menu_historial 
         Caption         =   "Ver Historial"
      End
      Begin VB.Menu menu_desarrollo 
         Caption         =   "Ver Desarollo"
      End
      Begin VB.Menu menu_archivos_asociados 
         Caption         =   "Ver Archivos Asociados"
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
Private Usuario As clsUsuario
Private usuarios As New Collection
Private Sector As clsSector
Private sectores As New Collection
Dim col As Collection
Private detallesPlanos As Collection
Dim dto As clsFilaPlanoRow
Private m_rows As New Collection    'colección de clsFilaPlanoRow

Public Enum Cols
  cID = 1
  cItem = 2
  cTipo = 3
  cUM = 4
  cNombre = 5
  cCantPedida = 6
  cCantRecibida = 7
  cCantFabricada = 8
  cCantScrap = 9
  cFechaInicio = 10
  cFechaFin = 11
  cUsuarioRecibio = 12
  cProcesoSig = 13
  cEsConjunto = 14   ' oculto
End Enum


Private Sub btnCargarSector_Click(): llenarDataGrid: End Sub

Private Sub cboSectores_Click()

    llenarDataGrid
    
    Me.lblSectorGrande.caption = "SECTOR: " & Me.cboSectores.Text

End Sub


Private Sub cmdBuscar_Click()

    If Not IsNumeric(Me.txtOTNro.Text) Then Exit Sub
    Set m_ot = DAOOrdenTrabajo.FindById(Me.txtOTNro.Text)
    If m_ot Is Nothing Then
        MsgBox "La Orden de Trabajo Nº " & Me.txtOTNro.Text & " no existe.", vbInformation + vbOKOnly
        Exit Sub
    End If
    Set m_ot.detalles = DAODetalleOrdenTrabajo.FindAllByOrdenTrabajo(m_ot.Id)
    
    CargarDetallesOT
    
    llenarDataGrid              ' <<< volver a cargar
    
End Sub


Private Sub CargarDetallesOT()
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

    Set usuarios = DAOUsuarios.GetAll()
    Me.gridUsuarios.ItemCount = usuarios.count
    Set Me.gridDetalles.Columns("recibio").DropDownControl = Me.gridUsuarios

   DAOSectores.LlenarComboXtremeModulos Me.cboSectores
   Me.cboSectores.ListIndex = 0
    
    Set detallesPlanos = New Collection
    
End Sub


Private Sub Form_Resize()
   On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    
    With fraDatosOT
        .Left = 5500
        .Top = 100
        .Width = 7500
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
        If gridDetalles.Columns.count >= cCantScrap Then
            gridDetalles.Columns(cID).Width = 500
            gridDetalles.Columns(cItem).Width = 500
            gridDetalles.Columns(cUM).Width = 500
            gridDetalles.Columns(cNombre).Width = 4000
            gridDetalles.Columns(cCantPedida).Width = 800
            gridDetalles.Columns(cCantFabricada).Width = 800
            gridDetalles.Columns(cCantRecibida).Width = 800
            gridDetalles.Columns(cCantScrap).Width = 800
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
        ' Según versión, uno de estos compila:
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
    
    ConstruirPlano
       
    gridDetalles.ItemCount = detallesPlanos.count

    On Error Resume Next
    gridDetalles.ReBind
    gridDetalles.Refresh
    On Error GoTo 0
    
End Sub


Private Sub ConstruirPlano()
    Dim d As DetalleOrdenTrabajo
    
   Set detallesPlanos = New Collection

    For Each d In m_ot.detalles
        AgregarFilaDetalle d, 0
        If Not d.Pieza Is Nothing Then
            If d.Pieza.EsConjunto Then
                AgregarHijos d.Id, d.Pieza.Id, 1, CLng(d.CantidadPedida) 'factor raíz
            End If
        End If
    Next
End Sub


Private Sub AgregarFilaDetalle(ByVal d As DetalleOrdenTrabajo, ByVal Nivel As Integer)
    
    Dim r As clsFilaPlanoRow: Set r = New clsFilaPlanoRow
    
    r.item = CStr(d.item)
    r.IdTabla = d.Id
    r.CantPedida = d.CantidadPedida
    r.Nivel = Nivel

    If Not d.Pieza Is Nothing Then
        r.IdPiezaPedido = d.Pieza.Id
        r.nombre = d.Pieza.nombre
        r.UnidadMedida = d.Pieza.UnidadMedida
        r.EsConjunto = d.Pieza.EsConjunto
    Else
        On Error Resume Next
        r.IdPiezaPedido = NzLng(d.Id) ' si no existe la prop, quedará 0
        On Error GoTo 0
        r.nombre = IIf(LenB(NzStr(d.NombrePiezaHistorico)) > 0, d.NombrePiezaHistorico, "Pieza sin catálogo")
        r.UnidadMedida = "-"
        r.EsConjunto = False
    End If

    ' <<< NUEVO: si NO es conjunto, traer último avance simple
    If Not r.EsConjunto And r.IdPiezaPedido > 0 Then
        Dim sid As Long: sid = NzLng(Me.cboSectores.ItemData(Me.cboSectores.ListIndex))
        
        Dim av As AvanceSimpleDTO
        av = DAOProduccion.FindAvanceSimple(m_ot.Id, r.IdTabla, sid, False) ' True para fallback

        r.CantRecibida = av.CantRecibida
        r.CantFabricada = av.CantFabricada
        r.CantScrap = av.CantScrap
        r.FechaInicio = av.FechaInicio
        r.FechaFin = av.FechaFin
        r.UsuarioRecibio = av.Recibio
        r.ProcesoSiguiente = av.SiguienteProceso
    End If

    detallesPlanos.Add r
        
End Sub


Private Sub AgregarFilaDTO(ByVal dto As DetalleOTConjuntoDTO, _
                           ByVal Nivel As Integer, _
                           ByVal factor As Long)
                          
    Dim r As clsFilaPlanoRow
    Set r = New clsFilaPlanoRow
    
    r.item = CStr(dto.IdentificadorPosicion)
    
    ' id del registro en dpc (detalle del conjunto)
    r.IdTabla = dto.Id
    
    If Not dto.Pieza Is Nothing Then
        r.IdPiezaPedido = dto.Pieza.Id
        r.nombre = dto.Pieza.nombre
        r.UnidadMedida = dto.Pieza.UnidadMedida
        r.EsConjunto = dto.Pieza.EsConjunto
    End If
    
    r.CantPedida = dto.CantidadTotalStatic
    r.Nivel = Nivel
    
    r.CantRecibida = dto.CantidadRecibida
    r.CantFabricada = dto.CantidadFabricada
    r.CantScrap = dto.CantidadScrap
    
    r.FechaInicio = dto.FechaInicio
    r.FechaFin = dto.FechaFin
    
    r.UsuarioRecibio = dto.Recibio
    r.ProcesoSiguiente = dto.SiguienteProceso
    
    detallesPlanos.Add r
  
End Sub


Private Sub AgregarHijos(ByVal idDetallePedido As Long, _
                         ByVal idPiezaPadre As Long, _
                         ByVal Nivel As Integer, _
                         ByVal factor As Long)
                         
    Dim hijos As Collection
    Dim dto As DetalleOTConjuntoDTO
    
    Dim SectorID As Long: SectorID = NzLng(Me.cboSectores.ItemData(Me.cboSectores.ListIndex))
    
    Set hijos = DAOProduccion.FindAllConjuntoProduccion(idDetallePedido, idPiezaPadre, vbNullString, False, 0, SectorID)
    
    If hijos Is Nothing Then Exit Sub

    For Each dto In hijos
        AgregarFilaDTO dto, Nivel, factor
        If Not dto.Pieza Is Nothing Then
            If dto.Pieza.EsConjunto Then
                AgregarHijos idDetallePedido, dto.Pieza.Id, Nivel + 1, dto.CantidadTotalStatic
            End If
        End If
    Next
End Sub


Private Sub gridDetalles_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If detallesPlanos.count > 0 Then
    
        SeleccionarDetalle
        
        If Button = 2 Then
   
            Me.PopupMenu Me.menu

        End If
    End If
End Sub


Private Sub SeleccionarDetalle()
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
    Values(cCantPedida) = r.CantPedida
    Values(cEsConjunto) = IIf(r.EsConjunto, 1, 0)

    If r.EsConjunto Then
        Values(cCantRecibida) = Null
        Values(cCantFabricada) = Null
        Values(cCantScrap) = Null
        Values(cFechaInicio) = Null
        Values(cFechaFin) = Null
        Values(cUsuarioRecibio) = Null
        Values(cProcesoSig) = Null
    Else
        Values(cCantRecibida) = r.CantRecibida
        Values(cCantFabricada) = r.CantFabricada
        Values(cCantScrap) = r.CantScrap
        Values(cFechaInicio) = IIf(r.FechaInicio = 0, Null, r.FechaInicio)
        Values(cFechaFin) = IIf(r.FechaFin = 0, Null, r.FechaFin)
        Values(cUsuarioRecibio) = r.UsuarioRecibio
        Values(cProcesoSig) = NzStr(r.ProcesoSiguiente)
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
        .idSector = NzLng(Me.cboSectores.ItemData(Me.cboSectores.ListIndex))
        .IdTabla = Values(cID)
        .CantRecibida = NzDbl(Values(cCantRecibida))
        .CantFabricada = NzDbl(Values(cCantFabricada))
        .CantScrap = NzDbl(Values(cCantScrap))
        .FechaInicio = NzDate(Values(cFechaInicio))
        .FechaFin = NzDate(Values(cFechaFin))
        .UsuarioRecibio = NzLng(Values(cUsuarioRecibio))   ' <-- FIX
        .ProcesoSiguiente = NzStr(Values(cProcesoSig))

    End With
    
    Dim prev As AvanceSimpleDTO
    prev = DAOProduccion.FindAvanceSimple(m_ot.Id, r.IdTabla, r.idSector, True) ' True=fallback

    If Not DAOProduccion.Save(r) Then
        MsgBox "Hubo un error al guardar"
    Else
        frmAviso.mostrar "Guardando..."
        Call DAOProduccionHistorial.Agregar(r, "UPDATE", "DATO DE PRODUCCIÓN CARGADO", prev)
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
    If RowIndex <= usuarios.count Then
        Set Usuario = usuarios.item(RowIndex)
        Values(1) = Usuario.Id
        Values(2) = Usuario.Usuario
    End If
End Sub


Private Sub menu_historial_Click()
    If dto Is Nothing Then Exit Sub

    Dim SectorID As Long
    SectorID = NzLng(Me.cboSectores.ItemData(Me.cboSectores.ListIndex))

    Dim hist As Collection
    ' Ajustá la firma según tu DAO: conviene filtrar por pedido, detalle, pieza y sector
    Set hist = DAOProduccionHistorial.GetAllByPieza(dto.IdTabla)

    If hist Is Nothing Or hist.count = 0 Then
        MsgBox "Sin historial para este ítem.", vbInformation
        Exit Sub
    End If

    ' Si la propiedad es objeto, usá Set
    Set frmHistorialesProduccion.lista = hist
    frmHistorialesProduccion.Show
    
End Sub
