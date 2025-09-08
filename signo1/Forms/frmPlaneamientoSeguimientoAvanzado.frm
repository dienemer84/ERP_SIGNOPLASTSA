VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmPlaneamientoSeguimientoAvanzado 
   Caption         =   "Seguimiento de Producción"
   ClientHeight    =   9105
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16245
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9105
   ScaleWidth      =   16245
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1335
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   16455
      _Version        =   786432
      _ExtentX        =   29025
      _ExtentY        =   2355
      _StockProps     =   79
      Caption         =   "Párametros de búsqueda"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.GroupBox fraDatosOT 
         Height          =   855
         Left            =   6720
         TabIndex        =   3
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
            TabIndex        =   7
            Top             =   240
            Width           =   540
         End
         Begin VB.Label lblFechaEntrega 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Entrega:"
            Height          =   195
            Left            =   4305
            TabIndex        =   6
            Top             =   525
            Width           =   1095
         End
         Begin VB.Label lblFechaCreado 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Creada:"
            Height          =   195
            Left            =   165
            TabIndex        =   5
            Top             =   525
            Width           =   1050
         End
         Begin VB.Label lblCliente 
            AutoSize        =   -1  'True
            Caption         =   "Cliente:"
            Height          =   195
            Left            =   165
            TabIndex        =   4
            Top             =   240
            Width           =   525
         End
      End
      Begin XtremeSuiteControls.PushButton cmdBuscar 
         Height          =   345
         Left            =   3600
         TabIndex        =   8
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
         TabIndex        =   9
         Top             =   262
         Width           =   1935
         _Version        =   786432
         _ExtentX        =   3413
         _ExtentY        =   529
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   1
      End
      Begin VB.Label lblOT 
         AutoSize        =   -1  'True
         Caption         =   "Orden de Trabajo"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   315
         Width           =   1245
      End
   End
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   7455
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   16575
      _Version        =   786432
      _ExtentX        =   29236
      _ExtentY        =   13150
      _StockProps     =   68
      Color           =   64
      PaintManager.Position=   2
      PaintManager.BoldSelected=   -1  'True
      PaintManager.MultiRowFixedSelection=   -1  'True
      ItemCount       =   5
      Item(0).Caption =   "CORTE"
      Item(0).ControlCount=   3
      Item(0).Control(0)=   "gridDetalles"
      Item(0).Control(1)=   "btnGuardar"
      Item(0).Control(2)=   "btnExportar"
      Item(1).Caption =   "PLEGADO"
      Item(1).ControlCount=   0
      Item(2).Caption =   "HERRERIA"
      Item(2).ControlCount=   0
      Item(3).Caption =   "PINTURA"
      Item(3).ControlCount=   0
      Item(4).Caption =   "TERMINACION"
      Item(4).ControlCount=   0
      Begin XtremeSuiteControls.PushButton btnExportar 
         Height          =   495
         Left            =   11760
         TabIndex        =   12
         Top             =   6360
         Width           =   1815
         _Version        =   786432
         _ExtentX        =   3201
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Exportar"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnGuardar 
         Height          =   495
         Left            =   13800
         TabIndex        =   11
         Top             =   6360
         Width           =   1935
         _Version        =   786432
         _ExtentX        =   3413
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Guardar"
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
      Begin GridEX20.GridEX gridDetalles 
         Height          =   5895
         Left            =   120
         TabIndex        =   1
         Top             =   240
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
         HideSelection   =   1
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
         ColumnsCount    =   13
         Column(1)       =   "frmPlaneamientoSeguimientoAvanzado.frx":0000
         Column(2)       =   "frmPlaneamientoSeguimientoAvanzado.frx":01AC
         Column(3)       =   "frmPlaneamientoSeguimientoAvanzado.frx":030C
         Column(4)       =   "frmPlaneamientoSeguimientoAvanzado.frx":0450
         Column(5)       =   "frmPlaneamientoSeguimientoAvanzado.frx":0590
         Column(6)       =   "frmPlaneamientoSeguimientoAvanzado.frx":073C
         Column(7)       =   "frmPlaneamientoSeguimientoAvanzado.frx":0898
         Column(8)       =   "frmPlaneamientoSeguimientoAvanzado.frx":09F4
         Column(9)       =   "frmPlaneamientoSeguimientoAvanzado.frx":0B34
         Column(10)      =   "frmPlaneamientoSeguimientoAvanzado.frx":0CF8
         Column(11)      =   "frmPlaneamientoSeguimientoAvanzado.frx":0E5C
         Column(12)      =   "frmPlaneamientoSeguimientoAvanzado.frx":0FA4
         Column(13)      =   "frmPlaneamientoSeguimientoAvanzado.frx":1104
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmPlaneamientoSeguimientoAvanzado.frx":1278
         FormatStyle(2)  =   "frmPlaneamientoSeguimientoAvanzado.frx":13B0
         FormatStyle(3)  =   "frmPlaneamientoSeguimientoAvanzado.frx":1460
         FormatStyle(4)  =   "frmPlaneamientoSeguimientoAvanzado.frx":1514
         FormatStyle(5)  =   "frmPlaneamientoSeguimientoAvanzado.frx":15EC
         FormatStyle(6)  =   "frmPlaneamientoSeguimientoAvanzado.frx":16A4
         ImageCount      =   0
         PrinterProperties=   "frmPlaneamientoSeguimientoAvanzado.frx":1784
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
Dim col As Collection
Private detallesPlanos As Collection

Private m_rows As New Collection    'colección de clsFilaPlanoRow

Public Enum Cols
  cId = 1
  cItem = 2
  cUM = 3
  cNombre = 4
  cCantPedida = 5
  cCantRecibida = 6
  cCantFabricada = 7
  cCantScrap = 8
  cFechaInicio = 9
  cFechaFin = 10
  cUsuarioRecibio = 11
  cProcesoSig = 12
  cEsConjunto = 13  '<< OCULTA
End Enum




Private Sub cmdBuscar_Click()
    If Not IsNumeric(Me.txtOTNro.Text) Then Exit Sub
    
    Set m_ot = DAOOrdenTrabajo.FindById(Me.txtOTNro.Text)       'me la recargo por las dudas
    
    Set m_ot.detalles = DAODetalleOrdenTrabajo.FindAllByOrdenTrabajo(m_ot.Id)

   
    If m_ot Is Nothing Then
        MsgBox "La Orden de Trabajo Nº " & Me.txtOTNro.Text & " no existe.", vbInformation + vbOKOnly
    Else
    
        CargarDetallesOT
        
        llenarDataGrid
        
    End If
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

    EnsureConjuntoStyle          '<< agrega el estilo
    
    Me.txtOTNro = "5964"
    
    cmdBuscar_Click
    
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
    If m_ot Is Nothing Or m_ot.detalles Is Nothing Then
        gridDetalles.ItemCount = 0
        Exit Sub
    End If

    ConstruirPlano

    gridDetalles.ItemCount = detallesPlanos.count
    On Error Resume Next: gridDetalles.ReBind: On Error GoTo 0
    gridDetalles.Refresh
    
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
    Dim r As clsFilaPlanoRow
    Set r = New clsFilaPlanoRow
    
    r.item = CStr(d.item)
    
    If Not d.Pieza Is Nothing Then r.Id = d.Pieza.Id
    r.nombre = d.Pieza.nombre
    r.UnidadMedida = d.Pieza.UnidadMedida
    r.CantPedida = d.CantidadPedida
    r.Nivel = Nivel
    r.EsConjunto = (Not d.Pieza Is Nothing And d.Pieza.EsConjunto)
    
    detallesPlanos.Add r
    

    
End Sub


Private Sub AgregarFilaDTO(ByVal dto As DetalleOTConjuntoDTO, _
                           ByVal Nivel As Integer, _
                           ByVal factor As Long)
                           
    Dim r As clsFilaPlanoRow
    Set r = New clsFilaPlanoRow
    
    r.item = CStr(dto.IdentificadorPosicion)
    
    If Not dto.Pieza Is Nothing Then r.Id = dto.Pieza.Id
    r.Id = dto.Pieza.Id
    r.nombre = dto.Pieza.nombre
    r.UnidadMedida = dto.Pieza.UnidadMedida
    r.CantPedida = CLng(dto.Cantidad) * CLng(factor)   'multiplicado por el padre
    r.Nivel = Nivel
    
    r.EsConjunto = (Not dto.Pieza Is Nothing And dto.Pieza.EsConjunto)
    
    Debug.Print ("Pieza Hijo: " & dto.Pieza.Id)
     
    detallesPlanos.Add r
  

End Sub


Private Sub AgregarHijos(ByVal idDetallePedido As Long, _
                         ByVal idPiezaPadre As Long, _
                         ByVal Nivel As Integer, _
                         ByVal factor As Long)
    
    Dim hijos As Collection
    Dim dto As DetalleOTConjuntoDTO

    Set hijos = DAODetalleOrdenTrabajo.FindAllConjunto(idDetallePedido, idPiezaPadre)
  
    If hijos Is Nothing Then Exit Sub

    For Each dto In hijos
        AgregarFilaDTO dto, Nivel, factor
        If Not dto.Pieza Is Nothing Then
            If dto.Pieza.EsConjunto Then
                AgregarHijos idDetallePedido, dto.Pieza.Id, Nivel + 1, CLng(factor) * CLng(dto.Cantidad)
                

            End If
        End If
    Next
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
        .Height = 1000
    End With
    
    ' TabControl debajo del GroupBox
    With TabControl1
        .Left = 100
        .Top = GroupBox1.Top + GroupBox1.Height + 100
        .Width = Me.ScaleWidth - 200
        .Height = Me.ScaleHeight - 1200
    End With
    
    ' Grid ajustado al resto de la ventana
    With gridDetalles
        Top = TabControl1.Top + TabControl1.Height + 100
        .Width = TabControl1.Width - 400
        .Height = TabControl1.Height - 1600
    End With
    
    With Me.btnExportar
        .Left = gridDetalles.Width - 3800
        .Top = gridDetalles.Height + 500

    End With
    
    With Me.btnGuardar
        .Left = gridDetalles.Width - 1800
        .Top = gridDetalles.Height + 500

    End With

End Sub


Private Sub gridDetalles_RowFormat(RowBuffer As GridEX20.JSRowData)

    If RowBuffer.DisplayValue(cEsConjunto) = 1 Then
        RowBuffer.RowStyle = "ConjuntoBold"   'aplica negrita a toda la fila
    Else
        RowBuffer.RowStyle = ""               'sin estilo
    End If

End Sub



Private Sub gridDetalles_UnboundReadData(ByVal RowIndex As Long, _
                                         ByVal Bookmark As Variant, _
                                         ByVal Values As GridEX20.JSRowData)
                                         
    Dim r As clsFilaPlanoRow
    Set r = detallesPlanos.item(RowIndex)
    
    Values(cId) = r.Id
    Values(cItem) = r.item
    Values(cUM) = r.UnidadMedida
    Values(cNombre) = String$(r.Nivel * 3, " ") & r.nombre
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
        Values(cProcesoSig) = r.ProcesoSiguiente
    End If

End Sub


Private Sub gridDetalles_UnboundUpdate(ByVal RowIndex As Long, _
                                       ByVal Bookmark As Variant, _
                                       ByVal Values As GridEX20.JSRowData)
                                       
      If RowIndex < 1 Or RowIndex > detallesPlanos.count Then Exit Sub
      Dim r As clsFilaPlanoRow: Set r = detallesPlanos.item(RowIndex)
    
'''      'bloquear conjuntos
'''      If Values(cEsConjunto) = 1 Then
'''        Values(cCantFabricada) = r.CantFabricada
'''        Values(cCantRecibida) = r.CantRecibida
'''        Values(cCantScrap) = r.CantScrap
'''        Exit Sub
'''      End If
    
      'helpers
      Dim v As Variant
      v = Values(cCantFabricada): If IsNull(v) Or v = "" Then v = 0
      If Not IsNumeric(v) Or v < 0 Then Values(cCantFabricada) = r.CantFabricada Else r.CantFabricada = CCur(v)
    
      v = Values(cCantRecibida): If IsNull(v) Or v = "" Then v = 0
      If Not IsNumeric(v) Or v < 0 Then Values(cCantRecibida) = r.CantRecibida Else r.CantRecibida = CCur(v)
    
      v = Values(cCantScrap): If IsNull(v) Or v = "" Then v = 0
      If Not IsNumeric(v) Or v < 0 Then Values(cCantScrap) = r.CantScrap Else r.CantScrap = CCur(v)
    
      If IsDate(Values(cFechaInicio)) Then r.FechaInicio = CDate(Values(cFechaInicio))
      If IsDate(Values(cFechaFin)) Then r.FechaFin = CDate(Values(cFechaFin))
    
      r.UsuarioRecibio = CStr(Values(cUsuarioRecibio))
      r.ProcesoSiguiente = CStr(Values(cProcesoSig))
    
'''      'reflejar normalizado
'''      Values(cCantFabricada) = r.CantFabricada
'''      Values(cCantRecibida) = r.CantRecibida
'''      Values(cCantScrap) = r.CantScrap
'''      Values(cFechaInicio) = r.FechaInicio
'''      Values(cFechaFin) = r.FechaFin
'''
'''      'persistencia (ajusta firma DAO)
      
      On Error Resume Next
      
        Dim msg As String
        msg = "Fila: " & RowIndex & vbCrLf & _
              "ID pieza: " & r.Id & "   Item: " & r.item & vbCrLf & _
              "Nombre: " & r.nombre & vbCrLf & _
              "Cant. pedida: " & Format(r.CantPedida, "0.##") & vbCrLf & _
              "Cant. recibida: " & Format(r.CantRecibida, "0.##") & vbCrLf & _
              "Cant. fabricada: " & Format(r.CantFabricada, "0.##") & vbCrLf & _
              "Scrap: " & Format(r.CantScrap, "0.##") & vbCrLf & _
              "Inicio: " & IIf(r.FechaInicio = 0, "-", Format(r.FechaInicio, "dd/mm/yyyy hh:nn")) & vbCrLf & _
              "Fin: " & IIf(r.FechaFin = 0, "-", Format(r.FechaFin, "dd/mm/yyyy hh:nn")) & vbCrLf & _
              "Recibió: " & r.UsuarioRecibio & vbCrLf & _
              "Siguiente proceso: " & r.ProcesoSiguiente

        MsgBox msg, vbInformation, "Datos ingresados"
      On Error GoTo 0
    
End Sub


Private Sub btnGuardar_Click()
    On Error GoTo err1

    If gridDetalles.EditMode = jgexEditModeOn Then
        gridDetalles.Update
        If gridDetalles.EditMode = jgexEditModeOn Then
            MsgBox "Termine de editar la celda.", vbExclamation
            Exit Sub
        End If
    End If

    If DAOProduccion.SaveMany(m_rows) Then
        MsgBox "Registros guardados.", vbInformation
    Else
        Err.Raise 9001, "btnGuardar_Click", DAOProduccion.LastError
    End If
    Exit Sub
err1:
    MsgBox "Error al guardar (" & Err.Number & "): " & Err.Description, vbCritical
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
    
