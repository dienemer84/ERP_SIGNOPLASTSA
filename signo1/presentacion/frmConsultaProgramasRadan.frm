VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmConsultaProgramasRadan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de programas de Radan©"
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10230
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConsultaProgramasRadan.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   10230
   Begin GridEX20.GridEX grid 
      Height          =   5940
      Left            =   90
      TabIndex        =   3
      Top             =   1485
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   10478
      Version         =   "2.0"
      PreviewRowIndent=   200
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      PreviewColumn   =   "comentario"
      MethodHoldFields=   -1  'True
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      ColumnsCount    =   6
      Column(1)       =   "frmConsultaProgramasRadan.frx":000C
      Column(2)       =   "frmConsultaProgramasRadan.frx":00F8
      Column(3)       =   "frmConsultaProgramasRadan.frx":01F4
      Column(4)       =   "frmConsultaProgramasRadan.frx":0320
      Column(5)       =   "frmConsultaProgramasRadan.frx":040C
      Column(6)       =   "frmConsultaProgramasRadan.frx":050C
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmConsultaProgramasRadan.frx":062C
      FormatStyle(2)  =   "frmConsultaProgramasRadan.frx":0754
      FormatStyle(3)  =   "frmConsultaProgramasRadan.frx":0804
      FormatStyle(4)  =   "frmConsultaProgramasRadan.frx":08B8
      FormatStyle(5)  =   "frmConsultaProgramasRadan.frx":0990
      FormatStyle(6)  =   "frmConsultaProgramasRadan.frx":0A48
      ImageCount      =   0
      PrinterProperties=   "frmConsultaProgramasRadan.frx":0B28
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1275
      Left            =   105
      TabIndex        =   0
      Top             =   60
      Width           =   10005
      _Version        =   786432
      _ExtentX        =   17648
      _ExtentY        =   2249
      _StockProps     =   79
      Caption         =   "Opciones de búsqueda"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.CheckBox chkComentarios 
         Height          =   210
         Left            =   8040
         TabIndex        =   5
         Top             =   270
         Width           =   1800
         _Version        =   786432
         _ExtentX        =   3175
         _ExtentY        =   370
         _StockProps     =   79
         Caption         =   "Mostrar Comentarios"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnBuscar 
         Default         =   -1  'True
         Height          =   465
         Left            =   8565
         TabIndex        =   4
         Top             =   600
         Width           =   1245
         _Version        =   786432
         _ExtentX        =   2196
         _ExtentY        =   820
         _StockProps     =   79
         Caption         =   "Buscar"
         UseVisualStyle  =   -1  'True
      End
      Begin VB.TextBox txtPrograma 
         Height          =   300
         Left            =   1890
         TabIndex        =   2
         Top             =   345
         Width           =   2340
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nombre de programa"
         Height          =   195
         Left            =   285
         TabIndex        =   1
         Top             =   375
         Width           =   1515
      End
   End
End
Attribute VB_Name = "frmConsultaProgramasRadan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private dtosArchivo As Collection
Private dtoArchivo As DTOArchivoOT



Private Sub btnBuscar_Click()
    Set dtosArchivo = New Collection

    Dim filter As String

    If LenB(Me.txtPrograma.text) > 0 Then
        filter = " arch.nombre LIKE '%" & Me.txtPrograma.text & "%.drg'"
    Else
        filter = " arch.nombre LIKE '%.drg'"
    End If


    'busco archivos en piezas
    Dim archivos As Collection
    Dim piezas As Collection
    Dim Pieza As Pieza
    Dim archivo As archivo
    Set archivos = DAOArchivo.FindAll(OA_Piezas, filter)
    If archivos.count = 0 Then
        Set piezas = New Collection
    Else
        Set piezas = DAOPieza.FindAll(FL_0, "s.id IN (" & funciones.JoinCollectionValues(archivos, ", ", "IdReferencia") & ")")
    End If
    Dim Detalles As Collection
    Dim detalle As DetalleOrdenTrabajo
    Dim detallesConjunto As Collection
    Dim detalleConjunto As DetalleOTConjuntoDTO

    Dim piezasId As New Collection
    For Each Pieza In piezas
        piezasId.Add Pieza.Id
    Next Pieza
    If piezasId.count = 0 Then
        Set Detalles = New Collection
        Set detallesConjunto = New Collection
    Else
        Set Detalles = DAODetalleOrdenTrabajo.FindAllByPieza(piezasId)
        Set detallesConjunto = DAODetalleOrdenTrabajo.FindConjuntoByPiezas(piezasId)
    End If


    Dim found As Boolean
    For Each archivo In archivos
        found = False
        Set dtoArchivo = New DTOArchivoOT
        dtoArchivo.Comentario = archivo.Comentario
        dtoArchivo.Programa = archivo.nombre
        dtoArchivo.SubidoPor = archivo.usuario.usuario

        Set Pieza = piezas.item(CStr(archivo.IdReferencia))

        For Each detalle In Detalles
            If detalle.Pieza.Id = Pieza.Id Then
                found = True
                Set dtoArchivo = New DTOArchivoOT
                dtoArchivo.Comentario = archivo.Comentario
                dtoArchivo.Programa = archivo.nombre
                dtoArchivo.SubidoPor = archivo.usuario.usuario

                dtoArchivo.Pieza = detalle.Pieza.nombre
                dtoArchivo.Ot = detalle.OrdenTrabajo.Id
                dtoArchivo.item = detalle.item

                dtosArchivo.Add dtoArchivo
            End If
        Next detalle

        For Each detalleConjunto In detallesConjunto
            If detalleConjunto.Pieza.Id = Pieza.Id Then
                found = True

                Set dtoArchivo = New DTOArchivoOT
                dtoArchivo.Comentario = archivo.Comentario
                dtoArchivo.Programa = archivo.nombre
                dtoArchivo.SubidoPor = archivo.usuario.usuario

                dtoArchivo.Pieza = detalleConjunto.Pieza.nombre
                dtoArchivo.Ot = detalleConjunto.idpedido
                dtoArchivo.item = detalleConjunto.DetalleRaiz.item

                dtosArchivo.Add dtoArchivo
            End If
        Next detalleConjunto

        If Not found Then
            Set dtoArchivo = New DTOArchivoOT
            dtoArchivo.Comentario = archivo.Comentario
            dtoArchivo.Programa = archivo.nombre
            dtoArchivo.SubidoPor = archivo.usuario.usuario

            dtoArchivo.Pieza = Pieza.nombre

            dtosArchivo.Add dtoArchivo
        End If
    Next



    'busco por ot
    Set archivos = DAOArchivo.FindAll(OA_OrdenesTrabajo, filter)
    For Each archivo In archivos
        Set dtoArchivo = New DTOArchivoOT
        dtoArchivo.Comentario = archivo.Comentario
        dtoArchivo.Programa = archivo.nombre
        dtoArchivo.SubidoPor = archivo.usuario.usuario
        dtoArchivo.Ot = archivo.IdReferencia
        dtosArchivo.Add dtoArchivo
    Next archivo


    'busco por detalleot
    Set archivos = DAOArchivo.FindAll(OA_OrdenesTrabajoDetalle, filter)
    If archivos.count > 0 Then
        Set Detalles = DAODetalleOrdenTrabajo.FindAll("dp.id IN (" & funciones.JoinCollectionValues(archivos, ", ", "IdReferencia") & ")")
        'Set detallesConjunto = DAODetalleOrdenTrabajo.FindConjuntoByPiezas(piezasId)
    End If
    For Each archivo In archivos
        Set detalle = Detalles.item(CStr(archivo.IdReferencia))

        Set dtoArchivo = New DTOArchivoOT
        dtoArchivo.Comentario = archivo.Comentario
        dtoArchivo.Programa = archivo.nombre
        dtoArchivo.SubidoPor = archivo.usuario.usuario

        dtoArchivo.Pieza = detalle.Pieza.nombre
        dtoArchivo.Ot = detalle.OrdenTrabajo.Id
        dtoArchivo.item = detalle.item
        dtosArchivo.Add dtoArchivo
    Next archivo


    'busco por detalleconjuntoot
    Set archivos = DAOArchivo.FindAll(OA_OrdenesTrabajoDetalleConjunto, filter)
    If archivos.count > 0 Then
        Set detallesConjunto = DAODetalleOrdenTrabajo.FindAllConjunto(, , "dpc.id IN (" & funciones.JoinCollectionValues(archivos, ", ", "IdReferencia") & ")")
    End If
    For Each archivo In archivos
        Set detalleConjunto = detallesConjunto.item(CStr(archivo.IdReferencia))

        Set dtoArchivo = New DTOArchivoOT
        dtoArchivo.Comentario = archivo.Comentario
        dtoArchivo.Programa = archivo.nombre
        dtoArchivo.SubidoPor = archivo.usuario.usuario

        dtoArchivo.Pieza = detalleConjunto.Pieza.nombre
        dtoArchivo.Ot = detalleConjunto.idpedido
        dtoArchivo.item = detalleConjunto.DetalleRaiz.item
        dtosArchivo.Add dtoArchivo
    Next archivo


    Me.grid.ItemCount = 0
    Me.grid.ItemCount = dtosArchivo.count
End Sub

Private Sub chkComentarios_Click()
    If Me.chkComentarios.value = xtpChecked Then
        Me.grid.PreviewRowLines = 1
    Else
        Me.grid.PreviewRowLines = 0
    End If
End Sub

Private Sub Form_Load()
    Customize Me
    GridEXHelper.CustomizeGrid Me.grid, True
    Me.grid.ItemCount = 0
End Sub


Private Sub grid_ColumnHeaderClick(ByVal Column As GridEX20.JSColumn)
    GridEXHelper.ColumnHeaderClick Me.grid, Column
End Sub

Private Sub grid_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex > 0 And dtosArchivo.count > 0 Then
        Set dtoArchivo = dtosArchivo.item(RowIndex)

        Values(Me.grid.Columns("pieza").Index) = dtoArchivo.Pieza
        Values(Me.grid.Columns("programa").Index) = dtoArchivo.Programa
        Values(Me.grid.Columns("ot").Index) = dtoArchivo.Ot
        Values(Me.grid.Columns("item").Index) = dtoArchivo.item
        Values(Me.grid.Columns("subidopor").Index) = dtoArchivo.SubidoPor
        Values(Me.grid.Columns("comentario").Index) = dtoArchivo.Comentario
    End If

End Sub

