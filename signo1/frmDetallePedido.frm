VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmPlaneamientoPedidosDetalle 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11190
   ClipControls    =   0   'False
   Icon            =   "frmDetallePedido.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   11190
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "[ Condiciones Comerciales ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   5640
      TabIndex        =   22
      Top             =   4440
      Width           =   5415
      Begin VB.Label lblDetalle 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1200
         TabIndex        =   32
         Top             =   480
         Width           =   3975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         Caption         =   "Detalle"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblCliente 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1200
         TabIndex        =   30
         Top             =   240
         Width           =   3975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         Caption         =   "Anticipo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblAnticipo 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1200
         TabIndex        =   27
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         Caption         =   "F.P. Ant."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         Caption         =   "F.P. Saldo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label lblFormaPagoSaldo 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1200
         TabIndex        =   24
         Top             =   1200
         Width           =   3975
      End
      Begin VB.Label lblFormaPagoAnticipo 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1200
         TabIndex        =   23
         Top             =   960
         Width           =   3975
      End
   End
   Begin GridEX20.GridEX grilla 
      Height          =   4215
      Left            =   105
      TabIndex        =   21
      Top             =   120
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   7435
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      PreviewColumn   =   "pieza"
      PreviewRowLines =   1
      ColumnAutoResize=   -1  'True
      MethodHoldFields=   -1  'True
      ContScroll      =   -1  'True
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      BackColorHeader =   16744576
      ImageCount      =   1
      ImagePicture1   =   "frmDetallePedido.frx":000C
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   10
      Column(1)       =   "frmDetallePedido.frx":0326
      Column(2)       =   "frmDetallePedido.frx":0482
      Column(3)       =   "frmDetallePedido.frx":0576
      Column(4)       =   "frmDetallePedido.frx":069A
      Column(5)       =   "frmDetallePedido.frx":07E2
      Column(6)       =   "frmDetallePedido.frx":08FE
      Column(7)       =   "frmDetallePedido.frx":0A0A
      Column(8)       =   "frmDetallePedido.frx":0B06
      Column(9)       =   "frmDetallePedido.frx":0C0E
      Column(10)      =   "frmDetallePedido.frx":0D16
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmDetallePedido.frx":0E2E
      FormatStyle(2)  =   "frmDetallePedido.frx":0F66
      FormatStyle(3)  =   "frmDetallePedido.frx":1016
      FormatStyle(4)  =   "frmDetallePedido.frx":10CA
      FormatStyle(5)  =   "frmDetallePedido.frx":11A2
      FormatStyle(6)  =   "frmDetallePedido.frx":125A
      ImageCount      =   1
      ImagePicture(1) =   "frmDetallePedido.frx":133A
      PrinterProperties=   "frmDetallePedido.frx":1654
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Materializacion"
      Height          =   375
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "[ Orden de Trabajo ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   1
      Top             =   4440
      Width           =   5415
      Begin VB.Label lblFechaEntrega 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1440
         TabIndex        =   19
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         Caption         =   "Entrega"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   600
         TabIndex        =   18
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label lblFechaAprobado 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3600
         TabIndex        =   17
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label lblAprobador 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1440
         TabIndex        =   16
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblFechaModificado 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3600
         TabIndex        =   15
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label lblModificador 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1440
         TabIndex        =   14
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2880
         TabIndex        =   13
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         Caption         =   "Finalizador"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2880
         TabIndex        =   11
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         Caption         =   "Modificador"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         Caption         =   "Aprobador"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label LblFinalizado 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1440
         TabIndex        =   8
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lblCreador 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1440
         TabIndex        =   7
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         Caption         =   "Creador"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2880
         TabIndex        =   5
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2880
         TabIndex        =   4
         Top             =   480
         Width           =   615
      End
      Begin VB.Label lblFechaCreado 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3600
         TabIndex        =   3
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lblFechaFinalizado 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3600
         TabIndex        =   2
         Top             =   960
         Width           =   1695
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Menu m1 
      Caption         =   "Menu1"
      Visible         =   0   'False
      Begin VB.Menu VerDesarrollo 
         Caption         =   "Ver Desarrollo..."
      End
      Begin VB.Menu archivos 
         Caption         =   "Archivos Asociados..."
         Begin VB.Menu mnuArchivosPieza 
            Caption         =   "De Pieza..."
         End
         Begin VB.Menu mnuArchivosPedido 
            Caption         =   "Del Detalle..."
         End
      End
      Begin VB.Menu scanear 
         Caption         =   "Adquirir..."
         Begin VB.Menu mnuAdquirirPieza 
            Caption         =   "A Pieza..."
         End
         Begin VB.Menu mnuAdquirirPedido 
            Caption         =   "Al Detalle..."
         End
      End
      Begin VB.Menu verIncidencias 
         Caption         =   "Ver Incidencias..."
         Begin VB.Menu mnuInciPieza 
            Caption         =   "De Pieza..."
         End
         Begin VB.Menu mnuInciPedido 
            Caption         =   "Del Detalle..."
         End
      End
   End
End
Attribute VB_Name = "frmPlaneamientoPedidosDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m_Archivos As New Dictionary
Dim m_Archivos_detalle As New Dictionary
Dim idPieza As Long
Dim rectmp As DetalleOrdenTrabajo
Private m_pedido As OrdenTrabajo


Public Property Let Pedido(nvalue As OrdenTrabajo)
    Set m_pedido = nvalue
End Property

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    frmMaterializacion.Id = m_pedido.Id
    frmMaterializacion.Ot = True
    frmMaterializacion.otro = False
    frmMaterializacion.presu = False
    frmMaterializacion.Show

End Sub

Private Sub mostrarDetalles()
    Me.lblAnticipo = IIf(m_pedido.Anticipo = 0, "Sin Anticipo", m_pedido.Anticipo & "%")
    Me.lblFormaPagoAnticipo = m_pedido.CondicionesComercialesAnticipo
    Me.lblFormaPagoSaldo = m_pedido.CondicionesComercialesSaldo
    Me.lblCliente = m_pedido.cliente.razon
    Me.lblDetalle = m_pedido.descripcion


    Dim creador As String, aprobador As String, modificador As String, finalizador As String
    creador = m_pedido.usuario.usuario

    If m_pedido.UsuarioAprobado Is Nothing Then
        aprobador = vbNullString
    Else
        aprobador = m_pedido.UsuarioAprobado.usuario
    End If
    If m_pedido.UsuarioModificado Is Nothing Then
        modificador = vbNullString
    Else
        modificador = m_pedido.usuario.usuario
    End If
    If m_pedido.UsuarioFinalizado Is Nothing Then
        finalizador = vbNullString
    Else
        finalizador = m_pedido.UsuarioFinalizado.usuario
    End If
    Me.lblFechaCreado = m_pedido.fechaCreado
    Me.lblFechaEntrega = m_pedido.FechaEntrega
    If aprobador = vbNullString Then    'no está aprobado
        Me.lblAprobador = "No Aprobado"
        Me.lblFechaAprobado = "No Aprobado"
    Else
        Me.lblAprobador = aprobador
        Me.lblFechaAprobado = m_pedido.fechaAprobado
    End If
    If finalizador = vbNullString Then    'finalizador
        Me.LblFinalizado = "En proceso"
        Me.lblFechaFinalizado = "En proceso"
    Else
        Me.LblFinalizado = finalizador
        Me.lblFechaFinalizado = m_pedido.FechaCerrado
    End If
    If creador = vbNullString Then    'no está creado (ERROR EN BBDD)
        Me.lblCreador = "Error en BBDD"
        Me.lblFechaCreado = "Error en BBDD"
    Else
        Me.lblCreador = creador
        Me.lblModificador = m_pedido.fechaCreado
    End If
    If modificador = vbNullString Then    'no esta modificado
        Me.lblModificador = "No Modificado"
        Me.lblFechaModificado = "No Modificado"
    Else
        Me.lblFechaModificado = m_pedido.FechaModificado
        Me.lblModificador = modificador
    End If
End Sub

Private Sub Form_Load()
    FormHelper.Customize Me
    Set m_pedido.Detalles = DAODetalleOrdenTrabajo.FindAllByOrdenTrabajo(m_pedido.Id)
    Me.caption = "OT Nro. " & Format(m_pedido.Id, "0000")
    GridEXHelper.CustomizeGrid Me.grilla

    If m_pedido.EsHija Then
        Me.caption = Me.caption & " (Corresponde a OTA " & m_pedido.OTMarcoIdPadre & ")"
    End If

    mostrarDetalles
    llenarLista

    Set m_Archivos = DAOArchivo.GetCantidadArchivosPorReferencia(OA_Piezas)
    Set m_Archivos_detalle = DAOArchivo.GetCantidadArchivosPorReferencia(OA_OrdenesTrabajoDetalle)

    Me.grilla.Columns(4).Visible = Permisos.sistemaVerPrecios

    ''Me.caption = caption & " (" & Name & ")"


End Sub

Private Sub llenarLista()
    Me.grilla.ItemCount = m_pedido.Detalles.count
End Sub



Private Sub grilla_FetchIcon(ByVal RowIndex As Long, ByVal ColIndex As Integer, ByVal RowBookmark As Variant, ByVal IconIndex As GridEX20.JSRetInteger)
    On Error Resume Next
    Set rectmp = m_pedido.Detalles(RowIndex)    'grilla.RowIndex(grilla.row))

    If ColIndex = 8 And m_Archivos.item(rectmp.Pieza.Id) > 0 Then
        IconIndex = 1
    End If

    If ColIndex = 9 And m_archivos_Detalles.item(rectmp.Id) > 0 Then
        IconIndex = 1
    End If




End Sub

Private Sub grilla_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If m_pedido.Detalles.count = 0 Then Exit Sub
    Set rectmp = m_pedido.Detalles(grilla.RowIndex(grilla.row))
    If Button = 2 Then
        'If rectmp.pieza.EsConjunto Then
        '    Me.VerDesarrollo.Caption = "Ver Conjunto..."
        '    Me.VerDesarrollo.Tag = 0
        'Else
        Me.VerDesarrollo.caption = "Ver Desarrollo..."
        '    Me.VerDesarrollo.Tag = -1
        'End If

        'If m_pedido.estado = EstadoOT_EnProceso Then
        '     Me.modificar.Enabled = True
        '  Else
        '       Me.modificar.Enabled = False
        '    End If

        'If Not Permisos.planOTmodificar Then Me.modificar = False
        Me.archivos = Permisos.SistemaArchivosVer
        Me.PopupMenu Me.m1
    End If
End Sub
Private Sub grilla_SelectionChange()
    On Error Resume Next
    Set rectmp = m_pedido.Detalles(grilla.RowIndex(grilla.row))
End Sub

Private Sub grilla_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    On Error Resume Next
    Set rectmp = m_pedido.Detalles(RowIndex)

    With rectmp
        Values(1) = rectmp.item
        Values(2) = rectmp.Nota
        Values(3) = rectmp.CantidadPedida
        Values(4) = m_pedido.moneda.NombreCorto
        Values(5) = funciones.FormatearDecimales(rectmp.Precio)
        Values(6) = rectmp.FechaEntrega
        Values(7) = IIf(rectmp.Pieza.EsConjunto, "Conjunto", "Unidad")
        Values(8) = rectmp.Pieza.nombre

        Values(9) = m_Archivos.item(rectmp.Pieza.Id)
        Values(10) = m_Archivos_detalle.item(rectmp.Id)


    End With
End Sub

Private Sub mnuAdquirirPedido_Click()
    Dim archivos As New classArchivos
    archivos.escanearDocumento 111, idPieza
End Sub

Private Sub mnuAdquirirPieza_Click()
    Dim archivos As New classArchivos
    archivos.escanearDocumento 1, idPieza
End Sub
Private Sub mnuArchivosPedido_Click()
    Dim frmar2 As New frmArchivos2
    frmar2.Origen = OrigenArchivos.OA_OrdenesTrabajoDetalle
    frmar2.ObjetoId = rectmp.Id
    frmar2.caption = "OT Nº " & m_pedido.IdFormateado & " - Item " & rectmp.item & " [" & rectmp.Pieza.nombre & "]"
    frmar2.Show
End Sub
Private Sub mnuArchivosPieza_Click()
    Dim frmar1 As New frmArchivos2
    frmar1.Origen = OrigenArchivos.OA_Piezas
    frmar1.ObjetoId = rectmp.Pieza.Id
    frmar1.caption = "Pieza " & rectmp.Pieza.nombre
    frmar1.Show
End Sub
Private Sub mnuInciPedido_Click()
    Dim frminci1 As New frmVerIncidencias
    frminci1.referencia = rectmp.Id
    frminci1.Origen = OI_OrdenesTrabajoDetalles
    frminci1.Show
End Sub
Private Sub mnuInciPieza_Click()
    Dim frminci2 As New frmVerIncidencias
    frminci2.referencia = rectmp.Pieza.Id
    frminci2.Origen = OI_Piezas
    frminci2.Show
End Sub
Private Sub modificar_Click()

End Sub
Private Sub VerDesarrollo_Click()
    Dim F As New frmDesarrollo
    Load F
    F.CargarPieza rectmp.Pieza.Id
    F.Show
End Sub

