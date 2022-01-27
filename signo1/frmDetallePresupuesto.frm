VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmVentasPresupuestoDetalle 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle Presupuesto"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10455
   ClipControls    =   0   'False
   Icon            =   "frmDetallePresupuesto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   10455
   Begin XtremeSuiteControls.PushButton Command2 
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   4560
      Width           =   1935
      _Version        =   786432
      _ExtentX        =   3413
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Grabar"
      UseVisualStyle  =   -1  'True
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   4335
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   7646
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      PreviewColumn   =   5
      PreviewRowLines =   1
      ColumnAutoResize=   -1  'True
      MethodHoldFields=   -1  'True
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      BackColorHeader =   16761024
      DataMode        =   99
      HeaderFontName  =   "Tahoma"
      HeaderFontSize  =   9.75
      FontName        =   "Tahoma"
      ColumnHeaderHeight=   330
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   6
      Column(1)       =   "frmDetallePresupuesto.frx":000C
      Column(2)       =   "frmDetallePresupuesto.frx":0108
      Column(3)       =   "frmDetallePresupuesto.frx":01D8
      Column(4)       =   "frmDetallePresupuesto.frx":02AC
      Column(5)       =   "frmDetallePresupuesto.frx":037C
      Column(6)       =   "frmDetallePresupuesto.frx":043C
      FormatStylesCount=   7
      FormatStyle(1)  =   "frmDetallePresupuesto.frx":0530
      FormatStyle(2)  =   "frmDetallePresupuesto.frx":0658
      FormatStyle(3)  =   "frmDetallePresupuesto.frx":0708
      FormatStyle(4)  =   "frmDetallePresupuesto.frx":07BC
      FormatStyle(5)  =   "frmDetallePresupuesto.frx":0894
      FormatStyle(6)  =   "frmDetallePresupuesto.frx":094C
      FormatStyle(7)  =   "frmDetallePresupuesto.frx":0A2C
      ImageCount      =   0
      PrinterProperties=   "frmDetallePresupuesto.frx":0AB8
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FF8080&
      Caption         =   "[ Presupuesto ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   7440
      TabIndex        =   0
      Top             =   4560
      Width           =   2895
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FF8080&
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
         Left            =   360
         TabIndex        =   10
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label lblFechaEntrega 
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1200
         TabIndex        =   9
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label lblFechaFinalizado 
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1200
         TabIndex        =   8
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label lblFechaCreado 
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1200
         TabIndex        =   7
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FF8080&
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
         Left            =   480
         TabIndex        =   6
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FF8080&
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
         Left            =   480
         TabIndex        =   5
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FF8080&
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
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblCreado 
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1200
         TabIndex        =   3
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label LblFinalizado 
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1200
         TabIndex        =   2
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FF8080&
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
         TabIndex        =   1
         Top             =   840
         Width           =   975
      End
   End
   Begin XtremeSuiteControls.PushButton Command1 
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   4920
      Width           =   1935
      _Version        =   786432
      _ExtentX        =   3413
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Imprimir"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton Command4 
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   5280
      Width           =   1935
      _Version        =   786432
      _ExtentX        =   3413
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Salir"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton Command8 
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   5640
      Width           =   1935
      _Version        =   786432
      _ExtentX        =   3413
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Materiales"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton Command10 
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   6000
      Width           =   1935
      _Version        =   786432
      _ExtentX        =   3413
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Materialización"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton Command3 
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   6360
      Width           =   1935
      _Version        =   786432
      _ExtentX        =   3413
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Más Información"
      UseVisualStyle  =   -1  'True
   End
   Begin VB.Menu m1 
      Caption         =   "Menu1"
      Visible         =   0   'False
      Begin VB.Menu Ver 
         Caption         =   "Ver Desarrollo..."
      End
      Begin VB.Menu mnuDesarrolloHistorico 
         Caption         =   "Ver Desarrollo Historico..."
      End
      Begin VB.Menu archivos 
         Caption         =   "Archivos Asociados..."
         Begin VB.Menu mnuVerArchivosDePieza 
            Caption         =   "De Pieza..."
         End
         Begin VB.Menu mnuVerArchivosDePedido 
            Caption         =   "Del Detalle..."
         End
      End
      Begin VB.Menu scanear 
         Caption         =   "Adquirir..."
         Begin VB.Menu mnuAdquirirAPieza 
            Caption         =   "A Pieza..."
         End
         Begin VB.Menu mnuAdquirirADetalle 
            Caption         =   "Al Detalle"
         End
      End
      Begin VB.Menu VerIncidencias 
         Caption         =   "Ver Incidencias..."
         Begin VB.Menu mnuVerIncidenciasDePieza 
            Caption         =   "De Pieza..."
         End
         Begin VB.Menu mnuVerIncidenciasDeDetallePedido 
            Caption         =   "Del Detalle..."
         End
      End
   End
End
Attribute VB_Name = "frmVentasPresupuestoDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim archi As classArchivos
Dim presu As clsPresupuesto
Dim tmp As clsPresupuestoDetalle

Public Property Let presupuesto(T As clsPresupuesto)
    Set presu = DAOPresupuestos.GetById(T.Id)


End Property
Private Sub Command1_Click()
    If DAOPresupuestos.ImprimirPresupuesto(presu) Then
        DAOPresupuestoHistorial.agregar presu, "presupuesto impreso"
    End If
End Sub
Private Sub Command10_Click()
    Dim frmmat As New frmMaterializacion

    frmmat.Id = presu.Id
    frmmat.Ot = False
    frmmat.otro = False
    frmmat.presu = True
    frmmat.Show
End Sub

Private Sub Command2_Click()
    If DAOPresupuestos.exporta(presu) Then
        DAOPresupuestoHistorial.agregar presu, "Presupuesto exportado"
    End If
End Sub
Private Sub Command3_Click()
    frmVentasPresupuestoMasDetalles.presu = presu
    frmVentasPresupuestoMasDetalles.Show 1
End Sub
Private Sub Command4_Click()
    Unload Me
End Sub
Private Sub Command8_Click()
    Dim baseP As New classPlaneamiento
    A = baseP.informePiezaMateriales(presu.Id, 2, True)
End Sub

Private Sub llenarLista()
    Me.GridEX1.ItemCount = presu.DetallePresupuesto.count
End Sub
Private Sub mostrar()
    Me.caption = "Presupuesto Nro. " & presu.Id
    Me.lblCreado = presu.UsuarioCreado.usuario
    Me.lblFechaCreado = presu.fechaCreado

    If presu.UsuarioFinalizado Is Nothing Then
        Me.LblFinalizado = Empty
        Me.lblFechaFinalizado = Empty
    Else
        Me.LblFinalizado = presu.UsuarioFinalizado.usuario
        Me.lblFechaFinalizado = presu.FechaFinalizado
    End If



End Sub

Private Sub traerDetalles()
    Set presu.DetallePresupuesto = DAOPresupuestosDetalle.GetAllByPresupuesto(presu)
End Sub

Private Sub Form_Load()
    FormHelper.Customize Me
    GridEXHelper.CustomizeGrid Me.GridEX1
    traerDetalles
    mostrar
    llenarLista
    Me.GridEX1.Columns(6).Visible = Permisos.sistemaVerPrecios
    
        Me.caption = caption & " (" & Name & ")"
        
        
End Sub


Private Sub GridEX1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    On Error GoTo err1
    If Button = 2 Then
        Set tmp = presu.DetallePresupuesto(GridEX1.RowIndex(GridEX1.row))

        If tmp.Pieza.EsConjunto Then
            Me.ver.caption = "Ver Conjunto..."
            Me.ver.Tag = 0
        Else
            Me.ver.caption = "Ver Desarrollo..."
            Me.ver.Tag = -1
        End If

        Me.PopupMenu Me.m1
    End If

    Exit Sub
err1:
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Set tmp = presu.DetallePresupuesto(RowIndex)

    With tmp
        Values(1) = tmp.item
        Values(2) = tmp.Detalles
        Values(3) = tmp.Cantidad
        Values(4) = tmp.entrega
        Values(5) = tmp.Pieza.nombre
        Values(6) = funciones.FormatearDecimales(tmp.ValorManual)
    End With
End Sub
Private Sub mnuAdquirirADetalle_Click()
    Set arch = New classArchivos
    Set tmp = presu.DetallePresupuesto(GridEX1.RowIndex(GridEX1.row))
    archi.escanearDocumento OrigenArchivos.OA_PresupuestoDetalle, tmp.Id
End Sub

Private Sub mnuAdquirirAPieza_Click()
    Set archi = New classArchivos
    Set tmp = presu.DetallePresupuesto(GridEX1.RowIndex(GridEX1.row))
    archi.escanearDocumento OrigenArchivos.OA_Piezas, tmp.Pieza.Id

End Sub

Private Sub mnuDesarrolloHistorico_Click()
    Set tmp = presu.DetallePresupuesto(GridEX1.RowIndex(GridEX1.row))
    Dim F As New frmDesarrollo
    Load F
    F.CargarDetallePresupuesto tmp.Id
    F.Show
End Sub

Private Sub mnuVerArchivosDePedido_Click()
    Set tmp = presu.DetallePresupuesto(GridEX1.RowIndex(GridEX1.row))
    Dim frmarchi1 As New frmArchivos2
    frmarchi1.Origen = OrigenArchivos.OA_PresupuestoDetalle
    frmarchi1.ObjetoId = tmp.Id
    frmarchi1.caption = "Presupuesto Nº " & presu.IdFormateada & " - Item " & tmp.item & " [" & tmp.Pieza.nombre & "]"
    frmarchi1.Show
End Sub
Private Sub mnuVerArchivosDePieza_Click()
    Set tmp = presu.DetallePresupuesto(GridEX1.RowIndex(GridEX1.row))
    Dim frmarchi2 As New frmArchivos2
    frmarchi2.Origen = OrigenArchivos.OA_Piezas
    frmarchi2.ObjetoId = tmp.Pieza.Id
    frmarchi2.caption = "Pieza " & tmp.Pieza.nombre
    frmarchi2.Show
End Sub
Private Sub mnuVerIncidenciasDeDetallePedido_Click()
    Set tmp = presu.DetallePresupuesto(GridEX1.RowIndex(GridEX1.row))
    Dim inci1 As New frmVerIncidencias

    inci1.referencia = tmp.Id
    inci1.Origen = OI_DetallePresupuesto
    inci1.Show
End Sub

Private Sub mnuVerIncidenciasDePieza_Click()
    Set tmp = presu.DetallePresupuesto(GridEX1.RowIndex(GridEX1.row))
    Dim inci2 As New frmVerIncidencias
    inci2.referencia = tmp.Pieza.Id
    inci2.Origen = OI_Piezas
    inci2.Show


End Sub

Private Sub PushButton1_Click()

End Sub

Private Sub ver_Click()
    Set tmp = presu.DetallePresupuesto(GridEX1.RowIndex(GridEX1.row))

    Dim F As New frmDesarrollo
    Load F
    F.CargarPieza tmp.Pieza.Id
    F.Show


End Sub

