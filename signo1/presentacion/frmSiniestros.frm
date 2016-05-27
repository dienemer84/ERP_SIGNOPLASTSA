VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmSiniestros 
   Caption         =   "Siniestros"
   ClientHeight    =   7050
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10260
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSiniestros.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7050
   ScaleWidth      =   10260
   Begin GridEX20.GridEX grid 
      Height          =   6990
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   12330
      Version         =   "2.0"
      PreviewRowIndent=   200
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      PreviewColumn   =   "diagnostico"
      PreviewRowLines =   1
      MethodHoldFields=   -1  'True
      AllowEdit       =   0   'False
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   13
      Column(1)       =   "frmSiniestros.frx":000C
      Column(2)       =   "frmSiniestros.frx":0164
      Column(3)       =   "frmSiniestros.frx":028C
      Column(4)       =   "frmSiniestros.frx":0370
      Column(5)       =   "frmSiniestros.frx":046C
      Column(6)       =   "frmSiniestros.frx":0570
      Column(7)       =   "frmSiniestros.frx":0690
      Column(8)       =   "frmSiniestros.frx":07AC
      Column(9)       =   "frmSiniestros.frx":08C0
      Column(10)      =   "frmSiniestros.frx":09C4
      Column(11)      =   "frmSiniestros.frx":0AC0
      Column(12)      =   "frmSiniestros.frx":0BFC
      Column(13)      =   "frmSiniestros.frx":0D50
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmSiniestros.frx":0EB0
      FormatStyle(2)  =   "frmSiniestros.frx":0FD8
      FormatStyle(3)  =   "frmSiniestros.frx":1088
      FormatStyle(4)  =   "frmSiniestros.frx":113C
      FormatStyle(5)  =   "frmSiniestros.frx":1214
      FormatStyle(6)  =   "frmSiniestros.frx":12CC
      ImageCount      =   0
      PrinterProperties=   "frmSiniestros.frx":13AC
   End
   Begin VB.Menu mnuOpciones 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuEditarSiniestro 
         Caption         =   "Editar Siniestro"
      End
      Begin VB.Menu mnuArchivosSiniestro 
         Caption         =   "Archivos del siniestro"
      End
      Begin VB.Menu mnuAccidente 
         Caption         =   "Accidente"
      End
      Begin VB.Menu mnuImprimirInformeAccidente 
         Caption         =   "Imprimir Informe de Accidente"
      End
   End
End
Attribute VB_Name = "frmSiniestros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements ISuscriber

Dim siniestros As Collection
Dim sin As SiniestroPersonal
Dim CantArchivos As Dictionary


Private Sub Form_Load()
    Customize Me
    GridEXHelper.CustomizeGrid Me.grid, True
    Channel.AgregarSuscriptor Me, TS_InformeAccidente
    'Channel.AgregarSuscriptor Me, TS_Siniestro
    Cargar
End Sub

Private Sub Cargar()
    Set CantArchivos = DAOArchivo.GetCantidadArchivosPorReferencia(OA_Siniestros)
    Set siniestros = DAOSiniestroPersonal.FindAll()
    Me.grid.ItemCount = 0
    Me.grid.ItemCount = siniestros.count
    GridEXHelper.AutoSizeColumns Me.grid
End Sub

Private Sub Form_Resize()
    If Me.WindowState <> 1 Then
        Me.grid.Width = Me.ScaleWidth - 50
        Me.grid.Height = Me.ScaleHeight - 50
        GridEXHelper.AutoSizeColumns Me.grid
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Channel.RemoverSuscripcionTotal Me
End Sub



Private Sub grid_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 And Me.grid.rowcount > 0 Then
        Me.mnuEditarSiniestro.Enabled = Permisos.RRHHSiniestros
        Me.mnuArchivosSiniestro.Enabled = Permisos.RRHHSiniestros

        If Permisos.RRHHSiniestros Then
            Me.mnuAccidente.Enabled = True
        Else
            If Permisos.RRHHInformeAccidente Then
                If IsSomething(funciones.GetUserObj().Empleado) Then
                    Me.mnuAccidente.Enabled = (sin.Supervisor.id = funciones.GetUserObj.Empleado.id)
                Else
                    Me.mnuAccidente.Enabled = False
                End If
            End If
        End If

        If sin.InformeAccidenteConfeccionado Then
            Me.mnuAccidente.caption = "Editar Informe Accidente"
        Else
            Me.mnuAccidente.caption = "Crear Informe Accidente"
        End If

        Me.mnuImprimirInformeAccidente.Enabled = Permisos.RRHHSiniestros And sin.InformeAccidenteConfeccionado

        Me.PopupMenu Me.mnuOpciones
    End If
End Sub

Private Sub grid_SelectionChange()
    On Error Resume Next
    Set sin = siniestros.item(Me.grid.RowIndex(Me.grid.row))
End Sub

Private Sub grid_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex > 0 And siniestros.count > 0 Then
        Set sin = siniestros(RowIndex)
        Values(1) = sin.NroSiniestro
        Values(2) = sin.FechaHoraOcurrido
        Values(3) = sin.ART.nombre
        Values(4) = sin.Asegurado.NombreCompleto
        Values(5) = sin.Supervisor.NombreCompleto
        If LenB(sin.Diagnostico) > 0 Then
            Values(6) = "Diagnóstico: " & sin.Diagnostico
        End If
        Values(7) = sin.PrestadorMedico
        If enums.TiposAccidente.Exists(CStr(sin.TipoAccidente)) Then
            Values(8) = enums.TiposAccidente.item(CStr(sin.TipoAccidente))
        End If
        If enums.TiposTratamiento.Exists(CStr(sin.TipoTratamiento)) Then
            Values(9) = enums.TiposTratamiento.item(CStr(sin.TipoTratamiento))
        End If
        If enums.TiposGravedad.Exists(CStr(sin.TipoGravedad)) Then
            Values(10) = enums.TiposGravedad.item(CStr(sin.TipoGravedad))
        End If
        If CDbl(sin.RenaudaTareas) > 0 Then Values(11) = sin.RenaudaTareas
        Values(12) = Val(CantArchivos(sin.id))
        Values(13) = sin.InformeAccidenteConfeccionado
    End If
End Sub

Private Property Get ISuscriber_id() As String
    Static id_suscriber As String
    If LenB(id_suscriber) = 0 Then id_suscriber = funciones.CreateGUID()
    ISuscriber_id = id_suscriber
End Property

Private Function ISuscriber_Notificarse(EVENTO As clsEventoObserver) As Variant
    Cargar
End Function

Private Sub mnuAccidente_Click()
    Dim F As New frmAccidente
    frmAccidente.Show
    frmAccidente.Cargar sin
End Sub

Private Sub mnuArchivosSiniestro_Click()
    If Not sin Is Nothing Then
        Dim frmArchi As New frmArchivos2
        frmArchi.Origen = OrigenArchivos.OA_Siniestros
        frmArchi.ObjetoId = sin.id
        frmArchi.caption = "Siniestro Nº" & sin.NroSiniestro
        frmArchi.Show
    End If
End Sub

Private Sub mnuEditarSiniestro_Click()
    If Not sin Is Nothing Then
        Dim F As New frmSiniestro
        F.Show
        Set F.sin = sin
        F.Cargar
    End If
End Sub

Private Sub mnuImprimirInformeAccidente_Click()
    With dsrInformeAccidente.Sections("cabeza").Controls
        .item("lblNroSiniestro").caption = sin.NroSiniestro
        .item("lblART").caption = sin.ART.nombre
    End With

    With dsrInformeAccidente.Sections("main").Controls
        .item("lblAccidentado").caption = sin.Asegurado.NombreCompleto
        .item("lblLegajo").caption = sin.Asegurado.legajo
        .item("lblFechaOcurrido").caption = Format(sin.FechaHoraOcurrido, "dddd dd/mm/yyyy a la\s hh:nn")
        .item("lblHsExtras").caption = IIf(sin.InformeAccidente.HsExtras, "Si", "No")
        .item("lblSector").caption = sin.Sector.Sector
        .item("lblPuesto").caption = sin.InformeAccidente.Puesto
        .item("lblTestigos").caption = sin.InformeAccidente.NombreTestigos
        .item("lblDescripcionCaso").caption = sin.InformeAccidente.DescripcionHecho

        .item("lblFallaMaquinas").caption = sin.InformeAccidente.FallaMaquinasEquipos
        .item("lblFaltaElementos").caption = sin.InformeAccidente.FaltaElementosProteccionPersonal
        .item("lblActoInseguro").caption = sin.InformeAccidente.ActoInseguro
        .item("lblOtros").caption = sin.InformeAccidente.Otros

        .item("lblSupervisor").caption = sin.Supervisor.NombreCompleto

        .item("lblNaturalezaLesion").caption = sin.InformeAccidente.NaturalezaLesion
        .item("lblUbicacionLesion").caption = sin.InformeAccidente.UbicacionLesion
        .item("lblFormaAccidente").caption = sin.InformeAccidente.FormaAccidente
        .item("lblAgenteMaterial").caption = sin.InformeAccidente.AgenteMaterial
        .item("lblFechaAlta").caption = sin.RenaudaTareas
        .item("lblCantDiasPerdidos").caption = DateDiff("d", sin.FechaHoraOcurrido, sin.RenaudaTareas)
        .item("lblRecomendaciones").caption = sin.InformeAccidente.RecomendacionParaEvitarRepeticion
    End With

    Dim r As Recordset
    Set r = conectar.RSFactory("SELECT 1")
    Set dsrInformeAccidente.DataSource = r
    dsrInformeAccidente.Show

End Sub
