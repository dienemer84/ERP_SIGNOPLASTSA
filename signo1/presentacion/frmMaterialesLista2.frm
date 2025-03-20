VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmMaterialesLista2 
   Caption         =   "Lista de materiales"
   ClientHeight    =   7950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   18135
   Icon            =   "frmMaterialesLista2.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7950
   ScaleWidth      =   18135
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1455
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   17820
      _Version        =   786432
      _ExtentX        =   31432
      _ExtentY        =   2566
      _StockProps     =   79
      Caption         =   "Parámetros de búsqueda"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.CheckBox chkAtributos 
         Height          =   195
         Left            =   6510
         TabIndex        =   9
         Top             =   450
         Width           =   1830
         _Version        =   786432
         _ExtentX        =   3228
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Mostrar Atributos"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cboRubros 
         Height          =   315
         Left            =   960
         TabIndex        =   7
         Top             =   720
         Width           =   3135
         _Version        =   786432
         _ExtentX        =   5530
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.PushButton cmdFiltrar 
         Default         =   -1  'True
         Height          =   375
         Left            =   5040
         TabIndex        =   4
         Top             =   360
         Width           =   1095
         _Version        =   786432
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Buscar"
         UseVisualStyle  =   -1  'True
      End
      Begin VB.TextBox txtFiltro 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Top             =   360
         Width           =   3735
      End
      Begin XtremeSuiteControls.PushButton Command2 
         Height          =   375
         Left            =   5040
         TabIndex        =   5
         Top             =   840
         Width           =   1095
         _Version        =   786432
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Imprimir"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton CMDsINCliente 
         Height          =   255
         Left            =   4200
         TabIndex        =   8
         Top             =   735
         Width           =   465
         _Version        =   786432
         _ExtentX        =   820
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "X"
         BackColor       =   12632256
         UseVisualStyle  =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rubro"
         Height          =   195
         Left            =   390
         TabIndex        =   6
         Top             =   750
         Width           =   435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filtro"
         Height          =   195
         Left            =   495
         TabIndex        =   3
         Top             =   375
         Width           =   330
      End
   End
   Begin GridEX20.GridEX grilla 
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   17955
      _ExtentX        =   31671
      _ExtentY        =   10398
      Version         =   "2.0"
      PreviewRowIndent=   200
      HoldSortSettings=   -1  'True
      DefaultGroupMode=   1
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      PreviewColumn   =   "attrs"
      CalendarTodayText=   "Hoy"
      CalendarNoneText=   "Nada"
      MultiSelect     =   -1  'True
      MethodHoldFields=   -1  'True
      ForeColorInfoText=   -2147483639
      BackColorInfoText=   8421504
      GroupByBoxInfoText=   "Arrastre el encabezado de una columna para agrupar"
      BackColorGBBox  =   8421504
      BackColorHeader =   16761024
      ImageCount      =   1
      ImagePicture1   =   "frmMaterialesLista2.frx":000C
      DataMode        =   99
      HeaderFontName  =   "Tahoma"
      FontName        =   "Tahoma"
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   17
      Column(1)       =   "frmMaterialesLista2.frx":0326
      Column(2)       =   "frmMaterialesLista2.frx":041A
      Column(3)       =   "frmMaterialesLista2.frx":0506
      Column(4)       =   "frmMaterialesLista2.frx":05F2
      Column(5)       =   "frmMaterialesLista2.frx":06CA
      Column(6)       =   "frmMaterialesLista2.frx":07C2
      Column(7)       =   "frmMaterialesLista2.frx":0912
      Column(8)       =   "frmMaterialesLista2.frx":0A6A
      Column(9)       =   "frmMaterialesLista2.frx":0B5E
      Column(10)      =   "frmMaterialesLista2.frx":0C2E
      Column(11)      =   "frmMaterialesLista2.frx":0D1A
      Column(12)      =   "frmMaterialesLista2.frx":0E46
      Column(13)      =   "frmMaterialesLista2.frx":0F16
      Column(14)      =   "frmMaterialesLista2.frx":0FE6
      Column(15)      =   "frmMaterialesLista2.frx":10D2
      Column(16)      =   "frmMaterialesLista2.frx":11E2
      Column(17)      =   "frmMaterialesLista2.frx":1312
      FormatStylesCount=   7
      FormatStyle(1)  =   "frmMaterialesLista2.frx":144A
      FormatStyle(2)  =   "frmMaterialesLista2.frx":1572
      FormatStyle(3)  =   "frmMaterialesLista2.frx":1622
      FormatStyle(4)  =   "frmMaterialesLista2.frx":16D6
      FormatStyle(5)  =   "frmMaterialesLista2.frx":17AE
      FormatStyle(6)  =   "frmMaterialesLista2.frx":1866
      FormatStyle(7)  =   "frmMaterialesLista2.frx":1946
      ImageCount      =   1
      ImagePicture(1) =   "frmMaterialesLista2.frx":1A22
      PrinterProperties=   "frmMaterialesLista2.frx":1D3C
   End
   Begin VB.Menu mnu 
      Caption         =   "menu"
      Visible         =   0   'False
      Begin VB.Menu codigo 
         Caption         =   "codigo"
         Enabled         =   0   'False
      End
      Begin VB.Menu editar 
         Caption         =   "Editar"
      End
      Begin VB.Menu mnuAprobar 
         Caption         =   "Aprobar"
      End
      Begin VB.Menu mnuHistoricoaAprobacion 
         Caption         =   "Ver Historial"
      End
      Begin VB.Menu mnuCopiar 
         Caption         =   "Copiar como..."
      End
      Begin VB.Menu histerico 
         Caption         =   "Historico de precios"
      End
      Begin VB.Menu mnuArchivos 
         Caption         =   "Archivos Asociados..."
      End
   End
End
Attribute VB_Name = "frmMaterialesLista2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Implements ISuscriber
Dim filtro As String
Dim vId As String
Dim rows As Long
Dim archivos As Dictionary
Dim Materiales As New Collection
Dim rectemp As clsMaterial
Private buscando As Boolean

Private Sub chkAtributos_Click()
    If Me.chkAtributos.value = xtpChecked Then
        Me.grilla.PreviewRowLines = 1
    Else
        Me.grilla.PreviewRowLines = 0
    End If
End Sub

Private Sub cmdFiltrar_Click()
    filtro = " 1=1 "
    If LenB(Me.txtFiltro.Text) > 0 Then
        filtro = filtro & " AND (" & DAOMateriales.TABLA_MATERIALES & "." & DAOMateriales.CAMPO_DESCRIPCION & " LIKE '%" & Trim(Me.txtFiltro.Text) & "%' or " & DAOMateriales.TABLA_MATERIALES & "." & DAOMateriales.CAMPO_CODIGO & " LIKE '%" & Trim(Me.txtFiltro) & "%')"
    End If

    If Me.cboRubros.ListIndex > -1 Then
        filtro = filtro & " and " & DAOMateriales.TABLA_MATERIALES & ".id_rubro" & "  = " & Me.cboRubros.ItemData(Me.cboRubros.ListIndex)
    End If

    Set Materiales = DAOMateriales.FindAll(filtro)
    Set archivos = DAOArchivo.GetCantidadArchivosPorReferencia(OA_Materiales)

    grilla.ItemCount = 0
    grilla.ItemCount = Materiales.count
    grilla.ReBind

    Dim P As Integer
    If Me.cboRubros.ListIndex > -1 Then
        P = Me.cboRubros.ItemData(Me.cboRubros.ListIndex)
    Else
        P = -1
    End If
    Filtros.FiltroBusquedaMaterial Me.txtFiltro, P
    GridEXHelper.AutoSizeColumns Me.grilla, True
End Sub

Private Sub CMDsINCliente_Click()
    Me.cboRubros.ListIndex = -1
End Sub


Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()
    Dim elegidos As Boolean
    If grilla.SelectedItems.count > 1 Then
        elegidos = True
    Else
        elegidos = False
    End If

    With Me.grilla.PrinterProperties
        .FitColumns = True
        .RepeatHeaders = True
        .Orientation = jgexPPLandscape
        .HeaderString(jgexHFCenter) = "Lista de Materiales"
        .FooterString(jgexHFCenter) = Now
    End With
    Load frmPrintPreview
    frmPrintPreview.Move Me.Left, Me.Top, Me.Width, Me.Height
    grilla.PrintPreview frmPrintPreview.GEXPreview1, elegidos
    frmPrintPreview.Show 1
End Sub
Private Sub Command3_Click()
    Unload Me
End Sub
Private Sub editar_Click()
    editamos
End Sub
Private Sub Form_Activate()

    Me.grilla.ReBind
End Sub

Private Sub Form_Load()
    FormHelper.Customize Me
    GridEXHelper.CustomizeGrid Me.grilla, True
    grilla.ItemCount = 0
    DAORubros.LlenarComboExtremeSuite Me.cboRubros
    Me.cboRubros.ListIndex = -1
    rows = 1
    Set archivos = DAOArchivo.GetCantidadArchivosPorReferencia(OA_Materiales)

    vId = funciones.CreateGUID

    Me.txtFiltro = Filtros.vFiltroBusquedaMaterial.nombre
    Me.cboRubros.ListIndex = funciones.PosIndexCbo(Filtros.vFiltroBusquedaMaterial.rubro, Me.cboRubros)
    Channel.AgregarSuscriptor Me, Materiales_
    cmdFiltrar_Click
End Sub
Private Sub Form_Resize()
    On Error Resume Next
    Me.grilla.Height = Me.ScaleHeight - 1800
    Me.grilla.Width = Me.ScaleWidth - 220
    Me.GroupBox1.Width = Me.grilla.Width
    Me.grilla.ColumnAutoResize = True

    Me.grilla.ReBind
End Sub
Private Sub Form_Terminate()
    Channel.RemoverSuscripcionTotal Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Channel.RemoverSuscripcionTotal Me
End Sub

Private Sub grilla_BeforeGroupChange(ByVal Group As GridEX20.JSGroup, ByVal ChangeOperation As GridEX20.jgexGroupChange, ByVal GroupPosition As Integer, ByVal Cancel As GridEX20.JSRetBoolean)
    Me.grilla.CollapseAll
End Sub

Private Sub grilla_BeforePrintPage(ByVal PageNumber As Long, ByVal nPages As Long)
    grilla.PrinterProperties.FooterString(jgexHFRight) = "Página " & PageNumber & " de " & nPages
End Sub
Private Sub grilla_ColumnHeaderClick(ByVal Column As GridEX20.JSColumn)
    GridEXHelper.ColumnHeaderClick Me.grilla, Column
End Sub
Private Sub grilla_DblClick()
'    editamos

End Sub

Private Sub grilla_FetchIcon(ByVal rowIndex As Long, ByVal ColIndex As Integer, ByVal RowBookmark As Variant, ByVal IconIndex As GridEX20.JSRetInteger)
    On Error Resume Next
    Set rectemp = Materiales(grid.rowIndex(RowPosition))

    If ColIndex = 16 And archivos.item(rectemp.Id) > 0 Then
        IconIndex = 1
    End If

End Sub

Private Sub grilla_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 67 And Shift = 2 Then    'CTRL + C
        GridEXHelper.Grid2Clipboard Me.grilla
        DoEvents
        MsgBox "La lista ha sido copiada al portapapeles.", vbInformation + vbOKOnly
    End If
End Sub

Private Sub grilla_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    grilla_SelectionChange
    If grilla.rowIndex(grilla.row) = 0 Then Exit Sub
    If Button = 2 Then
        Me.codigo.caption = "[ " & rectemp.codigo & " ]"
        Me.mnuAprobar.Enabled = Not rectemp.Aprobado
        Me.PopupMenu mnu
    End If
End Sub
Private Sub grilla_RowFormat(RowBuffer As GridEX20.JSRowData)
    If RowBuffer.value(8) = 0 Or Not RowBuffer.value(17) Then
        RowBuffer.RowStyle = "Valor0"
    End If
End Sub
Private Sub grilla_SelectionChange()
    On Error GoTo err1
    rows = grilla.rowIndex(grilla.row)
    Set rectemp = Materiales.item(grilla.rowIndex(grilla.row))
    Exit Sub
err1:
End Sub
Private Sub grilla_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Set rectemp = Materiales.item(rowIndex)
    With rectemp
        Values(1) = .codigo
        Values(2) = .Grupo.rubros.rubro
        Values(3) = .Grupo.Grupo
        Values(4) = .descripcion
        Values(5) = .Espesor
        Values(6) = enums.enumUnidades(.unidad)
        Values(7) = .PesoXUnidad
        Values(8) = .Valor
        Values(9) = .moneda.NombreCorto
        Values(10) = .FechaValor
        Values(11) = .Cantidad
        Values(12) = .almacen.almacen
        Values(13) = enums.enumEstadoMaterial(.estado)
        If .Tipo <> 0 Then Values(14) = enums.EnumTipoMaterial(.Tipo)
        Values(15) = funciones.JoinCollectionValues(.Atributos, ", ")
        Values(17) = .Aprobado

    End With
End Sub
Private Sub llenar_Grilla()
    Set Materiales = DAOMateriales.FindAll
    grilla.ItemCount = Materiales.count
    'grilla.ReBind
End Sub
Private Sub editamos()
    If grilla.rowcount > 0 Then
        A = grilla.rowIndex(grilla.row)
        If A = 0 Then Exit Sub

        Dim frm1 As New frmMaterialesNuevo
        frm1.Material = Materiales(grilla.rowIndex(grilla.row))
        frm1.Show
    End If
End Sub

Private Property Get ISuscriber_id() As String
    ISuscriber_id = vId
End Property

Private Function ISuscriber_Notificarse(EVENTO As clsEventoObserver) As Variant
    Dim tmp As clsMaterial
    If EVENTO.EVENTO = agregar_ Then
        cmdFiltrar_Click
    ElseIf EVENTO.EVENTO = modificar_ Then

        Set tmp = EVENTO.Elemento

        For i = Materiales.count To 1 Step -1
            If Materiales(i).Id = tmp.Id Then
                Set rectemp = Materiales(i)
                rectemp.Id = tmp.Id
                rectemp.Valor = tmp.Valor
                rectemp.descripcion = tmp.descripcion
                rectemp.Grupo = tmp.Grupo
                rectemp.estado = tmp.estado
                rectemp.almacen = tmp.almacen
                rectemp.Cantidad = tmp.Cantidad
                rectemp.FechaValor = tmp.FechaValor
                rectemp.Aprobado = tmp.Aprobado
                grilla.RefreshRowIndex i
                'GridEXHelper.AutoSizeColumns Me.grilla
                Exit For
            End If
        Next

    End If

End Function

Private Sub mnuAprobar_Click()
    On Error GoTo err1
    If MsgBox("¿Está seguro de aprobar el material?", vbYesNo) = vbYes Then

        If Not rectemp.Aprobado Then
            If DAOMateriales.aprobar(rectemp) Then
                MsgBox "Material aprobado con éxito!", vbExclamation

            End If
        End If
    End If
    Exit Sub

err1:
    MsgBox "Error al aprobar: " & Err.Description
End Sub

Private Sub mnuArchivos_Click()
    grilla_SelectionChange
    Dim frmarchi1 As New frmArchivos2
    frmarchi1.Origen = OA_Materiales
    frmarchi1.ObjetoId = rectemp.Id
    frmarchi1.caption = "Material  " & rectemp.descripcion
    frmarchi1.Show
End Sub

Private Sub mnuCopiar_Click()
    If grilla.rowcount > 0 Then
        A = grilla.rowIndex(grilla.row)
        If A = 0 Then Exit Sub

        Dim m1 As clsMaterial
        Set m1 = Materiales(grilla.rowIndex(grilla.row))

        Dim nombre As String
        nombre = InputBox("Ingrese el nombre del material")
        If LenB(nombre) = 0 Then
            MsgBox "Debe ingresar un nombre para el material", vbExclamation
        Else
            m1.Id = 0       'asi hace insert
            m1.descripcion = nombre
            m1.codigo = Me.BuildCodigoMaterial(m1)
            If DAOMateriales.crear(m1) Then
                cmdFiltrar_Click
                MsgBox "El material ha sido copiado.", vbInformation
            Else
                MsgBox "No se pudo copiar el material.", vbCritical
            End If
        End If

    End If
End Sub


Public Function BuildCodigoMaterial(MAT As clsMaterial) As String
    Dim cod As String

    cod = MAT.Grupo.rubros.iniciales
    cod = cod & Format(MAT.Grupo.Id, "000")

    BuildCodigoMaterial = cod
End Function

Private Sub mnuHistoricoaAprobacion_Click()
    Dim F As New frmHistoricoMateriales
    F.IdMaterial = rectemp.Id
    F.Show
End Sub
