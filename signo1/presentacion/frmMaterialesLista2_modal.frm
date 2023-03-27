VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmMaterialesLista2_modal 
   Caption         =   "Seleccione Material"
   ClientHeight    =   8010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16710
   Icon            =   "frmMaterialesLista2_modal.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   16710
   StartUpPosition =   3  'Windows Default
   Begin GridEX20.GridEX grilla 
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   16455
      _ExtentX        =   29025
      _ExtentY        =   10398
      Version         =   "2.0"
      HoldSortSettings=   -1  'True
      DefaultGroupMode=   1
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      CalendarTodayText=   "Hoy"
      CalendarNoneText=   "Nada"
      MultiSelect     =   -1  'True
      MethodHoldFields=   -1  'True
      ForeColorInfoText=   -2147483639
      BackColorInfoText=   8421504
      GroupByBoxInfoText=   "Arrastre el encabezado de una columna para agrupar"
      BackColorGBBox  =   8421504
      BackColorHeader =   16761024
      DataMode        =   99
      HeaderFontName  =   "Tahoma"
      FontName        =   "Tahoma"
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   13
      Column(1)       =   "frmMaterialesLista2_modal.frx":000C
      Column(2)       =   "frmMaterialesLista2_modal.frx":0100
      Column(3)       =   "frmMaterialesLista2_modal.frx":01EC
      Column(4)       =   "frmMaterialesLista2_modal.frx":02D8
      Column(5)       =   "frmMaterialesLista2_modal.frx":03B0
      Column(6)       =   "frmMaterialesLista2_modal.frx":0480
      Column(7)       =   "frmMaterialesLista2_modal.frx":05A8
      Column(8)       =   "frmMaterialesLista2_modal.frx":06D8
      Column(9)       =   "frmMaterialesLista2_modal.frx":07A4
      Column(10)      =   "frmMaterialesLista2_modal.frx":0874
      Column(11)      =   "frmMaterialesLista2_modal.frx":0960
      Column(12)      =   "frmMaterialesLista2_modal.frx":0A8C
      Column(13)      =   "frmMaterialesLista2_modal.frx":0B5C
      FormatStylesCount=   7
      FormatStyle(1)  =   "frmMaterialesLista2_modal.frx":0C2C
      FormatStyle(2)  =   "frmMaterialesLista2_modal.frx":0D54
      FormatStyle(3)  =   "frmMaterialesLista2_modal.frx":0E04
      FormatStyle(4)  =   "frmMaterialesLista2_modal.frx":0EB8
      FormatStyle(5)  =   "frmMaterialesLista2_modal.frx":0F90
      FormatStyle(6)  =   "frmMaterialesLista2_modal.frx":1048
      FormatStyle(7)  =   "frmMaterialesLista2_modal.frx":1128
      ImageCount      =   0
      PrinterProperties=   "frmMaterialesLista2_modal.frx":1204
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1455
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   16455
      _Version        =   786432
      _ExtentX        =   29025
      _ExtentY        =   2566
      _StockProps     =   79
      Caption         =   "Parámetros de búsqueda"
      UseVisualStyle  =   -1  'True
      Begin VB.TextBox txtFiltro 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Top             =   360
         Width           =   3735
      End
      Begin XtremeSuiteControls.ComboBox cboRubros 
         Height          =   315
         Left            =   960
         TabIndex        =   3
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
         TabIndex        =   6
         Top             =   750
         Width           =   480
         _Version        =   786432
         _ExtentX        =   847
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "X"
         BackColor       =   12632256
         UseVisualStyle  =   -1  'True
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Rubro"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Filtro"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmMaterialesLista2_modal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Implements ISuscriber
Dim filtro As String
Dim vId As String
Dim rows As Long
Dim Materiales As New Collection
Dim rectemp As clsMaterial
Private buscando As Boolean
Public Usable As Boolean
Private Sub cmdFiltrar_Click()
    filtro = " 1=1 "
    If LenB(Me.txtFiltro.text) > 0 Then
        filtro = filtro & " AND (" & DAOMateriales.TABLA_MATERIALES & "." & DAOMateriales.CAMPO_DESCRIPCION & " LIKE '%" & Trim(Me.txtFiltro.text) & "%' or " & DAOMateriales.TABLA_MATERIALES & "." & DAOMateriales.CAMPO_CODIGO & " LIKE '%" & Trim(Me.txtFiltro) & "%')"
    End If

    If Me.cboRubros.ListIndex > -1 Then
        filtro = filtro & " and " & DAOMateriales.TABLA_MATERIALES & ".id_rubro" & "  = " & Me.cboRubros.ItemData(Me.cboRubros.ListIndex)
    End If
    Set Materiales = DAOMateriales.FindAll(filtro)

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
    Set Selecciones.Material = Nothing
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
    grilla_SelectionChange
    If Usable Then
        Set Selecciones.Material = rectemp
        Unload Me
    End If

End Sub

Private Sub grilla_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 67 And Shift = 2 Then    'CTRL + C
        GridEXHelper.Grid2Clipboard Me.grilla
        DoEvents
        MsgBox "La lista ha sido copiada al portapapeles.", vbInformation + vbOKOnly
    End If
End Sub

Private Sub grilla_RowFormat(RowBuffer As GridEX20.JSRowData)
    If RowBuffer.value(8) = 0 Then
        RowBuffer.RowStyle = "Valor0"
    End If
End Sub
Private Sub grilla_SelectionChange()
    On Error GoTo err1
    rows = grilla.RowIndex(grilla.row)
    Set rectemp = Materiales.item(grilla.RowIndex(grilla.row))
    Exit Sub
err1:
End Sub
Private Sub grilla_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Set rectemp = Materiales.item(RowIndex)
    With rectemp
        Values(1) = .codigo
        Values(2) = .Grupo.rubros.rubro
        Values(3) = .Grupo.Grupo
        Values(4) = .descripcion
        Values(5) = .Espesor
        Values(6) = .unidad
        Values(7) = .PesoXUnidad
        Values(8) = .Valor
        Values(9) = .moneda.NombreCorto
        Values(10) = .FechaValor
        Values(11) = .Cantidad
        Values(12) = .almacen.almacen
        Values(13) = enums.enumEstadoMaterial(.estado)
    End With
End Sub
Private Sub llenar_Grilla()
    Set Materiales = DAOMateriales.FindAll
    grilla.ItemCount = Materiales.count
    'grilla.ReBind
End Sub
Private Sub editamos()
    If grilla.rowcount > 0 Then
        A = grilla.RowIndex(grilla.row)
        If A = 0 Then Exit Sub

        Dim frm1 As New frmMaterialesNuevo
        frm1.Material = Materiales(grilla.RowIndex(grilla.row))
        frm1.Show
    End If
End Sub

Private Property Get ISuscriber_id() As String
    ISuscriber_id = vId
End Property

Private Function ISuscriber_Notificarse(EVENTO As clsEventoObserver) As Variant
    Dim tmp As clsMaterial
    If EVENTO.EVENTO = agregar_ Then
        llenar_Grilla
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

                grilla.RefreshRowIndex i
                Exit For
            End If
        Next

    End If

End Function


