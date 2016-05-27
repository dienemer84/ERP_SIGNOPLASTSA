VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmComprasRequesProcesar 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Procesar requerimiento"
   ClientHeight    =   7545
   ClientLeft      =   1395
   ClientTop       =   5715
   ClientWidth     =   14505
   Icon            =   "frmComprasRequesProcesar.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   14505
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Agregar proveedor al item"
      Height          =   345
      Left            =   4305
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7035
      UseMaskColor    =   -1  'True
      Width           =   2145
   End
   Begin VB.ComboBox cboProveedores 
      Height          =   315
      Left            =   150
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   7035
      Width           =   4095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Guardar"
      Height          =   375
      Left            =   12735
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4725
      Width           =   1620
   End
   Begin GridEX20.GridEX grilla_materiales 
      Height          =   3480
      Left            =   120
      TabIndex        =   11
      Top             =   1185
      Width           =   14250
      _ExtentX        =   25135
      _ExtentY        =   6138
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      ColumnAutoResize=   -1  'True
      MethodHoldFields=   -1  'True
      AllowEdit       =   0   'False
      BorderStyle     =   3
      GroupByBoxVisible=   0   'False
      BackColorHeader =   16761024
      DataMode        =   99
      HeaderFontBold  =   -1  'True
      HeaderFontWeight=   700
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   13
      Column(1)       =   "frmComprasRequesProcesar.frx":000C
      Column(2)       =   "frmComprasRequesProcesar.frx":012C
      Column(3)       =   "frmComprasRequesProcesar.frx":0220
      Column(4)       =   "frmComprasRequesProcesar.frx":030C
      Column(5)       =   "frmComprasRequesProcesar.frx":03F8
      Column(6)       =   "frmComprasRequesProcesar.frx":04FC
      Column(7)       =   "frmComprasRequesProcesar.frx":05F4
      Column(8)       =   "frmComprasRequesProcesar.frx":06F0
      Column(9)       =   "frmComprasRequesProcesar.frx":07D4
      Column(10)      =   "frmComprasRequesProcesar.frx":08B8
      Column(11)      =   "frmComprasRequesProcesar.frx":099C
      Column(12)      =   "frmComprasRequesProcesar.frx":0A90
      Column(13)      =   "frmComprasRequesProcesar.frx":0B84
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmComprasRequesProcesar.frx":0C78
      FormatStyle(2)  =   "frmComprasRequesProcesar.frx":0DB0
      FormatStyle(3)  =   "frmComprasRequesProcesar.frx":0E60
      FormatStyle(4)  =   "frmComprasRequesProcesar.frx":0F14
      FormatStyle(5)  =   "frmComprasRequesProcesar.frx":0FEC
      FormatStyle(6)  =   "frmComprasRequesProcesar.frx":10A4
      ImageCount      =   0
      PrinterProperties=   "frmComprasRequesProcesar.frx":1184
   End
   Begin GridEX20.GridEX grilla_proveedores 
      Height          =   2205
      Left            =   135
      TabIndex        =   12
      Top             =   4770
      Width           =   6285
      _ExtentX        =   11086
      _ExtentY        =   3889
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      ColumnAutoResize=   -1  'True
      MultiSelect     =   -1  'True
      MethodHoldFields=   -1  'True
      AllowDelete     =   -1  'True
      AllowEdit       =   0   'False
      BorderStyle     =   3
      GroupByBoxVisible=   0   'False
      BackColorHeader =   16761024
      RowHeaders      =   -1  'True
      DataMode        =   99
      HeaderFontBold  =   -1  'True
      HeaderFontWeight=   700
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   1
      Column(1)       =   "frmComprasRequesProcesar.frx":135C
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmComprasRequesProcesar.frx":1438
      FormatStyle(2)  =   "frmComprasRequesProcesar.frx":1570
      FormatStyle(3)  =   "frmComprasRequesProcesar.frx":1620
      FormatStyle(4)  =   "frmComprasRequesProcesar.frx":16D4
      FormatStyle(5)  =   "frmComprasRequesProcesar.frx":17AC
      FormatStyle(6)  =   "frmComprasRequesProcesar.frx":1864
      ImageCount      =   0
      PrinterProperties=   "frmComprasRequesProcesar.frx":1944
   End
   Begin VB.Label lblDestino 
      BackColor       =   &H00FF8080&
      Height          =   255
      Left            =   1080
      TabIndex        =   7
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label lblSector 
      BackColor       =   &H00FF8080&
      Height          =   255
      Left            =   1080
      TabIndex        =   6
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label lblFecha 
      BackColor       =   &H00FF8080&
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label Label5 
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
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   855
   End
   Begin VB.Label lblNumero 
      BackColor       =   &H00FF8080&
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      Caption         =   "Destino"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      Caption         =   "Sector"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      Caption         =   "Número"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
   Begin VB.Menu procesar 
      Caption         =   "procesar"
      Visible         =   0   'False
      Begin VB.Menu proveedores 
         Caption         =   "Proveedores..."
      End
   End
End
Attribute VB_Name = "frmComprasRequesProcesar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As Recordset
Dim classC As New classCompras
Dim claseS As New classStock
Dim vIdReque As Long
Dim Concepto As Boolean
Dim vEstado As Integer
Dim tmpProveedor As clsProveedor
Dim vProveedores As Collection
Dim vProveedores_conceptos As Collection
Dim tmpMaterial As clsRequeMateriales
Dim vReque As clsRequerimiento
Dim vRequeTmp As clsRequerimiento
Public ReadOnly As Boolean

Public Property Let reque(nvalue As clsRequerimiento)
    Set vReque = nvalue
End Property
Public Property Let idReque(nIdReque)
    vIdReque = nIdReque
End Property
Private Sub Command1_Click()
    Unload Me
End Sub
Private Sub Command2_Click()
    'If vReque.ValidarProveedores Then
    If MsgBox("¿Está seguro de guardar?", vbYesNo + vbQuestion) = vbYes Then
        If Not DAORequerimiento.Save(vReque) Then
            MsgBox "Se produjo algún error, no se guardarán los cambios!", vbCritical, "Error"
        Else
            MsgBox "Guardado correctamente.", vbInformation
            Unload Me
        End If
    End If
    'Else
    '    MsgBox "Debe asignar al menos un proveedor a cada item del requerimiento!", vbCritical, "Error"
    'End If
End Sub

Private Sub Command3_Click()
    If Me.cboProveedores.ListIndex = -1 Then Exit Sub


    Dim IdElegido As Long
    Dim esta As Boolean
    'agregar un proveedor a la coleccion y mostrarlo
    esta = False
    IdElegido = Me.cboProveedores.ItemData(Me.cboProveedores.ListIndex)
    For h = 1 To vProveedores.count
        If vProveedores.item(h).id = IdElegido Then
            esta = True
            Exit For
        End If
    Next h
    If Not esta Then
        vProveedores.Add DAOProveedor.FindById(IdElegido)
        Me.grilla_proveedores.ItemCount = vProveedores.count
    End If
End Sub
Private Sub Form_Load()
    FormHelper.Customize Me
    GridEXHelper.CustomizeGrid Me.grilla_materiales
    GridEXHelper.CustomizeGrid Me.grilla_proveedores
    'DAOProveedor.LlenarCombo Me.cboProveedores, True
    Me.grilla_materiales.ItemCount = 0
    Me.grilla_proveedores.ItemCount = 0

    Me.cboProveedores.Enabled = Not ReadOnly
    Me.Command3.Enabled = Not ReadOnly
    Me.Command2.Enabled = Not ReadOnly
    Me.grilla_proveedores.AllowDelete = Not ReadOnly

    cargarDatosReque
End Sub
Private Sub cargarDatosReque()
    Set vRequeTmp = DAORequerimiento.FindById(vReque.id, True, True, True, True)
    If vRequeTmp.Guardado <> vReque.Guardado Then
        MsgBox "Error de inconcistencia de datos!", vbCritical, "Error"
        Exit Sub
    End If
    Me.caption = "Procesar proveedores requerimiento Nº " & vReque.id
    Me.lblSector = vReque.Sector.Sector
    Me.lblFecha = vReque.FechaCreado
    Me.lblNumero = vReque.id
    Me.lblDestino = vReque.StringDestino
    Me.grilla_materiales.ItemCount = vReque.Materiales.count
End Sub
Private Sub mostrarMenu()
    Me.PopupMenu Me.procesar
End Sub
Private Sub grilla_materiales_SelectionChange()
    a = grilla_materiales.RowIndex(grilla_materiales.row)
    Set tmpMaterial = vReque.Materiales.item(a)
    Me.grilla_proveedores.ItemCount = 0
    Set vProveedores = tmpMaterial.ListaProveedores
    Me.grilla_proveedores.ItemCount = vProveedores.count

    DAOProveedor.LlenarCombo Me.cboProveedores, True, , , tmpMaterial.Material.Grupo.rubros.id

    Me.cboProveedores.Enabled = (tmpMaterial.estado = EnProceso_ And Not ReadOnly)
    Me.Command3.Enabled = (tmpMaterial.estado = EnProceso_ And Not ReadOnly)
    Me.grilla_proveedores.AllowDelete = (tmpMaterial.estado = EnProceso_ And Not ReadOnly)
End Sub

Private Sub grilla_materiales_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Set tmpMaterial = vReque.Materiales.item(RowIndex)
    With tmpMaterial
        Values(1) = tmpMaterial.Cantidad
        Values(2) = tmpMaterial.Material.codigo
        Values(3) = tmpMaterial.Material.Grupo.rubros.Rubro
        Values(4) = tmpMaterial.Material.Grupo.Grupo
        Values(5) = tmpMaterial.Material.descripcion
        Values(6) = tmpMaterial.Material.Espesor
        Values(7) = tmpMaterial.Cantidad & "x" & tmpMaterial.Material.Largo & "x" & tmpMaterial.Material.Ancho
        Values(8) = funciones.FormatearDecimales(tmpMaterial.m2, 2)
        Values(9) = funciones.FormatearDecimales(tmpMaterial.ML, 2)
        Values(10) = funciones.FormatearDecimales(tmpMaterial.Kg, 2)
        Values(11) = enums.enumUnidades(tmpMaterial.Material.unidad)
        Values(12) = tmpMaterial.Observaciones
        Values(13) = enums.enumEstadoRequeCompra(tmpMaterial.estado)
    End With
End Sub
Private Sub grilla_proveedores_BeforeDelete(ByVal Cancel As GridEX20.JSRetBoolean)
    Cancel = (MsgBox("¿Está seguro de eliminar el Proveedor del Requisito?", vbYesNo, "Confirmación") = vbNo)
End Sub
Private Sub grilla_proveedores_UnboundDelete(ByVal RowIndex As Long, ByVal Bookmark As Variant)
    vProveedores.remove RowIndex
End Sub
Private Sub grilla_proveedores_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Set tmpProveedor = vProveedores.item(RowIndex)
    With tmpProveedor
        Values(1) = .RazonSocial
    End With
End Sub


