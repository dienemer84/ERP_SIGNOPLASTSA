VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmMaterialesLista 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Lista de materiales.."
   ClientHeight    =   5925
   ClientLeft      =   60
   ClientTop       =   270
   ClientWidth     =   16485
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4746.124
   ScaleMode       =   0  'User
   ScaleWidth      =   18756.04
   Begin VB.CommandButton cmdFiltrar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Filtrar"
      Default         =   -1  'True
      Height          =   375
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5520
      Width           =   1095
   End
   Begin VB.TextBox txtFiltro 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   720
      TabIndex        =   0
      Top             =   5520
      Width           =   3735
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5520
      Width           =   1095
   End
   Begin GridEX20.GridEX grilla 
      Height          =   5415
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   16455
      _ExtentX        =   29025
      _ExtentY        =   9551
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
      Column(1)       =   "frmListaMateriales.frx":0000
      Column(2)       =   "frmListaMateriales.frx":00F4
      Column(3)       =   "frmListaMateriales.frx":01E0
      Column(4)       =   "frmListaMateriales.frx":02CC
      Column(5)       =   "frmListaMateriales.frx":03A4
      Column(6)       =   "frmListaMateriales.frx":0474
      Column(7)       =   "frmListaMateriales.frx":059C
      Column(8)       =   "frmListaMateriales.frx":06CC
      Column(9)       =   "frmListaMateriales.frx":0798
      Column(10)      =   "frmListaMateriales.frx":0868
      Column(11)      =   "frmListaMateriales.frx":0954
      Column(12)      =   "frmListaMateriales.frx":0A80
      Column(13)      =   "frmListaMateriales.frx":0B50
      FormatStylesCount=   7
      FormatStyle(1)  =   "frmListaMateriales.frx":0C20
      FormatStyle(2)  =   "frmListaMateriales.frx":0D48
      FormatStyle(3)  =   "frmListaMateriales.frx":0DF8
      FormatStyle(4)  =   "frmListaMateriales.frx":0EAC
      FormatStyle(5)  =   "frmListaMateriales.frx":0F84
      FormatStyle(6)  =   "frmListaMateriales.frx":103C
      FormatStyle(7)  =   "frmListaMateriales.frx":111C
      ImageCount      =   0
      PrinterProperties=   "frmListaMateriales.frx":11F8
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Filtro"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   5520
      Width           =   975
   End
   Begin VB.Menu mnu 
      Caption         =   "menu"
      Visible         =   0   'False
      Begin VB.Menu codigo 
         Caption         =   "codigo"
         Enabled         =   0   'False
      End
      Begin VB.Menu editar 
         Caption         =   "Editar..."
      End
      Begin VB.Menu histerico 
         Caption         =   "Histórico de precios..."
      End
   End
End
Attribute VB_Name = "frmMaterialesLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Implements ISuscriber
Dim vid As String
Dim rows As Long
Dim Materiales As New Collection
Dim rectemp As clsMaterial
Private buscando As Boolean


Private Sub cmdFiltrar_Click()
    Set Materiales = DAOMateriales.FindAll("(" & DAOMateriales.TABLA_MATERIALES & "." & DAOMateriales.CAMPO_DESCRIPCION & " LIKE '%" & Trim(Me.txtFiltro.text) & "%' or " & DAOMateriales.TABLA_MATERIALES & "." & DAOMateriales.CAMPO_CODIGO & " LIKE '%" & Trim(Me.txtFiltro) & "%')")

    grilla.ItemCount = Materiales.count
    grilla.ReBind
    GridEXHelper.AutoSizeColumns Me.grilla, True
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

 '   If rows > 0 Then grilla.RefreshRowIndex rows
  '  grilla.ItemCount = Materiales.count
    
    Me.grilla.ReBind
End Sub
Private Sub Form_Load()
    FormHelper.Customize Me
    GridEXHelper.customizeGrid Me.grilla, True
    grilla.ItemCount = 0

    rows = 1
    llenar_Grilla
    vid = funciones.CreateGUID
    Channel.AgregarSuscriptor Me, Materiales_
End Sub
Private Sub Form_Resize()
    On Error Resume Next
    Me.grilla.Height = Me.ScaleHeight - (Me.Command2.Height + (500 - Me.Command2.Height))
    Me.grilla.Width = Me.ScaleWidth
    Me.grilla.ColumnAutoResize = True
    Me.Command2.Top = Me.ScaleHeight - 400
    'Me.Command3.Top = Me.Command2.Top
    Me.cmdFiltrar.Top = Me.Command2.Top
    Me.txtFiltro.Top = Me.Command2.Top
  Me.Label1.Top = Me.txtFiltro.Top
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
Private Sub grilla_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If grilla.RowIndex(grilla.Row) = 0 Then Exit Sub
    If Button = 2 Then
        Me.PopupMenu mnu
    End If
End Sub
Private Sub grilla_RowFormat(RowBuffer As GridEX20.JSRowData)
    If RowBuffer.value(8) = 0 Then
        RowBuffer.RowStyle = "Valor0"
    End If
End Sub
Private Sub grilla_SelectionChange()
    rows = grilla.RowIndex(grilla.Row)
End Sub
Private Sub grilla_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Set rectemp = Materiales.item(RowIndex)
    With rectemp
        Values(1) = .codigo
        Values(2) = .grupo.rubros.Rubro
        Values(3) = .grupo.grupo
        Values(4) = .Descripcion
        Values(5) = .Espesor
        Values(6) = .unidad
        Values(7) = .PesoXUnidad
        Values(8) = .valor
        Values(9) = .Moneda.NombreCorto
        Values(10) = .FechaValor
        Values(11) = .cantidad
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
    If grilla.RowCount > 0 Then
        a = grilla.RowIndex(grilla.Row)
        If a = 0 Then Exit Sub
        
        Dim frm1 As New frmMaterialesNuevo
        frm1.Material = Materiales(grilla.RowIndex(grilla.Row))
        frm1.Show
    End If
End Sub

Private Property Get ISuscriber_id() As String
ISuscriber_id = vid
End Property

Private Function ISuscriber_Notificarse(EVENTO As clsEventoObserver) As Variant
Dim tmp As clsMaterial
If EVENTO.EVENTO = agregar_ Then
    llenar_Grilla
ElseIf EVENTO.EVENTO = modificar_ Then

            Set tmp = EVENTO.Elemento
    
        For i = Materiales.count To 1 Step -1
            If Materiales(i).id = tmp.id Then
                Set rectemp = Materiales(i)
                rectemp.id = tmp.id
                rectemp.valor = tmp.valor
                rectemp.Descripcion = tmp.Descripcion
                rectemp.grupo = tmp.grupo
                rectemp.estado = tmp.estado
                rectemp.almacen = tmp.almacen
                rectemp.cantidad = tmp.cantidad
                rectemp.FechaValor = tmp.FechaValor
                
                grilla.RefreshRowIndex i
                Exit For
            End If
        Next

End If

End Function
