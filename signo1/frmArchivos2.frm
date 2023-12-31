VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{E027DF0E-25C2-4A43-8A2F-914B8380CFF5}#1.0#0"; "Dialogos.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~3.OCX"
Begin VB.Form frmArchivos2 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FF8080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Archivos"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10065
   ClipControls    =   0   'False
   Icon            =   "frmArchivos2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   4440
   ScaleWidth      =   10065
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.GroupBox GroupBox3 
      Height          =   2655
      Left            =   5640
      TabIndex        =   12
      Top             =   120
      Width           =   4335
      _Version        =   786432
      _ExtentX        =   7646
      _ExtentY        =   4683
      _StockProps     =   79
      Caption         =   "Seleccionar Archivos"
      UseVisualStyle  =   -1  'True
      Begin VB.TextBox txtNuevoComentario 
         Height          =   855
         Left            =   120
         MultiLine       =   -1  'True
         OLEDropMode     =   1  'Manual
         TabIndex        =   16
         Top             =   1200
         Width           =   4095
      End
      Begin VB.TextBox txtNombre 
         Height          =   285
         Left            =   1080
         OLEDropMode     =   1  'Manual
         TabIndex        =   15
         Top             =   600
         Width           =   3120
      End
      Begin VB.TextBox txtRuta 
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         OLEDropMode     =   1  'Manual
         TabIndex        =   14
         Top             =   240
         Width           =   3120
      End
      Begin XtremeSuiteControls.PushButton Command1 
         Height          =   375
         Left            =   1920
         TabIndex        =   13
         Top             =   2160
         Width           =   1095
         _Version        =   786432
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Subir"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton Command2 
         Height          =   375
         Left            =   3120
         TabIndex        =   17
         Top             =   2160
         Width           =   1095
         _Version        =   786432
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Abrir"
         UseVisualStyle  =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "Nombre"
         Height          =   195
         Left            =   450
         TabIndex        =   20
         Top             =   615
         Width           =   555
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "Archivo"
         Height          =   195
         Left            =   465
         TabIndex        =   19
         Top             =   270
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "Comentario"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   960
         Width           =   795
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox2 
      Height          =   4575
      Left            =   120
      TabIndex        =   4
      Top             =   4440
      Width           =   9855
      _Version        =   786432
      _ExtentX        =   17383
      _ExtentY        =   8070
      _StockProps     =   79
      Caption         =   "Buscador de Archivos"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.PushButton cmdSeleccionarDirectorio 
         Height          =   255
         Left            =   4560
         TabIndex        =   11
         Top             =   240
         Width           =   375
         _Version        =   786432
         _ExtentX        =   661
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "..."
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton cmdIniciarBusqueda 
         Height          =   255
         Left            =   3600
         TabIndex        =   10
         Top             =   600
         Width           =   1335
         _Version        =   786432
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Buscar"
         UseVisualStyle  =   -1  'True
      End
      Begin VB.TextBox txtUbicacionInicial 
         Height          =   285
         Left            =   1425
         TabIndex        =   7
         Text            =   "\\servidor\produccion\"
         Top             =   240
         Width           =   3015
      End
      Begin VB.TextBox txtPatron 
         Height          =   285
         Left            =   1425
         TabIndex        =   6
         Text            =   "*.pdf"
         Top             =   600
         Width           =   2055
      End
      Begin GridEX20.GridEX gridBuscador 
         Height          =   3405
         Left            =   120
         TabIndex        =   5
         Top             =   1005
         Width           =   9600
         _ExtentX        =   16933
         _ExtentY        =   6006
         Version         =   "2.0"
         HoldSortSettings=   -1  'True
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         ColumnAutoResize=   -1  'True
         DetectRowDrag   =   -1  'True
         HideSelection   =   2
         MethodHoldFields=   -1  'True
         ContScroll      =   -1  'True
         AllowCardSizing =   0   'False
         AllowColumnDrag =   0   'False
         AutomaticArrange=   0   'False
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         BackColorHeader =   16761024
         ImageCount      =   1
         ImagePicture1   =   "frmArchivos2.frx":000C
         ItemCount       =   1
         DataMode        =   99
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   2
         Column(1)       =   "frmArchivos2.frx":01E6
         Column(2)       =   "frmArchivos2.frx":031E
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmArchivos2.frx":0476
         FormatStyle(2)  =   "frmArchivos2.frx":05AE
         FormatStyle(3)  =   "frmArchivos2.frx":065E
         FormatStyle(4)  =   "frmArchivos2.frx":0712
         FormatStyle(5)  =   "frmArchivos2.frx":07EA
         FormatStyle(6)  =   "frmArchivos2.frx":08A2
         ImageCount      =   1
         ImagePicture(1) =   "frmArchivos2.frx":0982
         PrinterProperties=   "frmArchivos2.frx":0B5C
      End
      Begin VB.Label lblUbicacionInicial 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "Ubicacion Inicial"
         Height          =   195
         Left            =   195
         TabIndex        =   9
         Top             =   285
         Width           =   1170
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "Patron B�squeda"
         Height          =   195
         Left            =   150
         TabIndex        =   8
         Top             =   645
         Width           =   1230
      End
   End
   Begin XtremeSuiteControls.PushButton cmdBuscadorArchivos 
      Height          =   375
      Left            =   5760
      TabIndex        =   3
      Top             =   2880
      Width           =   3135
      _Version        =   786432
      _ExtentX        =   5530
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Buscar"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton Command3 
      Height          =   375
      Left            =   9000
      TabIndex        =   2
      Top             =   2880
      Width           =   975
      _Version        =   786432
      _ExtentX        =   1720
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Salir"
      UseVisualStyle  =   -1  'True
   End
   Begin GridEX20.GridEX grid 
      Height          =   4245
      Left            =   105
      TabIndex        =   1
      Top             =   120
      Width           =   5475
      _ExtentX        =   9657
      _ExtentY        =   7488
      Version         =   "2.0"
      PreviewRowIndent=   500
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      PreviewColumn   =   4
      PreviewRowLines =   2
      MethodHoldFields=   -1  'True
      AllowColumnDrag =   0   'False
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      BackColorHeader =   16761024
      ItemCount       =   1
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   5
      Column(1)       =   "frmArchivos2.frx":0D34
      Column(2)       =   "frmArchivos2.frx":0E48
      Column(3)       =   "frmArchivos2.frx":0FA0
      Column(4)       =   "frmArchivos2.frx":1070
      Column(5)       =   "frmArchivos2.frx":1164
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmArchivos2.frx":1240
      FormatStyle(2)  =   "frmArchivos2.frx":1378
      FormatStyle(3)  =   "frmArchivos2.frx":1428
      FormatStyle(4)  =   "frmArchivos2.frx":14DC
      FormatStyle(5)  =   "frmArchivos2.frx":15B4
      FormatStyle(6)  =   "frmArchivos2.frx":166C
      ImageCount      =   0
      PrinterProperties=   "frmArchivos2.frx":174C
   End
   Begin Dialogo.UserControl1 UserControl11 
      Height          =   375
      Left            =   7920
      TabIndex        =   0
      Top             =   3720
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   5640
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Image picture1 
      Height          =   4695
      Left            =   10080
      Stretch         =   -1  'True
      Top             =   120
      Width           =   3975
   End
   Begin VB.Menu archivos 
      Caption         =   "archivos"
      Visible         =   0   'False
      Begin VB.Menu export 
         Caption         =   "Exportar..."
      End
      Begin VB.Menu abrir 
         Caption         =   "Abrir..."
      End
   End
End
Attribute VB_Name = "frmArchivos2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vorigen As OrigenArchivos
Dim nombre As String
Dim idPieza As Long
Dim vruta As String
Dim clasea As New classArchivos
Private initialHeight As Long
Private normalHeight As Long
Private archivosEncontrados As Collection
Private tmpFileMetadataDTO As FileMetadataDTO
Private m_Archivos As New Collection
Private archivoActual As Archivo

Const LARGO_PREVIEW = 14340
Const LARGO_NORMAL = 10155

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
                                      (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
                                       ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Public Property Let origen(nOrigen As OrigenArchivos)
    vorigen = nOrigen
End Property
Public Property Let ruta(nRuta As String)
    vruta = nRuta
End Property
Public Property Let lblIdPieza(nIdPieza As Long)
    idPieza = nIdPieza
End Property

Private Sub abrir_Click()
    grid_DblClick
End Sub

Private Sub cmdCancelarBusqueda_Click()
    buscandoArchivos = False
End Sub



Private Sub cmdIniciarBusqueda_Click()
    Set archivosEncontrados = New Collection
    If Right$(Me.txtUbicacionInicial.text, 1) <> "\" Then Me.txtUbicacionInicial.text = Me.txtUbicacionInicial.text & "\"
    DoEvents
    funciones.ListFiles archivosEncontrados, Me.txtUbicacionInicial.text, Me.txtPatron.text
    Me.gridBuscador.ItemCount = archivosEncontrados.count
    GridEXHelper.AutoSizeColumns Me.gridBuscador, True
    Me.gridBuscador.Refresh
End Sub

Private Sub cmdSeleccionarDirectorio_Click()
    Dim T As String
    T = BrowseForDirectory("Seleccione el directorio en donde buscar")
    If LenB(T) > 0 Then
        Me.txtUbicacionInicial.text = T
    End If
End Sub

Private Sub Command1_Click()
    If Me.txtRuta <> Empty Then
        If MsgBox("�Desea subir el archivo?", vbYesNo + vbQuestion, "Confirmaci�n") = vbYes Then
            If (clasea.grabarArchivo(idPieza, nombre, vruta, Me.txtNuevoComentario, CInt(vorigen))) Then
                llenarLST
                Me.Refresh
                
            Else
                MsgBox "Se produjo un error al guardar el archivo!", vbCritical, "Error"
            End If
        End If
    End If
End Sub

Private Sub llenarLST()
    Dim filter As String
    filter = DAOArchivo.TABLA_ARCHIVO & "." & DAOArchivo.CAMPO_ID_REFERENCIA & "=" & idPieza

    Set m_Archivos = DAOArchivo.FindAll(vorigen, filter)
    Me.grid.ItemCount = 0
    Me.grid.ItemCount = m_Archivos.count

    Me.Refresh
    GridEXHelper.AutoSizeColumns Me.grid, True

End Sub

Private Sub Command2_Click()

    On Error GoTo err1
    CD.ShowOpen
    vruta = CD.FileName
    nombre = funciones.GetFileName(vruta)

    Me.txtNombre = nombre
    Me.txtRuta = vruta

err1:
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub


Private Sub cmdBuscadorArchivos_Click()
    
    

If Me.Height = initialHeight Then
         Me.Height = normalHeight
Else
        Me.Height = initialHeight
End If



End Sub
Private Sub export_Click()
    If Not archivoActual Is Nothing Then
        'Dim Id As Long
        'Id = CLng(Me.ListView1.SelectedItem)
        descargar archivoActual.id
    End If
End Sub

Private Sub Form_Load()
FormHelper.Customize Me
    buscandoArchivos = False
    normalHeight = 4920

    'Me.Width = 10125
    initialHeight = Me.Height
    Me.Height = normalHeight

    Me.gridBuscador.ItemCount = 0

    On Error GoTo err123
    If Trim(vruta) <> Empty Then
        Me.txtRuta = vruta
        CD.FileName = vruta
        nombre = funciones.GetFileName(vruta)
        Me.txtNombre = nombre
    End If

    GridEXHelper.customizeGrid Me.grid
    GridEXHelper.customizeGrid Me.gridBuscador
    llenarLST
    Exit Sub
err123:
End Sub



Private Sub drag_drop(Data)
    vruta = Data.Files(1)
    Me.txtRuta = vruta
    nombre = funciones.GetFileName(vruta)
    Me.txtNombre = nombre


End Sub






Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    drag_drop Data
End Sub


Private Sub fraBuscador_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    drag_drop Data
End Sub

Private Sub Frame2_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    drag_drop Data
End Sub

Private Sub Frame3_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    drag_drop Data
End Sub

Private Sub grid_ColumnHeaderClick(ByVal Column As GridEX20.JSColumn)
    GridEXHelper.ColumnHeaderClick Me.grid, Column
End Sub

Private Sub grid_DblClick()
    If Not archivoActual Is Nothing Then
        grid_SelectionChange
        abrir_documento archivoActual.id
    End If
End Sub

Private Sub grid_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 And Not archivoActual Is Nothing Then
        Me.PopupMenu archivos
    End If
End Sub

Private Sub grid_OLEDragDrop(Data As GridEX20.JSDataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    drag_drop Data
End Sub

Private Sub grid_SelectionChange()
    Dim idx As Long: idx = Me.grid.RowIndex(Me.grid.Row)

    If idx > 0 Then
        Set archivoActual = m_Archivos(idx)

        Me.Picture1 = clasea.previewImage(archivoActual.id)

        If Me.Picture1 = 0 Then
            Me.Width = LARGO_NORMAL
        Else
            Me.Width = LARGO_PREVIEW
        End If

    Else
        Set archivoActual = Nothing
    End If

End Sub

Private Sub grid_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex > 0 And RowIndex <= m_Archivos.count Then
        'Private m_archivos As New Collection
        Set archivoActual = m_Archivos(RowIndex)
        With archivoActual
            Values(1) = .nombre
            Values(2) = .FileSizeInKB
            If .Usuario Is Nothing Then
                Values(3) = vbNullString
            Else
                Values(3) = .Usuario.Usuario
            End If
            Values(4) = .Comentario
            If CDbl(.FechaSubida) > 0 Then Values(5) = .FechaSubida
        End With
    End If
End Sub

Private Sub gridBuscador_ColumnHeaderClick(ByVal Column As GridEX20.JSColumn)
    GridEXHelper.ColumnHeaderClick Me.gridBuscador, Column
End Sub

Private Sub gridBuscador_DblClick()


    If Me.gridBuscador.Row > 0 Then
        Dim idx As Long
        idx = Me.gridBuscador.RowIndex(Me.gridBuscador.Row)

        Dim dto As FileMetadataDTO

        If idx <= archivosEncontrados.count Then
            Set dto = archivosEncontrados(idx)

            ShellExecute Me.hwnd, "open", dto.FullFilePath, "", "", 4
        End If
    End If


End Sub

Private Sub gridBuscador_FetchIcon(ByVal RowIndex As Long, ByVal ColIndex As Integer, ByVal RowBookmark As Variant, ByVal IconIndex As GridEX20.JSRetInteger)
    If ColIndex = 1 Then IconIndex = 1
End Sub

Private Sub gridBuscador_SelectionChange()
    If Me.gridBuscador.Row > 0 Then
        Dim idx As Long
        idx = Me.gridBuscador.RowIndex(Me.gridBuscador.Row)

        Dim dto As FileMetadataDTO

        If idx <= archivosEncontrados.count Then
            Set dto = archivosEncontrados(idx)

            Me.txtNuevoComentario.text = vbNullString

            vruta = dto.FullFilePath
            nombre = dto.FileName

            Me.txtNombre = nombre
            Me.txtRuta = vruta
        End If
    End If
End Sub

Private Sub gridBuscador_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)


    If archivosEncontrados.count > 0 Then
        Set tmpFileMetadataDTO = archivosEncontrados.item(RowIndex)
        With tmpFileMetadataDTO
            Values(1) = tmpFileMetadataDTO.DirectoryName & "\" & tmpFileMetadataDTO.FileName
            Values(2) = tmpFileMetadataDTO.FileSizeInKB
        End With
    End If
End Sub





Private Sub descargar(id As Long)

    If Not archivoActual Is Nothing Then
        On Error GoTo err44
        Set dia = frmPrincipal.CD
        clasea.ejecutar "select nombre from archivos where id=" & archivoActual.id
        a = archivoActual.nombre
        dia.FileName = a
        dia.ShowSave

        If dia.FileName <> a Then
            If MsgBox("�Est� seguro de exportar?", vbYesNo + vbQuestion, "Confirmaci�n") = vbYes Then
                clasea.exportarArchivo archivoActual.id, dia.FileName
                'claseA.GetDocumento CLng(Me.ListView1.SelectedItem), dia.FileName
            End If

        End If
    End If

    Exit Sub
err44:


End Sub


Private Sub loadFile(Data)

End Sub


Private Sub abrir_documento(id)
    If Not archivoActual Is Nothing Then
        Dim a As String
        On Error GoTo era
        clasea.ejecutar "select nombre from archivos where id=" & archivoActual.id
        'a = App.Path & "\" & archivoActual.nombre 'claseA.nombre
        clasea.exportarArchivo archivoActual.id, a, True
        Exit Sub
    End If
era:

End Sub


Private Sub txtNuevoComentario_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    vruta = Data.Files(1)
    Me.txtRuta = vruta
    nombre = funciones.GetFileName(vruta)
    Me.txtNombre = nombre
End Sub
Private Sub txtRuta_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    drag_drop Data
End Sub



