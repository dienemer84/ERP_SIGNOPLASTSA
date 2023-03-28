VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmArchivos2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Archivos"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   11265
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmArchivos2.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   11265
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   10020
      Top             =   4440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin XtremeSuiteControls.GroupBox grpArchivosASubir 
      Height          =   3450
      Left            =   6555
      TabIndex        =   2
      Top             =   45
      Width           =   4605
      _Version        =   786432
      _ExtentX        =   8123
      _ExtentY        =   6085
      _StockProps     =   79
      Caption         =   "Archivos a subir"
      UseVisualStyle  =   -1  'True
      Begin VB.CheckBox chkArchivosCompra 
         Caption         =   "Archivos de compra"
         Height          =   195
         Left            =   2220
         TabIndex        =   9
         Top             =   2640
         Width           =   1995
      End
      Begin VB.TextBox txtComentario 
         Height          =   1035
         Left            =   2220
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   1500
         Width           =   2280
      End
      Begin VB.ListBox lstArchivos 
         Height          =   2400
         Left            =   105
         OLEDropMode     =   1  'Manual
         TabIndex        =   4
         Top             =   495
         Width           =   2025
      End
      Begin XtremeSuiteControls.PushButton btnSeleccionar 
         Height          =   330
         Left            =   2190
         TabIndex        =   3
         Top             =   495
         Width           =   2325
         _Version        =   786432
         _ExtentX        =   4101
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "Seleccionar archivos a subir..."
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnLimpiarLista 
         Height          =   330
         Left            =   2190
         TabIndex        =   5
         Top             =   870
         Width           =   2325
         _Version        =   786432
         _ExtentX        =   4101
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "Limpiar lista archivos a subir"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnSubir 
         Height          =   330
         Left            =   105
         TabIndex        =   6
         Top             =   3000
         Width           =   4395
         _Version        =   786432
         _ExtentX        =   7752
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "Subir archivos de la lista"
         UseVisualStyle  =   -1  'True
      End
      Begin VB.Label lblRuta 
         Caption         =   "Ruta: "
         Height          =   195
         Left            =   135
         TabIndex        =   10
         Top             =   225
         Width           =   4290
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Comentario"
         Height          =   195
         Left            =   2235
         TabIndex        =   7
         Top             =   1260
         Width           =   825
      End
   End
   Begin XtremeSuiteControls.GroupBox grpArchivos 
      Height          =   7155
      Left            =   105
      TabIndex        =   0
      Top             =   45
      Width           =   6345
      _Version        =   786432
      _ExtentX        =   11192
      _ExtentY        =   12621
      _StockProps     =   79
      Caption         =   "Archivos cargados del objeto"
      UseVisualStyle  =   -1  'True
      Begin GridEX20.GridEX gridArchivos 
         Height          =   6810
         Left            =   120
         TabIndex        =   1
         Top             =   225
         Width           =   6090
         _ExtentX        =   10742
         _ExtentY        =   12012
         Version         =   "2.0"
         PreviewRowIndent=   200
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         PreviewColumn   =   "comentarionombre"
         PreviewRowLines =   4
         MethodHoldFields=   -1  'True
         GroupByBoxVisible=   0   'False
         DataMode        =   99
         ColumnHeaderHeight=   285
         IntProp1        =   0
         ColumnsCount    =   5
         Column(1)       =   "frmArchivos2.frx":000C
         Column(2)       =   "frmArchivos2.frx":0130
         Column(3)       =   "frmArchivos2.frx":0224
         Column(4)       =   "frmArchivos2.frx":0348
         Column(5)       =   "frmArchivos2.frx":0480
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmArchivos2.frx":05BC
         FormatStyle(2)  =   "frmArchivos2.frx":06E4
         FormatStyle(3)  =   "frmArchivos2.frx":0794
         FormatStyle(4)  =   "frmArchivos2.frx":0848
         FormatStyle(5)  =   "frmArchivos2.frx":0920
         FormatStyle(6)  =   "frmArchivos2.frx":09D8
         ImageCount      =   0
         PrinterProperties=   "frmArchivos2.frx":0AB8
      End
   End
   Begin VB.Image imgPreview 
      Height          =   3585
      Left            =   6570
      Stretch         =   -1  'True
      Top             =   3615
      Width           =   4575
   End
   Begin VB.Menu archivos 
      Caption         =   "archivos"
      Begin VB.Menu mnuAbrir 
         Caption         =   "Abrir..."
      End
      Begin VB.Menu mnuExportar 
         Caption         =   "Exportar..."
      End
      Begin VB.Menu mnuEnviar 
         Caption         =   "Enviar por mail..."
      End
   End
End
Attribute VB_Name = "frmArchivos2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements ISuscriber
Private extensionesSoportadas As New Collection
Private m_archivosASubir As Collection
Private m_Archivos As Collection
Public Origen As OrigenArchivos
Public ObjetoId As Long
Private archivoActual As archivo
Dim clasea As New classArchivos
Dim id_suscriber As String

Private Property Get ISuscriber_id() As String
    ISuscriber_id = id_suscriber
End Property

Private Function ISuscriber_Notificarse(EVENTO As clsEventoObserver) As Variant

End Function
Private Sub LimpiarLista()
    Set m_archivosASubir = New Collection
    Me.lstArchivos.Clear
    Me.txtComentario.text = vbNullString
    Me.chkArchivosCompra.value = False
    Me.lblRuta.caption = "Ruta: "
End Sub

Private Sub btnSubir_Click()
    If m_archivosASubir.count = 0 Then
        MsgBox "Debe seleccionar como mínimo un archivo para subir.", vbExclamation
    Else
        Dim arch As Variant
        Me.btnSubir.Enabled = False
        For Each arch In m_archivosASubir
            ' If clasea.grabarArchivo(Me.ObjetoId, funciones.GetFileName(arch), CStr(arch), Me.txtComentario.text, CInt(Me.Origen), Me.chkArchivosCompra.value) Then

            If DAOArchivo.grabarArchivo(Me.ObjetoId, funciones.GetFileName(arch), CStr(arch), Me.txtComentario.text, CInt(Me.Origen), Me.chkArchivosCompra.value, Me) Then
                m_archivosASubir.remove CStr(arch)
            End If
        Next arch

        If m_archivosASubir.count > 0 Then
            LlenarListaArchivos
            MsgBox "Algunos archivos no pudieron ser subidos, los mismo permanecen en la lista a la espera.", vbExclamation
            Me.btnSubir.Enabled = True
        Else
            LimpiarLista
            MsgBox "Se subieron todos los archivos.", vbInformation
            Me.btnSubir.Enabled = True
        End If


        LlenarListaArchivosSubidos

    End If
End Sub

Private Sub Form_Load()
    extensionesSoportadas.Add "bmp", "bmp"
    extensionesSoportadas.Add "gif", "gif"
    extensionesSoportadas.Add "jpg", "jpg"
    extensionesSoportadas.Add "jpeg", "jpeg"

    Customize Me
    GridEXHelper.CustomizeGrid Me.gridArchivos
    LimpiarLista
    Me.gridArchivos.ItemCount = 0
    LlenarListaArchivosSubidos
    id_suscriber = funciones.CreateGUID
    'Channel.AgregarSuscriptor Me, TipoSuscripcion.EnvioMail_

End Sub

Private Sub LlenarListaArchivosSubidos()
    Set Me.imgPreview.Picture = Nothing

    Dim filter As String
    filter = DAOArchivo.TABLA_ARCHIVO & "." & DAOArchivo.CAMPO_ID_REFERENCIA & "=" & Me.ObjetoId

    Set m_Archivos = DAOArchivo.FindAll(Me.Origen, filter)
    Me.gridArchivos.ItemCount = 0
    Me.gridArchivos.ItemCount = m_Archivos.count

    Me.Refresh
    GridEXHelper.AutoSizeColumns Me.gridArchivos, True
    Me.gridArchivos.row = -1


End Sub

Private Sub gridArchivos_ColumnHeaderClick(ByVal Column As GridEX20.JSColumn)
    GridEXHelper.ColumnHeaderClick Me.gridArchivos, Column
End Sub

Private Sub gridArchivos_DblClick()
    If Not archivoActual Is Nothing Then
        gridArchivos_SelectionChange
        AbrirArchivo
    End If
End Sub

Private Sub gridArchivos_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 And Not archivoActual Is Nothing Then
        Me.PopupMenu archivos
    End If
End Sub

Private Sub gridArchivos_SelectionChange()
    On Error Resume Next

    Dim idx As Long: idx = Me.gridArchivos.RowIndex(Me.gridArchivos.row)

    If idx > 0 Then
        Dim ext As String
        Dim pos As Long
        Dim tmppath As String

        Set archivoActual = m_Archivos(idx)
        pos = InStrRev(archivoActual.nombre, ".")
        Set Me.imgPreview.Picture = Nothing

        If pos <> 0 Then
            ext = StrConv(Mid(archivoActual.nombre, pos + 1), vbLowerCase)
            If funciones.BuscarEnColeccion(extensionesSoportadas, ext) Then
                tmppath = clasea.exportarArchivo(archivoActual.Id)
                If LenB(tmppath) > 0 Then
                    Set Me.imgPreview.Picture = LoadPicture(tmppath)
                    Kill tmppath
                End If
            End If
        End If

    Else
        Set archivoActual = Nothing
    End If
End Sub

Private Sub gridArchivos_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex > 0 And RowIndex <= m_Archivos.count Then

        Set archivoActual = m_Archivos(RowIndex)
        With archivoActual
            Values(1) = .nombre
            Values(1) = .FileSizeInKB
            If .usuario Is Nothing Then
                Values(2) = vbNullString
            Else
                Values(2) = .usuario.usuario
            End If
            If CDbl(.FechaSubida) > 0 Then Values(3) = .FechaSubida

            Values(4) = "Nombre Archivo: " & .nombre
            If LenB(.Comentario) > 0 Then
                Values(4) = Values(4) & vbNewLine & "Comentario: " & .Comentario
            End If

            Values(5) = .DeCompra
        End With
    End If
End Sub

Private Sub lstArchivos_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    LimpiarLista

    Dim arch As Variant
    For Each arch In data.Files
        m_archivosASubir.Add arch, CStr(arch)
    Next arch

    LlenarListaArchivos
End Sub

Private Sub btnLimpiarLista_Click()
    LimpiarLista
End Sub



Private Sub btnSeleccionar_Click()
    On Error GoTo E

    Dim archArray() As String

    With Me.CommonDialog
        .MaxFileSize = 32767
        .Flags = cdlOFNFileMustExist + cdlOFNLongNames + cdlOFNExplorer + cdlOFNAllowMultiselect + cdlOFNPathMustExist
        .filename = vbNullString
        .CancelError = True
        .ShowOpen

        LimpiarLista

        archArray = Split(.filename, Chr(0))
        Dim i As Long

        If UBound(archArray) > 0 Then
            For i = 0 To UBound(archArray)
                If i > 0 Then
                    m_archivosASubir.Add archArray(0) & "\" & archArray(i), CStr(archArray(0) & "\" & archArray(i))
                End If
            Next i
        Else
            m_archivosASubir.Add .filename, CStr(.filename)
        End If

    End With

    LlenarListaArchivos

    Exit Sub
E:
    If Err.Number <> 32755 Then MsgBox Err.Description, vbCritical, "Error"
End Sub

Private Sub LlenarListaArchivos()
    Dim arch As Variant
    Me.lblRuta.caption = "Ruta: " & funciones.GetFilePath(m_archivosASubir(1))
    Me.lstArchivos.Clear
    For Each arch In m_archivosASubir
        Me.lstArchivos.AddItem funciones.GetFileName(CStr(arch))
    Next arch
End Sub

Private Sub AbrirArchivo()
    If Not archivoActual Is Nothing Then
        If archivoActual.DeCompra And Not Permisos.ArchivosDeCompras Then
            MsgBox "No tiene permisos para ver archivos de compras.", vbExclamation + vbOKOnly
        Else
            clasea.exportarArchivo archivoActual.Id, , True
        End If
    End If
End Sub



Private Sub GuardarArchivo()
    On Error GoTo err1
    If Not archivoActual Is Nothing Then
        If archivoActual.DeCompra And Not Permisos.ArchivosDeCompras Then
            MsgBox "No tiene permisos para ver archivos de compras.", vbExclamation + vbOKOnly
        Else
            Dim ruta As String

            frmPrincipal.CD.filename = archivoActual.nombre
            frmPrincipal.CD.ShowSave
            ruta = frmPrincipal.CD.filename

            If LenB(ruta) > 0 Then
                ruta = clasea.exportarArchivo(archivoActual.Id, ruta, False)
            End If
        End If
    End If
    Exit Sub

err1:
End Sub


Private Sub mnuAbrir_Click()
    gridArchivos_DblClick
End Sub

Private Sub mnuEnviar_Click()


    gridArchivos_SelectionChange
    If Not archivoActual Is Nothing Then
        If archivoActual.DeCompra And Not Permisos.ArchivosDeCompras Then
            MsgBox "No tiene permisos para ver archivos de compras.", vbExclamation + vbOKOnly
        Else

            Dim mail As String
            mail = InputBox("Ingrese dirección de email", "Envío de documentación")

            If LenB(mail) < 5 Then
                Exit Sub
            End If

            ERPHelper.SendMail "Envio de documentacion", "Envio de documentacion", mail, clasea.exportarArchivo(archivoActual.Id, , False)

            MsgBox "Se enviará por mail el documento seleccionado!", vbInformation
        End If
    End If

End Sub

Private Sub mnuExportar_Click()
    GuardarArchivo
End Sub
