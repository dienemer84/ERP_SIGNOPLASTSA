VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmListarStock 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Piezas"
   ClientHeight    =   9345
   ClientLeft      =   60
   ClientTop       =   240
   ClientWidth     =   18105
   Icon            =   "frmListarStock.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9345
   ScaleWidth      =   18105
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1335
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   12015
      _Version        =   786432
      _ExtentX        =   21193
      _ExtentY        =   2355
      _StockProps     =   79
      Caption         =   "Parámetros de búsqueda"
      UseVisualStyle  =   -1  'True
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1425
         TabIndex        =   6
         Top             =   750
         Width           =   2955
      End
      Begin XtremeSuiteControls.PushButton CMDsINCliente 
         Height          =   255
         Left            =   5745
         TabIndex        =   3
         Top             =   360
         Width           =   495
         _Version        =   786432
         _ExtentX        =   873
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "X"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cboCliente 
         Height          =   315
         Left            =   1440
         TabIndex        =   4
         Top             =   360
         Width           =   4215
         _Version        =   786432
         _ExtentX        =   7435
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Appearance      =   6
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.PushButton Command1 
         Default         =   -1  'True
         Height          =   390
         Left            =   4470
         TabIndex        =   5
         Top             =   720
         Width           =   1185
         _Version        =   786432
         _ExtentX        =   2090
         _ExtentY        =   688
         _StockProps     =   79
         Caption         =   "Buscar"
         Appearance      =   6
      End
      Begin VB.Label P 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
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
         Height          =   195
         Left            =   735
         TabIndex        =   8
         Top             =   405
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre pieza"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   7
         Top             =   765
         Width           =   1170
      End
   End
   Begin GridEX20.GridEX grid 
      Height          =   7635
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   12045
      _ExtentX        =   21246
      _ExtentY        =   13467
      Version         =   "2.0"
      DefaultGroupMode=   1
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      OLEDropMode     =   1
      ColumnAutoResize=   -1  'True
      MethodHoldFields=   -1  'True
      Options         =   -1
      RecordsetType   =   1
      ForeColorInfoText=   16777215
      BackColorInfoText=   8421504
      GroupByBoxInfoText=   "Arrastre una columna aqui para agrupar por dicha columna."
      AllowEdit       =   0   'False
      BackColorGBBox  =   8421504
      BackColorHeader =   16761024
      ImageCount      =   1
      ImagePicture1   =   "frmListarStock.frx":000C
      RowHeaders      =   -1  'True
      DataMode        =   99
      HeaderFontName  =   "Tahoma"
      FontName        =   "Tahoma"
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   8
      Column(1)       =   "frmListarStock.frx":0326
      Column(2)       =   "frmListarStock.frx":045A
      Column(3)       =   "frmListarStock.frx":0526
      Column(4)       =   "frmListarStock.frx":060A
      Column(5)       =   "frmListarStock.frx":06DA
      Column(6)       =   "frmListarStock.frx":07D2
      Column(7)       =   "frmListarStock.frx":089E
      Column(8)       =   "frmListarStock.frx":09D6
      FormatStylesCount=   11
      FormatStyle(1)  =   "frmListarStock.frx":0AAE
      FormatStyle(2)  =   "frmListarStock.frx":0BD6
      FormatStyle(3)  =   "frmListarStock.frx":0C86
      FormatStyle(4)  =   "frmListarStock.frx":0D3A
      FormatStyle(5)  =   "frmListarStock.frx":0E12
      FormatStyle(6)  =   "frmListarStock.frx":0ECA
      FormatStyle(7)  =   "frmListarStock.frx":0FAA
      FormatStyle(8)  =   "frmListarStock.frx":107E
      FormatStyle(9)  =   "frmListarStock.frx":1142
      FormatStyle(10) =   "frmListarStock.frx":121A
      FormatStyle(11) =   "frmListarStock.frx":12F2
      ImageCount      =   1
      ImagePicture(1) =   "frmListarStock.frx":13CA
      PrinterProperties=   "frmListarStock.frx":16E4
   End
   Begin VB.Label marcado 
      Caption         =   "Label1"
      Height          =   255
      Left            =   13170
      TabIndex        =   1
      Top             =   285
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblVisible 
      Caption         =   "Label1"
      Height          =   375
      Left            =   13155
      TabIndex        =   0
      Top             =   765
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Menu m1 
      Caption         =   "Menu1"
      Visible         =   0   'False
      Begin VB.Menu VerDesarrollo 
         Caption         =   "Ver Desarrollo..."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDesarrolloNuevo 
         Caption         =   "Ver Desarrollo..."
      End
      Begin VB.Menu verconjunto 
         Caption         =   "Ver Conjunto..."
         Visible         =   0   'False
      End
      Begin VB.Menu archivos 
         Caption         =   "Archivos Asociados..."
      End
      Begin VB.Menu scanear 
         Caption         =   "Adquirir..."
      End
      Begin VB.Menu verIncidencias 
         Caption         =   "Ver Incidencias..."
      End
      Begin VB.Menu adan1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNuevaRevision 
         Caption         =   "Nueva revisión"
      End
      Begin VB.Menu stockModif 
         Caption         =   "Modificar Stock..."
      End
      Begin VB.Menu MovStock 
         Caption         =   "Ver Movimientos..."
      End
      Begin VB.Menu historic 
         Caption         =   "Ver historial..."
      End
      Begin VB.Menu sEliminar 
         Caption         =   "Eliminar..."
      End
   End
End
Attribute VB_Name = "frmListarStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim base As classStock
Private m_piezas As New Collection
Private pieza_actual As Pieza
Private CantArchivos As Dictionary

Private Sub archivos_Click()
    If Not pieza_actual Is Nothing Then
        Dim frmArchi As New frmArchivos2
        frmArchi.Origen = OrigenArchivos.OA_Piezas
        frmArchi.ObjetoId = pieza_actual.Id
        frmArchi.caption = "Pieza " & pieza_actual.nombre
        frmArchi.Show
    End If
End Sub
Private Sub CMDsINCliente_Click()
    Me.cboCliente.ListIndex = -1
End Sub
Private Sub Command1_Click()
    llenarLista
End Sub
Private Sub llenarLista()

    Set pieza_actual = Nothing

    Dim filtro As String
    filtro = "1 = 1"

    If LenB(Me.Text1.text) > 0 Then
        filtro = filtro & " AND ({pieza}.{nombre} LIKE '%{valor_nombre}%'"

        If IsNumeric(Me.Text1.text) Then
            If Val(Me.Text1.text) <> 0 Then
                filtro = filtro & " OR {pieza}.id = " & Val(Me.Text1.text)
            End If
        End If

        filtro = filtro & ")"

        filtro = Replace(filtro, "{valor_nombre}", Me.Text1.text)
    End If
    If Me.cboCliente.ListIndex > -1 Then
        If Me.cboCliente.ItemData(Me.cboCliente.ListIndex) <> -1 Then
            filtro = filtro & " AND {pieza}.{cliente_id} = " & Me.cboCliente.ItemData(Me.cboCliente.ListIndex)
        End If
    End If

    filtro = Replace(filtro, "{cliente_id}", DAOPieza.CAMPO_ID_CLIENTE)
    filtro = Replace(filtro, "{pieza}", DAOPieza.TABLA_PIEZA)
    filtro = Replace(filtro, "{nombre}", DAOPieza.CAMPO_NOMBRE)
    Set CantArchivos = DAOArchivo.GetCantidadArchivosPorReferencia(OA_Piezas)
    Set m_piezas = DAOPieza.FindAll(0, filtro)
    Me.grid.ItemCount = 0
    Me.grid.ItemCount = m_piezas.count
    Me.caption = "Piezas [ Cantidad: " & m_piezas.count & " ]"
    GridEXHelper.AutoSizeColumns Me.grid, True
    Me.grid.Refresh
    grid_SelectionChange
    GridEXHelper.AutoSizeColumns Me.grid, True
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    FormHelper.Customize Me
    GridEXHelper.CustomizeGrid Me.grid, True
    Me.lblVisible = 0
    Set base = New classStock
    marcado = -1
    DAOCliente.llenarComboXtremeSuite Me.cboCliente, True, False, True
    Me.cboCliente.ListIndex = -1
    Me.grid.ItemCount = 0

End Sub




Private Sub Form_Resize()
    On Error Resume Next
    Me.grid.Width = Me.ScaleWidth - 250
    Me.grid.Height = Me.ScaleHeight - Me.grid.Top
    Me.GroupBox1.Width = Me.ScaleWidth - 250
End Sub

Private Sub grid_ColumnHeaderClick(ByVal Column As GridEX20.JSColumn)
    ColumnHeaderClick Me.grid, Column
End Sub
Private Sub grid_DblClick()
    If Not pieza_actual Is Nothing Then
        VerDesarrollo_Click
    End If
End Sub

Private Sub grid_FetchIcon(ByVal RowIndex As Long, ByVal ColIndex As Integer, ByVal RowBookmark As Variant, ByVal IconIndex As GridEX20.JSRetInteger)
    On Error Resume Next
    Set pieza_actual = m_piezas(grid.RowIndex(RowPosition))

    If ColIndex = 7 And CantArchivos.item(pieza_actual.Id) > 0 Then
        IconIndex = 1
    End If

End Sub

Private Sub grid_GroupByBoxHeaderClick(ByVal Group As GridEX20.JSGroup)
    GridEXHelper.GroupByBoxHeaderClick Group
End Sub

Private Sub grid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 67 And Shift = 2 Then    'CTRL + C
        GridEXHelper.Grid2Clipboard Me.grid
        DoEvents
        MsgBox "La lista ha sido copiada al portapapeles.", vbInformation + vbOKOnly
    End If
End Sub

Private Sub grid_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = 2 And Not pieza_actual Is Nothing Then
        grid_SelectionChange
        Me.verconjunto.Enabled = pieza_actual.EsConjunto
        Me.VerDesarrollo.Enabled = Not pieza_actual.EsConjunto
        If pieza_actual.Activa Then
            Me.sEliminar.caption = "Desactivar..."
        ElseIf es = 1 Then
            Me.sEliminar.caption = "Activar..."
        End If

        Me.archivos.Enabled = Permisos.SistemaArchivosVer

        frmListarStock.PopupMenu m1
    End If
End Sub

Private Sub grid_OLEDragDrop(data As GridEX20.JSDataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error GoTo err4
    Dim A As String
    'a = Data.GetData(1)
    Dim F As Boolean

    F = data.GetFormat(1)
    If Not F Then
        A = data.Files(1)



        If Not pieza_actual Is Nothing Then
            frmArchivos.Origen = 1
            frmArchivos.ruta = A
            frmArchivos.lblIdPieza = pieza_actual.Id
            frmArchivos.caption = "[ " & pieza_actual.nombre & " ]"
            frmArchivos.Show
        End If

    End If


err4:

End Sub
Private Sub grid_RowFormat(RowBuffer As GridEX20.JSRowData)
    If RowBuffer.value(5) > 0 Then RowBuffer.CellStyle(5) = "TieneIncidenciasArchivos"

    If RowBuffer.RowIndex > 0 Then
        Set pieza_actual = m_piezas.item(RowBuffer.RowIndex)
        If Not pieza_actual.Activa Then RowBuffer.RowStyle = "desactivado"
    End If


    If tpieza_actualcomplejidad = ComplejidadAlta Then
        RowBuffer.CellStyle(8) = "comp_alta"
    ElseIf pieza_actual.Complejidad = ComplejidadMedia Then
        RowBuffer.CellStyle(8) = "comp_media"
    ElseIf pieza_actual.Complejidad = ComplejidadBaja Then
        RowBuffer.CellStyle(8) = "comp_baja"


    End If

End Sub

Private Sub grid_SelectionChange()
    On Error Resume Next
    Dim RowPosition As Long
    RowPosition = grid.row
    If grid.RowIndex(RowPosition) > 0 Then
        Set pieza_actual = m_piezas(grid.RowIndex(RowPosition))
    Else
        Set pieza_actual = Nothing
    End If
End Sub

Private Sub grid_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
'On Error Resume Next
    If m_piezas.count > 0 Then
        Set pieza_actual = m_piezas.item(RowIndex)
        With pieza_actual
            Values(1) = .Id
            Values(2) = .nombre
            Values(3) = .Revision

            If .cliente Is Nothing Then
                Values(4) = vbNullString
            Else
                Values(4) = .cliente.razon
            End If
            Values(5) = .CantidadStock
            Values(6) = IIf(.EsConjunto, "Conjunto", "Unidad")
            Values(7) = " (" & Val(CantArchivos(pieza_actual.Id)) & ")"
            Values(8) = enums.EnumTiposComplejidad(.Complejidad)
        End With
    End If
End Sub

Private Sub historic_Click()
    If Not pieza_actual Is Nothing Then
        frmDesarrolloStockHistorial.idPieza = pieza_actual.Id
        frmDesarrolloStockHistorial.Show
    End If
End Sub

Private Sub mnuDesarrolloNuevo_Click()
    Dim F As New frmDesarrollo
    Load F
    F.CargarPieza pieza_actual.Id
    F.Show
End Sub

Private Sub mnuNuevaRevision_Click()
    On Error GoTo E

    If IsSomething(pieza_actual) Then
        If MsgBox("¿Desea realizar una nueva revisión de la pieza actual?", vbQuestion + vbYesNo) Then
            Dim nuevaRevision As String
            Dim piezaIdActual As Long
            Dim revisionActual As String
            revisionActual = pieza_actual.Revision

            If IsNumeric(revisionActual) Then
                nuevaRevision = Val(pieza_actual.Revision) + 1
            End If

            nuevaRevision = InputBox("Ingrese la revisión de la nueva pieza." & vbNewLine & "Revisión actual: " & revisionActual, , nuevaRevision)
            Dim result As Boolean
            If LenB(nuevaRevision) > 0 Then
                conectar.BeginTransaction

                pieza_actual.Revision = nuevaRevision
                result = DAOPieza.Save(pieza_actual)
                If Not result Then GoTo E

                If Not Versionar(DAOPieza.FindById(pieza_actual.Id, FL_4, True, True), pieza_actual.Id, revisionActual) Then GoTo E

                conectar.CommitTransaction

                MsgBox "Revisión " & nuevaRevision & " creada.", vbInformation
                llenarLista
            End If
        End If
    End If

    Exit Sub
E:
    conectar.RollBackTransaction
    llenarLista
End Sub


Private Function Versionar(Pieza As Pieza, idOriginal As Long, Revision As String) As Boolean
    On Error GoTo E

    Pieza.IdPiezaUltimaRevision = idOriginal
    Pieza.Revision = Revision
    Pieza.Id = 0    'fuerzo insert
    Pieza.Activa = False

    If DAOPieza.Save(Pieza, True) Then
        '    Set pieza.desarrollosManoObra = DAODesarrolloManoObra.FindAllByPiezaId(idOriginal)
        '    Set pieza.DesarrollosMaterial = DAODesarrolloMaterial.FindAllByPiezaId(idOriginal)

        Dim dmdo As DesarrolloManoObra
        Dim dm As DesarrolloMaterial
        Dim P As Pieza

        For Each dmdo In Pieza.desarrollosManoObra
            dmdo.Id = 0
            Set dmdo.Pieza = Pieza
            If Not DAODesarrolloManoObra.Save(dmdo, True) Then
                GoTo E
            End If
        Next dmdo

        For Each dm In Pieza.DesarrollosMaterial
            dm.Id = 0
            Set dm.Pieza = Pieza
            If Not DAODesarrolloMaterial.Save(dm, True) Then
                GoTo E
            End If
        Next dm

        For Each P In Pieza.PiezasHijas
            If Not Versionar(P, P.Id, P.Revision) Then GoTo E
            conectar.execute "INSERT INTO stockConjuntos_rev (idPiezaPadre, idPiezaHija, cantidad) VALUES (" & Pieza.Id & ", " & P.Id & ", " & P.Cantidad & ")"
        Next P

        Versionar = True
    Else
        Versionar = False
    End If

    Exit Function
E:
    Versionar = False
End Function




Private Sub MovStock_Click()
    If Not pieza_actual Is Nothing Then
        frmMovimientosStock.Frame1.caption = "[ " & pieza_actual.nombre & " ]"
        frmMovimientosStock.lblCliente = IIf(pieza_actual.cliente Is Nothing, vbNullString, pieza_actual.cliente.razon)
        frmMovimientosStock.lblid = pieza_actual.Id
        frmMovimientosStock.Show
    End If
End Sub
Private Sub scanear_Click()
    If Not pieza_actual Is Nothing Then
        Dim archivos As New classArchivos
        archivos.escanearDocumento OrigenArchivos.OA_Piezas, pieza_actual.Id
    End If

End Sub
Private Sub sEliminar_Click()
    If Not pieza_actual Is Nothing Then
        g = MsgBox("¿Seguro que desea cambiar el estado la pieza seleccionada?", vbYesNo, "Confirmacion")
        If g = 6 Then
            acc = base.cambiar_estado(pieza_actual.Id, CInt(pieza_actual.Activa) * -1)
        End If

        'base.llenar_lista_stock Me.lstStock, Me.cboCliente.ItemData(Me.cboCliente.ListIndex), Trim(Text1), , Cantidad
        llenarLista
    End If
End Sub
Private Sub stockModif_Click()
    If Permisos.DesaManejoStock And Not pieza_actual Is Nothing Then
        'modificar stock
        frmModificarStock.idPieza = pieza_actual.Id
        frmModificarStock.Show 1
    Else
        sinAcceso
    End If
End Sub
Private Sub Text1_GotFocus()
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1)
End Sub

Private Sub verconjunto_Click()
    If Not pieza_actual Is Nothing Then

        Dim F As New frmDesarrollo
        Load F
        F.CargarPieza pieza_actual.Id
        F.Show


    End If
End Sub
Private Sub VerDesarrollo_Click()
    If Not pieza_actual Is Nothing Then
        Dim F As New frmDesarrollo
        Load F
        F.CargarPieza pieza_actual.Id
        F.Show
    End If

End Sub
Private Sub verIncidencias_Click()
    If Not pieza_actual Is Nothing Then
        frmVerIncidencias.referencia = pieza_actual.Id
        frmVerIncidencias.Origen = OrigenIncidencias.OI_Piezas
        frmVerIncidencias.Show
    End If
End Sub

