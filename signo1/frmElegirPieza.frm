VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmElegirPieza 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Elegir Pieza..."
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10095
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   10095
   StartUpPosition =   2  'CenterScreen
   Begin GridEX20.GridEX GridEx1 
      Height          =   3735
      Left            =   120
      TabIndex        =   5
      Top             =   105
      Width           =   9780
      _ExtentX        =   17251
      _ExtentY        =   6588
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      ColumnAutoResize=   -1  'True
      MultiSelect     =   -1  'True
      MethodHoldFields=   -1  'True
      BackColorInfoText=   -2147483639
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      BackColorHeader =   16761024
      RowHeaders      =   -1  'True
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   2
      Column(1)       =   "frmElegirPieza.frx":0000
      Column(2)       =   "frmElegirPieza.frx":00F4
      FmtConditionsCount=   1
      FmtCondition(1) =   "frmElegirPieza.frx":01BC
      FormatStylesCount=   8
      FormatStyle(1)  =   "frmElegirPieza.frx":0280
      FormatStyle(2)  =   "frmElegirPieza.frx":03B8
      FormatStyle(3)  =   "frmElegirPieza.frx":0468
      FormatStyle(4)  =   "frmElegirPieza.frx":051C
      FormatStyle(5)  =   "frmElegirPieza.frx":05F4
      FormatStyle(6)  =   "frmElegirPieza.frx":06AC
      FormatStyle(7)  =   "frmElegirPieza.frx":078C
      FormatStyle(8)  =   "frmElegirPieza.frx":0894
      ImageCount      =   0
      PrinterProperties=   "frmElegirPieza.frx":092C
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   3960
      Width           =   6435
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Filtrar"
      Default         =   -1  'True
      Height          =   375
      Left            =   6780
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3945
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFC0C0&
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   8940
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3945
      Width           =   975
   End
   Begin VB.CommandButton bDePresu 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Agregar"
      Height          =   375
      Left            =   7860
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3945
      Width           =   975
   End
   Begin VB.Label lblidCliente 
      Caption         =   "Label1"
      Height          =   495
      Left            =   5160
      TabIndex        =   1
      Top             =   6360
      Width           =   1095
   End
End
Attribute VB_Name = "frmElegirPieza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim filtro As String
Dim loading As Boolean
Dim piezas As New Collection
Dim rectmp As Pieza
Dim EVENTO As clsEventoObserver
Dim colPiezas As New Collection
Dim vCliente As clsCliente
Dim vorigen As Integer
Public OtIdFilter As Long
Private dtos As New Collection
Private dto As DTOPiezaDetallePedido
Private dtosRetorno As New Collection

Public Property Let Origen(nOrigen As Integer)
    vorigen = nOrigen
End Property
Public Property Let cliente(ncli As clsCliente)
    Set vCliente = ncli
End Property
Private Function CrearEvento() As clsEventoObserver
    Set EVENTO = New clsEventoObserver
    Set EVENTO.Elemento = dtosRetorno
    EVENTO.EVENTO = agregarColeccion_
    EVENTO.Originador = vorigen
    Set CrearEvento = EVENTO
End Function

Private Sub bDePresu_Click()
    Dim A As JSSelectedItem
    Dim P As Pieza
    Dim detaOT As DetalleOrdenTrabajo
    Dim detas As Collection

    '
    '    If OtIdFilter > 0 Then
    '        Set detas = DAODetalleOrdenTrabajo.FindAllByOrdenTrabajo(OtIdFilter)
    '    End If


    For Each A In Me.GridEX1.SelectedItems
        Set dto = dtos.item(A.RowIndex)
        If Not dto.Pieza.Activa Then
            MsgBox "No puede agregar una pieza desactivada.", vbExclamation
            Debug.Print dto.Pieza.nombre
            Exit Sub
        End If
    Next A


    Set dtosRetorno = New Collection
    For Each A In Me.GridEX1.SelectedItems
        Set dto = dtos.item(A.RowIndex)
        dtosRetorno.Add dto
        '        Set P = piezas.Item(a.RowIndex)
        '
        '        If OtIdFilter > 0 Then
        '            For Each detaOT In detas
        '                If detaOT.pieza.id = P.id Then
        '                    P.Precio = detaOT.Precio
        '                    Exit For
        '                End If
        '            Next detaOT
        '        End If
        '
        '        colPiezas.Add P
    Next

    Me.Command1.default = True

    Channel.Notificar CrearEvento, IIf(vorigen = 1, NuevoPresupuesto_, TipoSuscripcion.NuevaOT_)

    Set EVENTO = Nothing
    Set colPiezas = Nothing
End Sub
Private Sub Command1_Click()
    llenarLista
End Sub
Private Sub llenarLista()
    If loading Then
        GridEX1.ItemCount = 0
        Exit Sub
    End If

    filtro = "s.id_cliente=" & vCliente.id
    If Trim(Me.Text1) <> Empty Then
        filtro = filtro & " and s.detalle LIKE '%" & Me.Text1.text & "%'"
    End If
    Me.GridEX1.ItemCount = 0
    Dim esta As Boolean
    Dim resfilter As String
    Dim pa As Pieza
    Set dtos = New Collection
    Dim dto As DTOPiezaDetallePedido
    If OtIdFilter > 0 Then
        'filtro = filtro & " and s.id IN (SELECT DISTINCT idPieza FROM detalles_pedidos WHERE idPedido = " & OtIdFilter & ")"

        Dim colDetalleOt As Collection
        Set colDetalleOt = DAODetalleOrdenTrabajo.FindAllByOrdenTrabajo(OtIdFilter)
        Dim deta As DetalleOrdenTrabajo
        For Each deta In colDetalleOt
            Set dto = New DTOPiezaDetallePedido
            Set dto.Pieza = deta.Pieza


            dto.idDetalleOt = deta.id
            dto.idOt = OtIdFilter
            dto.item = deta.item
            dto.Precio = deta.Precio
            dto.Disponibles = deta.MarcoCantidadDisponibles
            If Trim(Me.Text1) <> Empty Then
                esta = Not (InStr(dto.Pieza.nombre, UCase(Trim(Me.Text1.text))) = 0)
            Else
                esta = True
            End If


            If dto.Disponibles > 0 And esta Then
                dtos.Add dto
            End If

        Next deta

    Else
        Set piezas = DAOPieza.FindAll(FL_0, filtro)
        For Each pa In piezas
            Set dto = New DTOPiezaDetallePedido
            Set dto.Pieza = pa
            dtos.Add dto
        Next pa
    End If

    Me.GridEX1.ItemCount = dtos.count
    Me.bDePresu.default = True
End Sub
Private Sub Command3_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    FormHelper.Customize Me
    loading = True
    GridEXHelper.CustomizeGrid Me.GridEX1
    llenarLista
    loading = False
    
        ''Me.caption = caption & " (" & Name & ")"
        
        
End Sub
Private Sub GridEX1_ColumnHeaderClick(ByVal Column As GridEX20.JSColumn)
    GridEXHelper.ColumnHeaderClick Me.GridEX1, Column
End Sub
Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)
    If RowBuffer.value(2) = "Conjunto" Then RowBuffer.CellStyle(2) = "es_conjunto"
    If RowBuffer.RowIndex > 0 Then
        Set dto = dtos.item(RowBuffer.RowIndex)
        If Not dto.Pieza.Activa Then
            RowBuffer.RowStyle = "desactivado"
        End If
    End If
End Sub
Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Set dto = dtos.item(RowIndex)
    With rectmp
        If OtIdFilter = 0 Then
            Values(1) = dto.Pieza.nombre
        Else
            Values(1) = dto.item & " | " & dto.Pieza.nombre & " | Precio " & dto.Precio & " | Disp " & dto.Disponibles
        End If

        Values(2) = IIf(dto.Pieza.EsConjunto, "Conjunto", "Unidad")
    End With
End Sub

Private Sub Text1_GotFocus()
    foco Me.Text1
End Sub
