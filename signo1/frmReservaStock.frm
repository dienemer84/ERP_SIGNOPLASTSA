VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmReservaStock 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pre-Aprobación"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8475
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   8475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.PushButton Command1 
      Height          =   510
      Left            =   150
      TabIndex        =   4
      Top             =   4080
      Width           =   1065
      _Version        =   786432
      _ExtentX        =   1879
      _ExtentY        =   900
      _StockProps     =   79
      Caption         =   "Aprobar"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.ProgressBar progreso 
      Height          =   300
      Left            =   2760
      TabIndex        =   2
      Top             =   4290
      Visible         =   0   'False
      Width           =   3045
      _Version        =   786432
      _ExtentX        =   5371
      _ExtentY        =   529
      _StockProps     =   93
      Appearance      =   6
   End
   Begin GridEX20.GridEX grilla 
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   6800
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      PreviewColumn   =   7
      PreviewRowLines =   1
      ColumnAutoResize=   -1  'True
      MultiSelect     =   -1  'True
      MethodHoldFields=   -1  'True
      ContScroll      =   -1  'True
      AllowColumnDrag =   0   'False
      GroupByBoxVisible=   0   'False
      BackColorHeader =   16761024
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   7
      Column(1)       =   "frmReservaStock.frx":0000
      Column(2)       =   "frmReservaStock.frx":0134
      Column(3)       =   "frmReservaStock.frx":0248
      Column(4)       =   "frmReservaStock.frx":0360
      Column(5)       =   "frmReservaStock.frx":0470
      Column(6)       =   "frmReservaStock.frx":0540
      Column(7)       =   "frmReservaStock.frx":0658
      FormatStylesCount=   8
      FormatStyle(1)  =   "frmReservaStock.frx":073C
      FormatStyle(2)  =   "frmReservaStock.frx":0874
      FormatStyle(3)  =   "frmReservaStock.frx":0924
      FormatStyle(4)  =   "frmReservaStock.frx":09D8
      FormatStyle(5)  =   "frmReservaStock.frx":0AB0
      FormatStyle(6)  =   "frmReservaStock.frx":0B68
      FormatStyle(7)  =   "frmReservaStock.frx":0C48
      FormatStyle(8)  =   "frmReservaStock.frx":0D20
      ImageCount      =   0
      PrinterProperties=   "frmReservaStock.frx":0DC8
   End
   Begin XtremeSuiteControls.PushButton Command2 
      Height          =   510
      Left            =   1260
      TabIndex        =   5
      Top             =   4080
      Width           =   1065
      _Version        =   786432
      _ExtentX        =   1879
      _ExtentY        =   900
      _StockProps     =   79
      Caption         =   "Cancelar"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton Imprimir 
      Height          =   510
      Left            =   6225
      TabIndex        =   6
      Top             =   4080
      Width           =   1065
      _Version        =   786432
      _ExtentX        =   1879
      _ExtentY        =   900
      _StockProps     =   79
      Caption         =   "Imprimir"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton Command3 
      Height          =   510
      Left            =   7335
      TabIndex        =   7
      Top             =   4080
      Width           =   1065
      _Version        =   786432
      _ExtentX        =   1879
      _ExtentY        =   900
      _StockProps     =   79
      Caption         =   "No Definir"
      UseVisualStyle  =   -1  'True
   End
   Begin VB.Label lblAprobando 
      BackStyle       =   0  'Transparent
      Caption         =   "Aprobando..."
      Height          =   240
      Left            =   2775
      TabIndex        =   3
      Top             =   4080
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Label lblModoEdicion 
      BackColor       =   &H00FFC0C0&
      Caption         =   "MODO EDICION"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   1
      ToolTipText     =   "Presione <ENTER> para terminar de editar el campo"
      Top             =   4320
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Menu mnu 
      Caption         =   "mnuRoot"
      Visible         =   0   'False
      Begin VB.Menu mnuDefinirNoDefinir 
         Caption         =   "No Definir Procesos"
      End
   End
End
Attribute VB_Name = "frmReservaStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rectmp As DetalleOrdenTrabajo
Private deta As Collection
Public Ot As OrdenTrabajo
Private id_suscriber As String
Implements ISuscriber


Private Sub Command1_Click()
    Me.lblAprobando.Visible = True
    Me.progreso.Visible = True
    Dim A As Long
    If Me.grilla.EditMode = jgexEditModeOn Then
        MsgBox "Salga del modo edicion para poder guardar!", vbInformation, "Información"
    Else

        If MsgBox("¿Está seguro de continuar?", vbYesNo, "Confirmación") = vbYes Then
            A = grilla.RowIndex(grilla.row)
            If Not Ot.ValidarProcesos Then
                MsgBox "Por favor, defina todos los procesos para continuar!", vbCritical, "Error"
            Else
                If Not Ot.ValidarReservas Then
                    MsgBox "Por favor, controle las reservas para poder continuar!", vbCritical, "Error"
                Else
                    Me.Enabled = False
                    If DAOOrdenTrabajo.AprobarOT(Ot, Me.progreso) Then
                        MsgBox "Aprobación exitosa!", vbInformation, "Información"

                    End If
                    Me.Enabled = True
                End If

                Unload Me
            End If
        End If
    End If
    Me.progreso.Visible = False
    Me.lblAprobando.Visible = False
End Sub
Private Sub Command3_Click()
    On Error GoTo err1
    Dim j As JSSelectedItem

    For Each j In grilla.SelectedItems
        Set rectmp = Ot.detalles(j.RowIndex)
        rectmp.EstadoProceso = EstProcDetOT_ProcesoNoDefinido
        grilla.RefreshRowIndex j.RowIndex
    Next
    Exit Sub
err1:
End Sub
Private Sub Command2_Click()
    Unload Me
End Sub
Private Sub verModoEdicion()

    If grilla.EditMode = jgexEditModeOff Then
        Me.lblModoEdicion.Visible = False
    Else
        Me.lblModoEdicion.Visible = True
    End If
End Sub
Private Sub Form_Load()
    FormHelper.Customize Me
    GridEXHelper.CustomizeGrid Me.grilla, False, True
    Me.caption = "O/T " & Format(Id, "0000")
    Set Ot.detalles = DAODetalleOrdenTrabajo.FindAllByOrdenTrabajo(Ot.Id)
    Set deta = Ot.detalles
    llenarLista
    id_suscriber = funciones.CreateGUID

    'Me.caption = caption & " (" & Name & ")"

End Sub
Private Sub llenarLista()
    grilla.ItemCount = deta.count
End Sub

Private Sub grilla_Click()
    verModoEdicion
End Sub

Private Sub grilla_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then grilla.EditMode = jgexEditModeOff
    verModoEdicion
End Sub

Private Sub grilla_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If deta.count > 0 Then
        If Button = 2 Then
            If rectmp.EstadoProceso = EstProcDetOT_AunNoDefinido Then
                Me.mnuDefinirNoDefinir.caption = "Definir Proceso"
            ElseIf rectmp.EstadoProceso = EstProcDetOT_ProcesoDefinido Then
                Me.mnuDefinirNoDefinir.caption = "No Definir Proceso"
            ElseIf rectmp.EstadoProceso = EstProcDetOT_ProcesoNoDefinido Then
                Me.mnuDefinirNoDefinir.caption = "Definir Proceso"
            End If
            Me.PopupMenu Me.mnu
        End If
    End If
End Sub
Private Sub grilla_RowFormat(RowBuffer As GridEX20.JSRowData)
    If RowBuffer.value(5) > RowBuffer.value(4) Then
        RowBuffer.CellStyle(5) = "sin_stock"
    Else
        If RowBuffer.value(4) > 0 Then
            RowBuffer.CellStyle(4) = "hay"
        End If
    End If


End Sub
Private Sub grilla_SelectionChange()
    On Error Resume Next
    Set rectmp = deta.item(grilla.RowIndex(grilla.row))
End Sub
Private Sub grilla_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Set rectmp = deta.item(RowIndex)
    With rectmp
        Values(1) = rectmp.item
        Values(2) = rectmp.Nota
        Values(3) = rectmp.CantidadPedida
        Values(4) = rectmp.Pieza.CantidadStock
        Values(5) = rectmp.ReservaStock
        Values(6) = enums.enumEstadoProcesoDetalleOrdenTrabajo(rectmp.EstadoProceso)
        Values(7) = rectmp.Pieza.nombre
    End With
End Sub
Private Sub grilla_UnboundUpdate(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    On Error Resume Next
    Set rectmp = deta.item(RowIndex)
    If rectmp.Pieza.CantidadStock >= CLng(Values(5)) Then
        rectmp.ReservaStock = CLng(Values(5))
    End If
End Sub
Private Sub Imprimir_Click()
    imprimir_lista
End Sub
Private Sub imprimir_lista()
    Exit Sub
    Dim rs As Recordset
    Set rs = conectar.RSFactory("select idCliente from pedidos where id=" & Id)
    If Not rs.EOF And Not rs.BOF Then
        cli = rs!idCliente
    End If
    Set rs = conectar.RSFactory("select razon from clientes where id=" & cli)
    clie = rs!razon
    Printer.FontBold = True
    Printer.Font.Size = 12
    Printer.Font.Bold = True
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)
    Printer.Print
    Printer.Print "LISTA DE ELEMENTOS EN STOCK O/T " & Id
    Printer.Print
    Printer.Print "Cliente: " & cli & " - " & clie
    Printer.Print
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)
    Printer.Print
    Printer.Print "Pieza";
    Printer.Print Tab(50);
    Printer.Print "Cantidad";
    Printer.Print Tab(65);
    Printer.Print "En stock";
    Printer.Print Tab(80);
    Printer.Print "Reserva"
    Set rs = conectar.RSFactory("select dp.item,s.detalle_stock as donde,s.detalle,dp.cantidad as Pedidos,s.cantidad as EnStock, if (dp.cantidad>s.Cantidad,s.cantidad,dp.cantidad) as reserva from detalles_pedidos dp inner join stock s on dp.idPieza=s.id where dp.idPedido=" & Id & " and s.cantidad>0")
    While Not rs.EOF
        Printer.Print rs!item & " - " & rs!detalle;
        Printer.Print Tab(50);
        Printer.Print rs!pedidos;
        Printer.Print Tab(65);
        Printer.Print rs!enStock;
        Printer.Print Tab(80);
        Printer.Print rs!reserva
        Printer.Print Tab(10);
        Printer.Font.Bold = False
        Printer.Print rs!donde
        Printer.Font.Bold = True
        rs.MoveNext
    Wend
    Printer.EndDoc
End Sub

Private Property Get ISuscriber_id() As String
    ISuscriber_id = id_suscriber
End Property

Private Function ISuscriber_Notificarse(EVENTO As clsEventoObserver) As Variant

End Function
