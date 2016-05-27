VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVerDesarrollo 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ver Desarrollo..."
   ClientHeight    =   7200
   ClientLeft      =   390
   ClientTop       =   480
   ClientWidth     =   11085
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   11085
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command6 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   6840
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Costos"
      Height          =   255
      Left            =   3600
      TabIndex        =   20
      Top             =   6840
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cambiar"
      Height          =   255
      Left            =   2760
      TabIndex        =   19
      Top             =   6840
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Copiar"
      Height          =   255
      Left            =   1920
      TabIndex        =   18
      Top             =   6840
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Imprimir"
      Height          =   255
      Left            =   1080
      TabIndex        =   17
      Top             =   6840
      Width           =   735
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4680
      Top             =   6840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSComctlLib.ListView lstMateriales 
      Height          =   2535
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   4471
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   12
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Código"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Descripción"
         Object.Width           =   8643
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Pieza"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "X Pieza"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Y Pieza"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Scrap"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "Kg "
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Text            =   "M2/Ml"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Scrap"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "id_mater"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   11
         Text            =   "Costo"
         Object.Width           =   1587
      EndProperty
   End
   Begin MSComctlLib.ListView lstMdo 
      Height          =   3855
      Left            =   0
      TabIndex        =   2
      Top             =   2880
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   6800
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   11
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Código"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Cant OP"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Tiempo"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Sector"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "CPP"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Tarea"
         Object.Width           =   4057
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "idcpp"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Descripcion"
         Object.Width           =   4057
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "T.Total"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Costo"
         Object.Width           =   1587
      EndProperty
   End
   Begin VB.Label lblmdo 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   6120
      TabIndex        =   16
      Top             =   6840
      Width           =   735
   End
   Begin VB.Label lblcambio 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   7680
      TabIndex        =   15
      Top             =   6840
      Width           =   735
   End
   Begin VB.Label lblfijos 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   9000
      TabIndex        =   14
      Top             =   6840
      Width           =   615
   End
   Begin VB.Label lblTotalKg 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   6840
      TabIndex        =   13
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   8760
      TabIndex        =   12
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label lblTotalCosto 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   10440
      TabIndex        =   11
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label lblCostoMDO 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   10440
      TabIndex        =   10
      Top             =   6840
      Width           =   615
   End
   Begin VB.Label Label31 
      BackColor       =   &H00C0C0C0&
      Caption         =   "M.D.O"
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
      Left            =   5520
      TabIndex        =   9
      Top             =   6840
      Width           =   615
   End
   Begin VB.Label Label30 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cambio"
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
      Left            =   6960
      TabIndex        =   8
      Top             =   6840
      Width           =   735
   End
   Begin VB.Label Label29 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Fijos"
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
      Left            =   8520
      TabIndex        =   7
      Top             =   6840
      Width           =   495
   End
   Begin VB.Label Label17 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Total M2/Ml"
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
      Left            =   7680
      TabIndex        =   6
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label16 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Total KG"
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
      Left            =   6000
      TabIndex        =   5
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Costo $"
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
      Left            =   9720
      TabIndex        =   4
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Costo $"
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
      Left            =   9720
      TabIndex        =   3
      Top             =   6840
      Width           =   735
   End
   Begin VB.Label lblidPieza 
      Caption         =   "Label1"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3960
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "frmVerDesarrollo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m As New classListViewPrinter
Dim claseV As New classVentas
Dim base As classStock
Dim basene As New classNuevoElemento
Public Sub imprimir_lista(lstView As ListView)
    Dim Page As Integer
    Dim sngTotalPage As Single
    m.NumOfRowsPerPage = 10
    sngTotalPage = lstView.ListItems.count / m.NumOfRowsPerPage
    If sngTotalPage - Int(sngTotalPage) <> 0 Then sngTotalPage = Int(sngTotalPage) + 1
    frmVerDesarrollo.ScaleMode = vbPixels    'this must be done, the container [form1 in this case] must be in vbpixels scalemode
    Printer.ScaleMode = vbTwips
    Printer.PaperSize = vbPRPSA4
    Printer.Orientation = vbPRORLandscape
    Printer.Font = lstView.Font.Name
    Printer.FontSize = lstView.Font.Size
    While Not m.LastRowPrinted
        Page = Page + 1
        m.SetRows
        Printer.CurrentX = 700
        Printer.CurrentY = 900: Printer.FontSize = 18: Printer.FontName = "Times New Roman"
        Printer.Print "Desarrollo"
        Printer.FontSize = 8    ': Printer.FontName = "MS SANS SERIF"
        m.PrintHead Printer
        m.PrintBody Printer
        Printer.CurrentY = 4400
        Printer.CurrentX = 700
        Printer.Print "Fecha: " + str(Date)
        Printer.CurrentX = 700
        Printer.Print "Hora: " + str(time)
        Printer.CurrentX = 700
        Printer.Print "Página: " + str(Page) + " de " + str(sngTotalPage)
        Printer.NewPage
    Wend
    Printer.EndDoc
    m.LastRowPrinted = False
    frmVerDesarrollo.ScaleMode = vbTwips
End Sub
Private Sub Imprimir()
    Dim i As Integer, AnchoCol As Single, Espacio As Integer, X As Integer
    idP = CLng(Me.lblIdPieza)
    Set rs = conectar.RSFactory("select s.detalle,c.razon from sp.stock s inner join sp.clientes c on c.id=s.id_cliente and s.id=" & idP)
    Pie = rs!detalle
    cli = rs!Razon
    Printer.Orientation = 2
    AnchoCol = 0
    For i = 1 To lstMateriales.ColumnHeaders.count
        AnchoCol = AnchoCol + Me.lstMateriales.ColumnHeaders(i).Width
    Next

    Printer.Font.Size = 9.6
    Espacio = 0
    Printer.Print "Cliente: " & cli & " Elemento: " & Pie
    Printer.Print
    Printer.Print "MATERIALES"
    Printer.Print
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)
    With Me.lstMateriales
        For i = 1 To .ColumnHeaders.count
            Espacio = Espacio + CInt(.ColumnHeaders(i).Width * Printer.ScaleWidth / AnchoCol)
            If lstMateriales.ColumnHeaders(i).Width > 1 Then
                Printer.Print lstMateriales.ColumnHeaders(i).text;
            End If
            Printer.CurrentX = Espacio
        Next
        Printer.Print
        Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)
        Printer.Print
        For i = 1 To .ListItems.count
            Espacio = 0
            Set lItem = .ListItems(i)
            Printer.Print lItem.text;
            For X = 1 To .ColumnHeaders.count - 1
                Espacio = Espacio + CInt(.ColumnHeaders(X).Width * Printer.ScaleWidth / AnchoCol)
                Printer.CurrentX = Espacio
                If lstMateriales.ColumnHeaders(X + 1).Width > 1 Then
                    Printer.Print lItem.SubItems(X);
                End If
            Next
            Printer.Print
        Next
    End With
    Printer.Print
    AnchoCol = 0
    'Recorremos desde la primer columna hasta la última para almacenar el ancho total
    For i = 1 To Me.lstMdo.ColumnHeaders.count
        AnchoCol = AnchoCol + lstMdo.ColumnHeaders(i).Width
    Next
    Espacio = 0
    Printer.Print "Total Costo Materiales : $" & Me.lblTotalCosto
    Printer.Print
    Printer.Print "MANO DE OBRA"
    Printer.Print
    'Imprime una línea
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)
    With Me.lstMdo
        'Acá se imprimen los encabezados del ListView
        For i = 1 To .ColumnHeaders.count
            Espacio = Espacio + CInt(.ColumnHeaders(i).Width * Printer.ScaleWidth / AnchoCol)
            If lstMdo.ColumnHeaders(i).Width > 1 Then
                Printer.Print lstMdo.ColumnHeaders(i).text;
            End If
            Printer.CurrentX = Espacio
        Next
        Printer.Print
        'Imprime una línea
        Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)
        Printer.Print
        'Este bucle recorre los items y subitems del ListView  y los imprime
        For i = 1 To .ListItems.count
            Espacio = 0
            Set lItem = .ListItems(i)
            Printer.Print lItem.text;
            'Recorremos las columnas
            For X = 1 To .ColumnHeaders.count - 1
                Espacio = Espacio + CInt(.ColumnHeaders(X).Width * Printer.ScaleWidth / AnchoCol)
                Printer.CurrentX = Espacio
                If lstMdo.ColumnHeaders(X + 1).Width > 1 Then
                    Printer.Print lItem.SubItems(X);
                End If
            Next
            'Otro espacio en blanco
            Printer.Print
        Next
    End With
    Printer.Print
    Printer.Print "Total Costo Mano de obra : $" & Me.lblCostoMDO
    Printer.Print
    ''Imprime la línea de final de impresión
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)
    Printer.Print
    'Texto del pie
    Printer.Print Format(Date, "dd-mm-yyyy")
    'Comenzamos la impresión
    Printer.EndDoc
End Sub
Private Sub Command1_Click()
    On Error GoTo err411:
    Me.CommonDialog1.Copies = 1
    Me.CommonDialog1.ShowPrinter
    For i = 1 To Me.CommonDialog1.Copies
        Imprimir

    Next i
    Exit Sub
err411:
End Sub
Private Sub Command2_Click()
'a = ingresar1.mostrar("Referencia nueva pieza", "Copiar desarrollo")
    Dim pieza As String

    pieza = base.detalle_pieza(CLng(Me.lblIdPieza))
    a = funciones.ingreso(pieza)

    If Not IsEmpty(a) Then
        'si devuelve un string, procesarlo.
        If base.buscar_pieza(Trim(a)) = -1 Then
            If MsgBox("¿Desea proceder con la copia?", vbYesNo, "Confirmación") Then
                id = Me.lblIdPieza
                If base.CopiarPieza(id, a) Then MsgBox "Copia exitosa!", vbInformation, "Información"
            End If
        Else
            MsgBox "El detalle ya existe en la base de datos!", vbCritical, "Error"
        End If
    End If
End Sub
Private Sub Command3_Click()
    Set frm2 = New frmNuevoElemento
    frm2.lblidStock = Me.lblIdPieza
    frm2.lblidStock = Me.lblIdPieza
    frm2.txtNombreElemento = base.detalle_pieza(CInt(Me.lblIdPieza))

    'frmNuevoElemento.lblidStock = Me.lblIdPieza
    'frmNuevoElemento.lblidStock = Me.lblIdPieza
    'frmNuevoElemento.txtNombreElemento = base.detalle_pieza(CInt(Me.lblIdPieza))
    base.ejecutar "select detalle,id_cliente from stock where id=" & CLng(Me.lblIdPieza)
    frm2.txtIdCliente = base.idCliente
    basene.llenarListaMDO CLng(Me.lblIdPieza), frm2.ListView2
    basene.llenarLstmateriales CLng(Me.lblIdPieza), frm2.ListView1


    frm2.Caption = "Modificar desarrollo..."
    frm2.Command5.Visible = False
    frm2.btnModificar.Visible = True
    frm2.Show

    'frmNuevoElemento.Caption = "Modificar desarrollo..."
    'frmNuevoElemento.Command5.Visible = False
    'frmNuevoElemento.btnModificar.Visible = True
    'frmNuevoElemento.Show
End Sub
Private Sub Command4_Click()
    frmCostosIncidencia.Cliente = CLng(client)
    frmCostosIncidencia.idP = CLng(Me.lblIdPieza)
    frmCostosIncidencia.Show 1
End Sub



Private Sub Command6_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Set base = New classStock
    base.llenar_listas_desarrollo Me.lstMateriales, Me.lstMdo, CInt(Me.lblIdPieza)
    base.calcular_totales Me
    Me.Refresh
    Me.lstMateriales.Refresh
    Me.lstMdo.Refresh
End Sub

Private Sub lstMateriales_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim li As ListItem
    Set li = Me.lstMateriales.HitTest(X, Y)
    If li Is Nothing Then
        Me.lstMateriales.ToolTipText = ""
    Else
        Me.lstMateriales.ToolTipText = li.Tag
    End If

End Sub

Private Sub lstMdo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim li As ListItem
    Set li = Me.lstMdo.HitTest(X, Y)
    If li Is Nothing Then
        Me.lstMdo.ToolTipText = ""
    Else
        Me.lstMdo.ToolTipText = li.Tag
    End If

End Sub
