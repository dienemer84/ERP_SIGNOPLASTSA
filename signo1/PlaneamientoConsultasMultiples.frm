VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPlaneamientoConsultasMultiples 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Consultas Multiples"
   ClientHeight    =   7230
   ClientLeft      =   1230
   ClientTop       =   1215
   ClientWidth     =   7770
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   7770
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      Cancel          =   -1  'True
      Caption         =   "Command3"
      Height          =   495
      Left            =   6120
      TabIndex        =   7
      Top             =   8880
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1475
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ejecutar"
      Height          =   375
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1475
      Width           =   1215
   End
   Begin VB.ComboBox cboConsultas 
      Height          =   315
      ItemData        =   "PlaneamientoConsultasMultiples.frx":0000
      Left            =   120
      List            =   "PlaneamientoConsultasMultiples.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1470
      Width           =   4935
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ Resultado ]"
      Height          =   5295
      Left            =   0
      TabIndex        =   1
      Top             =   1920
      Width           =   7695
      Begin MSComctlLib.ListView lstResult 
         Height          =   4935
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   8705
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ OT Seleccionadas ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7695
      Begin MSComctlLib.ListView lstOTSelected 
         Height          =   975
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   1720
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
End
Attribute VB_Name = "frmPlaneamientoConsultasMultiples"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim clsP As New classPlaneamiento
Dim rs As Recordset
Dim strsql As String
Public OTElegidas As New Collection


Private Sub Command1_Click()
    Select Case Me.cboConsultas.ListIndex
    Case 0: piezasEnComun
    Case 1: MaterialesEnComun
    Case 3: DetallepiezasEnComun


    End Select
End Sub


Private Sub piezasEnComun()
'armo el encabezado
    Me.lstResult.ColumnHeaders.Clear
    Me.lstResult.ListItems.Clear
    Me.lstResult.ColumnHeaders.Add = "Cantidad"
    Me.lstResult.ColumnHeaders(1).Width = 1100
    Me.lstResult.ColumnHeaders(1).Alignment = lvwColumnLeft
    Me.lstResult.ColumnHeaders.Add = "Pieza"
    Me.lstResult.ColumnHeaders(2).Width = 6000
    Me.lstResult.ColumnHeaders(2).Alignment = lvwColumnLeft

    strsql = "select s.detalle,sum(d.cantidad) as CantTotal from detalles_pedidos d inner join stock s on d.idPieza=s.id where "
    For x = 1 To Me.lstOTSelected.ListItems.count
        If x = Me.lstOTSelected.ListItems.count Then
            strsql = strsql & " d.idpedido=" & Me.lstOTSelected.ListItems(x)
        Else
            strsql = strsql & " d.idpedido=" & Me.lstOTSelected.ListItems(x) & " or "
        End If
    Next x
    strsql = strsql & " group by s.id  order by s.id"
    Set rs = conectar.RSFactory(strsql)

    Dim y As ListItem
    While Not rs.EOF
        Set y = Me.lstResult.ListItems.Add(, , rs!cantTotal)
        y.SubItems(1) = rs!detalle
        rs.MoveNext
    Wend

End Sub

Private Sub DetallepiezasEnComun()
'armo el encabezado
    Me.lstResult.ColumnHeaders.Clear
    Me.lstResult.ListItems.Clear
    Me.lstResult.ColumnHeaders.Add = "Cantidad"
    Me.lstResult.ColumnHeaders(1).Width = 1100
    Me.lstResult.ColumnHeaders(1).Alignment = lvwColumnLeft
    Me.lstResult.ColumnHeaders.Add = "Pieza"
    Me.lstResult.ColumnHeaders(2).Width = 6000
    Me.lstResult.ColumnHeaders(2).Alignment = lvwColumnLeft

    strsql = "select ss.detalle, sum(d.cantidad) * sum(sc.cantidad) as cantTotal from stockConjuntos sc INNER JOIN stock s ON sc.idPiezaPadre = s.id INNER JOIN stock ss ON sc.idPiezaHija = ss.id INNER JOIN detalles_pedidos d ON sc.idPiezaPadre = d.idPieza where "
    
    For x = 1 To Me.lstOTSelected.ListItems.count
        If x = Me.lstOTSelected.ListItems.count Then
            strsql = strsql & " d.idpedido=" & Me.lstOTSelected.ListItems(x)
        Else
            strsql = strsql & " d.idpedido=" & Me.lstOTSelected.ListItems(x) & " or "
        End If
    Next x
    strsql = strsql & " group by ss.detalle order by ss.detalle"
    Set rs = conectar.RSFactory(strsql)

    Dim y As ListItem
    While Not rs.EOF
        Set y = Me.lstResult.ListItems.Add(, , rs!cantTotal)
        y.SubItems(1) = rs!detalle
        rs.MoveNext
    Wend

End Sub


Private Sub MaterialesEnComun()
    DAOOrdenTrabajo.informePiezaMateriales 0, 4, True, OTElegidas
End Sub
Private Sub Command2_Click()
    imprimirResu
End Sub
Private Sub Command3_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    FormHelper.Customize Me
    
    Me.lstOTSelected.ListItems.Clear
    
    Dim tmpOrden As OrdenTrabajo
    For x = 1 To Me.OTElegidas.count
        Set tmpOrden = Me.OTElegidas(x)
        Me.lstOTSelected.ListItems.Add , , tmpOrden.Id    'col(X)
    Next x
    
    Me.lstOTSelected.Refresh
    
End Sub

Private Sub imprimirResu()
    Printer.Orientation = 1
    Printer.FontBold = True
    Printer.Font.Size = 10

    AnchoCol = 0
    For i = 1 To Me.lstResult.ColumnHeaders.count
        AnchoCol = AnchoCol + Me.lstResult.ColumnHeaders(i).Width
    Next
    Espacio = 0
    Printer.Font.Bold = True
    Printer.Print UCase(Me.cboConsultas.text)
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)
    Printer.Print
    Printer.CurrentX = 500
    Printer.Print "O/T Intervinientes:"
    For i = 1 To Me.lstOTSelected.ListItems.count
        Printer.Print Tab(10);
        Printer.Print Me.lstOTSelected.ListItems(i)    '& " - " & clsP.clientePedido(CLng(Me.lstOTSelected.ListItems(i)))
    Next

    Printer.Print
    Printer.Font.Size = 8
    Printer.CurrentX = 500
    'Acá se imprimen los encabezados del ListView

    For i = 1 To Me.lstResult.ColumnHeaders.count

        Espacio = Espacio + CInt(Me.lstResult.ColumnHeaders(i).Width * Printer.ScaleWidth / AnchoCol)
        If Me.lstResult.ColumnHeaders(i).Width > 1 Then

            Printer.Print truncar(Me.lstResult.ColumnHeaders(i).text, 85);
        End If
        Printer.CurrentX = Espacio + 50
    Next
    Printer.Font.Bold = False
    Printer.Print
    'Imprime una línea
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)
    Printer.Print

    'Este bucle recorre los items y subitems del ListView  y los imprime

    For i = 1 To Me.lstResult.ListItems.count
        Printer.CurrentX = 500
        Espacio = 0

        Set lItem = Me.lstResult.ListItems(i)
        Printer.Print lItem.text;
        'Recorremos las columnas
        For x = 1 To Me.lstResult.ColumnHeaders.count - 1
            Espacio = Espacio + CInt(Me.lstResult.ColumnHeaders(x).Width * Printer.ScaleWidth / AnchoCol)
            Printer.CurrentX = Espacio
            If Me.lstResult.ColumnHeaders(x + 1).Width > 1 Then
                Printer.Print truncar(lItem.SubItems(x), 85);
            End If

            Printer.Print
        Next


    Next
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)
    Printer.Print "Fecha emisión " & Format(Date, "dd-mm-yyyy")
    'Comenzamos la impresión
    Printer.EndDoc

End Sub







