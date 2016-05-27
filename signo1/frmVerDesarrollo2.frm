VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{9AB27FA7-EB2B-426F-82AB-75F983B55258}#2.0#0"; "ingresamos.ocx"
Begin VB.Form frmVerDesarrollo 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ver Desarrollo..."
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   11295
   ShowInTaskbar   =   0   'False
   Begin ingreso.ingresamos ingresar2 
      Height          =   375
      Left            =   4080
      TabIndex        =   25
      Top             =   8400
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   661
   End
   Begin ingreso.ingresamos ingresar1 
      Height          =   495
      Left            =   2880
      TabIndex        =   24
      Top             =   8520
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   873
   End
   Begin VB.CommandButton Command5 
      Cancel          =   -1  'True
      Caption         =   "Command5"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   8640
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ Detalles ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11295
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "[ Acciones ]"
         Height          =   615
         Left            =   120
         TabIndex        =   14
         Top             =   7320
         Width           =   3615
         Begin VB.CommandButton Command4 
            Caption         =   "Costos"
            Height          =   255
            Left            =   2640
            TabIndex        =   22
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Cambiar"
            Height          =   255
            Left            =   1800
            TabIndex        =   21
            Top             =   240
            Width           =   735
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Copiar"
            Height          =   255
            Left            =   960
            TabIndex        =   16
            Top             =   240
            Width           =   735
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Imprimir"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   735
         End
      End
      Begin MSComctlLib.ListView lstMateriales 
         Height          =   2535
         Left            =   120
         TabIndex        =   2
         Top             =   480
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
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Código"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Descripción"
            Object.Width           =   8643
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Pieza"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
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
         Left            =   120
         TabIndex        =   3
         Top             =   3480
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
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Código"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Cant OP"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Tiempo"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Sector"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "CPP"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
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
            Alignment       =   2
            SubItemIndex    =   8
            Text            =   "Descripcion"
            Object.Width           =   4057
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
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
      Begin VB.Label lblCostoMDO 
         BackColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   10560
         TabIndex        =   20
         Top             =   7680
         Width           =   615
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
         Left            =   9840
         TabIndex        =   19
         Top             =   7680
         Width           =   735
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
         Left            =   9840
         TabIndex        =   18
         Top             =   3120
         Width           =   735
      End
      Begin VB.Label lblTotalCosto 
         BackColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   10560
         TabIndex        =   17
         Top             =   3120
         Width           =   615
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   8880
         TabIndex        =   13
         Top             =   3120
         Width           =   735
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
         Left            =   6120
         TabIndex        =   12
         Top             =   3120
         Width           =   855
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
         Left            =   7800
         TabIndex        =   11
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label lblTotalKg 
         BackColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   6960
         TabIndex        =   10
         Top             =   3120
         Width           =   615
      End
      Begin VB.Label lblfijos 
         BackColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   9120
         TabIndex        =   9
         Top             =   7680
         Width           =   615
      End
      Begin VB.Label lblmdo 
         BackColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   6240
         TabIndex        =   8
         Top             =   7680
         Width           =   735
      End
      Begin VB.Label lblcambio 
         BackColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   7800
         TabIndex        =   7
         Top             =   7680
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
         Left            =   8640
         TabIndex        =   6
         Top             =   7680
         Width           =   495
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
         Left            =   7080
         TabIndex        =   5
         Top             =   7680
         Width           =   735
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
         Left            =   5640
         TabIndex        =   4
         Top             =   7680
         Width           =   615
      End
   End
   Begin VB.Label lblidPieza 
      Caption         =   "Label1"
      Height          =   375
      Left            =   120
      TabIndex        =   1
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
Dim m As New ListViewPrinter
Dim claseV As New classVentas
Dim base As classStock
Dim basene As New classNuevoElemento

Public Sub imprimir_lista(lstView As ListView)

Dim Page As Integer
Dim sngTotalPage As Single
m.NumOfRowsPerPage = 10

sngTotalPage = lstView.ListItems.count / m.NumOfRowsPerPage
If sngTotalPage - Int(sngTotalPage) <> 0 Then sngTotalPage = Int(sngTotalPage) + 1

frmVerDesarrollo.ScaleMode = vbPixels 'this must be done, the container [form1 in this case] must be in vbpixels scalemode
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
        Printer.FontSize = 8: Printer.FontName = "MS SANS SERIF"
        
        m.PrintHead Printer
        m.PrintBody Printer
        
        Printer.CurrentY = 4400
        Printer.CurrentX = 700
        Printer.Print "Fecha: " + str(Date)
        Printer.CurrentX = 700
        Printer.Print "Hora: " + str(Time)
        Printer.CurrentX = 700
        Printer.Print "Página: " + str(Page) + " de " + str(sngTotalPage)
        Printer.NewPage
Wend
       Printer.EndDoc
m.LastRowPrinted = False
frmVerDesarrollo.ScaleMode = vbTwips
End Sub

Private Sub Command1_Click()
If MsgBox("¿Seguro de imprimir?", vbYesNo, "Confirmación") = vbYes Then

Dim i As Integer, AnchoCol As Single, Espacio As Integer, X As Integer
idP = CLng(Me.lblidPieza)
Set rs = base.CrearRS("select s.detalle,c.razon from sp.stock s inner join sp.clientes c on c.id=s.id_cliente and s.id=" & idP)
pie = rs!detalle
cli = rs!razon
  AnchoCol = 0
  For i = 1 To lstMateriales.ColumnHeaders.count
     AnchoCol = AnchoCol + Me.lstMateriales.ColumnHeaders(i).Width
  Next
  Espacio = 0
  Printer.Print "Cliente: " & cli & " Elemento: " & pie
  Printer.Print
  Printer.Print "MATERIALES"
  Printer.Print
  Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)
  
  With Me.lstMateriales
  For i = 1 To .ColumnHeaders.count
      Espacio = Espacio + CInt(.ColumnHeaders(i).Width * Printer.ScaleWidth / AnchoCol)
      If lstMateriales.ColumnHeaders(i).Width > 1 Then
        Printer.Print lstMateriales.ColumnHeaders(i).Text;
      End If
      Printer.CurrentX = Espacio
  Next
  Printer.Print
  Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)
  Printer.Print
   For i = 1 To .ListItems.count
       Espacio = 0
       Set litem = .ListItems(i)
       Printer.Print litem.Text;
       For X = 1 To .ColumnHeaders.count - 1
             Espacio = Espacio + CInt(.ColumnHeaders(X).Width * Printer.ScaleWidth / AnchoCol)
             Printer.CurrentX = Espacio
             If lstMateriales.ColumnHeaders(X + 1).Width > 1 Then
                Printer.Print litem.SubItems(X);
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
        Printer.Print lstMdo.ColumnHeaders(i).Text;
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
       
       Set litem = .ListItems(i)
       Printer.Print litem.Text;
       'Recorremos las columnas
       For X = 1 To .ColumnHeaders.count - 1
             Espacio = Espacio + CInt(.ColumnHeaders(X).Width * Printer.ScaleWidth / AnchoCol)
             Printer.CurrentX = Espacio
             If lstMdo.ColumnHeaders(X + 1).Width > 1 Then
                Printer.Print litem.SubItems(X);
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
End If
End Sub





Private Sub Command2_Click()
A = ingresar1.Mostrar("Referencia nueva pieza", "Copiar desarrollo")
If A <> -1 Then
'si devuelve un string, procesarlo.
If base.buscar_pieza(Trim(A)) = -1 Then
    If MsgBox("¿Desea proceder con la copia?", vbYesNo, "Confirmación") Then
        id = Me.lblidPieza
        If base.CopiarPieza(id, A) Then MsgBox "Copia exitosa!", vbInformation, "Información"
    End If
Else
 MsgBox "El detalle ya existe en la base de datos!", vbCritical, "Error"
End If

End If
End Sub

Private Sub Command3_Click()
Set frm2 = frmNuevoElemento
frmNuevoElemento.lblidStock = Me.lblidPieza
frmNuevoElemento.lblidStock = Me.lblidPieza
frmNuevoElemento.txtNombreElemento = base.detalle_pieza(CInt(Me.lblidPieza))
base.ejecutar "select detalle,id_cliente from stock where id=" & CLng(Me.lblidPieza)
frm2.txtIdCliente = base.idCliente
basene.llenarListaMDO CLng(Me.lblidPieza), frmNuevoElemento.ListView2
basene.llenarLstmateriales CLng(Me.lblidPieza), frmNuevoElemento.ListView1
                frmNuevoElemento.btnModificar.Visible = True
                frmNuevoElemento.Command5.Visible = False
                frmNuevoElemento.Caption = "Modificar desarrollo..."
                frmNuevoElemento.Show

'
frmNuevoElemento.Caption = "Modificar desarrollo..."
frmNuevoElemento.Command5.Visible = False
frmNuevoElemento.btnModificar.Visible = True
frmNuevoElemento.Show
End Sub

Private Sub Command4_Click()
frmCostosIncidencia.cliente = CLng(client)
frmCostosIncidencia.idP = CLng(Me.lblidPieza)
frmCostosIncidencia.Show 1
End Sub

Private Sub Command5_Click()
Unload Me
End Sub

Private Sub Form_Activate()
 Set base = New classStock
base.llenar_listas_desarrollo Me.lstMateriales, Me.lstMdo, CInt(Me.lblidPieza)
base.calcular_totales Me

End Sub


