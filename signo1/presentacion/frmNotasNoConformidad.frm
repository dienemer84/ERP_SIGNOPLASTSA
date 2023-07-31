VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmNotasNoConformidad 
   Caption         =   "Notas de No Conformidad"
   ClientHeight    =   6375
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11655
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNotasNoConformidad.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6375
   ScaleWidth      =   11655
   Begin VB.TextBox txtOT 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   540
      TabIndex        =   3
      Top             =   150
      Width           =   855
   End
   Begin GridEX20.GridEX gridNotas 
      Height          =   5730
      Left            =   15
      TabIndex        =   0
      Top             =   615
      Width           =   11640
      _ExtentX        =   20532
      _ExtentY        =   10107
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      ColumnAutoResize=   -1  'True
      MethodHoldFields=   -1  'True
      ContScroll      =   -1  'True
      AllowEdit       =   0   'False
      DataMode        =   99
      HeaderFontName  =   "MS Sans Serif"
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   9
      Column(1)       =   "frmNotasNoConformidad.frx":000C
      Column(2)       =   "frmNotasNoConformidad.frx":0114
      Column(3)       =   "frmNotasNoConformidad.frx":01F8
      Column(4)       =   "frmNotasNoConformidad.frx":02E4
      Column(5)       =   "frmNotasNoConformidad.frx":03D0
      Column(6)       =   "frmNotasNoConformidad.frx":04E0
      Column(7)       =   "frmNotasNoConformidad.frx":05F8
      Column(8)       =   "frmNotasNoConformidad.frx":06EC
      Column(9)       =   "frmNotasNoConformidad.frx":07E0
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmNotasNoConformidad.frx":08E8
      FormatStyle(2)  =   "frmNotasNoConformidad.frx":0A10
      FormatStyle(3)  =   "frmNotasNoConformidad.frx":0AC0
      FormatStyle(4)  =   "frmNotasNoConformidad.frx":0B74
      FormatStyle(5)  =   "frmNotasNoConformidad.frx":0C4C
      FormatStyle(6)  =   "frmNotasNoConformidad.frx":0D04
      ImageCount      =   0
      PrinterProperties=   "frmNotasNoConformidad.frx":0DE4
   End
   Begin XtremeSuiteControls.PushButton cmdBuscar 
      Default         =   -1  'True
      Height          =   405
      Left            =   1965
      TabIndex        =   1
      Top             =   90
      Width           =   1560
      _Version        =   786432
      _ExtentX        =   2752
      _ExtentY        =   714
      _StockProps     =   79
      Caption         =   "Buscar"
      BackColor       =   16744576
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin VB.Label lblOt 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "OT"
      Height          =   195
      Left            =   225
      TabIndex        =   2
      Top             =   165
      Width           =   210
   End
   Begin VB.Menu mnuContextual 
      Caption         =   "mnuContextual"
      Visible         =   0   'False
      Begin VB.Menu mnuResolver 
         Caption         =   "Resolver"
      End
   End
End
Attribute VB_Name = "frmNotasNoConformidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private notas As New Collection
Private Nota As NotaNoConformidad

Private Sub cmdBuscar_Click()
    Buscar
End Sub

Private Sub Buscar()
    Dim F As String
    F = "1 = 1"
    If LenB(Me.txtOt.text) > 0 Then
        F = F & " AND PlaneamientoTiemposProcesos.idPedido = " & Me.txtOt.text
    End If
    Me.gridNotas.ItemCount = 0
    Set notas = DAONotaNoConformidad.FindAll(F)
    Me.gridNotas.ItemCount = notas.count
    AutoSizeColumns Me.gridNotas, True
End Sub

Private Sub Form_Load()
    Customize Me
    GridEXHelper.CustomizeGrid Me.gridNotas, True, False
    Me.gridNotas.ItemCount = 0
End Sub

Private Sub Form_Resize()
    Me.gridNotas.Width = Me.ScaleWidth
    Me.gridNotas.Height = Me.ScaleHeight - 650
End Sub

Private Sub gridNotas_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 And Me.gridNotas.ItemCount > 0 Then
        Me.PopupMenu Me.mnuContextual
    End If
End Sub

Private Sub gridNotas_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If notas.count > 0 And rowIndex <= notas.count Then
        Set Nota = notas.item(rowIndex)
        Values(1) = Nota.numero
        Values(2) = Nota.TiempoProceso.idpedido
        Values(3) = Nota.TiempoProceso.DetalleOt.item
        Values(4) = Nota.TiempoProceso.DetalleOt.Pieza.nombre
        Values(5) = Nota.FechaCreacion
        If CDbl(Nota.FechaResolucion) > 0 Then Values(6) = Nota.FechaResolucion
        If IsSomething(Nota.usuarioCreador) Then Values(7) = Nota.usuarioCreador.usuario
        Values(8) = enums.EnumEstadoNNC(Nota.estado)
        Values(9) = Nota.TareaOrigen.Tarea
    End If
End Sub

Private Sub mnuResolver_Click()
    If Me.gridNotas.rowIndex(Me.gridNotas.row) > 0 Then
        Set Nota = notas.item(Me.gridNotas.rowIndex(Me.gridNotas.row))
        Dim F As New frmNotaNoConformidad
        F.nnc = Nota
        F.Show
    End If
End Sub

Private Sub txtOT_Validate(Cancel As Boolean)
    Cancel = Not IsNumeric(Me.txtOt.text) And LenB(Me.txtOt.text) > 0
End Sub
