VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~3.OCX"
Begin VB.Form frmPieza 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nueva Pieza"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   12180
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1140
      Left            =   60
      TabIndex        =   1
      Top             =   120
      Width           =   12060
      _Version        =   786432
      _ExtentX        =   21272
      _ExtentY        =   2011
      _StockProps     =   79
      Caption         =   "Datos Principales"
      UseVisualStyle  =   -1  'True
      Begin nucleo.ctrlCboIDPersona ctrlCboIDPersona1 
         Height          =   330
         Left            =   1455
         TabIndex        =   25
         Top             =   630
         Width           =   10560
         _ExtentX        =   18627
         _ExtentY        =   582
      End
      Begin VB.TextBox txtNuevoElemento 
         Height          =   285
         Left            =   1470
         TabIndex        =   2
         Top             =   315
         Width           =   10515
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Centro de costos"
         Height          =   300
         Left            =   120
         TabIndex        =   4
         Top             =   675
         Width           =   1275
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   315
         Left            =   345
         TabIndex        =   3
         Top             =   270
         Width           =   1005
         _Version        =   786432
         _ExtentX        =   1773
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Nombre"
         Alignment       =   1
      End
   End
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   6225
      Left            =   45
      TabIndex        =   0
      Top             =   1320
      Width           =   12075
      _Version        =   786432
      _ExtentX        =   21299
      _ExtentY        =   10980
      _StockProps     =   68
      Appearance      =   10
      Color           =   32
      ItemCount       =   4
      Item(0).Caption =   "Materiales"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "GroupBox2"
      Item(1).Caption =   "Mano de obra"
      Item(1).ControlCount=   0
      Item(2).Caption =   "Historico Fabricación"
      Item(2).ControlCount=   0
      Item(3).Caption =   "Historico Cotización"
      Item(3).ControlCount=   0
      Begin XtremeSuiteControls.GroupBox GroupBox2 
         Height          =   5715
         Left            =   105
         TabIndex        =   5
         Top             =   435
         Width           =   11865
         _Version        =   786432
         _ExtentX        =   20929
         _ExtentY        =   10081
         _StockProps     =   79
         Caption         =   "Datos generales"
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.PushButton PushButton1 
            Height          =   300
            Left            =   135
            TabIndex        =   39
            Top             =   5325
            Width           =   915
            _Version        =   786432
            _ExtentX        =   1614
            _ExtentY        =   529
            _StockProps     =   79
            Caption         =   "Quitar"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton Command1 
            Height          =   255
            Left            =   135
            TabIndex        =   38
            Top             =   1380
            Width           =   2040
            _Version        =   786432
            _ExtentX        =   3598
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Agregar"
            Appearance      =   1
         End
         Begin GridEX20.GridEX GridEx1 
            Height          =   3450
            Left            =   150
            TabIndex        =   22
            Top             =   1740
            Width           =   11625
            _ExtentX        =   20505
            _ExtentY        =   6085
            Version         =   "2.0"
            BoundColumnIndex=   ""
            ReplaceColumnIndex=   ""
            PreviewColumn   =   "descripcion"
            PreviewRowLines =   1
            ColumnAutoResize=   -1  'True
            MethodHoldFields=   -1  'True
            AllowDelete     =   -1  'True
            GroupByBoxVisible=   0   'False
            RowHeaders      =   -1  'True
            DataMode        =   99
            ColumnHeaderHeight=   285
            IntProp1        =   0
            ColumnsCount    =   12
            Column(1)       =   "frmPieza.frx":0000
            Column(2)       =   "frmPieza.frx":0120
            Column(3)       =   "frmPieza.frx":023C
            Column(4)       =   "frmPieza.frx":0330
            Column(5)       =   "frmPieza.frx":0450
            Column(6)       =   "frmPieza.frx":0564
            Column(7)       =   "frmPieza.frx":0678
            Column(8)       =   "frmPieza.frx":0790
            Column(9)       =   "frmPieza.frx":08A4
            Column(10)      =   "frmPieza.frx":09B4
            Column(11)      =   "frmPieza.frx":0AB8
            Column(12)      =   "frmPieza.frx":0BC8
            FormatStylesCount=   6
            FormatStyle(1)  =   "frmPieza.frx":0CE0
            FormatStyle(2)  =   "frmPieza.frx":0E18
            FormatStyle(3)  =   "frmPieza.frx":0EC8
            FormatStyle(4)  =   "frmPieza.frx":0F7C
            FormatStyle(5)  =   "frmPieza.frx":1054
            FormatStyle(6)  =   "frmPieza.frx":110C
            ImageCount      =   0
            PrinterProperties=   "frmPieza.frx":11EC
         End
         Begin XtremeSuiteControls.GroupBox GroupBox3 
            Height          =   960
            Left            =   2280
            TabIndex        =   12
            Top             =   615
            Width           =   1740
            _Version        =   786432
            _ExtentX        =   3069
            _ExtentY        =   1693
            _StockProps     =   79
            Caption         =   "Medida Hoja"
            UseVisualStyle  =   -1  'True
            Begin VB.TextBox txtLargoTerm 
               Height          =   285
               Left            =   735
               TabIndex        =   14
               Top             =   225
               Width           =   840
            End
            Begin VB.TextBox txtAnchoTerm 
               Height          =   285
               Left            =   750
               TabIndex        =   13
               Top             =   600
               Width           =   825
            End
            Begin VB.Label Label4 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Largo"
               Height          =   255
               Left            =   165
               TabIndex        =   16
               Top             =   225
               Width           =   615
            End
            Begin VB.Label Label5 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Ancho"
               Height          =   255
               Left            =   165
               TabIndex        =   15
               Top             =   600
               Width           =   615
            End
         End
         Begin VB.TextBox txtCantidad 
            Height          =   285
            Left            =   840
            TabIndex        =   7
            Top             =   660
            Width           =   1320
         End
         Begin VB.TextBox txtScrap 
            Height          =   285
            Left            =   840
            TabIndex        =   6
            ToolTipText     =   "Valor porcentual"
            Top             =   1020
            Width           =   1335
         End
         Begin XtremeSuiteControls.PushButton Command6 
            Height          =   285
            Left            =   5475
            TabIndex        =   8
            Top             =   285
            Width           =   300
            _Version        =   786432
            _ExtentX        =   529
            _ExtentY        =   494
            _StockProps     =   79
            Caption         =   "..."
            BackColor       =   16744576
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.GroupBox GroupBox4 
            Height          =   960
            Left            =   4065
            TabIndex        =   17
            Top             =   615
            Width           =   1740
            _Version        =   786432
            _ExtentX        =   3069
            _ExtentY        =   1693
            _StockProps     =   79
            Caption         =   "Medida Hoja"
            UseVisualStyle  =   -1  'True
            Begin VB.TextBox txtAnchoPieza 
               Height          =   285
               Left            =   735
               TabIndex        =   21
               Top             =   600
               Width           =   840
            End
            Begin VB.TextBox txtLargoPieza 
               Height          =   285
               Left            =   735
               TabIndex        =   20
               Top             =   255
               Width           =   840
            End
            Begin VB.Label Label15 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Ancho"
               Height          =   255
               Left            =   165
               TabIndex        =   19
               Top             =   600
               Width           =   615
            End
            Begin VB.Label Label8 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Largo"
               Height          =   255
               Left            =   165
               TabIndex        =   18
               Top             =   240
               Width           =   615
            End
         End
         Begin XtremeSuiteControls.ComboBox cboMateriales 
            Height          =   315
            Left            =   840
            TabIndex        =   23
            Top             =   270
            Width           =   4515
            _Version        =   786432
            _ExtentX        =   7964
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Sorted          =   -1  'True
            Style           =   2
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.GroupBox GroupBox5 
            Height          =   1380
            Left            =   5880
            TabIndex        =   24
            Top             =   195
            Width           =   5910
            _Version        =   786432
            _ExtentX        =   10425
            _ExtentY        =   2434
            _StockProps     =   79
            Caption         =   "Descripción Material"
            UseVisualStyle  =   -1  'True
            Begin VB.Label lblDescripcion 
               BackColor       =   &H00FFC0C0&
               Height          =   255
               Left            =   1260
               TabIndex        =   37
               Top             =   975
               Width           =   4560
            End
            Begin VB.Label Label14 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FF8080&
               Caption         =   "Espesor"
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
               Left            =   3810
               TabIndex        =   36
               Top             =   240
               Width           =   975
            End
            Begin VB.Label lblEspesor 
               BackColor       =   &H00FFC0C0&
               Height          =   255
               Left            =   4845
               TabIndex        =   35
               Top             =   240
               Width           =   1005
            End
            Begin VB.Label Label13 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FF8080&
               Caption         =   "Unidad"
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
               Left            =   3810
               TabIndex        =   34
               Top             =   480
               Width           =   975
            End
            Begin VB.Label lblUnidad 
               BackColor       =   &H00FFC0C0&
               Height          =   255
               Left            =   4830
               TabIndex        =   33
               Top             =   480
               Width           =   1005
            End
            Begin VB.Label lblKgM2 
               BackColor       =   &H00FFC0C0&
               Height          =   255
               Left            =   1230
               TabIndex        =   32
               Top             =   720
               Width           =   4605
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FF8080&
               Caption         =   "Kg x M2/Ml"
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
               Left            =   90
               TabIndex        =   31
               Top             =   720
               Width           =   1095
            End
            Begin VB.Label Label9 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FF8080&
               Caption         =   "Descripción "
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
               Left            =   45
               TabIndex        =   30
               Top             =   975
               Width           =   1215
            End
            Begin VB.Label lblGrupo 
               BackColor       =   &H00FFC0C0&
               Height          =   255
               Left            =   1245
               TabIndex        =   29
               Top             =   480
               Width           =   2430
            End
            Begin VB.Label Label7 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FF8080&
               Caption         =   "Grupo "
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
               Left            =   30
               TabIndex        =   28
               Top             =   480
               Width           =   1215
            End
            Begin VB.Label lblRubro 
               BackColor       =   &H00FFC0C0&
               Height          =   255
               Left            =   1245
               TabIndex        =   27
               Top             =   255
               Width           =   2445
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FF8080&
               Caption         =   "Rubro "
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
               Left            =   30
               TabIndex        =   26
               Top             =   255
               Width           =   1215
            End
         End
         Begin VB.Label Label3 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Código"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   300
            Width           =   615
         End
         Begin VB.Label Label11 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Cantidad"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   660
            Width           =   735
         End
         Begin VB.Label Label12 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Scrap"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   1020
            Width           =   735
         End
      End
   End
End
Attribute VB_Name = "frmPieza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim IdMoneda As Long
Dim desamat As DesarrolloMaterial
Dim vmaterial As clsMaterial
Dim colmat As New Collection

Private Sub cboMateriales_Click()
    buscarMaterial
End Sub

Private Sub limpiar()
    Me.GridEX1.ItemCount = 0
    Me.lblDescripcion = Empty
    Me.lblGrupo = Empty
    Me.lblKgM2 = Empty
    Me.lblRubro = Empty
    Me.lblEspesor = Empty
    Me.lblUnidad = Empty
End Sub


Private Sub MostrarMaterial(vmat As clsMaterial)
    Me.lblDescripcion = vmat.descripcion
    Me.lblRubro = vmat.Grupo.rubros.rubro
    Me.lblEspesor = vmat.Espesor
    Me.lblGrupo = vmat.Grupo.Grupo
    Me.lblKgM2 = vmat.PesoXUnidad
    Me.lblUnidad = enums.enumUnidades(vmat.unidad)
    validarControles
End Sub
Private Function buscarMaterial() As Boolean
    If Me.cboMateriales.ListIndex = -1 Then
        Set vmaterial = Nothing
        Exit Function
    End If
    buscarMaterial = True
    Set material = DAOMateriales.FindById(Me.cboMateriales.ItemData(Me.cboMateriales.ListIndex))
    If IsSomething(material) Then
        MostrarMaterial material
        Me.lblUnidad = enums.enumUnidades(material.unidad)
        buscarMaterial = True
    Else
        limpiar
        buscarMaterial = False
        MsgBox "No se encontro material con ese codigo", vbExclamation
    End If
End Function

Private Sub Command1_Click()
    If IsSomething(material) Then
        Set desamat = New DesarrolloMaterial
        desamat.Ancho = Val(Me.txtAnchoPieza)
        desamat.AnchoTerm = Val(Me.txtAnchoTerm)
        desamat.Largo = Val(Me.txtLargoPieza)
        desamat.LargoTerm = Val(Me.txtLargoTerm)
        desamat.Cantidad = Val(Me.txtCantidad)
        desamat.Scrap = Val(Me.txtScrap)
        Set desamat.material = material

        colmat.Add desamat
        LlenarListaMateriales
    End If
    LlenarListaMateriales

End Sub


Private Sub LlenarListaMateriales()
    Me.GridEX1.ItemCount = colmat.count
End Sub
Private Sub Command6_Click()
    Dim frm As New frmMaterialesLista2_modal
    frm.Usable = True
    frm.Show 1

    If IsSomething(Selecciones.material) Then
        Set material = Selecciones.material
        MostrarMaterial material
        Me.cboMateriales.ListIndex = funciones.PosIndexCbo(material.id, Me.cboMateriales)
        Set Selecciones.material = Nothing
    End If

End Sub

Private Sub validarControles()

    If IsSomething(material) Then
        Me.Command1.Enabled = True


        If material.unidad = m2_ Then
            'habilito los campos necesarios para M2
            Me.txtAnchoTerm = material.Ancho
            Me.txtLargoTerm = material.Largo
            Me.txtAnchoPieza.Enabled = True
            Me.txtAnchoTerm.Enabled = True
            Me.txtLargoPieza.Enabled = True
            Me.txtLargoTerm.Enabled = True
            Me.txtScrap.Enabled = True
        ElseIf material.unidad = kg_ Or material.unidad = un_ Then
            'si es un elemento unitario o por Kg deshabilito todo
            Me.txtAnchoPieza = 0
            Me.txtLargoPieza = 0
            Me.txtAnchoTerm = 0
            Me.txtLargoTerm = 0
            Me.txtAnchoPieza.Enabled = False
            Me.txtAnchoTerm.Enabled = False
            Me.txtLargoPieza.Enabled = False
            Me.txtLargoTerm.Enabled = False
            Me.txtScrap = 0
            Me.txtScrap.Enabled = False
        ElseIf material.unidad = Ml_ Then
            'habilito lo  necesitan los ml
            Me.txtAnchoPieza.Enabled = False
            Me.txtAnchoTerm.Enabled = False
            Me.txtLargoPieza.Enabled = True
            Me.txtLargoTerm.Enabled = True
            Me.txtAnchoTerm = 0
            Me.txtAnchoPieza = 0
            Me.txtLargoTerm = material.Largo

            Me.txtScrap.Enabled = True
        End If
    Else
        Me.Command1.Enabled = False
        '        Me.frame3.Enabled = False
        '        Me.Frame4.Enabled = False
    End If




End Sub



Private Sub LlenarComboMateriales()
    Me.cboMateriales.Clear
    Dim mattt As clsMaterial
    For Each mattt In DAOMateriales.FindAll()
        Me.cboMateriales.AddItem mattt.codigo & " - " & mattt.descripcion
        Me.cboMateriales.ItemData(Me.cboMateriales.NewIndex) = mattt.id
    Next mattt

End Sub

Private Sub Form_Load()
    Customize Me
    Me.ctrlCboIDPersona1.Personas = DAOCliente.FindAll(DAOCliente.TABLA_CLIENTE & "." & DAOCliente.CAMPO_ESTADO & "=" & EstadoCliente.activo & " ORDER BY " & DAOCliente.CAMPO_RAZON_SOCIAL)
    GridEXHelper.CustomizeGrid Me.GridEX1, False, True
    LlenarComboMateriales
    limpiar
End Sub


Private Sub GridEx1_BeforeDelete(ByVal Cancel As GridEX20.JSRetBoolean)
    Cancel = Not MsgBox("¿Seguro de eliminar el elemento?", vbYesNo, "Consulta") = vbYes
End Sub

Private Sub GridEX1_BeforeUpdate(ByVal Cancel As GridEX20.JSRetBoolean)



    Cancel = Not (IsNumeric(GridEX1.value(2)) And IsNumeric(GridEX1.value(5)) And IsNumeric(GridEX1.value(6)) And IsNumeric(GridEX1.value(7)) And IsNumeric(GridEX1.value(8)) And IsNumeric(GridEX1.value(9)))
End Sub

Private Sub GridEX1_SelectionChange()
    Dim it As Long: it = Me.GridEX1.RowIndex(Me.GridEX1.row)
    If it > 0 Then
        Set desamat = colmat.item(it)
        GridEX1.Columns(5).EditType = jgexEditTextBox
        GridEX1.Columns(6).EditType = jgexEditTextBox
        GridEX1.Columns(7).EditType = jgexEditTextBox
        GridEX1.Columns(8).EditType = jgexEditTextBox
        GridEX1.Columns(9).EditType = jgexEditTextBox


        If desamat.material.unidad = un_ Or desamat.material.unidad = kg_ Or desamat.material.unidad = litro_ Then
            GridEX1.Columns(5).EditType = jgexEditNone
            GridEX1.Columns(6).EditType = jgexEditNone
            GridEX1.Columns(7).EditType = jgexEditNone
            GridEX1.Columns(8).EditType = jgexEditNone
            GridEX1.Columns(9).EditType = jgexEditNone
        End If


        If desamat.material.unidad = Ml_ Then
            GridEX1.Columns(6).EditType = jgexEditNone
            GridEX1.Columns(8).EditType = jgexEditNone
        End If



    Else
        Set desamat = Nothing
    End If




End Sub

Private Sub GridEx1_UnboundDelete(ByVal RowIndex As Long, ByVal Bookmark As Variant)
    colmat.remove (RowIndex)
    LlenarListaMateriales
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Set desamat = colmat.item(RowIndex)
    Values(1) = desamat.material.codigo & " | " & desamat.material.descripcion
    Values(2) = desamat.Cantidad
    Values(3) = desamat.detalle
    Values(4) = desamat.material.Espesor
    Values(5) = desamat.Largo
    Values(6) = desamat.Ancho
    Values(7) = desamat.LargoTerm
    Values(8) = desamat.AnchoTerm
    Values(9) = desamat.Scrap
    Values(10) = desamat.Kg
    IdMoneda = DAOMoneda.FindFirstByPatronOrDefault.id
    Values(11) = desamat.CalcularDatosMaterial(IdMoneda).m2
    Values(12) = desamat.material.Moneda.NombreCorto & "  " & desamat.CalcularDatosMaterial(desamat.material.Moneda.id).costo
End Sub

Private Sub GridEX1_UnboundUpdate(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Set desamat = colmat.item(RowIndex)
    desamat.Cantidad = Values(2)
    desamat.Largo = Values(5)
    desamat.Ancho = Values(6)
    desamat.LargoTerm = Values(7)
    desamat.AnchoTerm = Values(8)
    desamat.Scrap = Values(9)
    desamat.detalle = Values(3)

End Sub

Private Sub txtAnchoPieza_GotFocus()
    foco Me.txtAnchoPieza
End Sub
Private Sub txtAnchoTerm_GotFocus()
    foco Me.txtAnchoTerm
End Sub
Private Sub txtLargoPieza_GotFocus()
    foco Me.txtLargoPieza
End Sub
Private Sub txtLargoTerm_GotFocus()
    foco Me.txtLargoTerm
End Sub
