VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmAdminFacturasNCElegirFC 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Elegir Factura"
   ClientHeight    =   5940
   ClientLeft      =   2625
   ClientTop       =   2295
   ClientWidth     =   6555
   ClipControls    =   0   'False
   Icon            =   "frmAdminFacturasNCElegirFC.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   6555
   Begin XtremeSuiteControls.GroupBox grpResultados 
      Height          =   3975
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   6255
      _Version        =   786432
      _ExtentX        =   11033
      _ExtentY        =   7011
      _StockProps     =   79
      Caption         =   "Resultados"
      UseVisualStyle  =   -1  'True
      Begin MSComctlLib.ListView lstFacturas 
         Height          =   3615
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   6376
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "FC"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Cliente"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "FF"
            Object.Width           =   1941
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Estado"
            Object.Width           =   1587
         EndProperty
      End
   End
   Begin XtremeSuiteControls.GroupBox GroFiltrosDe 
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      _Version        =   786432
      _ExtentX        =   11033
      _ExtentY        =   2778
      _StockProps     =   79
      Caption         =   "Filtros de Búsqueda"
      UseVisualStyle  =   -1  'True
      Begin VB.TextBox txtNumero 
         Height          =   330
         Left            =   960
         TabIndex        =   6
         Top             =   352
         Width           =   1740
      End
      Begin XtremeSuiteControls.PushButton btnBorrarNumero 
         Height          =   375
         Index           =   0
         Left            =   2880
         TabIndex        =   2
         Top             =   330
         Width           =   375
         _Version        =   786432
         _ExtentX        =   661
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "X"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnBorrarCliente 
         Height          =   375
         Index           =   1
         Left            =   3960
         TabIndex        =   3
         Top             =   960
         Width           =   375
         _Version        =   786432
         _ExtentX        =   661
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "X"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cboClientes 
         Height          =   315
         Left            =   960
         TabIndex        =   4
         Top             =   990
         Width           =   2850
         _Version        =   786432
         _ExtentX        =   5027
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Text            =   "ComboBox"
      End
      Begin XtremeSuiteControls.PushButton btnBuscar 
         Default         =   -1  'True
         Height          =   465
         Left            =   4560
         TabIndex        =   5
         Top             =   915
         Width           =   1410
         _Version        =   786432
         _ExtentX        =   2487
         _ExtentY        =   820
         _StockProps     =   79
         Caption         =   "Buscar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
      End
      Begin VB.Label lblNúmero 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Número"
         Height          =   225
         Left            =   120
         TabIndex        =   9
         Top             =   405
         Width           =   735
      End
      Begin VB.Label lblCliente 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente"
         Height          =   255
         Left            =   210
         TabIndex        =   8
         Top             =   1020
         Width           =   645
      End
      Begin VB.Label lblBuscarEn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   3600
         TabIndex        =   7
         Top             =   390
         Width           =   2445
      End
   End
End
Attribute VB_Name = "frmAdminFacturasNCElegirFC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strsql As String

Public idCliente As Long
Dim tipos As String
Dim estados As String
Public TiposDocs As New Collection    'of  tipoDocumentoContable
Public EstadosDocs As New Collection    'of Estado Factura


Private Sub llenarLST()

    Me.lstFacturas.ListItems.Clear

    Dim F As String
    
    F = "1 = 1 "

    If idCliente > 0 Then
        F = F & " and AdminFacturas.idCliente=" & idCliente
    End If

    If LenB(Me.txtNumero) > 0 Then
        F = F & " and AdminFacturas.NroFactura like '%" & Me.txtNumero & "%'"
    End If

    F = F & " and acftd.tipo_documento IN (" & tipos & ")"

    F = F & " and AdminFacturas.estado  IN (" & estados & ")"

    Dim facs As Collection
    Dim fac As Factura
    Set facs = DAOFactura.FindAll(F)

    For Each fac In facs
        Set x = Me.lstFacturas.ListItems.Add(, , fac.GetShortDescription(False, True))
        x.SubItems(1) = fac.cliente.razon
        x.SubItems(2) = fac.FechaEmision
        x.SubItems(3) = enums.EnumEstadoDocumentoContable(fac.estado)

        x.Tag = fac.Id
        
    Next fac
    
End Sub



Private Sub btnBorrarCliente_Click(Index As Integer)
    Me.cboClientes.ListIndex = -1
End Sub

Private Sub btnBorrarNumero_Click(Index As Integer)
    Me.txtNumero.text = ""
    
    llenarLST
End Sub

Private Sub Form_Load()
    FormHelper.Customize Me
    DAOCliente.llenarComboXtremeSuite Me.cboClientes
    If idCliente > 0 Then
        Me.cboClientes.ListIndex = funciones.PosIndexCbo(Me.idCliente, cboClientes)
    Else
        Me.cboClientes.ListIndex = -1
    End If


    tipos = funciones.JoinCollectionValues(TiposDocs, ",")
    estados = funciones.JoinCollectionValues(EstadosDocs, ",")

    Dim T As String
    For Each i In TiposDocs
        T = T & enums.EnumTipoDocumentoContable(i) & ","

    Next i

    Me.lblBuscarEn = "Busca en : " & Mid(T, 1, Len(T) - 1)

    llenarLST

    Me.cboClientes.Enabled = False
    Me.btnBorrarCliente(1).Enabled = False

End Sub

Private Sub Form_Terminate()
    Set Selecciones.Factura = Nothing

End Sub

Private Sub lstFacturas_DblClick()
    If Me.lstFacturas.ListItems.count > 0 Then

        Set Selecciones.Factura = DAOFactura.FindById(Me.lstFacturas.selectedItem.Tag, True)
        Unload Me
    End If
End Sub

Private Sub btnBuscar_Click()
    llenarLST
End Sub


