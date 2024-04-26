VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmAdminFacturasNCElegirFC 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Elegir Factura"
   ClientHeight    =   5430
   ClientLeft      =   2625
   ClientTop       =   2295
   ClientWidth     =   6270
   ClipControls    =   0   'False
   Icon            =   "frmAdminFacturasNCElegirFC.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   6270
   Begin XtremeSuiteControls.PushButton PushButton2 
      Height          =   330
      Left            =   4065
      TabIndex        =   5
      Top             =   495
      Width           =   420
      _Version        =   786432
      _ExtentX        =   741
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "X"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.ComboBox cboClientes 
      Height          =   315
      Left            =   1170
      TabIndex        =   4
      Top             =   495
      Width           =   2850
      _Version        =   786432
      _ExtentX        =   5027
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.PushButton PushButton1 
      Height          =   345
      Left            =   345
      TabIndex        =   3
      Top             =   990
      Width           =   1170
      _Version        =   786432
      _ExtentX        =   2064
      _ExtentY        =   609
      _StockProps     =   79
      Caption         =   "Buscar"
      UseVisualStyle  =   -1  'True
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   1155
      TabIndex        =   1
      Top             =   105
      Width           =   1740
   End
   Begin MSComctlLib.ListView lstFacturas 
      Height          =   3015
      Left            =   45
      TabIndex        =   0
      Top             =   1770
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   5318
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
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
   Begin VB.Label lblBuscarEn 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   1575
      TabIndex        =   7
      Top             =   1065
      Width           =   45
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Cliente"
      Height          =   255
      Left            =   255
      TabIndex        =   6
      Top             =   555
      Width           =   885
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "N?mero"
      Height          =   225
      Left            =   390
      TabIndex        =   2
      Top             =   165
      Width           =   750
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
    Dim F As String
    F = "1 = 1 "
    '    f = f & "acftd.tipo_documento = " & tipoDocumentoContable.Factura
    '    f = f & " or acftd.tipo_documento = " & tipoDocumentoContable.NotaDebito
    '
    '
    '
    '    f = f & " AND AdminFacturas.estado = " & EstadoFacturaCliente.Aprobada



    If idCliente > 0 Then
        F = F & " and AdminFacturas.idCliente=" & idCliente
    End If


    If LenB(Me.Text1) > 0 Then
        F = F & " and AdminFacturas.NroFactura like '%" & Me.Text1 & "%'"
    End If




    F = F & " and acftd.tipo_documento IN (" & tipos & ")"




    F = F & " and AdminFacturas.estado  IN (" & estados & ")"


    Dim facs As Collection
    Dim fac As Factura
    Set facs = DAOFactura.FindAll(F)



    For Each fac In facs

        Set x = Me.lstFacturas.ListItems.Add(, , fac.GetShortDescription(False, True))
        x.SubItems(1) = fac.Cliente.razon
        x.SubItems(2) = fac.FechaEmision
        x.SubItems(3) = enums.EnumEstadoDocumentoContable(fac.estado)

        x.Tag = fac.Id
    Next fac
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
    Me.PushButton2.Enabled = False



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

Private Sub PushButton1_Click()
    llenarLST
End Sub

Private Sub PushButton2_Click()
    Me.cboClientes.ListIndex = -1
End Sub
                                                                                                                                     
