VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmComprasProveedoresLista 
   BackColor       =   &H00FF8080&
   Caption         =   "Proveedores"
   ClientHeight    =   7470
   ClientLeft      =   60
   ClientTop       =   270
   ClientWidth     =   17790
   Icon            =   "frmListaProveedores.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7470
   ScaleWidth      =   17790
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   17415
      _Version        =   786432
      _ExtentX        =   30718
      _ExtentY        =   3201
      _StockProps     =   79
      Caption         =   "Búsqueda"
      UseVisualStyle  =   -1  'True
      Begin VB.TextBox TextFantasia 
         Height          =   285
         Left            =   1320
         TabIndex        =   13
         Top             =   960
         Width           =   3975
      End
      Begin VB.ListBox lstEstados 
         Height          =   735
         Left            =   6240
         Style           =   1  'Checkbox
         TabIndex        =   10
         Top             =   600
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1320
         TabIndex        =   4
         Top             =   600
         Width           =   3975
      End
      Begin VB.TextBox txtCuit 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1320
         TabIndex        =   2
         Top             =   240
         Width           =   2295
      End
      Begin XtremeSuiteControls.ComboBox cboRubros 
         Height          =   315
         Left            =   1320
         TabIndex        =   7
         Top             =   1350
         Width           =   4005
         _Version        =   786432
         _ExtentX        =   7064
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.PushButton cmdSinRubro 
         Height          =   255
         Left            =   5400
         TabIndex        =   8
         Top             =   1395
         Width           =   420
         _Version        =   786432
         _ExtentX        =   741
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "X"
         BackColor       =   12632256
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton Command1 
         Default         =   -1  'True
         Height          =   375
         Left            =   9600
         TabIndex        =   11
         Top             =   1320
         Width           =   1335
         _Version        =   786432
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Filtrar"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblEstados 
         Height          =   195
         Index           =   1
         Left            =   6240
         TabIndex        =   12
         Top             =   360
         Width           =   615
         _Version        =   786432
         _ExtentX        =   1085
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Estados:"
         BackColor       =   12632256
         Alignment       =   1
         AutoSize        =   -1  'True
      End
      Begin VB.Label P 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Rubros:"
         Height          =   255
         Left            =   480
         TabIndex        =   9
         Top             =   1375
         Width           =   735
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   1005
         Width           =   1020
         _Version        =   786432
         _ExtentX        =   1799
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Nom Fantasia:"
         BackColor       =   12632256
         Alignment       =   1
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Left            =   165
         TabIndex        =   5
         Top             =   615
         Width           =   1095
         _Version        =   786432
         _ExtentX        =   1931
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Razón Social:"
         BackColor       =   12632256
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label lblCuit 
         Height          =   195
         Index           =   0
         Left            =   840
         TabIndex        =   3
         Top             =   285
         Width           =   420
         _Version        =   786432
         _ExtentX        =   741
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "CUIT:"
         BackColor       =   12632256
         Alignment       =   1
         AutoSize        =   -1  'True
      End
   End
   Begin GridEX20.GridEX grilla 
      Height          =   5280
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   17415
      _ExtentX        =   30718
      _ExtentY        =   9313
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      MethodHoldFields=   -1  'True
      Options         =   -1
      RecordsetType   =   1
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      BackColorHeader =   16761024
      DataMode        =   99
      HeaderFontBold  =   -1  'True
      HeaderFontWeight=   700
      FontName        =   "Tahoma"
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   12
      Column(1)       =   "frmListaProveedores.frx":000C
      Column(2)       =   "frmListaProveedores.frx":0114
      Column(3)       =   "frmListaProveedores.frx":020C
      Column(4)       =   "frmListaProveedores.frx":0308
      Column(5)       =   "frmListaProveedores.frx":03F4
      Column(6)       =   "frmListaProveedores.frx":04E0
      Column(7)       =   "frmListaProveedores.frx":05C8
      Column(8)       =   "frmListaProveedores.frx":06AC
      Column(9)       =   "frmListaProveedores.frx":07A8
      Column(10)      =   "frmListaProveedores.frx":0898
      Column(11)      =   "frmListaProveedores.frx":097C
      Column(12)      =   "frmListaProveedores.frx":0A68
      FormatStylesCount=   7
      FormatStyle(1)  =   "frmListaProveedores.frx":0B74
      FormatStyle(2)  =   "frmListaProveedores.frx":0C9C
      FormatStyle(3)  =   "frmListaProveedores.frx":0D4C
      FormatStyle(4)  =   "frmListaProveedores.frx":0E00
      FormatStyle(5)  =   "frmListaProveedores.frx":0ED8
      FormatStyle(6)  =   "frmListaProveedores.frx":0F90
      FormatStyle(7)  =   "frmListaProveedores.frx":1070
      ImageCount      =   0
      PrinterProperties=   "frmListaProveedores.frx":1090
   End
   Begin VB.Menu m2 
      Caption         =   "m2"
      Visible         =   0   'False
      Begin VB.Menu editar 
         Caption         =   "Editar"
      End
      Begin VB.Menu con_tacto 
         Caption         =   "Cont?ctos..."
      End
   End
End
Attribute VB_Name = "frmComprasProveedoresLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Implements ISuscriber

Dim suscriber_id As String
Dim vSeleccionar As Boolean
Dim rows As Long
Dim rectemp As clsProveedor
'Dim id_rubro As Long
Dim proveedores As Collection
Dim Proveedor As clsProveedor


Public Property Let seleccionar(nvalue As Boolean)
    vSeleccionar = nvalue
End Property

'Private Sub cboRubro_Change()
'    Command1_Click
'End Sub

'Private Sub cmdSalir_Click()
'    Unload Me
'End Sub

Private Sub cmdSinRubro_Click()
    Me.cboRubros.ListIndex = -1
End Sub

Private Sub Command1_Click()
    Buscar
End Sub

Private Sub Buscar()

    Dim filtro As String
    filtro = "1 = 1 "

    If LenB(Me.Text1.Text) > 0 Then
        filtro = filtro & " and razon like '%" & Trim(Me.Text1.Text) & "%'"
    End If

    If Me.cboRubros.ListIndex > -1 Then
        filtro = filtro & " And  asignacion.id_rubro =  " & Me.cboRubros.ItemData(Me.cboRubros.ListIndex)
    End If

    If LenB(Me.TextFantasia.Text) > 0 Then
        filtro = filtro & " and razon_fantasia like '%" & Trim(Me.TextFantasia.Text) & "%'"
    End If

    If LenB(Me.txtCuit) > 0 Then
        filtro = filtro & " and cuit like '%" & Trim(Me.txtCuit) & "%'"
    End If

    Dim ctacte As Boolean
    Dim contado As Boolean
    Dim elim As Boolean
    ctacte = Me.lstEstados.Selected(PosIndexLST(EstadoProveedor.EstadoProveedorCuentaCorriente, Me.lstEstados))
    contado = Me.lstEstados.Selected(PosIndexLST(EstadoProveedor.EstadoProveedorContado, Me.lstEstados))
    elim = Me.lstEstados.Selected(PosIndexLST(EstadoProveedor.EstadoProveedorEliminado, Me.lstEstados))
    Set proveedores = DAOProveedor.FindAll(filtro, False, , , ctacte, contado, elim, False)
    grilla.ItemCount = 0
    grilla.ItemCount = proveedores.count
    grilla.ReBind
End Sub

Private Sub con_tacto_Click()
    If grilla.rowcount > 0 Then
        Set rectemp = proveedores(grilla.row)
        frmVentasClientesNuevoContacto.Proveedor = rectemp
        frmVentasClientesNuevoContacto.Show
    End If
End Sub


Private Sub editar_Click()
    If grilla.rowcount > 0 Then
        Set Proveedor = proveedores(grilla.row)
        Dim F As New frmComprasProveedoresModifica
        F.Proveedor = Proveedor
        F.Show
    End If
End Sub


Private Sub Form_Load()
    FormHelper.Customize Me
    GridEXHelper.CustomizeGrid Me.grilla
    Me.grilla.ItemCount = 0
    DAORubros.LlenarComboExtremeSuite Me.cboRubros
    Me.cboRubros.ListIndex = -1
    suscriber_id = funciones.CreateGUID
    Channel.AgregarSuscriptor Me, Proveedores_
    rows = 1
    llenarEstados


    Buscar

    ''Me.caption = caption & " (" & Name & ")"

End Sub

Private Sub llenarEstados()
    Me.lstEstados.AddItem enums.EnumEstadoProveedor(EstadoProveedor.EstadoProveedorContado)
    Me.lstEstados.ItemData(Me.lstEstados.NewIndex) = EstadoProveedor.EstadoProveedorContado
    Me.lstEstados.AddItem enums.EnumEstadoProveedor(EstadoProveedor.EstadoProveedorCuentaCorriente)
    Me.lstEstados.ItemData(Me.lstEstados.NewIndex) = EstadoProveedor.EstadoProveedorCuentaCorriente
    Me.lstEstados.AddItem enums.EnumEstadoProveedor(EstadoProveedor.EstadoProveedorEliminado)
    Me.lstEstados.ItemData(Me.lstEstados.NewIndex) = EstadoProveedor.EstadoProveedorEliminado

End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Me.grilla.Width = Me.ScaleWidth - 300
    Me.grilla.Height = Me.Height - 2700
    Me.grilla.ColumnAutoResize = True
    Me.GroupBox1.Width = Me.ScaleWidth - 300

End Sub

Private Sub Form_Terminate()
    Channel.RemoverSuscripcionTotal Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Channel.RemoverSuscripcionTotal Me
End Sub

Private Sub grilla_ColumnHeaderClick(ByVal Column As GridEX20.JSColumn)
    ordenar_grilla Column, Me.grilla
End Sub

Private Sub grilla_DblClick()
    A = grilla.rowIndex(grilla.row)
    If vSeleccionar And A > 0 Then
        Selecciones.proveedorElegido = proveedores(A)
    Else
        If grilla.rowcount > 0 And A > 0 Then

            Set Proveedor = proveedores(A)
            frmComprasProveedoresModifica.Proveedor = Proveedor
            frmComprasProveedoresModifica.Show
        End If
    End If
End Sub

Private Sub grilla_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Set rectemp = proveedores(grilla.rowIndex(grilla.row))
        Me.PopupMenu m2
    End If
End Sub

Private Sub grilla_SelectionChange()
    rows = grilla.rowIndex(grilla.row)
End Sub

Private Sub grilla_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Set rectemp = proveedores.item(rowIndex)  ' mcData.Item(RowIndex)
    With rectemp
        Values(1) = Format(.Id, "0000")
        Values(2) = .RazonSocial
        Values(3) = .razonFantasia
        Values(4) = .Cuit
        Values(5) = .IIBB
        Values(6) = .tel
        Values(7) = .Fax
        Values(8) = .direccion
        Values(9) = .email
        Values(10) = .CBU
        Values(11) = .ALIAS
        Values(12) = .TitularCta

    End With
End Sub
'
'Private Sub VerDetalles_Click()
'    If Me.grilla.rowcount > 0 Then
'        Set rectemp = proveedores(grilla.row)
'        frmComprasProveedoresModifica.tipoOperacion = Ver
'        frmComprasProveedoresModifica.Proveedor = rectemp
'        frmComprasProveedoresModifica.Show
'    End If
'End Sub

Private Property Get ISuscriber_id() As String
    ISuscriber_id = suscriber_id
End Property

Private Function ISuscriber_Notificarse(EVENTO As clsEventoObserver) As Variant
    On Error GoTo err1
    Dim tmp As clsProveedor
    If EVENTO.EVENTO = agregar_ Then
        proveedores.Add EVENTO.Elemento
        grilla.ItemCount = proveedores.count
    ElseIf EVENTO.EVENTO = modificar_ Then
        Set tmp = EVENTO.Elemento

        For i = proveedores.count To 1 Step -1
            If proveedores(i).Id = tmp.Id Then
                Set Proveedor = proveedores(i)
                Proveedor.estado = tmp.estado
                Proveedor.razonFantasia = tmp.razonFantasia
                Proveedor.RazonSocial = tmp.RazonSocial
                Proveedor.Cuit = tmp.Cuit
                Proveedor.IIBB = tmp.IIBB
                Proveedor.direccion = tmp.direccion

                grilla.RefreshRowIndex i
                Exit For
            End If
        Next
        grilla.RefreshRowIndex EVENTO.Elemento.Id

    End If
    Exit Function
err1:

End Function

