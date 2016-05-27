VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~3.OCX"
Begin VB.Form frmComprasProveedoresLista 
   BackColor       =   &H00FF8080&
   Caption         =   "Proveedores"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   270
   ClientWidth     =   13335
   Icon            =   "frmListaProveedores.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5865
   ScaleWidth      =   13335
   Begin VB.TextBox txtCuit 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   465
      TabIndex        =   10
      Top             =   5430
      Width           =   2295
   End
   Begin XtremeSuiteControls.ComboBox cboRubros 
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   4680
      Width           =   3135
      _Version        =   786432
      _ExtentX        =   5530
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.PushButton Command1 
      Default         =   -1  'True
      Height          =   375
      Left            =   11370
      TabIndex        =   4
      Top             =   4635
      Width           =   1335
      _Version        =   786432
      _ExtentX        =   2355
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Filtrar"
      UseVisualStyle  =   -1  'True
   End
   Begin VB.ListBox lstEstados 
      Height          =   735
      Left            =   4020
      Style           =   1  'Checkbox
      TabIndex        =   3
      Top             =   4695
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   7320
      TabIndex        =   0
      Top             =   4680
      Width           =   3975
   End
   Begin GridEX20.GridEX grilla 
      Height          =   4560
      Left            =   15
      TabIndex        =   2
      Top             =   15
      Width           =   17415
      _ExtentX        =   30718
      _ExtentY        =   8043
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
      ColumnsCount    =   9
      Column(1)       =   "frmListaProveedores.frx":000C
      Column(2)       =   "frmListaProveedores.frx":0138
      Column(3)       =   "frmListaProveedores.frx":0214
      Column(4)       =   "frmListaProveedores.frx":02F4
      Column(5)       =   "frmListaProveedores.frx":03C0
      Column(6)       =   "frmListaProveedores.frx":048C
      Column(7)       =   "frmListaProveedores.frx":0560
      Column(8)       =   "frmListaProveedores.frx":0628
      Column(9)       =   "frmListaProveedores.frx":06FC
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmListaProveedores.frx":07C8
      FormatStyle(2)  =   "frmListaProveedores.frx":08F0
      FormatStyle(3)  =   "frmListaProveedores.frx":09A0
      FormatStyle(4)  =   "frmListaProveedores.frx":0A54
      FormatStyle(5)  =   "frmListaProveedores.frx":0B2C
      FormatStyle(6)  =   "frmListaProveedores.frx":0BE4
      ImageCount      =   0
      PrinterProperties=   "frmListaProveedores.frx":0CC4
   End
   Begin XtremeSuiteControls.PushButton cmdSinRubro 
      Height          =   255
      Left            =   3345
      TabIndex        =   6
      Top             =   4710
      Width           =   495
      _Version        =   786432
      _ExtentX        =   873
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "X"
      BackColor       =   12632256
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.ComboBox cboFantasia 
      Height          =   315
      Left            =   8280
      TabIndex        =   7
      Top             =   5415
      Width           =   3525
      _Version        =   786432
      _ExtentX        =   6218
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Sorted          =   -1  'True
      Text            =   "cboProveedores"
   End
   Begin XtremeSuiteControls.PushButton btnClearFantasia 
      Height          =   255
      Left            =   11895
      TabIndex        =   8
      Top             =   5475
      Width           =   420
      _Version        =   786432
      _ExtentX        =   741
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "X"
      BackColor       =   12632256
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.Label lblCuit 
      Height          =   195
      Left            =   30
      TabIndex        =   11
      Top             =   5475
      Width           =   375
      _Version        =   786432
      _ExtentX        =   661
      _ExtentY        =   344
      _StockProps     =   79
      Caption         =   "CUIT"
      BackColor       =   12632256
      AutoSize        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   195
      Left            =   7260
      TabIndex        =   9
      Top             =   5490
      Width           =   975
      _Version        =   786432
      _ExtentX        =   1720
      _ExtentY        =   344
      _StockProps     =   79
      Caption         =   "Nom Fantasia"
      BackColor       =   12632256
      AutoSize        =   -1  'True
   End
   Begin VB.Label P 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Rubros"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   4320
      Width           =   495
   End
   Begin VB.Menu m2 
      Caption         =   "m2"
      Visible         =   0   'False
      Begin VB.Menu editar 
         Caption         =   "Editar"
      End
      Begin VB.Menu con_tacto 
         Caption         =   "Contáctos..."
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
Dim id_rubro As Long
Dim proveedores As Collection
Dim Proveedor As clsProveedor
Public Property Let seleccionar(nvalue As Boolean)
    vSeleccionar = nvalue
End Property
Private Sub cboRubro_Change()
    Command1_Click
End Sub
Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub btnClearFantasia_Click()
    Me.cboFantasia.ListIndex = -1
End Sub

Private Sub cmdSinRubro_Click()
    Me.cboRubros.ListIndex = -1
End Sub

Private Sub Command1_Click()
    Buscar
End Sub

Private Sub Buscar()

    Dim filtro As String
    filtro = "1 = 1 "

    If LenB(Me.Text1.text) > 0 Then
        filtro = filtro & " and razon like '%" & Trim(Me.Text1.text) & "%'"
    End If
    If Me.cboRubros.ListIndex > -1 Then
        filtro = filtro & " And  asignacion.id_rubro =  " & Me.cboRubros.ItemData(Me.cboRubros.ListIndex)
    End If

    If Me.cboFantasia.ListIndex > -1 Then
        filtro = filtro & " And  proveedores.id =  " & Me.cboFantasia.ItemData(Me.cboFantasia.ListIndex)
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
Private Sub estado_Click()
    Set rectemp = proveedores(grilla.row)
    If MsgBox("¿Seguro que desea cambiar el estado del proveedor seleccionado?", vbYesNo, "Confirmacion") = vbYes Then
        If DAOProveedor.CambiarEstado(rectemp) Then
            MsgBox "Cambio exitoso!", vbInformation, "Información"
        Else
            MsgBox "Se produjo algún error. No se realizó el cambio!", vbCritical, "Error"
        End If
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

    Dim P As clsProveedor
    For Each P In DAOProveedor.FindAll()

        If LenB(Trim$(P.razonFantasia)) > 0 Then
            Me.cboFantasia.AddItem P.razonFantasia
            Me.cboFantasia.ItemData(Me.cboFantasia.NewIndex) = P.id
        End If
    Next P
    Me.cboFantasia.ListIndex = -1

    Buscar
End Sub

Private Sub llenarEstados()
    Me.lstEstados.AddItem enums.EnumEstadoProveedor(EstadoProveedor.EstadoProveedorContado)
    Me.lstEstados.ItemData(Me.lstEstados.NewIndex) = EstadoProveedor.EstadoProveedorContado
    Me.lstEstados.AddItem enums.EnumEstadoProveedor(EstadoProveedor.EstadoProveedorCuentaCorriente)
    Me.lstEstados.ItemData(Me.lstEstados.NewIndex) = EstadoProveedor.EstadoProveedorCuentaCorriente
    Me.lstEstados.AddItem enums.EnumEstadoProveedor(EstadoProveedor.EstadoProveedorEliminado)
    Me.lstEstados.ItemData(Me.lstEstados.NewIndex) = EstadoProveedor.EstadoProveedorEliminado
    'Me.lstEstados.Selected(PosIndexLST(EstadoProveedor.EstadoProveedorCuentaCorriente, Me.lstEstados)) = True
    'Me.lstEstados.ListIndex = -1 'EstadoProveedor.EstadoProveedorCuentaCorriente

End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Me.grilla.Width = Me.ScaleWidth
    Me.grilla.Height = Me.Height - (Me.cboRubros.Height + (1500 - Me.cboRubros.Height))
    Me.grilla.ColumnAutoResize = True
    Me.cboRubros.Top = Me.Height - 1350
    Me.Text1.Top = Me.cboRubros.Top
    Me.Command1.Top = Me.cboRubros.Top
    Me.lstEstados.Top = Me.cboRubros.Top    '+ 400
    cmdSinRubro.Top = Me.cboRubros.Top
    Me.lblCuit.Top = Me.cboRubros.Top + 500
    Me.txtCuit.Top = Me.lblCuit.Top



    Me.cboFantasia.Top = Me.cboRubros.Top + 400
    Me.Label3.Top = Me.cboFantasia.Top
    Me.btnClearFantasia.Top = Me.cboFantasia.Top
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
    a = grilla.RowIndex(grilla.row)
    If vSeleccionar And a > 0 Then
        Selecciones.proveedorElegido = proveedores(a)
    Else
        If grilla.rowcount > 0 And a > 0 Then

            Set Proveedor = proveedores(a)
            frmComprasProveedoresModifica.Proveedor = Proveedor
            frmComprasProveedoresModifica.Show
        End If
    End If
End Sub
Private Sub grilla_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Set rectemp = proveedores(grilla.RowIndex(grilla.row))
        Me.PopupMenu m2
    End If
End Sub
Private Sub grilla_SelectionChange()
    rows = grilla.RowIndex(grilla.row)
End Sub
Private Sub grilla_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Set rectemp = proveedores.item(RowIndex)  ' mcData.Item(RowIndex)
    With rectemp
        Values(1) = Format(.id, "0000")
        Values(2) = .RazonSocial
        Values(3) = .razonFantasia
        Values(4) = .Cuit
        Values(5) = .IIBB
        Values(6) = .tel
        Values(7) = .Fax
        Values(8) = .direccion
        Values(9) = .email
    End With
End Sub
Private Sub VerDetalles_Click()
    If Me.grilla.rowcount > 0 Then
        Set rectemp = proveedores(grilla.row)
        frmComprasProveedoresModifica.tipoOperacion = ver
        frmComprasProveedoresModifica.Proveedor = rectemp
        frmComprasProveedoresModifica.Show
    End If
End Sub


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
            If proveedores(i).id = tmp.id Then
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




        grilla.RefreshRowIndex EVENTO.Elemento.id
    End If
    Exit Function
err1:

End Function
