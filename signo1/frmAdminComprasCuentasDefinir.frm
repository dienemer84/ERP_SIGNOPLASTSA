VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmAdminComprasCuentasDefinir 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Definir Cuentas a un proveedor..."
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   480
   ClientWidth     =   6255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   6255
   Begin XtremeSuiteControls.ComboBox cboProveedores 
      Height          =   315
      Left            =   1200
      TabIndex        =   9
      Top             =   240
      Width           =   4695
      _Version        =   786432
      _ExtentX        =   8281
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboCuentas 
      Height          =   315
      Left            =   1200
      TabIndex        =   8
      Top             =   1800
      Width           =   3615
      _Version        =   786432
      _ExtentX        =   6376
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Text            =   "ComboBox1"
   End
   Begin GridEX20.GridEX grilla 
      Height          =   2205
      Left            =   150
      TabIndex        =   7
      Top             =   2310
      Width           =   5970
      _ExtentX        =   10530
      _ExtentY        =   3889
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      ColumnAutoResize=   -1  'True
      MethodHoldFields=   -1  'True
      AllowDelete     =   -1  'True
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   2
      Column(1)       =   "frmAdminComprasCuentasDefinir.frx":0000
      Column(2)       =   "frmAdminComprasCuentasDefinir.frx":0118
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminComprasCuentasDefinir.frx":0214
      FormatStyle(2)  =   "frmAdminComprasCuentasDefinir.frx":034C
      FormatStyle(3)  =   "frmAdminComprasCuentasDefinir.frx":03FC
      FormatStyle(4)  =   "frmAdminComprasCuentasDefinir.frx":04B0
      FormatStyle(5)  =   "frmAdminComprasCuentasDefinir.frx":0588
      FormatStyle(6)  =   "frmAdminComprasCuentasDefinir.frx":0640
      ImageCount      =   0
      PrinterProperties=   "frmAdminComprasCuentasDefinir.frx":0720
   End
   Begin XtremeSuiteControls.PushButton Command3 
      Height          =   330
      Left            =   4065
      TabIndex        =   6
      Top             =   4590
      Width           =   1005
      _Version        =   786432
      _ExtentX        =   1773
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Salir"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton Command2 
      Height          =   345
      Left            =   5115
      TabIndex        =   5
      Top             =   4575
      Width           =   1020
      _Version        =   786432
      _ExtentX        =   1799
      _ExtentY        =   609
      _StockProps     =   79
      Caption         =   "Guardar"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton Command1 
      Height          =   360
      Left            =   4845
      TabIndex        =   4
      Top             =   1815
      Width           =   945
      _Version        =   786432
      _ExtentX        =   1667
      _ExtentY        =   635
      _StockProps     =   79
      Caption         =   "Agregar"
      UseVisualStyle  =   -1  'True
   End
   Begin VB.TextBox txtCodigo 
      Height          =   285
      Left            =   1200
      TabIndex        =   2
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      Caption         =   "Cuenta"
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
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      Caption         =   "Código"
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
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      Caption         =   "Proveedor"
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
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frmAdminComprasCuentasDefinir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vcta As clsCuentaContable
Dim vcodigo As Long
Public vProveedor As clsProveedor
Dim loading As Boolean
Private Sub llenarComboProveedores()
    Dim col As New Collection
    Set col = DAOProveedor.FindAll(, , True)
    For x = 1 To col.count
        Me.cboProveedores.AddItem UCase(col(x).RazonSocial)
        Me.cboProveedores.ItemData(cboProveedores.NewIndex) = col(x).Id
    Next x
    If Me.cboProveedores.ListCount > 0 Then Me.cboProveedores.ListIndex = 0
End Sub

Private Sub cboProveedores_Click()
    Me.txtCodigo = Me.cboProveedores.ItemData(Me.cboProveedores.ListIndex)
    marcar CLng(Me.txtCodigo)
End Sub
Private Sub Command1_Click()
    Dim Id As Long
    Dim esta As Boolean
    Dim cta As clsCuentaContable
    If cboCuentas.ListIndex <> -1 Then

        Id = CLng(Me.cboCuentas.ItemData(Me.cboCuentas.ListIndex))
    End If
    Set cta = DAOCuentaContable.GetById(Id)
    esta = False

    For Each vcta In vProveedor.cuentasContables
        If vcta.Id = cta.Id Then
            esta = True
            Exit For
        End If
    Next
    If esta Then Exit Sub
    vProveedor.cuentasContables.Add cta
    LlenarPlan2
End Sub

Private Sub Command2_Click()
    If MsgBox("¿Está seguro de guardar?", vbYesNo, "Confirmación") = vbYes Then


        If DAOCuentasProveedor.Save(vProveedor) Then
            MsgBox "Actualización exitosa!", vbInformation, "Información"
        Else
            MsgBox "Se produjo algun error. No se actualiza!", vbCritical, "Error"
        End If
    End If
End Sub
Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    loading = True
    GridEXHelper.CustomizeGrid Me.grilla, False, False
    FormHelper.Customize Me
    llenarComboProveedores
    llenarComboCuentas
    If IsSomething(vProveedor) Then
        Me.cboProveedores.ListIndex = funciones.PosIndexCbo(vProveedor.Id, Me.cboProveedores)
        Me.cboProveedores.Enabled = False
        Me.txtCodigo.Enabled = False
    End If
    Me.grilla.ItemCount = 0
    loading = False
End Sub
Private Sub mostarId()
    Me.txtCodigo = Me.cboProveedores.ItemData(Me.cboProveedores.ListIndex)
End Sub


Private Sub grilla_BeforeDelete(ByVal Cancel As GridEX20.JSRetBoolean)
    Cancel = Not (MsgBox("¿Seguro de eliminar?", vbInformation + vbYesNo, "Confirmar") = vbYes)
End Sub

Private Sub grilla_UnboundDelete(ByVal rowIndex As Long, ByVal Bookmark As Variant)
    vProveedor.cuentasContables.remove rowIndex
End Sub

Private Sub grilla_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    On Error Resume Next
    Set vcta = vProveedor.cuentasContables.item(rowIndex)
    Values(1) = vcta.codigo
    Values(2) = vcta.nombre
End Sub
Private Sub txtCodigo_Change()
    On Error Resume Next
    Me.cboProveedores.ListIndex = funciones.PosIndexCbo(CLng(Me.txtCodigo), Me.cboProveedores)
    marcar CLng(Me.txtCodigo)
End Sub
Public Sub marcar(nro As Long)
    If Not loading Then
        vcodigo = nro
        Set vProveedor = DAOProveedor.FindById(vcodigo)
        LlenarPlan2
    End If
End Sub
Private Sub llenarComboCuentas()
    Dim col As New Collection
    Set col = DAOCuentaContable.GetAll
    For x = 1 To col.count
        Me.cboCuentas.AddItem UCase(col(x).nombre)
        Me.cboCuentas.ItemData(cboCuentas.NewIndex) = col(x).Id
    Next x
    If Me.cboCuentas.ListCount > 0 Then Me.cboCuentas.ListIndex = 0
End Sub
Private Sub LlenarPlan2()
    Me.grilla.ItemCount = 0
    Me.grilla.ItemCount = vProveedor.cuentasContables.count
End Sub

