VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmAdminconfigCuentas 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cuentas Bancarias"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9165
   Icon            =   "frmAdminconfigCuentas.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   9165
   ShowInTaskbar   =   0   'False
   Begin GridEX20.GridEX grid 
      Height          =   3915
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   9150
      _ExtentX        =   16140
      _ExtentY        =   6906
      Version         =   "2.0"
      PreviewRowIndent=   500
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      PreviewRowLines =   2
      ColumnAutoResize=   -1  'True
      MethodHoldFields=   -1  'True
      ContScroll      =   -1  'True
      AllowColumnDrag =   0   'False
      GroupByBoxVisible=   0   'False
      BackColorHeader =   16761024
      RowHeaders      =   -1  'True
      ItemCount       =   1
      DataMode        =   99
      AllowAddNew     =   -1  'True
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   4
      Column(1)       =   "frmAdminconfigCuentas.frx":000C
      Column(2)       =   "frmAdminconfigCuentas.frx":013C
      Column(3)       =   "frmAdminconfigCuentas.frx":0228
      Column(4)       =   "frmAdminconfigCuentas.frx":0494
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminconfigCuentas.frx":05A8
      FormatStyle(2)  =   "frmAdminconfigCuentas.frx":06E0
      FormatStyle(3)  =   "frmAdminconfigCuentas.frx":0790
      FormatStyle(4)  =   "frmAdminconfigCuentas.frx":0844
      FormatStyle(5)  =   "frmAdminconfigCuentas.frx":091C
      FormatStyle(6)  =   "frmAdminconfigCuentas.frx":09D4
      ImageCount      =   0
      PrinterProperties=   "frmAdminconfigCuentas.frx":0AB4
   End
   Begin GridEX20.GridEX gridMonedas 
      Height          =   1320
      Left            =   4650
      TabIndex        =   2
      Top             =   1605
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   2328
      Version         =   "2.0"
      PreviewRowIndent=   500
      BoundColumnIndex=   "id"
      ReplaceColumnIndex=   "nombre"
      ActAsDropDown   =   -1  'True
      PreviewRowLines =   2
      ColumnAutoResize=   -1  'True
      HideSelection   =   2
      MethodHoldFields=   -1  'True
      ContScroll      =   -1  'True
      AllowColumnDrag =   0   'False
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      ColumnHeaders   =   0   'False
      BackColorHeader =   16761024
      ItemCount       =   1
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   2
      Column(1)       =   "frmAdminconfigCuentas.frx":0C8C
      Column(2)       =   "frmAdminconfigCuentas.frx":0DD0
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminconfigCuentas.frx":0EE4
      FormatStyle(2)  =   "frmAdminconfigCuentas.frx":101C
      FormatStyle(3)  =   "frmAdminconfigCuentas.frx":10CC
      FormatStyle(4)  =   "frmAdminconfigCuentas.frx":1180
      FormatStyle(5)  =   "frmAdminconfigCuentas.frx":1258
      FormatStyle(6)  =   "frmAdminconfigCuentas.frx":1310
      ImageCount      =   0
      PrinterProperties=   "frmAdminconfigCuentas.frx":13F0
   End
   Begin GridEX20.GridEX gridBancos 
      Height          =   1320
      Left            =   930
      TabIndex        =   1
      Top             =   1665
      Visible         =   0   'False
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   2328
      Version         =   "2.0"
      PreviewRowIndent=   500
      BoundColumnIndex=   "id"
      ReplaceColumnIndex=   "nombre"
      ActAsDropDown   =   -1  'True
      PreviewRowLines =   2
      ColumnAutoResize=   -1  'True
      HideSelection   =   2
      MethodHoldFields=   -1  'True
      ContScroll      =   -1  'True
      AllowColumnDrag =   0   'False
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      ColumnHeaders   =   0   'False
      BackColorHeader =   16761024
      ItemCount       =   1
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   2
      Column(1)       =   "frmAdminconfigCuentas.frx":15C8
      Column(2)       =   "frmAdminconfigCuentas.frx":170C
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminconfigCuentas.frx":1820
      FormatStyle(2)  =   "frmAdminconfigCuentas.frx":1958
      FormatStyle(3)  =   "frmAdminconfigCuentas.frx":1A08
      FormatStyle(4)  =   "frmAdminconfigCuentas.frx":1ABC
      FormatStyle(5)  =   "frmAdminconfigCuentas.frx":1B94
      FormatStyle(6)  =   "frmAdminconfigCuentas.frx":1C4C
      ImageCount      =   0
      PrinterProperties=   "frmAdminconfigCuentas.frx":1D2C
   End
End
Attribute VB_Name = "frmAdminconfigCuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cuentas As New Collection
Private cuenta As CuentaBancaria
Private bancos As New Collection
Private Banco As Banco
Private monedas As New Collection
Private Moneda As clsMoneda

Private Sub Form_Load()
    FormHelper.Customize Me

    GridEXHelper.CustomizeGrid Me.grid, False, True
    GridEXHelper.CustomizeGrid Me.gridBancos, False, False
    GridEXHelper.CustomizeGrid Me.gridMonedas, False, False

    Set monedas = DAOMoneda.GetAll()
    Me.gridMonedas.ItemCount = monedas.count
    Set grid.Columns("moneda").DropDownControl = Me.gridMonedas

    Set bancos = DAOBancos.GetAll()
    Me.gridBancos.ItemCount = bancos.count
    Set grid.Columns("banco").DropDownControl = Me.gridBancos

    CargarCuentas
End Sub

Private Sub CargarCuentas()
    Set cuentas = DAOCuentaBancaria.FindAll()
    Me.grid.ItemCount = cuentas.count
End Sub

Private Sub grid_BeforeDelete(ByVal Cancel As GridEX20.JSRetBoolean)
    'Cancel = (MsgBox("�Desea eliminar la cuenta?", vbYesNo + vbQuestion) = vbNo)
End Sub



Private Sub grid_UnboundAddNew(ByVal NewRowBookmark As GridEX20.JSRetVariant, ByVal Values As GridEX20.JSRowData)
    Set cuenta = New CuentaBancaria
    With cuenta
        If IsNumeric(Values(1)) And Not IsEmpty(Values(1)) Then Set .Banco = DAOBancos.GetById(Values(1))
        .numero = Values(2)
        .TipoCuenta = Values(3)
        If IsNumeric(Values(4)) And Not IsEmpty(Values(4)) Then Set .Moneda = DAOMoneda.GetById(Values(4))
    End With
    If DAOCuentaBancaria.Save(cuenta) Then
        cuentas.Add cuenta, CStr(cuenta.id)
    Else
        MsgBox "Hubo un error al guardar los valores"
    End If

End Sub


Private Sub grid_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex > cuentas.count Then Exit Sub
    Set cuenta = cuentas.item(RowIndex)
    With cuenta
        Values(1) = .Banco.nombre
        Values(2) = .numero
        Values(3) = .TipoCuenta
        If IsSomething(.Moneda) Then Values(4) = .Moneda.NombreCorto
    End With
End Sub

Private Sub grid_UnboundUpdate(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex > cuentas.count Then Exit Sub
    Set cuenta = cuentas.item(RowIndex)
    With cuenta
        If IsNumeric(Values(1)) And Not IsEmpty(Values(1)) Then Set .Banco = DAOBancos.GetById(Values(1))
        .numero = Values(2)
        .TipoCuenta = Values(3)
        If IsNumeric(Values(4)) And Not IsEmpty(Values(4)) Then Set .Moneda = DAOMoneda.GetById(Values(4))
    End With
    If Not DAOCuentaBancaria.Save(cuenta) Then MsgBox "Hubo un error al guardar los valores"
End Sub

Private Sub gridBancos_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex > bancos.count Then Exit Sub
    Set Banco = bancos.item(RowIndex)
    With Banco
        Values(1) = .id
        Values(2) = .nombre
    End With
End Sub



Private Sub gridMonedas_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex > monedas.count Then Exit Sub
    Set Moneda = monedas.item(RowIndex)
    With Moneda
        Values(1) = .id
        Values(2) = .NombreCorto
    End With
End Sub