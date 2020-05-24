VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmAdminconfigCuentas 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cuentas Bancarias"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11235
   Icon            =   "frmAdminconfigCuentas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   11235
   ShowInTaskbar   =   0   'False
   Begin GridEX20.GridEX grid 
      Height          =   3915
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   11190
      _ExtentX        =   19738
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
      ColumnsCount    =   5
      Column(1)       =   "frmAdminconfigCuentas.frx":000C
      Column(2)       =   "frmAdminconfigCuentas.frx":013C
      Column(3)       =   "frmAdminconfigCuentas.frx":0228
      Column(4)       =   "frmAdminconfigCuentas.frx":0494
      Column(5)       =   "frmAdminconfigCuentas.frx":05A8
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminconfigCuentas.frx":068C
      FormatStyle(2)  =   "frmAdminconfigCuentas.frx":07C4
      FormatStyle(3)  =   "frmAdminconfigCuentas.frx":0874
      FormatStyle(4)  =   "frmAdminconfigCuentas.frx":0928
      FormatStyle(5)  =   "frmAdminconfigCuentas.frx":0A00
      FormatStyle(6)  =   "frmAdminconfigCuentas.frx":0AB8
      ImageCount      =   0
      PrinterProperties=   "frmAdminconfigCuentas.frx":0B98
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
      Column(1)       =   "frmAdminconfigCuentas.frx":0D70
      Column(2)       =   "frmAdminconfigCuentas.frx":0EB4
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminconfigCuentas.frx":0FC8
      FormatStyle(2)  =   "frmAdminconfigCuentas.frx":1100
      FormatStyle(3)  =   "frmAdminconfigCuentas.frx":11B0
      FormatStyle(4)  =   "frmAdminconfigCuentas.frx":1264
      FormatStyle(5)  =   "frmAdminconfigCuentas.frx":133C
      FormatStyle(6)  =   "frmAdminconfigCuentas.frx":13F4
      ImageCount      =   0
      PrinterProperties=   "frmAdminconfigCuentas.frx":14D4
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
      Column(1)       =   "frmAdminconfigCuentas.frx":16AC
      Column(2)       =   "frmAdminconfigCuentas.frx":17F0
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminconfigCuentas.frx":1904
      FormatStyle(2)  =   "frmAdminconfigCuentas.frx":1A3C
      FormatStyle(3)  =   "frmAdminconfigCuentas.frx":1AEC
      FormatStyle(4)  =   "frmAdminconfigCuentas.frx":1BA0
      FormatStyle(5)  =   "frmAdminconfigCuentas.frx":1C78
      FormatStyle(6)  =   "frmAdminconfigCuentas.frx":1D30
      ImageCount      =   0
      PrinterProperties=   "frmAdminconfigCuentas.frx":1E10
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
Private moneda As clsMoneda

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
    'Cancel = (MsgBox("¿Desea eliminar la cuenta?", vbYesNo + vbQuestion) = vbNo)
End Sub



Private Sub grid_UnboundAddNew(ByVal NewRowBookmark As GridEX20.JSRetVariant, ByVal Values As GridEX20.JSRowData)
    Set cuenta = New CuentaBancaria
    With cuenta
        If IsNumeric(Values(1)) And Not IsEmpty(Values(1)) Then Set .Banco = DAOBancos.GetById(Values(1))
        .numero = Values(2)
        .TipoCuenta = Values(3)
        If IsNumeric(Values(4)) And Not IsEmpty(Values(4)) Then Set .moneda = DAOMoneda.GetById(Values(4))
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
        If IsSomething(.moneda) Then Values(4) = .moneda.NombreCorto
        Values(5) = .CBU
    End With
End Sub

Private Sub grid_UnboundUpdate(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex > cuentas.count Then Exit Sub
    Set cuenta = cuentas.item(RowIndex)
    With cuenta
        If IsNumeric(Values(1)) And Not IsEmpty(Values(1)) Then Set .Banco = DAOBancos.GetById(Values(1))
        .numero = Values(2)
        .TipoCuenta = Values(3)
        .CBU = Values(5)
        If IsNumeric(Values(4)) And Not IsEmpty(Values(4)) Then Set .moneda = DAOMoneda.GetById(Values(4))
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
    Set moneda = monedas.item(RowIndex)
    With moneda
        Values(1) = .id
        Values(2) = .NombreCorto
    End With
End Sub
