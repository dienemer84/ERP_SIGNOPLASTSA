VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~3.OCX"
Begin VB.Form frmRubroProveedor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asignación de proveedores a rubro"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5205
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRubroProveedor.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   5205
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   4695
      Left            =   150
      TabIndex        =   2
      Top             =   660
      Width           =   4890
      _Version        =   786432
      _ExtentX        =   8625
      _ExtentY        =   8281
      _StockProps     =   79
      Caption         =   "Proveedores asignados al rubro"
      UseVisualStyle  =   -1  'True
      Begin GridEX20.GridEX gridProveedores 
         Height          =   4350
         Left            =   90
         TabIndex        =   3
         Top             =   225
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   7673
         Version         =   "2.0"
         HoldSortSettings=   -1  'True
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         MethodHoldFields=   -1  'True
         AllowColumnDrag =   0   'False
         AllowDelete     =   -1  'True
         AllowEdit       =   0   'False
         ColumnHeaders   =   0   'False
         RowHeaders      =   -1  'True
         DataMode        =   99
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   1
         Column(1)       =   "frmRubroProveedor.frx":000C
         SortKeysCount   =   1
         SortKey(1)      =   "frmRubroProveedor.frx":00B0
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmRubroProveedor.frx":0118
         FormatStyle(2)  =   "frmRubroProveedor.frx":0240
         FormatStyle(3)  =   "frmRubroProveedor.frx":02F0
         FormatStyle(4)  =   "frmRubroProveedor.frx":03A4
         FormatStyle(5)  =   "frmRubroProveedor.frx":047C
         FormatStyle(6)  =   "frmRubroProveedor.frx":0534
         ImageCount      =   0
         PrinterProperties=   "frmRubroProveedor.frx":0614
      End
   End
   Begin XtremeSuiteControls.ComboBox cboRubro 
      Height          =   315
      Left            =   690
      TabIndex        =   0
      Top             =   195
      Width           =   4320
      _Version        =   786432
      _ExtentX        =   7620
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Sorted          =   -1  'True
      Style           =   2
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboProveedores 
      Height          =   315
      Left            =   165
      TabIndex        =   4
      Top             =   5790
      Width           =   3855
      _Version        =   786432
      _ExtentX        =   6800
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Sorted          =   -1  'True
      Style           =   2
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.PushButton cmdAsignar 
      Height          =   690
      Left            =   4125
      TabIndex        =   6
      Top             =   5490
      Width           =   930
      _Version        =   786432
      _ExtentX        =   1640
      _ExtentY        =   1217
      _StockProps     =   79
      Caption         =   "Asignar proveedor al rubro"
      BackColor       =   15786449
      UseVisualStyle  =   -1  'True
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Proveedores disponibles para asignar al rubro"
      Height          =   195
      Left            =   180
      TabIndex        =   5
      Top             =   5520
      Width           =   3285
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Rubro"
      Height          =   195
      Left            =   165
      TabIndex        =   1
      Top             =   225
      Width           =   435
   End
End
Attribute VB_Name = "frmRubroProveedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private proveedoresRubro As Collection
Private P As clsProveedor

Private Sub cboRubro_Click()
    If Me.cboRubro.ListIndex <> -1 Then
        Me.gridProveedores.ItemCount = 0
        Set proveedoresRubro = DAOProveedor.FindAllByRubro(Me.cboRubro.ItemData(Me.cboRubro.ListIndex))
        Me.gridProveedores.ItemCount = proveedoresRubro.count
    End If
End Sub

Private Sub cmdAsignar_Click()
    If Me.cboProveedores.ListIndex = -1 Then Exit Sub

    If IsSomething(proveedoresRubro) Then
        If funciones.BuscarEnColeccion(proveedoresRubro, CStr(Me.cboProveedores.ItemData(Me.cboProveedores.ListIndex))) Then
            MsgBox "Ese proveedor ya se encuentra asignado al rubro.", vbExclamation
        Else
            Dim q As String
            q = "INSERT INTO asignacion VALUES (NULL, " & Me.cboProveedores.ItemData(Me.cboProveedores.ListIndex) & ", " & Me.cboRubro.ItemData(Me.cboRubro.ListIndex) & ")"
            If conectar.execute(q) Then
                proveedoresRubro.Add DAOProveedor.FindById(Me.cboProveedores.ItemData(Me.cboProveedores.ListIndex), False, False, True), CStr(Me.cboProveedores.ItemData(Me.cboProveedores.ListIndex))
                Me.gridProveedores.ItemCount = 0
                Me.gridProveedores.ItemCount = proveedoresRubro.count
            Else
                MsgBox "No se pudo asignar el proveedor al rubro.", vbCritical
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    Customize Me

    DAORubros.LlenarComboExtremeSuite Me.cboRubro
    Me.cboRubro.ListIndex = -1
    DAOProveedor.llenarComboXtremeSuite Me.cboProveedores, True, True
    GridEXHelper.CustomizeGrid Me.gridProveedores
    Me.gridProveedores.ItemCount = 0
End Sub

Private Sub gridProveedores_UnboundDelete(ByVal RowIndex As Long, ByVal Bookmark As Variant)
    If RowIndex > 0 And proveedoresRubro.count > 0 Then
        If conectar.execute("DELETE FROM asignacion WHERE id_proveedor = " & Me.cboProveedores.ItemData(Me.cboProveedores.ListIndex) & " AND id_rubro = " & Me.cboRubro.ItemData(Me.cboRubro.ListIndex)) Then
            proveedoresRubro.remove RowIndex
        End If
    End If
End Sub

Private Sub gridProveedores_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex > 0 And proveedoresRubro.count > 0 Then
        Set P = proveedoresRubro.item(RowIndex)
        Values(1) = P.RazonSocial
    End If
End Sub
