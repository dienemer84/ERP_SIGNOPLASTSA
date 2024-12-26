VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmDefinirConjunto 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Conjuntos..."
   ClientHeight    =   5685
   ClientLeft      =   7545
   ClientTop       =   5040
   ClientWidth     =   7155
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   7155
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.ComboBox cccccccccccccc 
      Height          =   315
      Left            =   960
      TabIndex        =   15
      Top             =   75
      Width           =   6135
      _Version        =   786432
      _ExtentX        =   10821
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Text            =   "ComboBox1"
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Agregar"
      Enabled         =   0   'False
      Height          =   255
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox txtCantidad 
      Height          =   285
      Left            =   960
      TabIndex        =   8
      Text            =   "1"
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox txtPieza 
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   7
      ToolTipText     =   "Doble click para seleccionar pieza"
      Top             =   1200
      Width           =   6135
   End
   Begin VB.TextBox lblDetalle 
      Height          =   285
      Left            =   960
      TabIndex        =   4
      Top             =   480
      Width           =   6135
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Quitar"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5160
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Default         =   -1  'True
      Height          =   375
      Left            =   5910
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5145
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Crear"
      Height          =   375
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5160
      Width           =   1215
   End
   Begin MSComctlLib.ListView lstDetalleConj 
      Height          =   3015
      Left            =   0
      TabIndex        =   10
      Top             =   2040
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   5318
      View            =   3
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      OLEDropMode     =   1
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      OLEDropMode     =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Cant"
         Object.Width           =   1323
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Detalle"
         Object.Width           =   10583
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "id"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   2760
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line1 
      X1              =   4440
      X2              =   7080
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Elegir elementos"
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
      Left            =   0
      TabIndex        =   14
      Top             =   840
      Width           =   7335
   End
   Begin VB.Label idPieza 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Label4"
      Height          =   255
      Left            =   2640
      TabIndex        =   13
      Top             =   1560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cantidad"
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
      Left            =   0
      TabIndex        =   12
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Pi 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Pieza"
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
      Left            =   0
      TabIndex        =   11
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cliente"
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
      TabIndex        =   6
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Detalle"
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
      Left            =   240
      TabIndex        =   5
      Top             =   480
      Width           =   615
   End
End
Attribute VB_Name = "frmDefinirConjunto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strsql As String
Dim rs As Recordset
Dim vIdPieza As Long
Dim vAccion As Integer
Dim vbusqueda As String
Dim claseS As New classStock
Public Property Let idPiezaMadre(idPieza As Long)
    vIdPieza = idPieza
End Property
Public Property Let accion(accion As Integer)
    vAccion = accion
End Property
Public Property Let busqueda(busqueda As Integer)
    vbusqueda = busqueda
End Property
Private Sub Command1_Click()
    If claseS.buscar_pieza(Trim(Me.lblDetalle)) > 0 Then
        MsgBox "El nombre asignado ya existe en la base de datos", vbCritical, "Error"
    Else
        If Me.lstDetalleConj.ListItems.count > 0 Then
            Dim h As VbMsgBoxResult
            h = MsgBox("¿Está conforme con los datos ingresados?", vbYesNo, "Confirmación")
            If h = 6 Then
                Dim idcli As Long
                idcli = Me.cccccccccccccc.ItemData(Me.cccccccccccccc.ListIndex)
                If claseS.definirConjunto2(UCase(Me.lblDetalle), idcli, Me.lstDetalleConj) Then
                    MsgBox "Conjunto creado satisfactoriamente", vbInformation, "Información"
                End If
            End If
        Else
            MsgBox "No puede crear un conjunto con un elemento o menos", vbCritical, "Error"
        End If
    End If
End Sub
Private Sub Command2_Click()
    Dim x As ListItem
    Set x = Me.lstDetalleConj.ListItems.Add(, , Trim(Me.txtCantidad))
    x.SubItems(1) = Me.txtPieza
    x.SubItems(2) = Me.idPieza
End Sub

Private Sub Command3_Click()
    If Me.lstDetalleConj.ListItems.count > 0 Then
        Dim h As VbMsgBoxResult
        h = MsgBox("¿Está seguro de modificar el conjunto?", vbYesNo, "Confirmación")
        If h = 6 Then
            Dim idcli As Long
            idcli = Me.cccccccccccccc.ItemData(Me.cccccccccccccc.ListIndex)
            If claseS.modificarConjunto(UCase(Me.lblDetalle), idcli, Me.lstDetalleConj, vIdPieza) Then
                MsgBox "Conjunto modificador satisfactoriamente", vbInformation, "Información"
                Unload Me
            End If
        End If
    Else
        MsgBox "No puede crear un conjunto con un elemento o menos", vbCritical, "Error"
    End If

End Sub

Private Sub Command4_Click()
    For i = Me.lstDetalleConj.ListItems.count To 1 Step -1
        If Me.lstDetalleConj.ListItems(i).Checked = True Then
            Me.lstDetalleConj.ListItems.remove (i)
        End If
    Next i
End Sub

Private Sub Command5_Click()
    If MsgBox("¿Está seguro de salir?", vbYesNo, "Confirmación") = vbYes Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    FormHelper.Customize Me
    validar    'valido txt en bco al ppio
    lblDetalle_Change    'valido txt en bco al principio


    DAOCliente.llenarComboXtremeSuite Me.cccccccccccccc

    Dim idCliente As Long
    On Error Resume Next
    If vAccion = 1 Then
        Me.Command1.Visible = False
        Me.Command3.Visible = True
    ElseIf vAccion = 0 Then
        Me.Command3.Visible = False
        Me.Command1.Visible = True
    End If
    'muest,ro los datos del conjunto
    If vAccion = 1 Then    ' solo si es una modificación
        strsql = "select detalle,id_cliente from stock where id=" & vIdPieza
        Set rs = conectar.RSFactory(strsql)
        detalle = rs!detalle
        idCliente = rs!id_cliente
        Me.lblDetalle = detalle
        Me.cccccccccccccc.ListIndex = funciones.PosIndexCbo(idCliente, Me.cccccccccccccc)
        strsql = "select s.detalle,sc.cantidad,sc.idPiezaHija from stockConjuntos sc inner join stock s on s.id=sc.idPiezaHija where sc.idPiezaPadre=" & vIdPieza
        Set rs = conectar.RSFactory(strsql)
        Me.lstDetalleConj.ListItems.Clear
        If Not rs.EOF And Not rs.BOF Then    'solo llena si hay datos
            While Not rs.EOF
                Set x = Me.lstDetalleConj.ListItems.Add(, , rs!Cantidad)
                x.SubItems(1) = rs!detalle
                x.SubItems(2) = rs!idPiezaHija
                rs.MoveNext
            Wend
        End If
    End If

End Sub

Private Sub lblDetalle_Change()
    If Trim(Me.lblDetalle) = Empty Then
        Me.Command1.Enabled = False
        Me.Command3.Enabled = False
    Else
        Me.Command1.Enabled = True
        Me.Command3.Enabled = True
    End If
End Sub


Private Sub lstDetalleConj_OLEDragDrop(data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim rs As Recordset
    idp = data.GetData(1)
    strsql = "select detalle from stock where id=" & idp
    Set rs = conectar.RSFactory(strsql)
    If Not rs.EOF And Not rs.BOF Then
        Me.txtCantidad = 1
        Me.txtPieza = rs!detalle
        Me.idPieza = idp
        Dim A As ListItem
        Set A = Me.lstDetalleConj.ListItems.Add(, , Trim(Me.txtCantidad))
        A.SubItems(1) = Me.txtPieza
        A.SubItems(2) = Me.idPieza
        A.Tag = ipd

    End If


End Sub

Private Sub txtCantidad_Change()
    validar
End Sub
Private Sub txtCantidad_Validate(Cancel As Boolean)
    If Not IsNumeric(Me.txtCantidad) Then Cancel = True
End Sub
Private Sub txtPieza_Change()
    validar
End Sub
Private Sub txtPieza_DblClick()
    frmListarStock_seleccion.Text1 = funciones.quePiezaElegidabusqueda

    frmListarStock_seleccion.Show 1
    Me.txtPieza = funciones.quePiezaElegidaDetalle
    Me.idPieza = funciones.quePiezaElegida
End Sub
Private Sub validar()
    If Trim(Me.txtCantidad) = Empty Or Trim(Me.txtPieza) = Empty Then
        Me.Command2.Enabled = False
    Else
        Me.Command2.Enabled = True
    End If
End Sub
