VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmAgendaNuevaDetalles 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13155
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   13155
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.GroupBox GroupBox 
      Height          =   5895
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   12885
      _Version        =   786432
      _ExtentX        =   22728
      _ExtentY        =   10398
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin GridEX20.GridEX dgDatos 
         Height          =   5415
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   12555
         _ExtentX        =   22146
         _ExtentY        =   9551
         Version         =   "2.0"
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         ColumnAutoResize=   -1  'True
         MethodHoldFields=   -1  'True
         AllowDelete     =   -1  'True
         GroupByBoxVisible=   0   'False
         RowHeaders      =   -1  'True
         DataMode        =   99
         AllowAddNew     =   -1  'True
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   6
         Column(1)       =   "AdminFacturasEmisibles.frx":0000
         Column(2)       =   "AdminFacturasEmisibles.frx":0124
         Column(3)       =   "AdminFacturasEmisibles.frx":0218
         Column(4)       =   "AdminFacturasEmisibles.frx":0318
         Column(5)       =   "AdminFacturasEmisibles.frx":0418
         Column(6)       =   "AdminFacturasEmisibles.frx":0504
         FormatStylesCount=   6
         FormatStyle(1)  =   "AdminFacturasEmisibles.frx":05F4
         FormatStyle(2)  =   "AdminFacturasEmisibles.frx":072C
         FormatStyle(3)  =   "AdminFacturasEmisibles.frx":07DC
         FormatStyle(4)  =   "AdminFacturasEmisibles.frx":0890
         FormatStyle(5)  =   "AdminFacturasEmisibles.frx":0968
         FormatStyle(6)  =   "AdminFacturasEmisibles.frx":0A20
         ImageCount      =   0
         PrinterProperties=   "AdminFacturasEmisibles.frx":0B00
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox 
      Height          =   1695
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12885
      _Version        =   786432
      _ExtentX        =   22728
      _ExtentY        =   2990
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.PushButton PushButton 
         Height          =   375
         Index           =   3
         Left            =   9720
         TabIndex        =   16
         Top             =   960
         Width           =   375
         _Version        =   786432
         _ExtentX        =   661
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "X"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton 
         Height          =   375
         Index           =   2
         Left            =   9720
         TabIndex        =   15
         Top             =   360
         Width           =   375
         _Version        =   786432
         _ExtentX        =   661
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "X"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton 
         Height          =   375
         Index           =   1
         Left            =   4080
         TabIndex        =   14
         Top             =   960
         Width           =   375
         _Version        =   786432
         _ExtentX        =   661
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "X"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton 
         Height          =   375
         Index           =   0
         Left            =   4080
         TabIndex        =   13
         Top             =   360
         Width           =   375
         _Version        =   786432
         _ExtentX        =   661
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "X"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtEmail 
         Height          =   375
         Left            =   5640
         TabIndex        =   4
         Top             =   960
         Width           =   3975
         _Version        =   786432
         _ExtentX        =   7011
         _ExtentY        =   661
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XtremeSuiteControls.FlatEdit txtLocalidad 
         Height          =   375
         Left            =   5640
         TabIndex        =   3
         Top             =   360
         Width           =   3975
         _Version        =   786432
         _ExtentX        =   7011
         _ExtentY        =   661
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XtremeSuiteControls.PushButton btnGuardar 
         Height          =   495
         Left            =   10560
         TabIndex        =   5
         Top             =   840
         Width           =   2055
         _Version        =   786432
         _ExtentX        =   3625
         _ExtentY        =   873
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtDomicilio 
         Height          =   375
         Left            =   960
         TabIndex        =   2
         Top             =   960
         Width           =   3015
         _Version        =   786432
         _ExtentX        =   5318
         _ExtentY        =   661
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XtremeSuiteControls.FlatEdit txtEmpresa 
         Height          =   375
         Left            =   960
         TabIndex        =   1
         Top             =   360
         Width           =   3015
         _Version        =   786432
         _ExtentX        =   5318
         _ExtentY        =   661
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XtremeSuiteControls.Label Label 
         Height          =   375
         Index           =   3
         Left            =   4680
         TabIndex        =   12
         Top             =   960
         Width           =   855
         _Version        =   786432
         _ExtentX        =   1508
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Email"
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label Label 
         Height          =   375
         Index           =   2
         Left            =   4440
         TabIndex        =   11
         Top             =   360
         Width           =   1095
         _Version        =   786432
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Localidad"
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label Label 
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   735
         _Version        =   786432
         _ExtentX        =   1296
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Direccion"
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label Label 
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   615
         _Version        =   786432
         _ExtentX        =   1085
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Nombre"
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label lblCodigo 
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1815
         _Version        =   786432
         _ExtentX        =   3201
         _ExtentY        =   661
         _StockProps     =   79
      End
   End
End
Attribute VB_Name = "frmAgendaNuevaDetalles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vContacto As clsContactoPpal
Dim strsql As String
Public editar As Boolean
Public Usable As Boolean
Dim tmp As clsContactoPpalDetalle


Public Property Let contacto(nvalue As clsContactoPpal)
    Set vContacto = nvalue
End Property


Private Sub btnGuardar_Click()
    On Error GoTo err1
    If MsgBox("¿Seguro de guardar los cambios?", vbYesNo + vbQuestion, "Confirmación") = vbYes Then
    
        If vContacto Is Nothing Then
            Set vContacto = New clsContactoPpal
        Else
            vContacto.Id = vContacto.Id
        End If
        
        vContacto.Empresa = UCase(Me.txtEmpresa.Text)
        vContacto.localidad = UCase(Me.txtLocalidad.Text)
        vContacto.Email = Me.txtEmail.Text
        vContacto.direccion = UCase(Me.txtDomicilio.Text)
       
        
        If Not DAOContactoPpal.Save(vContacto, True, False) Then
            MsgBox "Se produjo algun error al guardar!", vbCritical, "Error"
        Else

            MsgBox "Guardado correctamente!", vbInformation, "Información"

            Me.dgDatos.ReBind

        End If
        
    End If
    
    Exit Sub
err1:
End Sub

Private Sub dgDatos_AfterUpdate()
    If Not noadd Then
        Me.dgDatos.ItemCount = vContacto.Detalles.count
    End If
End Sub


Private Sub dgDatos_BeforeDelete(ByVal Cancel As GridEX20.JSRetBoolean)
    Cancel = MsgBox("¿Está seguro de eliminar el detalle?", vbYesNo + vbInformation, "Confirmación") = vbNo

End Sub


Private Sub dgDatos_BeforeUpdate(ByVal Cancel As GridEX20.JSRetBoolean)
'    Cancel = (Me.dgDatos.value(5) < 0 Or Not IsNumeric(Me.dgDatos.value(5)))
End Sub


Private Sub dgDatos_DblClick()
'    On Error Resume Next
'    dgDatos_SelectionChange
'    Dim pos As Long
'    If Usable Then
'        Set Selecciones.RemitoElegido = vContactoPpal
'        Unload Me
'    End If
'
'    If editar Then
'        pos = Me.dgDatos.RowIndex(dgDatos.row)
'        If Remito.CantidadDeLineasActuales > funciones.itemsPorRemito Then
'            MsgBox "La cantidad de líneas superan a lo permitido"
'        Else
'
'        End If
'
'    End If
'
'    Me.dgDatos.RefreshRowIndex pos

End Sub


Private Sub dgDatos_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'    If Button = 2 And IsSomething(tmp) Then
''        Me.mnuNoFacturable.Enabled = Not tmp.Facturado And Remito.estado = EstadoRemito.RemitoAprobado And (Remito.EstadoFacturado = RemitoNoFacturado Or Remito.EstadoFacturado = RemitoFacturadoParcial)
''
''        If tmp.facturable Then
''            Me.mnuNoFacturable.caption = "Hacer No Facturable"
''        Else
''            Me.mnuNoFacturable.caption = "Hacer Facturable"
''        End If
''        Me.PopupMenu Me.mnuDetalleRemito
'    End If
End Sub


Private Sub dgDatos_RowFormat(RowBuffer As GridEX20.JSRowData)
'    On Error Resume Next
'    'xxxx
'    If RowBuffer.RowIndex > 0 And Remito.Detalles.count > 0 Then
'        Set tmp = Remito.Detalles(RowBuffer.RowIndex)
''        If tmp.facturable Then
''            If Not tmp.Facturado Then
''                RowBuffer.CellStyle(6) = "NoFacturado"
''            Else
''                RowBuffer.CellStyle(6) = "Facturado"
''            End If
''        Else
''            RowBuffer.CellStyle(6) = "NoFacturable"
''        End If
'    End If
    
End Sub


Private Sub dgDatos_SelectionChange()
'    Dim it As Long
'    it = Me.dgDatos.RowIndex(dgDatos.row)
'    If it > 0 And Remito.Detalles.count > 0 Then
'        Set tmp = Remito.Detalles.item(it)
'
'        If tmp.Origen = OrigenRemitoConcepto Then
'            grilla.Columns(2).EditType = jgexEditTextBox
'            grilla.Columns(4).EditType = jgexEditTextBox
'        Else
'            grilla.Columns(2).EditType = jgexEditNone
'            grilla.Columns(4).EditType = jgexEditTextBox
'        End If
'
'        If (Not tmp.facturable Or tmp.Facturado) Or (tmp.Origen <> OrigenRemitoConcepto And Not valorizable) Then
'            grilla.Columns(5).EditType = jgexEditNone
'        Else
'            grilla.Columns(5).EditType = jgexEditTextBox
'        End If
'    Else
'        grilla.Columns(2).EditType = jgexEditTextBox
'        grilla.Columns(4).EditType = jgexEditTextBox
'        grilla.Columns(5).EditType = jgexEditTextBox
'    End If
End Sub


Private Sub dgDatos_UnboundAddNew(ByVal NewRowBookmark As GridEX20.JSRetVariant, ByVal Values As GridEX20.JSRowData)

        Set tmp = New clsContactoPpalDetalle

        tmp.detalle = UCase(Values(2))
        tmp.Telefono1 = Values(3)
        tmp.Telefono2 = Values(4)
        tmp.Mail = Values(5)
        tmp.Mas = Values(6)
        
        vContacto.Detalles.Add tmp

End Sub


Private Sub dgDatos_UnboundDelete(ByVal RowIndex As Long, ByVal Bookmark As Variant)
    If RowIndex > 0 And vContacto.Detalles.count > 0 Then
        Set tmp = vContacto.Detalles(RowIndex)
        If 1 = 1 Then
            If DAOContactoPpalDetalles.Delete(tmp) Then
                vContacto.Detalles.remove RowIndex
            Else
                MsgBox "Se produjo algún error!", vbCritical
            End If
        Else
            vContacto.Detalles.remove RowIndex
        End If
    End If
    
End Sub


Private Sub dgDatos_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    On Error Resume Next
    If RowIndex > 0 And vContacto.Detalles.count > 0 Then
    Debug.Print (vContacto.Detalles.count)
        Set tmp = vContacto.Detalles(RowIndex)

        With Values
            .value(1) = tmp.Id
            .value(2) = tmp.detalle
            .value(3) = tmp.Telefono1
            .value(4) = tmp.Telefono2
            .value(5) = tmp.Mail
            .value(6) = tmp.Mas
            
        End With
    End If
End Sub


Private Sub dgDatos_UnboundUpdate(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex > 0 And vContacto.Detalles.count > 0 Then
        Set tmp = vContacto.Detalles.item(RowIndex)
        tmp.detalle = UCase(Values(2))
        tmp.Telefono1 = Values(3)
        tmp.Telefono2 = Values(4)
        tmp.Mail = Values(5)
        tmp.Mas = Values(6)

    End If
End Sub

Private Sub Form_Load()

    FormHelper.Customize Me
    
    GridEXHelper.CustomizeGrid Me.dgDatos, False, True
    
    Me.btnGuardar.caption = "Agregar"
    Me.caption = "Agregar Contacto..."

    If Not vContacto Is Nothing Then

        Me.btnGuardar.caption = "Modificar"
        Me.caption = "Modificar Contacto..."
        Me.txtEmpresa.Text = vContacto.Empresa
        Me.txtDomicilio.Text = vContacto.direccion
        Me.txtLocalidad.Text = vContacto.localidad
        Me.txtEmail.Text = vContacto.Email
        
        Me.caption = "Modificar Contacto: " & vContacto.Id & " Nombre: " & vContacto.Empresa
        
        Me.dgDatos.AllowEdit = True
        Me.dgDatos.AllowAddNew = True
        
        llenarForm
        
        llenarLista
          
    End If

End Sub


Private Sub llenarLista()

    If IsSomething(vContacto) Then
        
        Set vContacto.Detalles = DAOContactoPpalDetalles.FindAllByContactoPpal(vContacto.Id)

    End If

    Me.dgDatos.ItemCount = 0
    Me.dgDatos.ItemCount = vContacto.Detalles.count
    
End Sub


Private Sub llenarForm()
    On Error GoTo err1
    With vContacto
        Me.txtEmpresa = .Empresa
        Me.txtDomicilio = .direccion
        Me.txtLocalidad = .localidad
        Me.txtEmail = .Email
            

    End With

    Exit Sub
err1:
    Debug.Print Err.Description

End Sub

Private Sub PushButton_Click(Index As Integer)
    If Index = 0 Then
        Me.txtEmpresa.Text = ""
    ElseIf Index = 1 Then
        Me.txtDomicilio.Text = ""
    ElseIf Index = 2 Then
        Me.txtLocalidad.Text = ""
    ElseIf Index = 3 Then
        Me.txtEmail.Text = ""
    End If
End Sub
