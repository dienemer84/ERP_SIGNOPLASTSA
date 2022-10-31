VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmVentasClienteNuevo 
   BackColor       =   &H00FF8080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Nuevo Cliente..."
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   7965
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.ComboBox cboProvincias 
      Height          =   315
      Left            =   1065
      TabIndex        =   25
      Top             =   1185
      Width           =   3675
      _Version        =   786432
      _ExtentX        =   6482
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Text            =   "ComboBox1"
   End
   Begin VB.CheckBox chkValido 
      Caption         =   "Válido para remitar y facturar"
      Height          =   225
      Left            =   1065
      TabIndex        =   21
      Top             =   4920
      Value           =   1  'Checked
      Width           =   3495
   End
   Begin VB.TextBox txtDetalleFP 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   4065
      TabIndex        =   19
      Top             =   4125
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   0
      Left            =   1080
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   6735
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   1
      Left            =   1080
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   480
      Width           =   6735
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   4
      Left            =   1050
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   2295
      Width           =   6735
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   5
      Left            =   1050
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   2655
      Width           =   6735
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   6
      Left            =   1050
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   3015
      Width           =   6735
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   7
      Left            =   1050
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   3375
      Width           =   6735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Guadar"
      Height          =   375
      Left            =   5670
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4845
      Width           =   975
   End
   Begin VB.ComboBox CboIVA 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1050
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   3735
      Width           =   6735
   End
   Begin VB.TextBox txtFP 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1065
      TabIndex        =   7
      Top             =   4125
      Width           =   1380
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   6750
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4845
      Width           =   975
   End
   Begin XtremeSuiteControls.ComboBox cboMonedas 
      Height          =   315
      Left            =   1065
      TabIndex        =   23
      Top             =   4470
      Width           =   1380
      _Version        =   786432
      _ExtentX        =   2434
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Style           =   2
      Appearance      =   6
      Text            =   "cboMoneda"
      DropDownItemCount=   3
   End
   Begin XtremeSuiteControls.ComboBox cboPaises 
      Height          =   315
      Left            =   1065
      TabIndex        =   27
      Top             =   825
      Width           =   3675
      _Version        =   786432
      _ExtentX        =   6482
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboLocalidades 
      Height          =   315
      Left            =   1065
      TabIndex        =   29
      Top             =   1545
      Width           =   3675
      _Version        =   786432
      _ExtentX        =   6482
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Sorted          =   -1  'True
      Text            =   "ComboBox1"
      AutoComplete    =   -1  'True
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "País"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   180
      TabIndex        =   28
      Top             =   885
      Width           =   855
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Localidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   150
      TabIndex        =   26
      Top             =   1590
      Width           =   855
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Provincia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   165
      TabIndex        =   24
      Top             =   1245
      Width           =   855
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      Caption         =   "Moneda"
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
      Left            =   150
      TabIndex        =   22
      Top             =   4500
      Width           =   855
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
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
      Left            =   3105
      TabIndex        =   20
      Top             =   4125
      Width           =   855
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FF8080&
      Caption         =   "Días"
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
      Left            =   2565
      TabIndex        =   18
      Top             =   4140
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      Caption         =   "Nombre "
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
      TabIndex        =   17
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      Caption         =   "Domicilio "
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
      TabIndex        =   16
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      Caption         =   "Teléfonos "
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
      Left            =   90
      TabIndex        =   15
      Top             =   2325
      Width           =   975
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      Caption         =   "Fax "
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
      Left            =   90
      TabIndex        =   14
      Top             =   2670
      Width           =   975
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      Caption         =   "E-Mail "
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
      Left            =   90
      TabIndex        =   13
      Top             =   3015
      Width           =   975
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      Caption         =   "CUIT "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   90
      TabIndex        =   12
      Top             =   3375
      Width           =   975
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      Caption         =   "IVA "
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
      Left            =   90
      TabIndex        =   11
      Top             =   3735
      Width           =   975
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      Caption         =   "F.Pago "
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
      Left            =   225
      TabIndex        =   10
      Top             =   4125
      Width           =   855
   End
End
Attribute VB_Name = "frmVentasClienteNuevo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vCliente As clsCliente
Dim strsql As String

Public Property Let cliente(nvalue As clsCliente)
    Set vCliente = nvalue
End Property
Private Sub Guardar()
    Dim Cuit
    Dim EVENTO As clsEventoObserver
    
    On Error GoTo err2
    
    razon = UCase(Text1(0))
    Domicilio = UCase(Text1(1))
    '    localidad = UCase(Text1(2))
    '    cp = UCase(Text1(3))
    telefono = UCase(Text1(4))
    Fax = UCase(Text1(5))
    email = UCase(Text1(6))
    ivan = Me.cboIVA.ItemData(Me.cboIVA.ListIndex)
    Cuit = Trim(Text1(7))

    FP = UCase(Me.txtFP)
    '    provincia = UCase(Text1(9))
    '    pais = UCase(Text1(10))
    FP_detalle = UCase(Me.txtDetalleFP)
    valido = Val(Me.chkValido.value)



    If MsgBox("¿Está conforme con los datos?", vbYesNo, "Confirmación") = vbYes Then
        If vCliente Is Nothing Then
            ErrorCode = 0

            If Not IsNumeric(Text1(7)) Then
                ErrorCode = 1
                errorCode2 = 1
            End If

                If ErrorCode > 0 Then
                    aa = "Debe introducir datos correctos para: "
                    If errorCode2 = 1 Then
                        aa = aa & Chr(10) & "CUIT"
                    End If
                    MsgBox aa, vbCritical, "Error"
                Else
            
                Dim cliente As New clsCliente

'31.10.22- SE AGREGA ESTA LINEA PARA QUE TOME EL VALOR DEL IVA
                Set cliente.TipoIVA = DAOTipoIva.GetById(ivan)

                cliente.Cuit = Cuit
                cliente.Domicilio = Domicilio
                cliente.email = email
                cliente.estado = EstadoCliente.activo
                cliente.Fax = Fax
                cliente.FP = FP


                cliente.PasswordSistema = 0
                
                Set cliente.provincia = DAOProvincias.FindById(Me.cboProvincias.ItemData(Me.cboProvincias.ListIndex))
                Set cliente.localidad = DAOLocalidades.FindById(Me.cboLocalidades.ItemData(Me.cboLocalidades.ListIndex))

                cliente.razon = razon
                cliente.FormaPago = FP_detalle
                cliente.telefono = telefono
                cliente.ValidoRemitoFactura = valido

                'cliente.Cuit = Cuil
                
                cliente.idMonedaDefault = Me.cboMonedas.ItemData(Me.cboMonedas.ListIndex)

                Dim F As String
                F = "c.cuit = " & Escape(Text1(7))
                
                If IsSomething(vCliente) Then
                F = F & " AND c.id <> " & vCliente.Id
                End If

               If DAOCliente.FindAll(F).count > 0 Then
                MsgBox "Ya existe un cliente con ese Nº de CUIT.", vbCritical, "Error"
                
                Else
                

                 
                If DAOCliente.crear(cliente) Then
                    MsgBox "Alta Exitosa!", vbInformation, "Información"
                       
                    Set EVENTO = New clsEventoObserver
                    Set EVENTO.Elemento = cliente
                    EVENTO.EVENTO = agregar_
                    Set EVENTO.Originador = Me
                    Channel.Notificar EVENTO, Clientes_
                    
                    Unload Me
                    
                    Else
                        MsgBox "Se produjo algún error, no se realizan cambios!", vbCritical, "Error"
                    End If
                
               
                
                End If




            End If
        Else
            'se modifica

            Set vCliente.TipoIVA = DAOTipoIva.GetById(ivan)

            vCliente.Cuit = Cuit
            vCliente.Domicilio = Domicilio
            vCliente.email = email
            vCliente.estado = EstadoCliente.activo
            vCliente.Fax = Fax
            vCliente.FP = FP


            vCliente.PasswordSistema = 0
            
            Set vCliente.provincia = DAOProvincias.FindById(Me.cboProvincias.ItemData(Me.cboProvincias.ListIndex))
            Set vCliente.localidad = DAOLocalidades.FindById(Me.cboLocalidades.ItemData(Me.cboLocalidades.ListIndex))
            
            vCliente.razon = razon
            vCliente.telefono = telefono
            vCliente.FormaPago = FP_detalle
            vCliente.ValidoRemitoFactura = valido
            vCliente.idMonedaDefault = Me.cboMonedas.ItemData(Me.cboMonedas.ListIndex)
            
            If DAOCliente.modificar(vCliente) Then
                MsgBox "Modificación Exitosa!", vbInformation, "Información"
                
                
                Set EVENTO = New clsEventoObserver
                Set EVENTO.Elemento = cliente
                EVENTO.EVENTO = modificar_
                Set EVENTO.Originador = Me
                Channel.Notificar EVENTO, Clientes_
                
                Unload Me
                
            Else
                MsgBox "Se produjo algún error, no se realizan cambios!", vbCritical, "Error"
            End If
            
        End If
        End If
    Exit Sub
err2:

End Sub

Private Sub cboPaises_Click()
    cboProvincias.Clear
    'cboLocalidades.Clear
    'cboPartidos.Clear

    Dim Id As Long
    If cboPaises.ListIndex >= 0 Then
        Id = Me.cboPaises.ItemData(Me.cboPaises.ListIndex)
        DAOProvincias.LlenarCombo Me.cboProvincias, Id
    End If

    cboProvincias_Click
End Sub




Private Sub cboProvincias_Click()
    Dim Id As Long
    If cboProvincias.ListIndex >= 0 Then
        Id = Me.cboProvincias.ItemData(Me.cboProvincias.ListIndex)
        DAOLocalidades.LlenarCombo Me.cboLocalidades, Id
    End If

End Sub

Private Sub Command1_Click()
    Guardar
End Sub

Private Sub Command2_Click()
    If MsgBox("¿Está seguro de salir?", vbYesNo, "Confirmación") = vbYes Then
        Unload Me
    End If

End Sub



Private Sub Form_Load()
    On Error Resume Next
    FormHelper.Customize Me
    For x = 0 To 10
        Text1(x) = Empty
    Next x
    DAOTipoIva.LlenarCombo Me.cboIVA
    Command1.caption = "Agregar"
    Me.caption = "Agregar Cliente..."
    DAOMoneda.llenarComboXtremeSuite Me.cboMonedas
    DAOPais.LlenarCombo Me.cboPaises
    If Not vCliente Is Nothing Then
        llenarForm
        Command1.caption = "Modificar"
        Me.caption = "Modificar Cliente..."
    End If


    'Me.caption = caption & "(" & Name & ")"


End Sub

Private Sub Form_Terminate()
    frmVentasClientesLista.llenar_Grilla
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmVentasClientesLista.llenar_Grilla
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    Text1(Index).SelStart = 0
    Text1(Index).SelLength = Len(Text1(Index))
End Sub
Private Sub txtFP_Validate(Cancel As Boolean)
    ValidarTextBox Me.txtFP, Cancel
End Sub
Private Sub llenarForm()
    On Error GoTo err1
    With vCliente
        Text1(0) = .razon
        Text1(1) = .Domicilio


        Text1(4) = .telefono
        Text1(5) = .Fax
        Text1(6) = .email
        Text1(7) = .Cuit



        'aca posiciono el combo

        Me.cboPaises.ListIndex = funciones.PosIndexCbo(.provincia.pais.Id, Me.cboPaises)
        Me.cboProvincias.ListIndex = funciones.PosIndexCbo(.provincia.Id, Me.cboProvincias)
        Me.cboLocalidades.ListIndex = funciones.PosIndexCbo(.localidad.Id, Me.cboLocalidades)


        Me.chkValido.value = Escape(.ValidoRemitoFactura)
        txtFP = .FP
        Me.txtDetalleFP = .FormaPago
        cboIVA.ListIndex = funciones.PosIndexCbo(.TipoIVA.idIVA, cboIVA)
        Me.cboMonedas.ListIndex = funciones.PosIndexCbo(vCliente.idMonedaDefault, Me.cboMonedas)
        
    End With

    Exit Sub
err1:
    Debug.Print Err.Description

End Sub
