VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmVentasClienteNuevo 
   BackColor       =   &H00FF8080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Nuevo Cliente..."
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   7830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtIDImpositivo 
      Height          =   285
      Left            =   1560
      TabIndex        =   34
      Top             =   2760
      Width           =   3735
   End
   Begin VB.TextBox txtCuitPais 
      Height          =   285
      Left            =   1560
      TabIndex        =   33
      Top             =   2400
      Width           =   3735
   End
   Begin VB.TextBox txtCP 
      Height          =   285
      Left            =   1560
      TabIndex        =   5
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   31
      Top             =   6960
      Width           =   7575
      Begin XtremeSuiteControls.PushButton btnGuardar 
         Height          =   495
         Left            =   5880
         TabIndex        =   15
         Top             =   240
         Width           =   1455
         _Version        =   786432
         _ExtentX        =   2566
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Guardar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnSalir 
         Height          =   495
         Left            =   240
         TabIndex        =   16
         Top             =   240
         Width           =   1335
         _Version        =   786432
         _ExtentX        =   2355
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Salir"
         UseVisualStyle  =   -1  'True
      End
   End
   Begin XtremeSuiteControls.ComboBox cboProvincias 
      Height          =   315
      Left            =   1560
      TabIndex        =   3
      Top             =   1200
      Width           =   3675
      _Version        =   786432
      _ExtentX        =   6482
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Text            =   "ComboBox1"
   End
   Begin VB.CheckBox chkValido 
      Caption         =   "V�lido para remitar y facturar"
      Height          =   225
      Left            =   4200
      TabIndex        =   14
      Top             =   6600
      Value           =   1  'Checked
      Width           =   3495
   End
   Begin VB.TextBox txtDetalleFP 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1560
      TabIndex        =   13
      Top             =   5880
      Width           =   5535
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   0
      Left            =   1560
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   5535
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   1
      Left            =   1560
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   480
      Width           =   5535
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   4
      Left            =   1560
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   3255
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   5
      Left            =   1560
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   3600
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   6
      Left            =   1560
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   3975
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   7
      Left            =   1560
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   4335
      Width           =   4935
   End
   Begin VB.ComboBox CboIVA 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   4695
      Width           =   4935
   End
   Begin VB.TextBox txtFP 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1560
      TabIndex        =   11
      Top             =   5085
      Width           =   1380
   End
   Begin XtremeSuiteControls.ComboBox cboMonedas 
      Height          =   315
      Left            =   1560
      TabIndex        =   12
      Top             =   5430
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
      Left            =   1560
      TabIndex        =   2
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
      Left            =   1560
      TabIndex        =   4
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
   Begin VB.Label lblIDImpositivo 
      Alignment       =   1  'Right Justify
      Caption         =   "ID Impositivo"
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
      Left            =   0
      TabIndex        =   36
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label lblCuitPais 
      Alignment       =   1  'Right Justify
      Caption         =   "Cuit Pais"
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
      TabIndex        =   35
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "C�digo Postal"
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
      Left            =   -120
      TabIndex        =   32
      Top             =   1935
      Width           =   1455
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Pa�s"
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
      Left            =   480
      TabIndex        =   30
      Top             =   870
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
      Left            =   480
      TabIndex        =   29
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
      Left            =   480
      TabIndex        =   28
      Top             =   1230
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
      Left            =   480
      TabIndex        =   27
      Top             =   5460
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
      Left            =   600
      TabIndex        =   26
      Top             =   5895
      Width           =   735
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FF8080&
      Caption         =   "D�as"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3120
      TabIndex        =   25
      Top             =   4440
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
      Left            =   480
      TabIndex        =   24
      Top             =   135
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
      Left            =   480
      TabIndex        =   23
      Top             =   495
      Width           =   975
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      Caption         =   "Tel�fonos "
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
      Left            =   360
      TabIndex        =   22
      Top             =   3285
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
      Left            =   360
      TabIndex        =   21
      Top             =   3630
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
      Left            =   360
      TabIndex        =   20
      Top             =   3990
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
      Height          =   255
      Left            =   360
      TabIndex        =   19
      Top             =   4350
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
      Left            =   360
      TabIndex        =   18
      Top             =   4725
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
      Left            =   480
      TabIndex        =   17
      Top             =   5100
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

Public Property Let Cliente(nvalue As clsCliente)
    Set vCliente = nvalue
End Property


Private Sub Guardar()
    Dim Cuit
    Dim EVENTO As clsEventoObserver

    On Error GoTo err2

    razon = UCase(Text1(0))
    Domicilio = UCase(Text1(1))
    telefono = UCase(Text1(4))
    Fax = UCase(Text1(5))
    Email = UCase(Text1(6))
    ivan = Me.cboIVA.ItemData(Me.cboIVA.ListIndex)
    Cuit = Trim(Text1(7))
    FP = UCase(Me.txtFP)
    FP_detalle = UCase(Me.txtDetalleFP)
    valido = Val(Me.chkValido.value)
    
    IDImpositivo = UCase(Me.txtIDImpositivo)
    CuitPais = UCase(Me.txtCuitPais)

    CodigoPOS = UCase(txtCP)
    
    If razon = "" Then
        MsgBox "Debe introducir Raz�n Social.", vbCritical, "Error"
        Exit Sub
    End If
    
    If Domicilio = "" Then
        MsgBox "Debe introducir Domicilio.", vbCritical, "Error"
        Exit Sub
    End If
    
    If CodigoPOS = "" Then
        MsgBox "Debe introducir el c�digo postal.", vbCritical, "Error"
        Exit Sub
    End If

    
    If MsgBox("�Est� conforme con los datos?", vbYesNo, "Confirmaci�n") = vbYes Then
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

                Dim Cliente As New clsCliente

                '31.10.22- SE AGREGA ESTA LINEA PARA QUE TOME EL VALOR DEL IVA
                Set Cliente.TipoIVA = DAOTipoIva.GetById(ivan)

                Cliente.Cuit = Cuit
                Cliente.Domicilio = Domicilio
                Cliente.Email = Email
                Cliente.estado = EstadoCliente.activo
                Cliente.Fax = Fax

                Cliente.PasswordSistema = 0
                Cliente.razon = razon
                Cliente.FormaPago = FP_detalle
                Cliente.telefono = telefono
                Cliente.ValidoRemitoFactura = valido
                Cliente.idMonedaDefault = Me.cboMonedas.ItemData(Me.cboMonedas.ListIndex)
                Cliente.CodigoPostal = CodigoPOS
                
                                Cliente.FP = FP
                                
                Cliente.IDImpositivo = IDImpositivo
                Cliente.CuitPais = CuitPais
                
                Set Cliente.provincia = DAOProvincias.FindById(Me.cboProvincias.ItemData(Me.cboProvincias.ListIndex))
                Set Cliente.localidad = DAOLocalidades.FindById(Me.cboLocalidades.ItemData(Me.cboLocalidades.ListIndex))

                Dim F As String
                F = "c.cuit = " & Escape(Text1(7))

                If IsSomething(vCliente) Then
                    F = F & " AND c.id <> " & vCliente.Id
                End If

                If DAOCliente.FindAll(F).count > 0 Then
                    MsgBox "Ya existe un cliente con ese N� de CUIT.", vbCritical, "Error"
                Else

                    If DAOCliente.crear(Cliente) Then
                        MsgBox "Alta Exitosa!", vbInformation, "Informaci�n"

                        Set EVENTO = New clsEventoObserver
                        Set EVENTO.Elemento = Cliente
                        EVENTO.EVENTO = agregar_
                        Set EVENTO.Originador = Me
                        Channel.Notificar EVENTO, Clientes_

                        Unload Me
                        
                    Else
                        MsgBox "Se produjo alg�n error, no se realizan cambios!", vbCritical, "Error"
                    End If
                End If
            End If
        Else
            'se modifica

            Set vCliente.TipoIVA = DAOTipoIva.GetById(ivan)

            vCliente.Cuit = Cuit
            vCliente.Domicilio = Domicilio
            vCliente.Email = Email
            vCliente.estado = EstadoCliente.activo
            vCliente.Fax = Fax
            vCliente.FP = FP
            vCliente.PasswordSistema = 0
            vCliente.razon = razon
            vCliente.telefono = telefono
            vCliente.FormaPago = FP_detalle
            vCliente.ValidoRemitoFactura = valido
            vCliente.idMonedaDefault = Me.cboMonedas.ItemData(Me.cboMonedas.ListIndex)
            
            vCliente.CuitPais = CuitPais
            vCliente.IDImpositivo = IDImpositivo
            
            vCliente.CodigoPostal = CodigoPOS

            Set vCliente.provincia = DAOProvincias.FindById(Me.cboProvincias.ItemData(Me.cboProvincias.ListIndex))
            Set vCliente.localidad = DAOLocalidades.FindById(Me.cboLocalidades.ItemData(Me.cboLocalidades.ListIndex))


            If DAOCliente.modificar(vCliente) Then
                MsgBox "Modificaci�n Exitosa!", vbInformation, "Informaci�n"


                Set EVENTO = New clsEventoObserver
                Set EVENTO.Elemento = Cliente
                EVENTO.EVENTO = modificar_
                Set EVENTO.Originador = Me
                Channel.Notificar EVENTO, Clientes_

                Unload Me

            Else
                MsgBox "Se produjo alg�n error, no se realizan cambios!", vbCritical, "Error"
            End If

        End If
    End If
    Exit Sub
err2:

End Sub

Private Sub btnGuardar_Click()
    Guardar
End Sub

Private Sub btnSalir_Click()
    If MsgBox("�Est� seguro de salir?", vbYesNo, "Confirmaci�n") = vbYes Then
        Unload Me
    End If

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
    If MsgBox("�Est� seguro de salir?", vbYesNo, "Confirmaci�n") = vbYes Then
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
    
    Me.txtCuitPais.Text = "-"
    Me.txtIDImpositivo = "-"
    
    Me.Text1(4) = "-"
    Me.Text1(5) = "-"
    Me.Text1(6) = "-"
    Me.txtFP = "0"
    Me.txtDetalleFP = "-"

    ''Me.caption = caption & "(" & Name & ")"


End Sub

Private Sub Form_Terminate()
'    frmVentasClientesLista.llenar_Grilla
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    frmVentasClientesLista.llenar_Grilla
    
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
        Text1(6) = .Email
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
