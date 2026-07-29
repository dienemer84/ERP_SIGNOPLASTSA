VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmVentasClienteNuevo 
   BackColor       =   &H00FF8080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Nuevo Cliente..."
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   7785
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.PushButton cmdConsultarARCA 
      Height          =   345
      Left            =   4320
      TabIndex        =   37
      Top             =   90
      Width           =   2775
      _Version        =   786432
      _ExtentX        =   4895
      _ExtentY        =   609
      _StockProps     =   79
      Caption         =   "Cargar desde ARCA"
      UseVisualStyle  =   -1  'True
   End
   Begin VB.TextBox txtIDImpositivo 
      Height          =   285
      Left            =   1560
      TabIndex        =   34
      Top             =   3120
      Width           =   3735
   End
   Begin VB.TextBox txtCuitPais 
      Height          =   285
      Left            =   1560
      TabIndex        =   33
      Top             =   2760
      Width           =   3735
   End
   Begin VB.TextBox txtCP 
      Height          =   285
      Left            =   1560
      TabIndex        =   5
      Top             =   2280
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
      Top             =   1560
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
      Top             =   480
      Width           =   5535
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   1
      Left            =   1560
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   840
      Width           =   5535
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   4
      Left            =   1560
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   3615
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   5
      Left            =   1560
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   3960
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   6
      Left            =   1560
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   4335
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
      Top             =   135
      Width           =   2655
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
      Top             =   1185
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
      Top             =   1905
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
      Top             =   3120
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
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Código Postal"
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
      Top             =   2295
      Width           =   1455
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
      Left            =   480
      TabIndex        =   30
      Top             =   1230
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
      Top             =   1950
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
      Top             =   1590
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
      Height          =   195
      Left            =   3000
      TabIndex        =   25
      Top             =   5160
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
      Top             =   495
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
      Top             =   855
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
      Left            =   360
      TabIndex        =   22
      Top             =   3645
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
      Top             =   3990
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
      Top             =   4350
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
      Left            =   480
      TabIndex        =   19
      Top             =   150
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
Private mCargandoUbicacion As Boolean
Private mFormularioFacturaOrigen As frmAdminFacturasEdicion

Public Property Let Cliente(nValue As clsCliente)
    Set vCliente = nValue
End Property

Public Property Set FormularioFacturaOrigen( _
    ByVal valor As frmAdminFacturasEdicion _
)

    Set mFormularioFacturaOrigen = valor

End Property


Private Sub Guardar()

    Dim cuit As String
    Dim EVENTO As clsEventoObserver
    Dim razon As String, Domicilio As String, telefono As String, Fax As String, email As String
    Dim FP As String, FP_detalle As String
    Dim ivan As Long
    Dim valido As Long
    Dim pais As Long
    Dim IDImpositivo As String, CuitPais As String, CodigoPOS As String
    Dim ErrorCode As Long, errorCode2 As Long
    Dim aa As String
    Dim F As String

    On Error GoTo err2

    '=========================
    ' Toma datos de pantalla
    '=========================
    razon = UCase$(Trim$(Text1(0)))
    Domicilio = UCase$(Trim$(Text1(1)))
    telefono = UCase$(Trim$(Text1(4)))
    Fax = UCase$(Trim$(Text1(5)))
    email = UCase$(Trim$(Text1(6)))

    ivan = Me.CboIVA.ItemData(Me.CboIVA.ListIndex)
    
    cuit = NormalizarCuit(Text1(7))

    FP = UCase$(Trim$(Me.txtFP))
    FP_detalle = UCase$(Trim$(Me.txtDetalleFP))
    valido = val(Me.chkValido.value)

    pais = Me.cboPaises.ItemData(Me.cboPaises.ListIndex)

    IDImpositivo = UCase$(Trim$(Me.txtIDImpositivo))
    CuitPais = UCase$(Trim$(Me.txtCuitPais))
    CodigoPOS = UCase$(Trim$(txtCP))

    '=========================
    ' Validaciones básicas
    '=========================
    If razon = "" Then
        MsgBox "Debe introducir Razón Social.", vbCritical, "Error"
        Exit Sub
    End If

    If Domicilio = "" Then
        MsgBox "Debe introducir Domicilio.", vbCritical, "Error"
        Exit Sub
    End If

    If CodigoPOS = "" Then
        MsgBox "Debe introducir el código postal.", vbCritical, "Error"
        Exit Sub
    End If
    
    

    'Confirmación
    If MsgBox("¿Está conforme con los datos?", vbYesNo + vbQuestion, "Confirmación") <> vbYes Then
        Exit Sub
    End If

    '=========================
    ' Validaciones (CUIT)
    '=========================

    ErrorCode = 0
    errorCode2 = 0
    
    If Not CuitTieneFormatoBasico(cuit) Then
        ErrorCode = 1
        errorCode2 = 1
    End If
    
    If ErrorCode > 0 Then
        aa = "Debe introducir datos correctos para: "
        If errorCode2 = 1 Then aa = aa & vbCrLf & "CUIT (11 dígitos, sin letras)"
        MsgBox aa, vbCritical, "Error"
        Exit Sub
    End If

'Si querés, podés reescribir el textbox ya normalizado:
Text1(7) = cuit

    '=========================
    ' ALTA
    '=========================
    If vCliente Is Nothing Then

        Dim Cliente As clsCliente
        Set Cliente = New clsCliente

        Set Cliente.TipoIVA = DAOTipoIva.GetById(ivan)

        Cliente.cuit = cuit
        Cliente.Domicilio = Domicilio
        Cliente.email = email
        Cliente.estado = EstadoCliente.activo
        Cliente.Fax = Fax
        Cliente.PasswordSistema = 0
        Cliente.razon = razon
        Cliente.FormaPago = FP_detalle
        Cliente.telefono = telefono
        Cliente.ValidoRemitoFactura = valido
        Cliente.idMonedaDefault = _
            Me.cboMonedas.ItemData(Me.cboMonedas.ListIndex)

        Cliente.CodigoPostal = CodigoPOS
        Cliente.FP = FP
        Cliente.IDImpositivo = IDImpositivo
        Cliente.CuitPais = CuitPais

        Set Cliente.provincia = _
            DAOProvincias.FindById( _
                Me.cboProvincias.ItemData( _
                    Me.cboProvincias.ListIndex _
                ) _
            )

        Set Cliente.localidad = _
            DAOLocalidades.FindById( _
                Me.cboLocalidades.ItemData( _
                    Me.cboLocalidades.ListIndex _
                ) _
            )

        Set Cliente.pais = _
            DAOPais.FindById( _
                Me.cboPaises.ItemData( _
                    Me.cboPaises.ListIndex _
                ) _
            )

        '==========================================
        ' Control de CUIT duplicado para Argentina
        '==========================================
        F = "c.cuit = " & Escape(cuit)

        If Cliente.pais.Id = 1 Then

            If DAOCliente.FindAll(F).count > 0 Then

                MsgBox _
                    "Ya existe un cliente con ese Nº de CUIT.", _
                    vbCritical, _
                    "Error"

                Exit Sub

            End If

        End If

        '==========================================
        ' Guardar cliente nuevo
        '==========================================
        If Not DAOCliente.crear(Cliente) Then

            MsgBox _
                "Se produjo algún error, no se realizan cambios.", _
                vbCritical, _
                "Error"

            Exit Sub

        End If

        'Actualizar directamente el formulario de factura.
        If Not mFormularioFacturaOrigen Is Nothing Then

            mFormularioFacturaOrigen.SeleccionarClienteCreado _
                Cliente.Id

        End If

        'Mantener la notificación para otros formularios.
        Set EVENTO = New clsEventoObserver
        Set EVENTO.Elemento = Cliente
        EVENTO.EVENTO = agregar_
        Set EVENTO.Originador = Me

        Channel.Notificar EVENTO, Clientes_

        MsgBox _
            "Alta Exitosa!", _
            vbInformation, _
            "Información"

        Unload Me

        'Muy importante:
        'Unload no detiene automáticamente el procedimiento.
        Exit Sub

    Else

        '=========================
        ' MODIFICAR
        '=========================
        Set vCliente.TipoIVA = DAOTipoIva.GetById(ivan)

        vCliente.cuit = cuit
        vCliente.Domicilio = Domicilio
        vCliente.email = email
        vCliente.estado = EstadoCliente.activo
        vCliente.Fax = Fax
        vCliente.FP = FP
        vCliente.PasswordSistema = 0
        vCliente.razon = razon
        vCliente.telefono = telefono
        vCliente.FormaPago = FP_detalle
        vCliente.ValidoRemitoFactura = valido

        vCliente.idMonedaDefault = _
            Me.cboMonedas.ItemData(Me.cboMonedas.ListIndex)

        vCliente.CuitPais = CuitPais
        vCliente.IDImpositivo = IDImpositivo
        vCliente.CodigoPostal = CodigoPOS

        Set vCliente.provincia = _
            DAOProvincias.FindById( _
                Me.cboProvincias.ItemData( _
                    Me.cboProvincias.ListIndex _
                ) _
            )

        Set vCliente.localidad = _
            DAOLocalidades.FindById( _
                Me.cboLocalidades.ItemData( _
                    Me.cboLocalidades.ListIndex _
                ) _
            )

        Set vCliente.pais = _
            DAOPais.FindById( _
                Me.cboPaises.ItemData( _
                    Me.cboPaises.ListIndex _
                ) _
            )

        '==========================================
        ' Control de CUIT duplicado al modificar
        '==========================================
        F = "c.cuit = " & Escape(cuit)
        F = F & " AND c.id <> " & CStr(vCliente.Id)

        If vCliente.pais.Id = 1 Then

            If DAOCliente.FindAll(F).count > 0 Then

                MsgBox _
                    "Ya existe otro cliente con ese Nº de CUIT.", _
                    vbCritical, _
                    "Error"

                Exit Sub

            End If

        End If

        '==========================================
        ' Guardar modificación
        '==========================================
        If Not DAOCliente.modificar(vCliente) Then

            MsgBox _
                "Se produjo algún error, no se realizan cambios.", _
                vbCritical, _
                "Error"

            Exit Sub

        End If

        Set EVENTO = New clsEventoObserver
        Set EVENTO.Elemento = vCliente
        EVENTO.EVENTO = modificar_
        Set EVENTO.Originador = Me

        Channel.Notificar EVENTO, Clientes_

        MsgBox _
            "Modificación Exitosa!", _
            vbInformation, _
            "Información"

        Unload Me
        Exit Sub

    End If

    Exit Sub

err2:

    MsgBox _
        "Error al guardar: " & _
        CStr(Err.Number) & " - " & Err.Description, _
        vbCritical, _
        "Error"

End Sub


Private Sub btnGuardar_Click()
    Guardar
End Sub

Private Sub btnSalir_Click()
    If MsgBox("¿Está seguro de salir?", vbYesNo, "Confirmación") = vbYes Then
        Unload Me
    End If

End Sub





Private Sub cboPaises_Click()
    If mCargandoUbicacion Then
        Exit Sub
    End If

    On Error GoTo ManejarError

    CargarProvinciasPaisSeleccionado

    Exit Sub

ManejarError:

    MsgBox _
        "Error al cargar las provincias." & vbCrLf & vbCrLf & _
        "Origen: " & Err.Source & vbCrLf & _
        "Error " & CStr(Err.Number) & ": " & Err.Description, _
        vbCritical, _
        "Clientes"
End Sub


Private Sub cboProvincias_Click()
    If mCargandoUbicacion Then
        Exit Sub
    End If

    On Error GoTo ManejarError

    CargarLocalidadesProvinciaSeleccionada

    Exit Sub

ManejarError:

    MsgBox _
        "Error al cargar las localidades." & vbCrLf & vbCrLf & _
        "Origen: " & Err.Source & vbCrLf & _
        "Error " & CStr(Err.Number) & ": " & Err.Description, _
        vbCritical, _
        "Clientes"

End Sub

Private Sub Command1_Click()
    Guardar
End Sub

Private Sub Command2_Click()
    If MsgBox("¿Está seguro de salir?", vbYesNo, "Confirmación") = vbYes Then
        Unload Me
    End If

End Sub



Private Sub cmdConsultarARCA_Click()

    On Error GoTo ManejarError

    Dim consulta As clsConsultaARCA
    Dim cuitIngresado As String
    Dim respuesta As VbMsgBoxResult
    Dim provinciaEncontrada As Boolean
    Dim localidadEncontrada As Boolean
    Dim condicionIVAEncontrada As Boolean

    cuitIngresado = NormalizarCuit(Text1(7).Text)

    If Len(Trim$(cuitIngresado)) = 0 Then
        MsgBox "Debe ingresar primero el CUIT que desea consultar.", _
               vbExclamation, _
               "Consulta ARCA"

        Text1(7).SetFocus
        Exit Sub
    End If

    If Len(Trim$(Text1(0).Text)) > 0 Or _
       Len(Trim$(Text1(1).Text)) > 0 Then

        respuesta = MsgBox( _
            "La consulta reemplazará la razón social, el domicilio, " & _
            "el código postal, la provincia, la localidad y su condición frente al IVA." & _
            vbCrLf & vbCrLf & _
            "¿Desea continuar?", _
            vbYesNo + vbQuestion, _
            "Consultar ARCA" _
        )

        If respuesta <> vbYes Then
            Exit Sub
        End If

    End If

    cmdConsultarARCA.Enabled = False
    cmdConsultarARCA.caption = "Consultando..."

    Screen.MousePointer = vbHourglass
    DoEvents

    Set consulta = New clsConsultaARCA

    If Not consulta.Consultar(cuitIngresado) Then
    
        Dim mensajeErrorARCA As String
    
        mensajeErrorARCA = consulta.UltimoError
    
        LimpiarFormularioPorFalloARCA
    
        MsgBox _
            mensajeErrorARCA & _
            vbCrLf & vbCrLf & _
            "No se obtuvieron datos desde ARCA." & _
            vbCrLf & _
            "Los campos del formulario fueron reseteados.", _
            vbExclamation, _
            "Consulta ARCA"
    
        GoTo salir
    
    End If

    '=========================================
    ' Datos generales devueltos por ARCA
    '=========================================

    Text1(7).Text = consulta.cuit
    Text1(0).Text = UCase$(Trim$(consulta.RazonSocial))
    Text1(1).Text = UCase$(Trim$(consulta.direccion))
    txtCP.Text = UCase$(Trim$(consulta.CodigoPostal))
    condicionIVAEncontrada = _
        SeleccionarCondicionIVAARCA(consulta.CondicionIVA)
        
    'Si ARCA no devolvió la dirección por separado,
    'utilizamos el domicilio completo.
    If Len(Trim$(Text1(1).Text)) = 0 Then
        Text1(1).Text = UCase$(Trim$(consulta.Domicilio))
    End If

    '=========================================
    ' País, provincia y localidad
    '=========================================

    If Not SeleccionarPaisArgentina() Then
    
        MsgBox _
            "No se pudo seleccionar Argentina en el combo de países.", _
            vbExclamation, _
            "Consulta ARCA"
    
        GoTo salir
    
    End If
    
    provinciaEncontrada = SeleccionarProvinciaARCA( _
                                consulta.provincia _
                            )

    If provinciaEncontrada Then

        localidadEncontrada = SeleccionarLocalidadARCA( _
                                    consulta.localidad _
                                )

    Else

        localidadEncontrada = False

    End If

    '=========================================
    ' Resultado
    '=========================================

    If Not provinciaEncontrada Then

        MsgBox _
            "ARCA devolvió correctamente los datos del cliente," & _
            vbCrLf & _
            "pero no se encontró la provincia en el combo:" & _
            vbCrLf & vbCrLf & _
            consulta.provincia & _
            vbCrLf & vbCrLf & _
            "Seleccione la provincia y la localidad manualmente.", _
            vbExclamation, _
            "Consulta ARCA"

    ElseIf Not localidadEncontrada Then

        MsgBox _
            "ARCA devolvió correctamente los datos del cliente," & _
            vbCrLf & _
            "pero no se encontró la localidad en el combo:" & _
            vbCrLf & vbCrLf & _
            consulta.localidad & _
            vbCrLf & vbCrLf & _
            "Seleccione la localidad manualmente.", _
            vbExclamation, _
            "Consulta ARCA"
            
    ElseIf Not condicionIVAEncontrada Then
    
        MsgBox _
            "ARCA devolvió la condición frente al IVA:" & _
            vbCrLf & vbCrLf & _
            consulta.CondicionIVA & _
            vbCrLf & vbCrLf & _
            "pero no existe una opción equivalente en el combo IVA." & _
            vbCrLf & _
            "Seleccione la condición manualmente.", _
            vbExclamation, _
            "Consulta ARCA"

    Else

        MsgBox _
            "Los datos del cliente se obtuvieron correctamente desde ARCA." & _
            vbCrLf & vbCrLf & _
            "Condición IVA: " & consulta.CondicionIVA, _
            vbInformation, _
            "Consulta ARCA"

    End If

salir:

    Screen.MousePointer = vbDefault

    cmdConsultarARCA.Enabled = True
    cmdConsultarARCA.caption = "Consultar ARCA"

    Set consulta = Nothing
    Exit Sub

ManejarError:

    numeroError = Err.Number
    descripcionError = Err.Description

    Screen.MousePointer = vbDefault

    cmdConsultarARCA.Enabled = True
    cmdConsultarARCA.caption = "Consultar ARCA"

    LimpiarFormularioPorFalloARCA

    MsgBox _
        "Error al consultar ARCA: " & _
        CStr(numeroError) & " - " & descripcionError & _
        vbCrLf & vbCrLf & _
        "No se obtuvieron datos desde ARCA." & _
        vbCrLf & _
        "Los campos del formulario fueron limpiados.", _
        vbCritical, _
        "Consulta ARCA"

    Set consulta = Nothing

End Sub



Private Sub Form_Load()

    Dim paso As String
    Dim posicionArgentina As Long

    On Error GoTo ManejarError

    mCargandoUbicacion = True

    paso = "Personalizar formulario"

    On Error Resume Next
    FormHelper.Customize Me
    Err.Clear
    On Error GoTo ManejarError

    paso = "Limpiar controles"

    Text1(0) = ""
    Text1(1) = ""
    Text1(4) = ""
    Text1(5) = ""
    Text1(6) = ""
    Text1(7) = ""

    paso = "Cargar condición de IVA"
    DAOTipoIva.LlenarCombo Me.CboIVA

    paso = "Cargar monedas"
    DAOMoneda.llenarComboXtremeSuite Me.cboMonedas

    paso = "Cargar países"

    'Mientras se llena el combo no permitimos que su evento
    'Click cargue provincias automáticamente.
    DAOPais.LlenarCombo Me.cboPaises

    If vCliente Is Nothing Then

'''        Command1.caption = "Agregar"
        Me.caption = "Agregar Cliente..."

        txtCuitPais = "-"
        txtIDImpositivo = "-"

        Text1(4) = "-"
        Text1(5) = "-"
        Text1(6) = "-"

        txtFP = "0"
        txtDetalleFP = "-"

        paso = "Seleccionar Argentina"

        posicionArgentina = BuscarPosicionPorItemData( _
                                Me.cboPaises, _
                                1 _
                            )

        If posicionArgentina >= 0 Then
            Me.cboPaises.ListIndex = posicionArgentina
        End If

        'Terminó la carga inicial.
        mCargandoUbicacion = False

        paso = "Cargar provincias de Argentina"
        CargarProvinciasPaisSeleccionado

    Else

        'llenarForm necesita que los eventos funcionen para
        'cargar provincia y localidad del cliente existente.
        mCargandoUbicacion = False

        paso = "Cargar cliente existente"
        llenarForm

        Command1.caption = "Modificar"
        Me.caption = "Modificar Cliente..."

    End If

    Exit Sub

ManejarError:

    mCargandoUbicacion = False

    MsgBox _
        "Error al cargar el formulario de clientes." & _
        vbCrLf & vbCrLf & _
        "Paso: " & paso & vbCrLf & _
        "Origen: " & Err.Source & vbCrLf & _
        "Error " & CStr(Err.Number) & ": " & _
        Err.Description, _
        vbCritical, _
        "Clientes"

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
        Text1(7) = .cuit
        Me.txtCP = .CodigoPostal
        Me.txtCuitPais.Text = .CuitPais
        Me.txtIDImpositivo.Text = .IDImpositivo

        'aca posiciono el combo
        Me.cboPaises.ListIndex = funciones.PosIndexCbo(.provincia.pais.Id, Me.cboPaises)
        Me.cboProvincias.ListIndex = funciones.PosIndexCbo(.provincia.Id, Me.cboProvincias)
        Me.cboLocalidades.ListIndex = funciones.PosIndexCbo(.localidad.Id, Me.cboLocalidades)

        Me.chkValido.value = Escape(.ValidoRemitoFactura)
        Me.txtFP = .FP
        Me.txtDetalleFP = .FormaPago
        Me.CboIVA.ListIndex = funciones.PosIndexCbo(.TipoIVA.idIVA, CboIVA)
        Me.cboMonedas.ListIndex = funciones.PosIndexCbo(vCliente.idMonedaDefault, Me.cboMonedas)
    End With

    Exit Sub
err1:
    Debug.Print Err.Description
End Sub

Private Function SeleccionarPaisArgentina() As Boolean

    Dim posicion As Long

    On Error GoTo ManejarError

    SeleccionarPaisArgentina = False

    posicion = BuscarPosicionPorItemData( _
                    Me.cboPaises, _
                    1 _
                )

    If posicion < 0 Then
        Exit Function
    End If

    'Impedimos que ListIndex dispare automáticamente
    'cboPaises_Click.
    mCargandoUbicacion = True
    Me.cboPaises.ListIndex = posicion
    mCargandoUbicacion = False

    'Cargamos las provincias de forma controlada.
    CargarProvinciasPaisSeleccionado

    SeleccionarPaisArgentina = True
    Exit Function

ManejarError:

    mCargandoUbicacion = False

    Err.Raise _
        Err.Number, _
        "SeleccionarPaisArgentina", _
        Err.Description

End Function

Private Function SeleccionarProvinciaARCA( _
    ByVal nombreProvincia As String _
) As Boolean

    Dim posicion As Long

    SeleccionarProvinciaARCA = False

    If Len(Trim$(nombreProvincia)) = 0 Then
        Exit Function
    End If

    posicion = BuscarTextoEnCombo( _
                    cboProvincias, _
                    nombreProvincia _
                )

    If posicion < 0 Then

        'Equivalencias frecuentes para CABA.
        Select Case NormalizarTextoComparacion(nombreProvincia)

            Case "CIUDAD AUTONOMA BUENOS AIRES", _
                 "CIUDAD AUTONOMA DE BUENOS AIRES", _
                 "CAPITAL FEDERAL", _
                 "CABA"

                posicion = BuscarTextoEnCombo( _
                                cboProvincias, _
                                "CAPITAL FEDERAL" _
                            )

                If posicion < 0 Then

                    posicion = BuscarTextoEnCombo( _
                                    cboProvincias, _
                                    "CIUDAD AUTONOMA" _
                                )

                End If

        End Select

    End If

    If posicion < 0 Then
        Exit Function
    End If

    mCargandoUbicacion = True
    cboProvincias.ListIndex = posicion
    mCargandoUbicacion = False
    
    'Carga las localidades una sola vez y de manera controlada.
    CargarLocalidadesProvinciaSeleccionada
    
    SeleccionarProvinciaARCA = True

End Function

Private Function SeleccionarLocalidadARCA( _
    ByVal nombreLocalidad As String _
) As Boolean

    Dim posicion As Long

    SeleccionarLocalidadARCA = False

    If Len(Trim$(nombreLocalidad)) = 0 Then
        Exit Function
    End If

    posicion = BuscarTextoEnCombo( _
                    cboLocalidades, _
                    nombreLocalidad _
                )

    If posicion < 0 Then
        Exit Function
    End If

    cboLocalidades.ListIndex = posicion

    SeleccionarLocalidadARCA = True

End Function

Private Function BuscarTextoEnCombo( _
    ByVal combo As Object, _
    ByVal textoBuscado As String _
) As Long

    Dim i As Long
    Dim textoARCA As String
    Dim TextoCombo As String

    BuscarTextoEnCombo = -1

    textoARCA = NormalizarTextoComparacion(textoBuscado)

    If Len(textoARCA) = 0 Then
        Exit Function
    End If

    'Primero intentamos una coincidencia exacta.
    For i = 0 To combo.ListCount - 1

        TextoCombo = NormalizarTextoComparacion( _
                            CStr(combo.list(i)) _
                        )

        If TextoCombo = textoARCA Then

            BuscarTextoEnCombo = i
            Exit Function

        End If

    Next i

    'Después intentamos coincidencia parcial.
    For i = 0 To combo.ListCount - 1

        TextoCombo = NormalizarTextoComparacion( _
                            CStr(combo.list(i)) _
                        )

        If InStr(1, TextoCombo, textoARCA, vbTextCompare) > 0 Or _
           InStr(1, textoARCA, TextoCombo, vbTextCompare) > 0 Then

            BuscarTextoEnCombo = i
            Exit Function

        End If

    Next i

End Function

Private Function NormalizarTextoComparacion( _
    ByVal valor As String _
) As String

    Dim resultado As String

    resultado = UCase$(Trim$(valor))

    resultado = Replace(resultado, "Á", "A")
    resultado = Replace(resultado, "É", "E")
    resultado = Replace(resultado, "Í", "I")
    resultado = Replace(resultado, "Ó", "O")
    resultado = Replace(resultado, "Ú", "U")
    resultado = Replace(resultado, "Ü", "U")

    resultado = Replace(resultado, ".", " ")
    resultado = Replace(resultado, ",", " ")
    resultado = Replace(resultado, "-", " ")
    resultado = Replace(resultado, "_", " ")

    Do While InStr(resultado, "  ") > 0
        resultado = Replace(resultado, "  ", " ")
    Loop

    NormalizarTextoComparacion = Trim$(resultado)

End Function


Private Function BuscarPosicionPorItemData( _
    ByVal combo As Object, _
    ByVal idBuscado As Long _
) As Long

    Dim i As Long

    BuscarPosicionPorItemData = -1

    For i = 0 To combo.ListCount - 1

        If CLng(combo.ItemData(i)) = idBuscado Then

            BuscarPosicionPorItemData = i
            Exit Function

        End If

    Next i

End Function

Private Sub CargarProvinciasPaisSeleccionado()

    Dim idPais As Long
    Dim numeroError As Long
    Dim descripcionError As String

    On Error GoTo ManejarError

    Me.cboProvincias.Clear
    Me.cboLocalidades.Clear

    If Me.cboPaises.ListIndex < 0 Then
        Exit Sub
    End If

    idPais = CLng( _
                Me.cboPaises.ItemData( _
                    Me.cboPaises.ListIndex _
                ) _
             )

    'DAOProvincias puede establecer ListIndex = 0.
    'Mientras carga, impedimos que se dispare nuevamente
    'cboProvincias_Click.
    mCargandoUbicacion = True

    DAOProvincias.LlenarCombo _
        Me.cboProvincias, _
        idPais

    mCargandoUbicacion = False

    'Una vez finalizada la carga, cargamos las localidades
    'de manera controlada y una sola vez.
    If Me.cboProvincias.ListIndex >= 0 Then
        CargarLocalidadesProvinciaSeleccionada
    End If

    Exit Sub

ManejarError:

    numeroError = Err.Number
    descripcionError = Err.Description

    mCargandoUbicacion = False

    Err.Raise _
        numeroError, _
        "CargarProvinciasPaisSeleccionado", _
        descripcionError

End Sub


Private Sub CargarLocalidadesProvinciaSeleccionada()

    Dim idProvincia As Long
    Dim numeroError As Long
    Dim descripcionError As String

    On Error GoTo ManejarError

    Me.cboLocalidades.Clear

    If Me.cboProvincias.ListIndex < 0 Then
        Exit Sub
    End If

    idProvincia = CLng( _
                    Me.cboProvincias.ItemData( _
                        Me.cboProvincias.ListIndex _
                    ) _
                  )

    mCargandoUbicacion = True

    DAOLocalidades.LlenarCombo _
        Me.cboLocalidades, _
        idProvincia

    mCargandoUbicacion = False

    Exit Sub

ManejarError:

    numeroError = Err.Number
    descripcionError = Err.Description

    mCargandoUbicacion = False

    Err.Raise _
        numeroError, _
        "CargarLocalidadesProvinciaSeleccionada", _
        descripcionError

End Sub

Private Function SeleccionarCondicionIVAARCA( _
    ByVal condicionARCA As String _
) As Boolean

    Dim condicionNormalizada As String
    Dim TextoCombo As String
    Dim posicion As Long

    SeleccionarCondicionIVAARCA = False

    condicionNormalizada = _
        NormalizarTextoComparacion(condicionARCA)

    Select Case condicionNormalizada

        Case "MONOTRIBUTO"

            TextoCombo = "Monotributo"

        Case "EXENTO", _
             "IVA EXENTO"

            TextoCombo = "Exento"

        Case "RESP INSCRIPTO", _
             "RESPONSABLE INSCRIPTO", _
             "IVA RESPONSABLE INSCRIPTO"

            TextoCombo = "Resp. Inscripto"

        Case "SIN DATOS", ""

            TextoCombo = "Sin Datos"

        Case Else

            'Intentamos buscar exactamente lo enviado por ARCA.
            TextoCombo = condicionARCA

    End Select

    posicion = BuscarTextoEnCombo( _
                    Me.CboIVA, _
                    TextoCombo _
                )

    If posicion < 0 Then
        Exit Function
    End If

    Me.CboIVA.ListIndex = posicion

    SeleccionarCondicionIVAARCA = True

End Function

Private Sub LimpiarFormularioPorFalloARCA()

    On Error GoTo salir

    'Si estamos modificando un cliente existente,
    'restauramos sus datos guardados para no perderlos.
    If Not vCliente Is Nothing Then
        llenarForm
        Exit Sub
    End If

    mCargandoUbicacion = True

    'Datos generales.
    Text1(0).Text = ""       'Razón social
    Text1(1).Text = ""       'Domicilio
    Text1(4).Text = ""       'Teléfono
    Text1(5).Text = ""       'Fax
    Text1(6).Text = ""       'E-mail
    Text1(7).Text = ""       'CUIT

    txtCP.Text = ""
    txtCuitPais.Text = ""
    txtIDImpositivo.Text = ""

    txtFP.Text = ""
    txtDetalleFP.Text = ""

    'Ubicación.
    cboProvincias.Clear
    cboLocalidades.Clear

    If cboPaises.ListCount > 0 Then
        cboPaises.ListIndex = -1
    End If

    'Condición frente al IVA.
    If CboIVA.ListCount > 0 Then
        CboIVA.ListIndex = -1
    End If

    'Moneda.
    If cboMonedas.ListCount > 0 Then
        cboMonedas.ListIndex = -1
    End If

    chkValido.value = 0

salir:

    mCargandoUbicacion = False

End Sub
