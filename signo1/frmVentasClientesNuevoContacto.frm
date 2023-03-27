VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmVentasClientesNuevoContacto 
   BackColor       =   &H00FF8080&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Contáctos relacionados..."
   ClientHeight    =   8880
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11010
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   11010
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   8280
      Width           =   975
   End
   Begin GridEX20.GridEX grilla 
      Height          =   3615
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   6376
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      MethodHoldFields=   -1  'True
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      BackColorHeader =   16761024
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   11
      Column(1)       =   "frmVentasClientesNuevoContacto.frx":0000
      Column(2)       =   "frmVentasClientesNuevoContacto.frx":00F4
      Column(3)       =   "frmVentasClientesNuevoContacto.frx":01C4
      Column(4)       =   "frmVentasClientesNuevoContacto.frx":0290
      Column(5)       =   "frmVentasClientesNuevoContacto.frx":0360
      Column(6)       =   "frmVentasClientesNuevoContacto.frx":0434
      Column(7)       =   "frmVentasClientesNuevoContacto.frx":0508
      Column(8)       =   "frmVentasClientesNuevoContacto.frx":05AC
      Column(9)       =   "frmVentasClientesNuevoContacto.frx":0650
      Column(10)      =   "frmVentasClientesNuevoContacto.frx":06F4
      Column(11)      =   "frmVentasClientesNuevoContacto.frx":0798
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmVentasClientesNuevoContacto.frx":083C
      FormatStyle(2)  =   "frmVentasClientesNuevoContacto.frx":0974
      FormatStyle(3)  =   "frmVentasClientesNuevoContacto.frx":0A24
      FormatStyle(4)  =   "frmVentasClientesNuevoContacto.frx":0AD8
      FormatStyle(5)  =   "frmVentasClientesNuevoContacto.frx":0BB0
      FormatStyle(6)  =   "frmVentasClientesNuevoContacto.frx":0C68
      ImageCount      =   0
      PrinterProperties=   "frmVentasClientesNuevoContacto.frx":0D48
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FF8080&
      Caption         =   "[ Nuevo contácto ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   120
      TabIndex        =   11
      Top             =   3720
      Width           =   10695
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Nuevo"
         Height          =   375
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   4560
         Width           =   975
      End
      Begin VB.TextBox txtEmail 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1440
         TabIndex        =   8
         Top             =   3240
         Width           =   9135
      End
      Begin VB.TextBox txtCelular 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         Top             =   1080
         Width           =   9135
      End
      Begin VB.TextBox txtCargo 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1440
         TabIndex        =   3
         Top             =   1440
         Width           =   9135
      End
      Begin VB.TextBox txtTelefono 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Top             =   720
         Width           =   9135
      End
      Begin VB.CommandButton s 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Guardar"
         Default         =   -1  'True
         Height          =   375
         Left            =   9480
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   4560
         Width           =   1095
      End
      Begin VB.TextBox txtDetalle 
         BackColor       =   &H00FFFFFF&
         Height          =   765
         Left            =   1440
         TabIndex        =   9
         Top             =   3600
         Width           =   9135
      End
      Begin VB.TextBox txtPaís 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1440
         TabIndex        =   7
         Top             =   2880
         Width           =   9135
      End
      Begin VB.TextBox txtProvincia 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1440
         TabIndex        =   6
         Top             =   2520
         Width           =   9135
      End
      Begin VB.TextBox txtLocalidad 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1440
         TabIndex        =   5
         Top             =   2160
         Width           =   9135
      End
      Begin VB.TextBox txtDireccion 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Top             =   1800
         Width           =   9135
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1440
         TabIndex        =   0
         Top             =   360
         Width           =   9135
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FF8080&
         Caption         =   "EMail"
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
         TabIndex        =   21
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FF8080&
         Caption         =   "Celular"
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
         TabIndex        =   20
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FF8080&
         Caption         =   "Cargo"
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
         TabIndex        =   19
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FF8080&
         Caption         =   "Teléfono"
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
         TabIndex        =   18
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label6 
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
         Left            =   120
         TabIndex        =   17
         Top             =   3600
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FF8080&
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
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FF8080&
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
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FF8080&
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
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FF8080&
         Caption         =   "Nombre"
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
         TabIndex        =   13
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FF8080&
         Caption         =   "Dirección"
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
         TabIndex        =   12
         Top             =   1800
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmVentasClientesNuevoContacto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rows As Long
Dim idPersona As Long
Dim rectemp As clsContacto
Private vCliente As clsCliente
Private vProveedor As clsProveedor
Dim vContacto As clsContacto
Dim contactos As Collection

Public Property Let cliente(nvalue As clsCliente)
    Set vCliente = nvalue
End Property

Public Property Let Proveedor(nvalue As clsProveedor)
    Set vProveedor = nvalue
End Property


Private Sub Command1_Click()
    If MsgBox("¿Está seguro de salir?", vbYesNo, "Confirmación") = vbYes Then
        Unload Me
    End If
End Sub


Private Sub Command3_Click()
    Me.txtNombre = Empty
    Me.txtTelefono = Empty
    Me.txtDireccion = Empty
    Me.txtLocalidad = Empty
    Me.txtProvincia = Empty
    Me.txtPaís = Empty
    Me.txtDetalle = Empty
    Me.txtEmail = Empty
    Me.txtCargo = Empty
    Me.txtCelular = Empty
    Set vContacto = Nothing

End Sub

Private Sub Form_Load()
    FormHelper.Customize Me
    GridEXHelper.CustomizeGrid Me.grilla
    Dim rs As New Recordset
    llenar_Grilla

    ''Me.caption = caption & " (" & Name & ")"

End Sub

Private Sub agregarContacto()

    Dim vConta As New clsContacto
    If vContacto Is Nothing Then

        If Trim(Me.txtNombre) = Empty Then
            error1 = True
        Else
            nombre = UCase(Trim(Me.txtNombre))
        End If
        vConta.nombre = nombre
        vConta.telefono = UCase(Trim(Me.txtTelefono))
        vConta.Domicilio = UCase(Trim(Me.txtDireccion))
        vConta.localidad = UCase(Trim(Me.txtLocalidad))
        vConta.provincia = UCase(Trim(Me.txtProvincia))
        vConta.pais = UCase(Trim(Me.txtPaís))
        vConta.detalle = UCase(Trim(Me.txtDetalle))
        vConta.email = UCase(Trim(Me.txtEmail))
        vConta.celular = UCase(Trim(Me.txtCelular))
        vConta.Cargo = UCase(Trim(Me.txtCargo))


        vConta.idPersona = idPersona

        If Not error1 Then
            If MsgBox("¿Está seguro de agregar el contácto?", vbYesNo, "Confirmación") = vbYes Then

                If DAOContacto.agregar(vConta) Then
                    MsgBox "Alta exitosa!", vbInformation, "Información"
                    contactos.Add vConta
                    llenar_Grilla
                    grilla.MoveLast
                Else
                    MsgBox "Se produjo algun error!", vbCritical, "Error"

                End If
            End If
        End If
    Else
        If Trim(Me.txtNombre) = Empty Then
            error1 = True
        Else
            nombre = UCase(Trim(Me.txtNombre))
        End If
        vContacto.nombre = nombre
        vContacto.telefono = UCase(Trim(Me.txtTelefono))
        vContacto.Domicilio = UCase(Trim(Me.txtDireccion))
        vContacto.localidad = UCase(Trim(Me.txtLocalidad))
        vContacto.provincia = UCase(Trim(Me.txtProvincia))
        vContacto.pais = UCase(Trim(Me.txtPaís))
        vContacto.detalle = UCase(Trim(Me.txtDetalle))
        vContacto.email = UCase(Trim(Me.txtEmail))
        vContacto.celular = UCase(Trim(Me.txtCelular))
        vContacto.Cargo = UCase(Trim(Me.txtCargo))
        vContacto.idPersona = vCliente.Id


        If DAOContacto.modificar(vContacto) Then
            MsgBox "Modificación exitosa!", vbInformation, "Información"
            grilla.RefreshRowIndex rows
        Else
            MsgBox "Se produjo algún error!", vbCritical, "Error"
        End If


    End If
End Sub



Private Sub llenar_Grilla()
    If Not vCliente Is Nothing Then
        Set contactos = DAOContacto.FindAll(cliente_, "idCliente = " & vCliente.Id)
        idPersona = vCliente.Id
    ElseIf Not vProveedor Is Nothing Then
        Set contactos = DAOContacto.FindAll(proveedor_, "idCliente = " & vProveedor.Id)
        idPersona = vProveedor.Id
    End If
    grilla.ItemCount = contactos.count
End Sub

Private Sub grilla_SelectionChange()
    If grilla.rowcount = 0 Then Exit Sub
    rows = grilla.RowIndex(grilla.row)
    mostrarContacto

End Sub

Private Sub grilla_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Set rectemp = contactos(RowIndex)
    With rectemp
        Values(1) = Format(.Id, "0000")
        Values(2) = UCase(.nombre)
        Values(3) = UCase(.Cargo)
        Values(4) = .celular
        Values(5) = .telefono
        Values(6) = .Domicilio
        Values(7) = .localidad
        Values(8) = .provincia
        Values(9) = .pais
        Values(10) = .email
    End With
End Sub
Private Sub mostrarContacto()
    Set vContacto = contactos(grilla.RowIndex(grilla.row))
    Me.txtNombre = vContacto.nombre
    Me.txtTelefono = vContacto.telefono
    Me.txtDireccion = vContacto.Domicilio
    Me.txtLocalidad = vContacto.localidad
    Me.txtProvincia = vContacto.provincia
    Me.txtPaís = vContacto.pais
    Me.txtDetalle = vContacto.detalle
    Me.txtEmail = vContacto.email
    Me.txtCargo = vContacto.Cargo
    Me.txtCelular = vContacto.celular

End Sub

Private Sub s_Click()
    agregarContacto
End Sub
