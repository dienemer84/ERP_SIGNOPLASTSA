VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmComprasProveedoresNuevoContacto 
   BackColor       =   &H00FF8080&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Contáctos relacionados..."
   ClientHeight    =   7365
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   10725
   ShowInTaskbar   =   0   'False
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
      Height          =   4815
      Left            =   0
      TabIndex        =   13
      Top             =   2520
      Width           =   10695
      Begin VB.TextBox txtEmail 
         Height          =   285
         Left            =   1440
         TabIndex        =   8
         Top             =   3240
         Width           =   9135
      End
      Begin VB.TextBox txtCelular 
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         Top             =   1080
         Width           =   9135
      End
      Begin VB.TextBox txtCargo 
         Height          =   285
         Left            =   1440
         TabIndex        =   3
         Top             =   1440
         Width           =   9135
      End
      Begin VB.TextBox txtTelefono 
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Top             =   720
         Width           =   9135
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Cancel          =   -1  'True
         Caption         =   "Salir"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   4320
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Agregar"
         Height          =   255
         Left            =   9600
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   4440
         Width           =   975
      End
      Begin VB.TextBox txtDetalle 
         Height          =   765
         Left            =   1440
         TabIndex        =   9
         Top             =   3600
         Width           =   9135
      End
      Begin VB.TextBox txtPaís 
         Height          =   285
         Left            =   1440
         TabIndex        =   7
         Top             =   2880
         Width           =   9135
      End
      Begin VB.TextBox txtProvincia 
         Height          =   285
         Left            =   1440
         TabIndex        =   6
         Top             =   2520
         Width           =   9135
      End
      Begin VB.TextBox txtLocalidad 
         Height          =   285
         Left            =   1440
         TabIndex        =   5
         Top             =   2160
         Width           =   9135
      End
      Begin VB.TextBox txtDireccion 
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Top             =   1800
         Width           =   9135
      End
      Begin VB.TextBox txtNombre 
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
         TabIndex        =   25
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
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   21
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   14
         Top             =   1800
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "[ Cliente ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   10695
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2055
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   3625
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Lista de contáctos"
         ColumnCount     =   10
         BeginProperty Column00 
            DataField       =   "nombre"
            Caption         =   "Nombre"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "tel"
            Caption         =   "Telefono"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "celular"
            Caption         =   "Celular"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "cargo"
            Caption         =   "Cargo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "Direccion"
            Caption         =   "Dirección"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "localidad"
            Caption         =   "Localidad"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "provincia"
            Caption         =   "Provincia"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "país"
            Caption         =   "País"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "detalle"
            Caption         =   "Detalles"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column09 
            DataField       =   "email"
            Caption         =   "E-Mail"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            RecordSelectors =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
            EndProperty
            BeginProperty Column04 
            EndProperty
            BeginProperty Column05 
            EndProperty
            BeginProperty Column06 
            EndProperty
            BeginProperty Column07 
            EndProperty
            BeginProperty Column08 
            EndProperty
            BeginProperty Column09 
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lstContactos 
         Height          =   2175
         Left            =   240
         TabIndex        =   20
         Top             =   2880
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   3836
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nombre"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Telefono"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Dirección"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Localidad"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Provincia"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "País"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Detalle"
            Object.Width           =   6174
         EndProperty
      End
   End
End
Attribute VB_Name = "frmComprasProveedoresNuevoContacto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim classP As New classAdministracion


Dim vIdproveedor As Long
Dim vidModificar As Long
Public Property Let idProveedor(nIdProveedor As Long)
    vIdproveedor = nIdProveedor
End Property

Private Sub Command1_Click()
    If MsgBox("¿Está seguro de salir?", vbYesNo, "Confirmación") = vbYes Then
        Unload Me
    End If
End Sub

Private Sub Command2_Click()
    agregarContacto
End Sub

Private Sub Form_Load()
    FormHelper.Customize Me
    Dim rs As New Recordset
    Set rs = conectar.RSFactory("select razon from proveedores where id=" & vIdproveedor)
    If Not rs.EOF And Not rs.BOF Then
        Me.Frame1.caption = "[ " & UCase(rs!razon) & " ]"
    End If
    verContactos
End Sub

Private Sub verContactos()
'Dim X As ListItem
    Dim rs As Recordset
    'Me.lstContactos.ListItems.Clear
    Dim strsql As String
    strsql = "select cargo,email,celular,tel,direccion,localidad,provincia,país, detalle as deta, nombre from contactos where idCliente=" & vIdproveedor & " and tipo=2"
    Set rs = conectar.RSFactory(strsql)
    Set Me.DataGrid1.DataSource = rs

End Sub



Private Sub agregarContacto()
    Dim strsql As String
    If Trim(Me.txtNombre) = Empty Then
        error1 = True
    Else
        nombre = LCase(Trim(Me.txtNombre))
    End If
    tel = LCase(Trim(Me.txtTelefono))
    direccion = LCase(Trim(Me.txtDireccion))
    localidad = LCase(Trim(Me.txtLocalidad))
    provincia = LCase(Trim(Me.txtProvincia))
    pais = LCase(Trim(Me.txtPaís))
    detalle = LCase(Trim(Me.txtDetalle))
    email = LCase(Trim(Me.txtEmail))
    celular = LCase(Trim(Me.txtCelular))
    Cargo = LCase(Trim(Me.txtCargo))






    If Not error1 Then
        If MsgBox("¿Está seguro de agregar el contácto?", vbYesNo, "Confirmación") = vbYes Then
            strsql = "insert into contactos (idCliente, Nombre, Tel, Direccion, Localidad, Provincia, País, Detalle,cargo,celular,email,tipo) values( " & vIdproveedor & ",'" & nombre & "','" & tel & "','" & direccion & "','" & localidad & "','" & provincia & "','" & país & "','" & detalle & "','" & Cargo & "','" & celular & "','" & email & "',2)"
            If classP.ejecutarComando(strsql) Then
                MsgBox "Alta exitosa!", vbInformation, "Información"
                verContactos
            End If
        End If
    End If
End Sub

