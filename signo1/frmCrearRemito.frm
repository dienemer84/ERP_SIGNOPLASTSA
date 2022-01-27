VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmPlaneamientoRemitosNuevo 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Nuevo Remito..."
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   7725
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   4095
      Left            =   105
      TabIndex        =   0
      Top             =   60
      Width           =   7515
      _Version        =   786432
      _ExtentX        =   13256
      _ExtentY        =   7223
      _StockProps     =   79
      Caption         =   "GroupBox1"
      UseVisualStyle  =   -1  'True
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Definir entrega"
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
         Left            =   495
         TabIndex        =   3
         Top             =   1680
         Width           =   1695
      End
      Begin XtremeSuiteControls.GroupBox GroupBox2 
         Height          =   1740
         Left            =   255
         TabIndex        =   10
         Top             =   1725
         Width           =   7035
         _Version        =   786432
         _ExtentX        =   12409
         _ExtentY        =   3069
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.PushButton PushButton1 
            Height          =   315
            Left            =   5985
            TabIndex        =   15
            Top             =   390
            Width           =   780
            _Version        =   786432
            _ExtentX        =   1376
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "Filtrar"
            UseVisualStyle  =   -1  'True
         End
         Begin VB.TextBox Text1 
            Height          =   300
            Left            =   2160
            TabIndex        =   14
            Top             =   390
            Width           =   3735
         End
         Begin XtremeSuiteControls.ComboBox cboContactos 
            Height          =   315
            Left            =   1155
            TabIndex        =   11
            Top             =   840
            Width           =   5610
            _Version        =   786432
            _ExtentX        =   9895
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Sorted          =   -1  'True
            Style           =   2
            Appearance      =   6
            UseVisualStyle  =   -1  'True
            Text            =   "ComboBox1"
         End
         Begin VB.Label Label5 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Filtrar Contáctos por"
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
            TabIndex        =   13
            Top             =   435
            Width           =   2220
         End
         Begin VB.Label Label4 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Contácto"
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
            Left            =   270
            TabIndex        =   12
            Top             =   885
            Width           =   945
         End
      End
      Begin VB.TextBox txtRtoNro 
         Height          =   285
         Left            =   1065
         TabIndex        =   5
         Top             =   435
         Width           =   1215
      End
      Begin VB.TextBox txtDetalles 
         Height          =   285
         Left            =   1065
         TabIndex        =   4
         Top             =   795
         Width           =   6135
      End
      Begin XtremeSuiteControls.PushButton Command1 
         Height          =   420
         Left            =   2520
         TabIndex        =   1
         Top             =   3555
         Width           =   1125
         _Version        =   786432
         _ExtentX        =   1984
         _ExtentY        =   741
         _StockProps     =   79
         Caption         =   "Generar"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cboClientes 
         Height          =   315
         Left            =   1065
         TabIndex        =   2
         Top             =   1155
         Width           =   6135
         _Version        =   786432
         _ExtentX        =   10821
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.PushButton Command2 
         Height          =   420
         Left            =   3870
         TabIndex        =   6
         Top             =   3555
         Width           =   1125
         _Version        =   786432
         _ExtentX        =   1984
         _ExtentY        =   741
         _StockProps     =   79
         Caption         =   "Cerrar"
         UseVisualStyle  =   -1  'True
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Número"
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
         Left            =   225
         TabIndex        =   9
         Top             =   435
         Width           =   1095
      End
      Begin VB.Label Label2 
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
         Left            =   225
         TabIndex        =   8
         Top             =   1155
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Detalles"
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
         TabIndex        =   7
         Top             =   795
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmPlaneamientoRemitosNuevo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim claseP As New classPlaneamiento
Dim claseS As New classStock
Dim Remito As Remito
Dim cliente As clsCliente
Private Sub cboClientes_Click()
    Me.cboContactos.Clear
    llenarContactos
End Sub
Private Sub Check1_Click()
    If Me.Check1.value Then
        Me.cboContactos.Enabled = True
    Else
        Me.cboContactos.Enabled = False
    End If
End Sub
Private Sub Command1_Click()
    If DAORemitoS.FindByNumero(CLng(Me.txtRtoNro)) Is Nothing Then
        If MsgBox("¿Desea crear el Remito Nro. " & Format(CLng(Me.txtRtoNro), "0000") & "?", vbYesNo, "Confirmación") = vbYes Then

            Set Remito = New Remito
            Set Remito.cliente = DAOCliente.BuscarPorID(Me.cboClientes.ItemData(cboClientes.ListIndex))
            Remito.detalle = UCase(Me.txtDetalles)
            Remito.FEcha = Now
            Remito.estado = RemitoPendiente
            Remito.EstadoFacturado = RemitoNoFacturado
            Set Remito.usuarioAprobador = Nothing
            Set Remito.usuarioCreador = funciones.GetUserObj

            Remito.numero = CLng(Me.txtRtoNro)



            If Not Remito.cliente.CUITValido Or Not Remito.cliente.ValidoRemitoFactura Then
                MsgBox "El cliente no es válido para generar un remito!", vbCritical, "Error"
                Exit Sub
            End If

            If Me.Check1.value Then
                Set Remito.contacto = DAOContacto.FindById(TipoPersona.cliente_, (Me.cboContactos.ItemData(Me.cboContactos.ListIndex)))
            Else
                Set Remito.contacto = Nothing
            End If
            If Not DAORemitoS.Save(Remito) Then
                MsgBox "Se produjo un error!", vbCritical, "Error"
            Else
                MsgBox "Guardado correctamente!", vbInformation, "Información"
                Me.txtRtoNro = DAORemitoS.ProximoRemito
            End If



            Me.txtDetalles = Empty


        Else
            MsgBox "Se produjo un error!", vbCritical, "Error"
        End If
    End If
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    FormHelper.Customize Me
    Dim idcli As Long
    Dim rs As Recordset

    DAOCliente.llenarComboXtremeSuite Me.cboClientes, False, True, False
    llenarContactos
    Me.txtRtoNro = DAORemitoS.ProximoRemito
    verificar
    Check1_Click

End Sub
Private Sub llenarContactos()
    Dim idcli As Long
    Dim contacto As clsContacto

    idcli = CInt(Me.cboClientes.ItemData(Me.cboClientes.ListIndex))
    Dim c As New Collection
    Set c = DAOContacto.FindAll(cliente_, "nombre like '%" & Me.Text1 & "%' and idCliente=" & idcli)
    cboContactos.Clear

    For Each contacto In c
        If IsSomething(contacto) Then
            Me.cboContactos.AddItem contacto.nombre
            Me.cboContactos.ItemData(Me.cboContactos.NewIndex) = contacto.Id
        End If
    Next


    If Me.cboContactos.ListCount > 0 Then
        Me.cboContactos.ListIndex = 0
    End If

End Sub
Private Sub verificar()
    If Trim(Me.txtDetalles) = Empty Or Trim(Me.txtRtoNro) = Empty Then
        Me.Command1.Enabled = False
    Else
        Me.Command1.Enabled = True
    End If
End Sub

Private Sub PushButton1_Click()
    llenarContactos
End Sub

Private Sub txtDetalles_Change()
    verificar
End Sub
Private Sub txtRtoNro_Change()
    verificar
End Sub
