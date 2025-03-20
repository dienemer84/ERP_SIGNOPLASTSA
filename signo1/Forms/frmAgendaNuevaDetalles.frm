VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmAgendaNuevaDetalles 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ver contactos"
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   11805
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtEmail 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   5
      Text            =   "Text4"
      Top             =   1320
      Width           =   3975
   End
   Begin VB.TextBox txtLocalidad 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   4
      Text            =   "Text3"
      Top             =   480
      Width           =   3975
   End
   Begin VB.TextBox txtDireccion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   1320
      Width           =   4455
   End
   Begin VB.TextBox txtNombre 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   480
      Width           =   4455
   End
   Begin XtremeSuiteControls.PushButton btnModificar 
      Default         =   -1  'True
      Height          =   495
      Left            =   9480
      TabIndex        =   1
      Top             =   1260
      Width           =   1935
      _Version        =   786432
      _ExtentX        =   3413
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Modificar"
      UseVisualStyle  =   -1  'True
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   4335
      Left            =   360
      TabIndex        =   0
      Top             =   2040
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   7646
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      ColumnAutoResize=   -1  'True
      MethodHoldFields=   -1  'True
      AllowDelete     =   -1  'True
      GroupByBoxVisible=   0   'False
      DataMode        =   99
      AllowAddNew     =   -1  'True
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   5
      Column(1)       =   "frmAgendaNuevaDetalles.frx":0000
      Column(2)       =   "frmAgendaNuevaDetalles.frx":0118
      Column(3)       =   "frmAgendaNuevaDetalles.frx":021C
      Column(4)       =   "frmAgendaNuevaDetalles.frx":0320
      Column(5)       =   "frmAgendaNuevaDetalles.frx":040C
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAgendaNuevaDetalles.frx":0508
      FormatStyle(2)  =   "frmAgendaNuevaDetalles.frx":0640
      FormatStyle(3)  =   "frmAgendaNuevaDetalles.frx":06F0
      FormatStyle(4)  =   "frmAgendaNuevaDetalles.frx":07A4
      FormatStyle(5)  =   "frmAgendaNuevaDetalles.frx":087C
      FormatStyle(6)  =   "frmAgendaNuevaDetalles.frx":0934
      ImageCount      =   0
      PrinterProperties=   "frmAgendaNuevaDetalles.frx":0A14
   End
   Begin VB.Label Label4 
      Caption         =   "Dirección"
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   1080
      Width           =   2895
   End
   Begin VB.Label Label3 
      Caption         =   "Email"
      Height          =   255
      Left            =   5160
      TabIndex        =   8
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Localidad"
      Height          =   255
      Left            =   5160
      TabIndex        =   7
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "frmAgendaNuevaDetalles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Id As Long
Dim contacto_ As clsContactoPpal
Dim contacto_detalles As clsContactoPpalDetalle
Dim detalles As Collection

Public Property Let Contacto(nvalue As clsContactoPpal)
    Set contacto_ = nvalue
End Property


Private Sub Form_Load()
    FormHelper.Customize Me
    
    Me.txtNombre.Text = contacto_.Empresa
    Me.txtDireccion.Text = contacto_.direccion
    Me.txtLocalidad.Text = contacto_.localidad
    Me.txtEmail.Text = contacto_.email
    
    llenarGrid

End Sub


Public Function llenarGrid()

    
    Set detalles = DAOContactoPpalDetalles.FindAllByContactoPpal(contacto_.Id)

    Me.GridEX1.ItemCount = 0
    Me.GridEX1.ItemCount = detalles.count


End Function


Private Sub GridEX1_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    On Error GoTo err1
    
    Set contacto_detalles = detalles.item(rowIndex)

    Values(1) = contacto_detalles.detalle
    Values(2) = contacto_detalles.Telefono1
    Values(3) = contacto_detalles.Telefono2
    Values(4) = contacto_detalles.mail
    Values(5) = contacto_detalles.Mas

    Exit Sub
err1:


End Sub
