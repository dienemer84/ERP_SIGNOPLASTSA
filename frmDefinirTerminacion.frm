VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~3.OCX"
Begin VB.Form frmDefinirTerminacion 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Terminación..."
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   4725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   4725
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.PushButton Command1 
      Default         =   -1  'True
      Height          =   375
      Left            =   840
      TabIndex        =   14
      Top             =   3240
      Width           =   1335
      _Version        =   786432
      _ExtentX        =   2355
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Actualizar"
      UseVisualStyle  =   -1  'True
   End
   Begin VB.ComboBox cboHorno 
      Height          =   315
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2400
      Width           =   2535
   End
   Begin VB.ComboBox cboApp 
      Height          =   315
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2040
      Width           =   2535
   End
   Begin VB.ComboBox cboSup 
      Height          =   315
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1680
      Width           =   2535
   End
   Begin VB.ComboBox cboFosf 
      Height          =   315
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   480
      Width           =   2655
   End
   Begin VB.ComboBox cboCant 
      Height          =   315
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   2655
   End
   Begin VB.ComboBox cboSector 
      Height          =   315
      Left            =   720
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2760
      Width           =   3855
   End
   Begin VB.ComboBox cboRubros 
      Height          =   315
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   840
      Width           =   3735
   End
   Begin XtremeSuiteControls.PushButton Command2 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   2520
      TabIndex        =   15
      Top             =   3240
      Width           =   1335
      _Version        =   786432
      _ExtentX        =   2355
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Salir"
      UseVisualStyle  =   -1  'True
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Sector"
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
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cantidad de pintura"
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
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Fosfatos"
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
      TabIndex        =   11
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Superficies"
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
      Left            =   960
      TabIndex        =   10
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Aplicación"
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
      Left            =   960
      TabIndex        =   9
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Horneado"
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
      Left            =   960
      TabIndex        =   8
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Rubro"
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
      TabIndex        =   7
      Top             =   840
      Width           =   615
   End
End
Attribute VB_Name = "frmDefinirTerminacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim notLoading As Boolean
Dim dto As DTOCuentasTerminacion
Private Sub cboRubros_Click()
    If Not notLoading Then Exit Sub
    Set dto.Rubro = DAORubros.FindById(Me.cboRubros.ItemData(Me.cboRubros.ListIndex))
    LlenarCuentasMAT
End Sub
Private Sub cboSector_Click()
    If Not notLoading Then Exit Sub
    On Error Resume Next
    Set dto.Sector = DAOSectores.GetById(Me.cboSector.ItemData(Me.cboSector.ListIndex))

    LlenarCuentasMDO
End Sub

Private Sub Command1_Click()
    Dim strsql As String
    Dim Rubro As Long
    Dim Sector As Long
    Dim idFosf As Long
    Dim idCant As Long
    Dim idSup As Long
    Dim idApp As Long
    Dim idHorno As Long
    idApp = Me.cboApp.ItemData(Me.cboApp.ListIndex)
    idHorno = Me.cboHorno.ItemData(Me.cboHorno.ListIndex)
    idSup = Me.cboSup.ItemData(Me.cboSup.ListIndex)
    idCant = Me.cboCant.ItemData(Me.cboCant.ListIndex)
    idFosf = Me.cboFosf.ItemData(Me.cboFosf.ListIndex)
    Sector = Me.cboSector.ItemData(Me.cboSector.ListIndex)
    Rubro = Me.cboRubros.ItemData(Me.cboRubros.ListIndex)


    Set dto.Rubro = DAORubros.FindById(Rubro)
    Set dto.CantidadFosfatos = DAOMateriales.FindById(idFosf)
    Set dto.CantidadPintura = DAOMateriales.FindById(idCant)
    Set dto.Sector = DAOSectores.GetById(Sector)
    Set dto.Aplicacion = DAOTareas.FindById(idApp)
    Set dto.Horneado = DAOTareas.FindById(idHorno)
    Set dto.Limpieza = DAOTareas.FindById(idSup)

    If MsgBox("¿Seguro de actualizar?", vbYesNo, "Confirmación") = 6 Then
        If DAOdtoCuentasTerminacion.SaveConfigTerminacion(dto) Then
            MsgBox "Actualización exitosa!", vbOKOnly
        Else
            MsgBox "Se produjo algún error!", vbCritical
        End If

    End If
End Sub

Private Sub Command2_Click()

    Unload Me
End Sub


Private Sub LlenarCuentasMDO()
    On Error Resume Next
    DAOTareas.LlenarComboPorSector Me.cboHorno, dto.Sector
    DAOTareas.LlenarComboPorSector Me.cboApp, dto.Sector
    DAOTareas.LlenarComboPorSector Me.cboSup, dto.Sector
    Me.cboHorno.ListIndex = funciones.PosIndexCbo(dto.Horneado.id, Me.cboHorno)
    Me.cboApp.ListIndex = funciones.PosIndexCbo(dto.Aplicacion.id, Me.cboApp)
    Me.cboSup.ListIndex = funciones.PosIndexCbo(dto.Limpieza.id, Me.cboSup)

End Sub

Private Sub LlenarCuentasMAT()
    On Error Resume Next
    DAOMateriales.LlenarComboPorRubro Me.cboCant, dto.Rubro
    DAOMateriales.LlenarComboPorRubro Me.cboFosf, dto.Rubro
    Me.cboCant.ListIndex = funciones.PosIndexCbo(dto.CantidadPintura.id, cboCant)
    Me.cboFosf.ListIndex = funciones.PosIndexCbo(dto.CantidadFosfatos.id, cboFosf)
End Sub

Private Sub LlenarRubros()
    DAORubros.LlenarCombo Me.cboRubros
    Me.cboRubros.ListIndex = funciones.PosIndexCbo(dto.Rubro.id, Me.cboRubros)
End Sub

Private Sub LlenarSector()
    DAOSectores.LlenarCombo Me.cboSector
    Me.cboSector.ListIndex = funciones.PosIndexCbo(dto.Sector.id, Me.cboSector)

End Sub


Private Sub Form_Load()
    Customize Me
    Set dto = DAOdtoCuentasTerminacion.GetConfigTerminacion
    LlenarRubros
    LlenarSector
    LlenarCuentasMDO
    LlenarCuentasMAT

    notLoading = True



End Sub
