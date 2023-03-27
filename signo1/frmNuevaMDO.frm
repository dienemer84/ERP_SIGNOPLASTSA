VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmNuevaMDO 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tarea"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   255
   ClientWidth     =   4035
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNuevaMDO.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   4035
   Begin MSComCtl2.DTPicker dtpFecha 
      Height          =   315
      Left            =   2160
      TabIndex        =   14
      Top             =   3045
      Visible         =   0   'False
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   556
      _Version        =   393216
      Format          =   58458113
      CurrentDate     =   40101
   End
   Begin VB.CommandButton cmdGuardar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Guardar"
      Default         =   -1  'True
      Height          =   375
      Left            =   2865
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2160
      Width           =   1095
   End
   Begin VB.ComboBox cboSectores 
      Height          =   315
      Left            =   735
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   105
      Width           =   3225
   End
   Begin VB.ComboBox cboCant 
      Height          =   315
      ItemData        =   "frmNuevaMDO.frx":000C
      Left            =   1785
      List            =   "frmNuevaMDO.frx":0036
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   510
      Width           =   2115
   End
   Begin VB.TextBox txtNombreTarea 
      Height          =   300
      Left            =   1275
      TabIndex        =   2
      Top             =   915
      Width           =   2610
   End
   Begin VB.TextBox txtDescripcion 
      Height          =   300
      Left            =   1065
      TabIndex        =   3
      Top             =   1320
      Width           =   2850
   End
   Begin VB.TextBox txtValor 
      Height          =   300
      Left            =   135
      TabIndex        =   6
      Text            =   "1"
      Top             =   3045
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox cboCategoria 
      Height          =   315
      Left            =   945
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1680
      Width           =   2970
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "Sector"
      Height          =   195
      Left            =   165
      TabIndex        =   13
      Top             =   150
      Width           =   465
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "Cantidad por proceso"
      Height          =   195
      Left            =   150
      TabIndex        =   12
      Top             =   540
      Width           =   1545
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "Nombre Tarea"
      Height          =   195
      Left            =   180
      TabIndex        =   11
      Top             =   945
      Width           =   1020
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "Descripción"
      Height          =   195
      Left            =   165
      TabIndex        =   10
      Top             =   1350
      Width           =   810
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "Valor"
      Height          =   195
      Left            =   -315
      TabIndex        =   9
      Top             =   3075
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "Fecha"
      Height          =   195
      Left            =   1620
      TabIndex        =   8
      Top             =   3090
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "Categoria"
      Height          =   195
      Left            =   150
      TabIndex        =   7
      Top             =   1710
      Width           =   705
   End
End
Attribute VB_Name = "frmNuevaMDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_tarea As clsTarea
Private m_categorias As Collection

Public Property Let tareaId(value As Long)
    Set m_tarea = DAOTareas.FindById(value)
    CargarTarea
End Property

Private Sub CargarTarea()
    Me.cboCant.ListIndex = PosIndexCbo(m_tarea.CantPorProc, Me.cboCant)
    Me.cboSectores.ListIndex = PosIndexCbo(m_tarea.Sector.Id, Me.cboSectores)

    If m_tarea.CategoriaSueldo Is Nothing Then
        Me.cboCategoria.ListIndex = -1
    Else
        Me.cboCategoria.ListIndex = PosIndexCbo(m_tarea.CategoriaSueldo.Id, Me.cboCategoria)
    End If

    Me.txtDescripcion.text = m_tarea.descripcion
    Me.dtpFecha.value = m_tarea.FEcha
    'Set m_tarea.Moneda = DAOMoneda.GetById(0)
    Me.txtNombreTarea.text = m_tarea.Tarea
    Me.txtValor.text = m_tarea.Valor
End Sub

Private Sub cmdBorrarCategoria_Click()
    Me.cboCategoria.ListIndex = -1
End Sub

Private Sub cmdGuardar_Click()
    On Error GoTo E

    If m_tarea Is Nothing Then Set m_tarea = New clsTarea

    Dim msg As String
    If Me.cboSectores.ListIndex = -1 Then msg = msg & vbNewLine & "- Debe seleccionar el sector"
    If Me.cboCant.ListIndex = -1 Then msg = msg & vbNewLine & "- Debe seleccionar la cantidad por proceso"
    If LenB(Me.txtNombreTarea.text) = 0 Then msg = msg & vbNewLine & "- Debe especificar el nombre de la tarea"
    If LenB(Me.txtDescripcion.text) = 0 Then msg = msg & vbNewLine & "- Debe especificar la descripcion"
    '    If LenB(Me.txtValor.text) = 0 Then MSG = MSG & vbNewLine & "- Debe especificar el valor"
    If Me.cboCategoria.ListIndex = -1 Then msg = msg & vbNewLine & "- Debe seleccionar la categoria"

    If LenB(msg) > 0 Then
        MsgBox msg, vbCritical
        Exit Sub
    End If

    m_tarea.CantPorProc = Me.cboCant.ItemData(Me.cboCant.ListIndex)
    If Me.cboCategoria.ListIndex = -1 Then
        Set m_tarea.CategoriaSueldo = Nothing
    Else
        Set m_tarea.CategoriaSueldo = m_categorias.item(CStr(Me.cboCategoria.ItemData(Me.cboCategoria.ListIndex)))
    End If
    m_tarea.descripcion = Me.txtDescripcion.text
    m_tarea.FEcha = Me.dtpFecha.value
    'Set m_tarea.moneda = DAOMoneda.GetById(0)  'scarlo del form
    Set m_tarea.Sector = DAOSectores.GetById(Me.cboSectores.ItemData(Me.cboSectores.ListIndex))
    m_tarea.Tarea = Me.txtNombreTarea.text
    m_tarea.Valor = 0    'CDbl(Me.txtValor.text)


    If DAOTareas.Save(m_tarea) Then

        Dim EVENTO As New clsEventoObserver
        Set EVENTO.Elemento = m_tarea

        If m_tarea.Id = 0 Then
            EVENTO.EVENTO = agregar_
        Else
            EVENTO.EVENTO = modificar_
        End If

        Set EVENTO.Originador = Me

        Channel.Notificar EVENTO, Tareas_

        MsgBox "Tarea guardada", vbInformation
        Unload Me
        Exit Sub
    Else
        MsgBox "Hubo un error al guardar la tarea", vbCritical
    End If

    Exit Sub
E:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub Form_Load()
    FormHelper.Customize Me
    Me.cboCant.ListIndex = 0
    Me.dtpFecha.value = Date

    DAOSectores.LlenarCombo Me.cboSectores

    Dim cat As CategoriaSueldo
    Me.cboCategoria.Clear
    Set m_categorias = DAOCategoriaSueldo.FindAll()
    For Each cat In m_categorias
        Me.cboCategoria.AddItem cat.nombre
        Me.cboCategoria.ItemData(Me.cboCategoria.NewIndex) = cat.Id
    Next cat

    ''Me.caption = caption & " (" & Name & ")"

End Sub


