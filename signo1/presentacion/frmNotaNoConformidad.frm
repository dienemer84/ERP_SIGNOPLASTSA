VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmNotaNoConformidad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nota No Conformidad"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7845
   Icon            =   "frmNotaNoConformidad.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   7845
   Begin XtremeSuiteControls.PushButton A 
      Height          =   405
      Left            =   4920
      TabIndex        =   23
      Top             =   7470
      Width           =   1500
      _Version        =   786432
      _ExtentX        =   2646
      _ExtentY        =   714
      _StockProps     =   79
      Caption         =   "Archivos"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton Imprimir 
      Height          =   405
      Left            =   3405
      TabIndex        =   19
      Top             =   7470
      Width           =   1500
      _Version        =   786432
      _ExtentX        =   2646
      _ExtentY        =   714
      _StockProps     =   79
      Caption         =   "Imprimir"
      UseVisualStyle  =   -1  'True
   End
   Begin VB.ComboBox cboOperario 
      Height          =   315
      Left            =   1545
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   2670
      Width           =   5790
   End
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   4125
      Left            =   195
      TabIndex        =   10
      Top             =   3165
      Width           =   7485
      _Version        =   786432
      _ExtentX        =   13203
      _ExtentY        =   7276
      _StockProps     =   68
      Appearance      =   10
      Color           =   32
      ItemCount       =   3
      SelectedItem    =   2
      Item(0).Caption =   "Falla"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "txtFalla"
      Item(1).Caption =   "Resolución"
      Item(1).ControlCount=   2
      Item(1).Control(0)=   "txtAccion"
      Item(1).Control(1)=   "lblAprobador"
      Item(2).Caption =   "Incidencias"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "txtIncidencias"
      Begin VB.TextBox txtIncidencias 
         Height          =   3495
         Left            =   105
         MultiLine       =   -1  'True
         TabIndex        =   20
         Top             =   405
         Width           =   7200
      End
      Begin VB.TextBox txtAccion 
         Height          =   2925
         Left            =   -69910
         MultiLine       =   -1  'True
         TabIndex        =   12
         Top             =   435
         Visible         =   0   'False
         Width           =   7200
      End
      Begin VB.TextBox txtFalla 
         Height          =   3525
         Left            =   -69835
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   420
         Visible         =   0   'False
         Width           =   7125
      End
      Begin VB.Label lblAprobador 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Por"
         Height          =   195
         Left            =   -69790
         TabIndex        =   13
         Top             =   3600
         Visible         =   0   'False
         Width           =   2160
      End
   End
   Begin VB.ComboBox cboTareasDisponibles 
      Height          =   315
      Left            =   1545
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   2250
      Width           =   6120
   End
   Begin XtremeSuiteControls.PushButton cmdQuitarOperario 
      Height          =   315
      Left            =   7335
      TabIndex        =   18
      Top             =   2655
      Width           =   300
      _Version        =   786432
      _ExtentX        =   529
      _ExtentY        =   556
      _StockProps     =   79
      Caption         =   "X"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton cmdTerminar 
      Height          =   390
      Left            =   1830
      TabIndex        =   21
      Top             =   7470
      Width           =   1500
      _Version        =   786432
      _ExtentX        =   2646
      _ExtentY        =   688
      _StockProps     =   79
      Caption         =   "Guardar"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton cmdCrear 
      Height          =   390
      Left            =   285
      TabIndex        =   22
      Top             =   7470
      Width           =   1500
      _Version        =   786432
      _ExtentX        =   2646
      _ExtentY        =   688
      _StockProps     =   79
      Caption         =   "Crear Resolución"
      UseVisualStyle  =   -1  'True
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Operario"
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
      Left            =   690
      TabIndex        =   17
      Top             =   2700
      Width           =   735
   End
   Begin VB.Label lblEstado 
      BackColor       =   &H00FFC0FF&
      Height          =   300
      Left            =   1530
      TabIndex        =   15
      Top             =   135
      Width           =   6120
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Estado"
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
      Left            =   840
      TabIndex        =   14
      Top             =   150
      Width           =   600
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tarea Origen"
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
      Left            =   315
      TabIndex        =   8
      Top             =   2295
      Width           =   1125
   End
   Begin VB.Label lblSectorDestino 
      BackColor       =   &H00FFC0FF&
      Height          =   300
      Left            =   1530
      TabIndex        =   7
      Top             =   1905
      Width           =   6120
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sector Destino "
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
      Left            =   105
      TabIndex        =   6
      Top             =   1890
      Width           =   1335
   End
   Begin VB.Label lblPiezaTarea 
      BackColor       =   &H00FFC0FF&
      Height          =   360
      Left            =   1530
      TabIndex        =   5
      Top             =   915
      Width           =   6120
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pieza / Tarea"
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
      Left            =   255
      TabIndex        =   4
      Top             =   975
      Width           =   1185
   End
   Begin VB.Label lblOriginador 
      BackColor       =   &H00FFC0FF&
      Height          =   300
      Left            =   1530
      TabIndex        =   3
      Top             =   1455
      Width           =   6120
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Originador"
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
      Left            =   555
      TabIndex        =   2
      Top             =   1455
      Width           =   885
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OT / Item"
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
      Left            =   600
      TabIndex        =   1
      Top             =   525
      Width           =   840
   End
   Begin VB.Label lblOTItem 
      BackColor       =   &H00FFC0FF&
      Height          =   300
      Left            =   1530
      TabIndex        =   0
      Top             =   525
      Width           =   6120
   End
End
Attribute VB_Name = "frmNotaNoConformidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private TiempoProceso As PlaneamientoTiempoProceso
Dim Pieza As Pieza
Dim Tarea As clsTarea
Dim n As New NotaNoConformidad
Public Property Let nnc(nvalue As NotaNoConformidad)
    Set n = nvalue
    Set TiempoProceso = nvalue.TiempoProceso    'DAOTiemposProceso.FindById(nValue)
    SetUpForm
    'si entra x aca es porque ya está creada y esta en edición
    If n.estado = NNC_EnEdicion Then
        'si esta en edición puede editarse la falla pero no la acción
        Me.txtFalla.Locked = False
        Me.txtAccion.Locked = True
    ElseIf n.estado = NNC_Pendiente Then
        'si esta pendiente puede modificarse la accion a tomar pero no la falla
        Me.txtFalla.Locked = True
        Me.txtAccion.Locked = False
    ElseIf n.estado = NNC_Resuelta Then
        'si esta resolved no puede hcaerse un joraca
        Me.txtFalla.Locked = True
        Me.txtAccion.Locked = True
        Me.cmdCrear.Enabled = False
        Me.cmdTerminar.Enabled = False
    End If



End Property


Private Sub SetUpForm()
    Set Pieza = DAOPieza.FindById(TiempoProceso.idPieza, FL_0)
    Set Tarea = DAOTareas.FindById(TiempoProceso.Tarea.Id)
End Sub
Public Property Let idTiempoProceso(nvalue As Long)
'si entra por acá es porque  es una nueva nnc
    Dim d As DetalleOTConjuntoDTO
    Set TiempoProceso = DAOTiemposProceso.FindById(nvalue)
    SetUpForm

    Dim Id As Long

    If TiempoProceso.EsConjunto Then

        Dim detalle As DetalleOTConjuntoDTO
        Set detalle = DAODetalleOrdenTrabajo.FindConjuntoById(TiempoProceso.idDetallePedido)
        If IsSomething(detalle) Then
            Id = detalle.idDetallePedido
        End If
    Else
        Id = TiempoProceso.idDetallePedido
    End If

    Set TiempoProceso.DetalleOt = DAODetalleOrdenTrabajo.FindById(Id)
    Set TiempoProceso.Tarea = Tarea


End Property



Private Sub A_Click()
    Dim frmarchi1 As New frmArchivos2
    frmarchi1.Origen = OrigenArchivos.OA_NotaNoConformidad
    frmarchi1.ObjetoId = n.Id
    frmarchi1.caption = "NNC Nº " & n.Id
    frmarchi1.Show
End Sub

Private Sub cmdCrear_Click()
    If LenB(Me.txtAccion) > 20 Then


        n.AccionTomada = Me.txtAccion
        n.FechaResolucion = Now
        Set n.UsuarioResolucionador = funciones.GetUserObj
        Me.lblAprobador = n.UsuarioResolucionador.Empleado.NombreCompleto
        n.estado = NNC_Resuelta

        If guardarNNC(n) Then
            MsgBox "Se creó correctamente la resolución de la NNC " & n.numero
            Unload Me
        Else

            n.FechaResolucion = Empty
            n.UsuarioResolucionador = Nothing
            n.AccionTomada = vbNullString
            MsgBox "Error. Consulte.", vbCritical, "Error"
        End If
    Else
        MsgBox "Escriba algo correctamente o se formateará la PC!", vbCritical, "Error"

    End If
End Sub

Private Function guardarNNC(n As NotaNoConformidad) As Boolean
    Dim ea As EstadoNotaNoConformidad
    guardarNNC = False
    ea = n.estado
    n.descripcion = Me.txtFalla
    n.Incidencias = Me.txtIncidencias
    If Me.cboTareasDisponibles.ListIndex <> -1 Then
        Set n.TareaOrigen = DAOTareas.FindById(Me.cboTareasDisponibles.ItemData(Me.cboTareasDisponibles.ListIndex))
    End If
    If Me.cboOperario.ListIndex <> -1 Then
        Set n.Operario = DAOEmpleados.GetById(Me.cboOperario.ItemData(Me.cboOperario.ListIndex))
    End If
    Set n.usuarioCreador = funciones.GetUserObj
    If DAONotaNoConformidad.Guardar(n) Then
        guardarNNC = True
        MsgBox "Nota de No Conformidad guardada.", vbOKOnly + vbInformation
    Else
        GoTo err1

    End If
    Exit Function
err1:
    n.estado = ea
    guardarNNC = False
End Function



Private Sub cmdQuitarOperario_Click()
    Me.cboOperario.ListIndex = -1
End Sub

Private Sub cmdTerminar_Click()
    If MsgBox("¿Está seguro de guardar la NNC?", vbYesNo, "Confirmación") = vbYes Then
        n.estado = NNC_Pendiente
        If guardarNNC(n) Then
            MsgBox "NNC " & n.numero & " Guardada correctamente!", vbExclamation, "Información"
            Unload Me
        End If

    End If
End Sub

Private Sub Form_Load()
    FormHelper.Customize Me
    If IsSomething(n) Then
        '  Me.caption = Me.caption & " " & n.numero
    Else
        Me.caption = "Nueva nota no conformidad"
    End If
    Me.TabControl1.selectedItem = 0

    Set n.TiempoProceso = TiempoProceso
    Me.txtFalla = n.descripcion
    Me.txtAccion = n.AccionTomada
    Me.lblEstado = enums.EnumEstadoNNC(n.estado)
    Me.txtIncidencias = n.Incidencias
    Me.lblPiezaTarea = Pieza.nombre & " / " & TiempoProceso.Tarea.Id & " - " & Tarea.Tarea
    Me.lblOTItem = TiempoProceso.idpedido & " / " & TiempoProceso.DetalleOt.item
    Me.lblOriginador = funciones.GetUserObj.Empleado.NombreCompleto
    Me.lblSectorDestino = Tarea.Sector.Sector & " (" & Tarea.Id & " - " & Tarea.Tarea & ")"
    Me.cboOperario.Clear
    Dim empleados As Collection
    Dim emp As clsEmpleado
    Set empleados = DAOEmpleados.GetAllByTareaId(TiempoProceso.Tarea.Id)
    For Each emp In empleados
        Me.cboOperario.AddItem emp.NombreCompleto
        Me.cboOperario.ItemData(Me.cboOperario.NewIndex) = emp.Id
    Next emp

    If IsSomething(n.UsuarioResolucionador) Then Me.lblAprobador = "Por " & n.UsuarioResolucionador.Id


    If IsSomething(n.Operario) Then
        Me.cboOperario.ListIndex = PosIndexCbo(n.Operario.Id, cboOperario)
    End If
    ' Me.cmdTerminar.Enabled = (n.estado = NNC_EnEdicion)
    'Me.cmdCrear.Enabled = (n.estado = NNC_Pendiente)
    'Me.cmdGuardar.Enabled = (n.estado = NNC_EnEdicion)

    llenarComboTareas



End Sub


Private Sub llenarComboTareas()
    Dim col As Collection

    Set col = DAOTareas.FindAll("t.id IN (SELECT tarea_id FROM empleado_tarea WHERE personal_id = " & funciones.GetUserObj.Empleado.Id & ")")

    Dim tar As clsTarea
    For Each tar In col
        Me.cboTareasDisponibles.AddItem tar.Tarea & " (" & tar.Sector.Sector & ")"
        Me.cboTareasDisponibles.ItemData(Me.cboTareasDisponibles.NewIndex) = tar.Id
    Next

    If col.count > 0 Then Me.cboTareasDisponibles.ListIndex = 0
End Sub

Private Sub Imprimir_Click()
    If n.Id > 0 Then
        DAONotaNoConformidad.Imprimir n
    Else
        MsgBox "Debe grabar la NNC antes de imprimir!", vbOKOnly + vbInformation
    End If
End Sub
