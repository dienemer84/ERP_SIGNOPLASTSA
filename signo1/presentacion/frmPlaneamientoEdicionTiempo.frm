VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmPlaneamientoEdicionTiempo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tiempo"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4440
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPlaneamientoEdicionTiempo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4440
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCant 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1110
      TabIndex        =   8
      Top             =   1980
      Width           =   870
   End
   Begin MSComCtl2.DTPicker dtpInicio 
      Height          =   330
      Left            =   1110
      TabIndex        =   5
      Top             =   1020
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy HH:mm:ss"
      Format          =   58458115
      CurrentDate     =   40140
   End
   Begin VB.ComboBox cboEmpleado 
      Height          =   315
      Left            =   1110
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   555
      Width           =   3120
   End
   Begin MSComCtl2.DTPicker dtpFin 
      Height          =   330
      Left            =   1110
      TabIndex        =   6
      Top             =   1500
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy HH:mm:ss"
      Format          =   58458115
      CurrentDate     =   40140
   End
   Begin XtremeSuiteControls.PushButton btnGuardar 
      Height          =   345
      Left            =   1170
      TabIndex        =   9
      Top             =   2460
      Width           =   1485
      _Version        =   786432
      _ExtentX        =   2619
      _ExtentY        =   609
      _StockProps     =   79
      Caption         =   "Guardar"
      ForeColor       =   9126421
      Appearance      =   6
   End
   Begin XtremeSuiteControls.PushButton btnCerrar 
      Cancel          =   -1  'True
      Height          =   345
      Left            =   2730
      TabIndex        =   10
      Top             =   2460
      Width           =   1485
      _Version        =   786432
      _ExtentX        =   2619
      _ExtentY        =   609
      _StockProps     =   79
      Caption         =   "Cerrar"
      ForeColor       =   9126421
      Appearance      =   6
   End
   Begin VB.Label lblCantProc2 
      AutoSize        =   -1  'True
      Caption         =   "Cant. Proc:"
      Height          =   195
      Left            =   195
      TabIndex        =   7
      Top             =   1995
      Width           =   825
   End
   Begin VB.Label lblFin2 
      AutoSize        =   -1  'True
      Caption         =   "Fin:"
      Height          =   195
      Left            =   750
      TabIndex        =   4
      Top             =   1545
      Width           =   270
   End
   Begin VB.Label lblInicio2 
      AutoSize        =   -1  'True
      Caption         =   "Inicio:"
      Height          =   195
      Left            =   585
      TabIndex        =   3
      Top             =   1065
      Width           =   435
   End
   Begin VB.Label lblTarea 
      AutoSize        =   -1  'True
      Caption         =   "Tarea:"
      Height          =   195
      Left            =   555
      TabIndex        =   2
      Top             =   150
      Width           =   480
   End
   Begin VB.Label llblEmpleado 
      AutoSize        =   -1  'True
      Caption         =   "Empleado:"
      Height          =   195
      Left            =   285
      TabIndex        =   0
      Top             =   585
      Width           =   750
   End
End
Attribute VB_Name = "frmPlaneamientoEdicionTiempo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private tarea_id As Long
Private planeamiento_tiempo_proceso_id As Long

Public deta As New PlaneamientoTiempoProcesoDetalle

Public Enum TipoEdicion
    iniciar
    finalizar
    IniciarFinalizar
    editar
End Enum

Public TEdicion As TipoEdicion

Private Sub ActivarControles()
    Me.dtpInicio.Enabled = (TEdicion = TipoEdicion.iniciar Or TEdicion = TipoEdicion.IniciarFinalizar Or TEdicion = TipoEdicion.editar)
    Me.dtpFin.Visible = (TEdicion = TipoEdicion.finalizar Or TEdicion = TipoEdicion.IniciarFinalizar Or TEdicion = TipoEdicion.editar)

    Me.dtpFin.value = Now

    Me.txtCant.Visible = (TEdicion = TipoEdicion.finalizar Or TEdicion = TipoEdicion.IniciarFinalizar Or TEdicion = TipoEdicion.editar)
    Me.cboEmpleado.Enabled = (TEdicion = TipoEdicion.iniciar Or TEdicion = TipoEdicion.IniciarFinalizar Or TEdicion = TipoEdicion.editar)

    Me.lblCantProc2.Visible = (TEdicion = TipoEdicion.finalizar Or TEdicion = TipoEdicion.IniciarFinalizar Or TEdicion = TipoEdicion.editar)
    Me.lblFin2.Visible = (TEdicion = TipoEdicion.finalizar Or TEdicion = TipoEdicion.IniciarFinalizar Or TEdicion = TipoEdicion.editar)
    'Me.lblInicio2.Visible = (TEdicion = TipoEdicion.Iniciar Or TEdicion = TipoEdicion.IniciarFinalizar)

End Sub

Public Property Set detalle(value As PlaneamientoTiempoProcesoDetalle)
    Set deta = value

    Me.cboEmpleado.Clear
    Me.cboEmpleado.AddItem value.Empleado.legajo & " - " & value.Empleado.NombreCompleto
    Me.cboEmpleado.ItemData(Me.cboEmpleado.NewIndex) = value.Empleado.Id
    Me.cboEmpleado.ListIndex = 0
    Me.cboEmpleado.Enabled = False

    Me.dtpInicio.value = value.FechaInicioTarea
    Me.dtpFin.value = Now
End Property

Public Property Set detalleEditar(value As PlaneamientoTiempoProcesoDetalle)
    Set deta = value

    Me.cboEmpleado.Clear
    Me.cboEmpleado.AddItem value.Empleado.legajo & " - " & value.Empleado.NombreCompleto
    Me.cboEmpleado.ItemData(Me.cboEmpleado.NewIndex) = value.Empleado.Id
    Me.cboEmpleado.ListIndex = 0
    Me.cboEmpleado.Enabled = False

    Me.txtCant.text = value.CantidadProcesada

    Me.dtpInicio.value = value.FechaInicioTarea
    Me.dtpFin.value = value.FechaFinTarea
End Property


Public Property Let PlaneamientoTiempoProcesoId(value As Long)
    planeamiento_tiempo_proceso_id = value

    Dim ptp As PlaneamientoTiempoProceso
    Set ptp = DAOTiemposProceso.FindById(planeamiento_tiempo_proceso_id)

    Me.dtpInicio.value = Now

    Dim col As Collection
    Set col = GetEmpleadosByTareaId(ptp.Tarea.Id)
    Me.cboEmpleado.Clear
    Dim empl As clsEmpleado
    For Each empl In col
        Me.cboEmpleado.AddItem empl.legajo & " - " & empl.NombreCompleto
        Me.cboEmpleado.ItemData(Me.cboEmpleado.NewIndex) = empl.Id
    Next
    If col.count = 0 Then MsgBox "No hay empleados asignados para poder realizar la tarea [" & ptp.Tarea.Id & " - " & ptp.Tarea.Tarea & "]." & vbNewLine & "Primero sectorice empleados y luego asignele tareas permitidas.", vbInformation

    Me.lblTarea.caption = "Tarea:  " & ptp.Tarea.Id & " - " & ptp.Tarea.Tarea

    ActivarControles
End Property


Private Sub btnCerrar_Click()
    TareaAgregada = False
    Unload Me
End Sub


Private Sub btnGuardar_Click()


    If Me.cboEmpleado.ListIndex = -1 Then
        MsgBox "Debe especificar un empleado.", vbInformation
        Exit Sub
    Else
        If TEdicion <> finalizar And TEdicion <> editar Then
            Dim dett As PlaneamientoTiempoProcesoDetalle
            Set dett = DAOTiemposProcesosDetalles.FindFirstWithoutFinishByEmpleadoId(Me.cboEmpleado.ItemData(Me.cboEmpleado.ListIndex))
            If Not dett Is Nothing Then
                If dett.PlaneamientoTiempoProceso.Tarea.Id <> DAOTiemposProceso.FindById(planeamiento_tiempo_proceso_id).Tarea.Id Then
                    MsgBox "El empleado [" & Me.cboEmpleado.text & "] tiene una tarea iniciada sin finalizar desde el " & dett.FechaInicioTarea, vbInformation + vbOKOnly
                    Exit Sub
                End If
            End If
        End If
    End If

    If IsNull(Me.dtpInicio.value) Then
        MsgBox "Debe especificar una fecha y hora.", vbInformation
        Exit Sub
    End If

    If (CDbl(Me.dtpInicio.value) - Int(CDbl(Me.dtpInicio.value))) = 0 Then
        MsgBox "Debe especificar fecha y hora.", vbInformation
        Exit Sub
    End If


    If Me.txtCant.Visible And Not IsNumeric(Me.txtCant) And TEdicion <> editar Then
        MsgBox "Debe especificar la cantidad procesada", vbInformation
        Exit Sub
    End If


    deta.FechaInicioTarea = Me.dtpInicio

    If Me.dtpFin.Visible Then
        If (CDbl(Me.dtpFin.value) - Int(CDbl(Me.dtpFin.value))) = 0 And TEdicion <> editar Then
            MsgBox "Debe especificar fecha y hora.", vbInformation
            Exit Sub
        End If

        If Me.dtpFin.value <= Me.dtpInicio.value And TEdicion <> editar Then
            MsgBox "Fecha fin no puede ser superior o igual a fecha inicio.", vbInformation
            Exit Sub
        End If

        deta.FechaFinTarea = Me.dtpFin.value
    Else
        deta.FechaFinTarea = Empty
    End If

    Set deta.Empleado = DAOEmpleados.GetById(Me.cboEmpleado.ItemData(Me.cboEmpleado.ListIndex))
    deta.CantidadProcesada = Val(Me.txtCant.text)
    deta.IdPlaneamientoTiempoProceso = planeamiento_tiempo_proceso_id

    If DAOTiemposProcesosDetalles.Save(deta) Then
        TareaAgregada = True
        Unload Me
    Else
        MsgBox "Ocurrió un error", vbCritical + vbOKOnly
    End If
End Sub



Private Sub Form_Load()

    TareaAgregada = False
    Customize Me

    Me.dtpFin.value = #12:00:00 AM#
End Sub


Private Sub txtCant_Validate(Cancel As Boolean)
    Cancel = (Not IsNumeric(Me.txtCant.text) And LenB(Me.txtCant.text) <> 0)
End Sub
