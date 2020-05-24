VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~3.OCX"
Begin VB.Form frmTiempoProcesoDetalle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Carga de Tiempo"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12765
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTiempoProcesoDetalle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   12765
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   2880
      Left            =   135
      TabIndex        =   5
      Top             =   2535
      Width           =   12435
      _Version        =   786432
      _ExtentX        =   21934
      _ExtentY        =   5080
      _StockProps     =   79
      Caption         =   "Info Proceso"
      ForeColor       =   9126421
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Tarea:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   285
         TabIndex        =   13
         Top             =   2205
         Width           =   1140
      End
      Begin VB.Label lblTarea 
         AutoSize        =   -1  'True
         Caption         =   "12 - Corte Chapa"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1530
         TabIndex        =   12
         Top             =   2220
         Width           =   2730
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Pieza:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   285
         TabIndex        =   11
         Top             =   1650
         Width           =   1095
      End
      Begin VB.Label lblPieza 
         AutoSize        =   -1  'True
         Caption         =   "bbañbñañbabkjsasdaasdasd"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1515
         TabIndex        =   10
         Top             =   1650
         Width           =   4455
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Item:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   285
         TabIndex        =   9
         Top             =   1095
         Width           =   1020
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         Caption         =   "999"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1470
         TabIndex        =   8
         Top             =   1110
         Width           =   585
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "OT:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   285
         TabIndex        =   7
         Top             =   540
         Width           =   630
      End
      Begin VB.Label lblOT 
         AutoSize        =   -1  'True
         Caption         =   "999999"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1065
         TabIndex        =   6
         Top             =   540
         Width           =   1170
      End
   End
   Begin VB.Line Line 
      BorderColor     =   &H00FFDBBF&
      DrawMode        =   9  'Not Mask Pen
      Index           =   1
      X1              =   12765
      X2              =   -15
      Y1              =   825
      Y2              =   825
   End
   Begin VB.Label lblTiempoProcesoID 
      AutoSize        =   -1  'True
      Caption         =   "999999"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2550
      TabIndex        =   4
      Top             =   1980
      Width           =   1170
   End
   Begin VB.Label lblEmpleado 
      AutoSize        =   -1  'True
      Caption         =   "10 - Raul Carlomagno"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1695
      TabIndex        =   3
      Top             =   1035
      Width           =   3465
   End
   Begin VB.Line Line 
      BorderColor     =   &H00FFDBBF&
      DrawMode        =   9  'Not Mask Pen
      Index           =   0
      X1              =   12525
      X2              =   195
      Y1              =   1695
      Y2              =   1695
   End
   Begin XtremeSuiteControls.Label lblMensajes 
      Height          =   825
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   12750
      _Version        =   786432
      _ExtentX        =   22490
      _ExtentY        =   1455
      _StockProps     =   79
      Caption         =   "El legajo no existe"
      ForeColor       =   9126421
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Nº Proceso:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   225
      TabIndex        =   1
      Top             =   1965
      Width           =   2145
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Legajo:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   225
      TabIndex        =   0
      Top             =   1020
      Width           =   1350
   End
End
Attribute VB_Name = "frmTiempoProcesoDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Empleado As clsEmpleado
Private proceso As PlaneamientoTiempoProceso
Private detalle As PlaneamientoTiempoProcesoDetalle

Private scannerBuffer As String
Private fromEmpleado As Boolean






Private Sub Form_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    'perform action
    'Debug.Print scannerBuffer
    
    If StrConv(Left(scannerBuffer, 1), vbUpperCase) = "E" Then
        GetEmpleado
    Else
        fromEmpleado = False
        GetProceso
    End If
    
    scannerBuffer = vbNullString
Else
    scannerBuffer = scannerBuffer + Chr(KeyAscii)
End If

End Sub

Private Sub GetProceso()
    Dim detallesNoFinalizados As Collection

    If Not fromEmpleado Then
        Set proceso = DAOTiemposProceso.FindById(Val(IIf(Len(scannerBuffer) > 8, Left(scannerBuffer, 8), scannerBuffer)))
    End If

    
    
    If IsSomething(proceso) Then
        CargarInfoProceso
        If Not proceso.Finalizado Then
            If IsSomething(Empleado) Then
                Set detalle = DAOTiemposProcesosDetalles.FindFirstWithoutFinishByEmpleadoIdAndTiempoProceso(Empleado.id, proceso.id)
                If IsSomething(detalle) Then
                    detalle.FechaFinTarea = Now
                    Dim Cant As String
                    Cant = InputBox("Ingrese la cantidad procesada para finalizar la tarea")
                    If LenB(Cant) = 0 Or Not IsNumeric(Cant) Then
                        ShowMessage "Debe ingresar la cantidad para finalizar la tarea"
                    Else
                        detalle.CantidadProcesada = Val(Cant)
                        If DAOTiemposProcesosDetalles.Save(detalle) Then
                            ShowMessage "La tarea ha sido finalizada (Duración: " & detalle.DiferenciaTiempoHorasMinutos & ")"
                        Else
                            ShowMessage "Hubo un error al finalizar la tarea"
                        End If
                    End If
                    If detalle.PlaneamientoTiempoProceso.tarea.id = proceso.tarea.id Then
                        GoTo inicia
                    Else
                        ShowMessage "Ya tiene una tarea iniciada de (" & detalle.PlaneamientoTiempoProceso.tarea.Description & ") el " & detalle.FechaInicioTarea
                    End If
                Else

               End If
            Else
                'puede iniciar
inicia:
                If DAOEmpleados.GetTareasIdAsignadasByPersonalId(Empleado.id).Exists(proceso.tarea.id) Then

                    Set detalle = New PlaneamientoTiempoProcesoDetalle
                    Set detalle.Empleado = Empleado
                    detalle.FechaCarga = Now
                    detalle.FechaInicioTarea = Now
                    detalle.IdPlaneamientoTiempoProceso = proceso.id
                    detalle.legajo = Empleado.legajo
                    If DAOTiemposProcesosDetalles.Save(detalle) Then
                        ShowMessage "La tarea ha sido iniciada"
                    Else
                        ShowMessage "Hubo un error al iniciar la tarea"
                    End If
                Else
                    ShowMessage "No puede realizar la tarea (" & proceso.tarea.Description & ")"
                End If
            End If
        Else
            ShowMessage "Ahora ingrese legajo"
        End If

        Else
            ShowMessage "El proceso ya esta finalizado"
        End If

Else
    LimpiarProceso
    ShowMessage "El proceso no existe"
End If

End Sub

Private Sub CargarInfoProceso()
    If IsSomething(proceso) Then
        
        Set proceso.DetalleOt = DAODetalleOrdenTrabajo.FindById(proceso.idDetallePedido)
        Me.lblTiempoProcesoID.caption = proceso.id
        Me.lblOT.caption = proceso.idpedido
        Me.lblItem.caption = proceso.DetalleOt.Item
        Me.lblPieza.caption = proceso.DetalleOt.pieza.nombre
        Me.lblTarea.caption = proceso.tarea.Description
        
    End If
End Sub

Private Sub Form_Load()
Customize Me
LimpiarData
End Sub

Private Sub LimpiarProceso()
Me.lblTiempoProcesoID.caption = vbNullString

Me.lblOT.caption = vbNullString
Me.lblItem.caption = vbNullString
Me.lblPieza.caption = vbNullString
Me.lblTarea.caption = vbNullString

Set proceso = Nothing
End Sub

Private Sub LimpiarMensaje()
Me.lblMensajes.caption = vbNullString
End Sub

Private Sub ShowMessage(MSG As String)
PintarMensaje
DoEvents
Me.lblMensajes.caption = MSG
DoEvents
Sleep 500
PintarMensaje
DoEvents
End Sub

Private Sub PintarMensaje()
If Me.lblMensajes.BackColor = vbWhite Then
    Me.lblMensajes.BackColor = FormHelper.LetraAzul
    Me.lblMensajes.ForeColor = vbWhite
Else
    Me.lblMensajes.BackColor = vbWhite
    Me.lblMensajes.ForeColor = FormHelper.LetraAzul
End If
End Sub

Private Sub LimpiarEmpleado()
Me.lblEmpleado.caption = vbNullString
Set Empleado = Nothing
End Sub

Private Sub LimpiarData()

LimpiarProceso
LimpiarMensaje
LimpiarEmpleado

Set detalle = Nothing
End Sub


Private Sub GetEmpleado()
    Dim leg As String
    leg = Right(scannerBuffer, Len(scannerBuffer) - 1)
    leg = Val(leg)
    Set Empleado = DAOEmpleados.GetByLegajo(leg)
    
    LimpiarMensaje
    
    If IsSomething(Empleado) Then
        Me.lblEmpleado.caption = Empleado.legajo & " - " & Empleado.NombreCompleto
        If IsSomething(proceso) Then
            fromEmpleado = True
            GetProceso
        Else
            ShowMessage "Ahora ingrese proceso"
        End If
    Else
        LimpiarEmpleado
        ShowMessage "El legajo no existe"
    End If
End Sub




