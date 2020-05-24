VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~3.OCX"
Begin VB.Form frmTareasAsignadasInicio 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Inicio de Tareas Asignadas"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14955
   Icon            =   "frmTareasAsignadasInicio.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "df"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   14955
   ShowInTaskbar   =   0   'False
   Begin GridEX20.GridEX gridProcesos 
      Height          =   4155
      Left            =   225
      TabIndex        =   17
      Top             =   1380
      Width           =   11235
      _ExtentX        =   19817
      _ExtentY        =   7329
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      TabKeyBehavior  =   1
      MethodHoldFields=   -1  'True
      ContScroll      =   -1  'True
      CardCaptionPrefix=   "Tarea: "
      GroupByBoxVisible=   0   'False
      View            =   1
      ItemCount       =   27
      DataMode        =   99
      CardSpacing     =   6
      CardWidth       =   180
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   3
      Column(1)       =   "frmTareasAsignadasInicio.frx":000C
      Column(2)       =   "frmTareasAsignadasInicio.frx":0138
      Column(3)       =   "frmTareasAsignadasInicio.frx":022C
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmTareasAsignadasInicio.frx":0320
      FormatStyle(2)  =   "frmTareasAsignadasInicio.frx":0458
      FormatStyle(3)  =   "frmTareasAsignadasInicio.frx":0508
      FormatStyle(4)  =   "frmTareasAsignadasInicio.frx":05BC
      FormatStyle(5)  =   "frmTareasAsignadasInicio.frx":0694
      FormatStyle(6)  =   "frmTareasAsignadasInicio.frx":074C
      ImageCount      =   0
      PrinterProperties=   "frmTareasAsignadasInicio.frx":082C
   End
   Begin VB.TextBox txtBackend 
      Height          =   360
      Left            =   5040
      TabIndex        =   16
      Top             =   7080
      Width           =   6270
   End
   Begin XtremeSuiteControls.PushButton cmd1 
      Height          =   900
      Left            =   11880
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   180
      Width           =   900
      _Version        =   786432
      _ExtentX        =   1587
      _ExtentY        =   1587
      _StockProps     =   79
      Caption         =   "1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton cmd4 
      Height          =   900
      Left            =   11880
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1155
      Width           =   900
      _Version        =   786432
      _ExtentX        =   1587
      _ExtentY        =   1587
      _StockProps     =   79
      Caption         =   "4"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton cmd7 
      Height          =   900
      Left            =   11880
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2130
      Width           =   900
      _Version        =   786432
      _ExtentX        =   1587
      _ExtentY        =   1587
      _StockProps     =   79
      Caption         =   "7"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton cmd2 
      Height          =   900
      Left            =   12885
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   180
      Width           =   900
      _Version        =   786432
      _ExtentX        =   1587
      _ExtentY        =   1587
      _StockProps     =   79
      Caption         =   "2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton cmd5 
      Height          =   900
      Left            =   12885
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1155
      Width           =   900
      _Version        =   786432
      _ExtentX        =   1587
      _ExtentY        =   1587
      _StockProps     =   79
      Caption         =   "5"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton cmd8 
      Height          =   900
      Left            =   12885
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2130
      Width           =   900
      _Version        =   786432
      _ExtentX        =   1587
      _ExtentY        =   1587
      _StockProps     =   79
      Caption         =   "8"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton cmd3 
      Height          =   900
      Left            =   13875
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   180
      Width           =   900
      _Version        =   786432
      _ExtentX        =   1587
      _ExtentY        =   1587
      _StockProps     =   79
      Caption         =   "3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton cmd6 
      Height          =   900
      Left            =   13890
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1155
      Width           =   900
      _Version        =   786432
      _ExtentX        =   1587
      _ExtentY        =   1587
      _StockProps     =   79
      Caption         =   "6"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton cmd9 
      Height          =   900
      Left            =   13890
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2130
      Width           =   900
      _Version        =   786432
      _ExtentX        =   1587
      _ExtentY        =   1587
      _StockProps     =   79
      Caption         =   "9"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton CMD0 
      CausesValidation=   0   'False
      Height          =   900
      Left            =   12900
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3120
      Width           =   900
      _Version        =   786432
      _ExtentX        =   1587
      _ExtentY        =   1587
      _StockProps     =   79
      Caption         =   "0"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton cmdEmpleado 
      Height          =   900
      Left            =   11895
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   4545
      Width           =   900
      _Version        =   786432
      _ExtentX        =   1587
      _ExtentY        =   1587
      _StockProps     =   79
      Caption         =   "Legajo"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton cmdEnter 
      Height          =   900
      Left            =   12900
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   4545
      Width           =   1890
      _Version        =   786432
      _ExtentX        =   3334
      _ExtentY        =   1587
      _StockProps     =   79
      Caption         =   "Confirma"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton cmdPunto 
      CausesValidation=   0   'False
      Height          =   900
      Left            =   11880
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3120
      Width           =   900
      _Version        =   786432
      _ExtentX        =   1587
      _ExtentY        =   1587
      _StockProps     =   79
      Caption         =   "."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton btnIniciarTarea 
      Height          =   420
      Left            =   8625
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   885
      Width           =   2835
      _Version        =   786432
      _ExtentX        =   5001
      _ExtentY        =   741
      _StockProps     =   79
      Caption         =   "Iniciar Tarea Seleccionada"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Legajo:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   210
      TabIndex        =   15
      Top             =   870
      Width           =   1185
   End
   Begin XtremeSuiteControls.Label lblMensajes 
      Height          =   840
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   11700
      _Version        =   786432
      _ExtentX        =   20637
      _ExtentY        =   1482
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
   Begin VB.Line Line 
      BorderColor     =   &H00FFDBBF&
      DrawMode        =   9  'Not Mask Pen
      Index           =   0
      X1              =   11430
      X2              =   180
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label lblEmpleado 
      AutoSize        =   -1  'True
      Caption         =   "10 - Raul Carlomagno"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1500
      TabIndex        =   13
      Top             =   885
      Width           =   3060
   End
   Begin VB.Line Line 
      BorderColor     =   &H00FFDBBF&
      DrawMode        =   9  'Not Mask Pen
      Index           =   1
      X1              =   11700
      X2              =   0
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line 
      BorderColor     =   &H00FFDBBF&
      DrawMode        =   9  'Not Mask Pen
      Index           =   2
      X1              =   11700
      X2              =   11700
      Y1              =   0
      Y2              =   5580
   End
End
Attribute VB_Name = "frmTareasAsignadasInicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const CONTECLADO = 15090
Private Const SINTECLADO = 11765
Private Empleado As clsEmpleado

Private procesos As Collection
Private proceso As PlaneamientoTiempoProceso
Private DetalleOt As DetalleOrdenTrabajo

Private tiempoProcesoDet As PlaneamientoTiempoProcesoDetalle

Private scannerBuffer As String
Private fromEmpleado As Boolean

Private finalizandoProceso As Boolean
Private lastkeypressMS As Double

Private tareasAsignadasSinIniciar As Collection



Private Sub MandarTecla(KeyCode As Integer)
    Form_KeyPress KeyCode
    EnfocarTextBox
End Sub

Private Sub btnIniciarTarea_Click()
    gridProcesos_DblClick
End Sub

Private Sub CMD0_Click()
    MandarTecla Asc("0")
End Sub

Private Sub cmd1_Click()
    MandarTecla Asc("1")
End Sub

Private Sub cmd2_Click()
    MandarTecla Asc("2")
End Sub

Private Sub cmd3_Click()
    MandarTecla Asc("3")
End Sub

Private Sub cmd4_Click()
    MandarTecla Asc("4")
End Sub

Private Sub cmd5_Click()
    MandarTecla Asc("5")
End Sub

Private Sub cmd6_Click()
    MandarTecla Asc("6")
End Sub

Private Sub cmd7_Click()
    MandarTecla Asc("7")
End Sub

Private Sub cmd8_Click()
    MandarTecla Asc("8")
End Sub

Private Sub cmd9_Click()
    MandarTecla Asc("9")
End Sub

Private Sub cmdEmpleado_Click()
    MandarTecla Asc("e")
End Sub

Private Sub cmdEnter_Click()
    MandarTecla vbKeyReturn
End Sub

Private Sub cmdPunto_Click()
    MandarTecla Asc(".")
End Sub



Private Sub EnfocarTextBox()
    Me.txtBackend.text = vbNullString
    Me.txtBackend.SetFocus
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)

    'si la ultima tecla fue hace mas de 10 segundos, empezar de vuelta
    If (GetTickCount - lastkeypressMS) > 30000 Then
        LimpiarData
        finalizandoProceso = False
    End If

    lastkeypressMS = GetTickCount

    If KeyAscii = vbKeyReturn Then
        'perform action

        If finalizandoProceso Then
            FinalizarProcesoPost
        Else
            If LenB(scannerBuffer) > 0 Then
                If StrConv(Left(scannerBuffer, 1), vbUpperCase) = "E" Then
                    GetEmpleado
                Else
                    If IsSomething(Empleado) Then
                        fromEmpleado = False
                        GetTareasAsignadasSinIniciar
                    End If
                End If
            Else
                LimpiarData
                finalizandoProceso = False
            End If

            scannerBuffer = vbNullString
        End If
    Else
        scannerBuffer = scannerBuffer + Chr(KeyAscii)
    End If

End Sub

Private Sub GetTareasAsignadasSinIniciar()
    Dim tmpTiempoProcesoDetalle As PlaneamientoTiempoProcesoDetalle


    Set tareasAsignadasSinIniciar = DAOTiemposProcesosDetalles.FindAllAsignedNotFinishedByEmpleado(Empleado.id)


    Me.gridProcesos.ItemCount = 0
    If tareasAsignadasSinIniciar.count > 0 Then
        ShowMessage "Elija una tarea para iniciar"

        Me.gridProcesos.ItemCount = tareasAsignadasSinIniciar.count
        Me.gridProcesos.row = -1
        Set proceso = Nothing
    Else
        ShowMessage "No tiene tareas asignadas sin finalizar"
    End If



    'For Each tmpTiempoProcesoDetalle In tareasAsignadasSinIniciar

    'Next tmpTiempoProcesoDetalle


    '    Dim tmpId As Long
    '    If Not fromEmpleado Then
    '        LimpiarDetalleOT
    '        'truncamos por overflow
    '
    '        scannerBuffer = Replace$(scannerBuffer, "e", vbNullString)
    '        scannerBuffer = Replace$(scannerBuffer, "E", vbNullString)
    '        scannerBuffer = Replace$(scannerBuffer, ".", vbNullString)
    '
    '        If Len(scannerBuffer) > 7 Then
    '            tmpId = Val(Right(scannerBuffer, 7))
    '        Else
    '            tmpId = Val(scannerBuffer)
    '        End If
    '
    '
    ''        Set DetalleOt = DAODetalleOrdenTrabajo.FindById(tmpId, False, False, False)
    ''        Set procesos = New Collection
    ''        If IsSomething(DetalleOt) Then
    ''            Set procesos = DAOTiemposProceso.FindAllByDetallePedidoIdAndPiezaId(DetalleOt.Id, DetalleOt.pieza.Id)
    ''
    ''
    ''            'ver esto para terminar el proceso
    ''            Dim procesosDetallesSinTerminar As Collection
    ''            Set procesosDetallesSinTerminar = DAOTiemposProcesosDetalles.FindAllWithoutFinishByEmpleado(Empleado.Id)
    ''
    ''            Dim ptpd As PlaneamientoTiempoProcesoDetalle
    ''            Dim tmpProc As PlaneamientoTiempoProceso
    ''            If IsSomething(procesosDetallesSinTerminar) Then
    ''                For Each ptpd In procesosDetallesSinTerminar
    ''                    For Each tmpProc In procesos
    ''                        If tmpProc.Id = ptpd.IdPlaneamientoTiempoProceso Then
    ''                            'tiene un detalle para terminar!
    ''                              Set tiempoProcesoDet = ptpd
    ''                              FinalizarProceso
    ''                              Exit For
    ''                        End If
    ''                    Next tmpProc
    ''                Next ptpd
    ''            End If
    ''
    ''        End If
    '
    '    End If

    '    If finalizandoProceso Then Exit Sub
    '
    '
    '    If IsSomething(DetalleOt) Then
    '        Me.gridProcesos.ItemCount = 0
    '        If procesos.count > 0 Then
    '            ShowMessage "Elija una tarea para iniciar"
    '
    '            Me.gridProcesos.ItemCount = procesos.count
    '            Me.gridProcesos.row = -1
    '            Set proceso = Nothing
    '        Else
    '            ShowMessage "La ruta no posee tareas"
    '        End If
    '    Else
    '        LimpiarDetalleOT
    '        If tmpId = 0 Then
    '            ShowMessage "No hay ruta seleccionada"
    '        Else
    '            ShowMessage "La ruta Nº " & tmpId & " no existe"
    '        End If
    '    End If


End Sub

Private Sub ClearEmpleadoProcesoObject()
    Set Empleado = Nothing
    Set proceso = Nothing
    Set procesos = New Collection
    Set tiempoProcesoDet = Nothing
End Sub

Private Sub TieneTarea(det As PlaneamientoTiempoProcesoDetalle)
    Set det.PlaneamientoTiempoProceso = DAOTiemposProceso.FindById(det.IdPlaneamientoTiempoProceso)
    ShowMessage "Ya tiene una tarea iniciada de (" & det.PlaneamientoTiempoProceso.Tarea.Description & ") el " & det.FechaInicioTarea
End Sub

Private Sub FinalizarProceso()
    ShowMessage "Ingrese la cantidad procesada"
    finalizandoProceso = True
End Sub


Private Sub FinalizarProcesoPost()
    tiempoProcesoDet.FechaFinTarea = Now
    Dim Cant As String
    Cant = scannerBuffer
    scannerBuffer = vbNullString
    If LenB(Cant) = 0 Or Not IsNumeric(Cant) Then
        ShowMessage "Debe ingresar la cantidad para finalizar la tarea"
    Else
        If MsgBox("La cantidad ingresada es " & Cant & vbNewLine & "¿Ese valor es correcto?", vbQuestion + vbYesNo) = vbYes Then
            tiempoProcesoDet.CantidadProcesada = Val(Cant)
            If DAOTiemposProcesosDetalles.Save(tiempoProcesoDet) Then
                ClearEmpleadoProcesoObject
                ShowMessage "La tarea ha sido finalizada (Duración: " & tiempoProcesoDet.DiferenciaTiempoHorasMinutos & ")"
            Else
                ShowMessage "Hubo un error al finalizar la tarea"
            End If
            finalizandoProceso = False
        Else
            ShowMessage "Reingrese la cantidad para finalizar la tarea"
        End If
    End If

End Sub

Private Sub IniciarProceso()


    If DAOEmpleados.GetTareasIdAsignadasByPersonalId(Empleado.id).Exists(tiempoProcesoDet.PlaneamientoTiempoProceso.Tarea.id) Then

        'me tengo que fijar que de los que tenga sin terminar sean el mismo tipo de tarea que la que va a iniciar
        Dim detallesSinTerminar As Collection
        Set detallesSinTerminar = DAOTiemposProcesosDetalles.FindAllWithoutFinishByEmpleado(Empleado.id)
        Dim tmpProcesoDet As PlaneamientoTiempoProcesoDetalle
        Dim tmpProcesoDetInconcluso As PlaneamientoTiempoProcesoDetalle

        Dim mismaTarea As Boolean: mismaTarea = True
        Dim tareaAsignadaYaIniciada As Boolean: tareaAsignadaYaIniciada = False
        If IsSomething(detallesSinTerminar) Then
            For Each tmpProcesoDet In detallesSinTerminar
                tareaAsignadaYaIniciada = tareaAsignadaYaIniciada Or (tmpProcesoDet.IdPlaneamientoTiempoProceso = tiempoProcesoDet.PlaneamientoTiempoProceso.id)

                mismaTarea = mismaTarea And (tiempoProcesoDet.PlaneamientoTiempoProceso.Tarea.id = tmpProcesoDet.PlaneamientoTiempoProceso.Tarea.id)
                If Not mismaTarea Then
                    Set tmpProcesoDetInconcluso = tmpProcesoDet
                    Exit For
                End If
            Next
        End If

        Dim newProcesoDet As New PlaneamientoTiempoProcesoDetalle

        If mismaTarea Then
            If tareaAsignadaYaIniciada Then
                ShowMessage "Esa tarea asignada ya esta iniciada"
            Else
                Set newProcesoDet.Empleado = Empleado
                newProcesoDet.FechaCarga = Now
                newProcesoDet.FechaInicioTarea = Now
                newProcesoDet.IdPlaneamientoTiempoProceso = tiempoProcesoDet.PlaneamientoTiempoProceso.id
                newProcesoDet.legajo = Empleado.legajo
                If DAOTiemposProcesosDetalles.Save(newProcesoDet) Then
                    Me.gridProcesos.ItemCount = 0
                    ShowMessage "La tarea [" & tiempoProcesoDet.PlaneamientoTiempoProceso.Tarea.Description & "] ha sido iniciada"
                    ClearEmpleadoProcesoObject
                Else
                    ShowMessage "Hubo un error al iniciar la tarea"
                End If
            End If

        Else
            ShowMessage "Ya tiene una tarea iniciada de (" & tmpProcesoDetInconcluso.PlaneamientoTiempoProceso.Tarea.Description & ") el " & tmpProcesoDetInconcluso.FechaInicioTarea
        End If
    Else
        ShowMessage "No puede realizar la tarea (" & tiempoProcesoDet.PlaneamientoTiempoProceso.Tarea.Description & ")"
    End If
End Sub
Private Sub CargarInfoRuta()
    If IsSomething(proceso) Then

        If Not proceso.EsConjunto Then
            Set proceso.DetalleOt = DAODetalleOrdenTrabajo.FindById(proceso.idDetallePedido)
            'Me.lblItem.caption = proceso.DetalleOt.item
            ' Me.lblPieza.caption = proceso.DetalleOt.pieza.nombre
        Else
            Set proceso.DetalleOtConjunto = DAODetalleOrdenTrabajo.FindConjuntoById(proceso.idDetallePedido)
            ' Me.lblItem.caption = vbNullString    'proceso.DetalleOtConjunto.
            '  Me.lblPieza.caption = proceso.DetalleOtConjunto.pieza.nombre
        End If

        'Me.lblTiempoProcesoID.caption = proceso.id
        ' Me.lblOT.caption = proceso.idpedido'

        'Me.lblTarea.caption = proceso.Tarea.'Description



    End If
End Sub

Private Sub Form_Load()
    Me.gridProcesos.ItemCount = 0

    Me.Width = CONTECLADO
    Customize Me
    LimpiarData

    GridEXHelper.CustomizeGrid Me.gridProcesos

End Sub

Private Sub LimpiarDetalleOT()
    Set DetalleOt = Nothing
End Sub

Private Sub LimpiarMensaje()
    Me.lblMensajes.caption = vbNullString
End Sub

Private Sub ShowMessage(msg As String)
    PintarMensaje
    DoEvents
    Me.lblMensajes.caption = msg
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

    LimpiarDetalleOT
    LimpiarMensaje
    LimpiarEmpleado


    Set tiempoProcesoDet = Nothing
    Me.gridProcesos.ItemCount = 0
End Sub


Private Sub GetEmpleado()
    Dim leg As String
    leg = Right(scannerBuffer, Len(scannerBuffer) - 1)
    leg = Val(leg)
    Set Empleado = DAOEmpleados.GetByLegajo(leg)

    LimpiarMensaje

    If IsSomething(Empleado) Then
        Me.lblEmpleado.caption = Empleado.legajo & " - " & Empleado.NombreCompleto
        LimpiarDetalleOT
        'If IsSomething(DetalleOt) Then
        fromEmpleado = True
        GetTareasAsignadasSinIniciar
        'Else
        'ShowMessage "Ahora ingrese ruta"
        'End If
    Else
        LimpiarEmpleado
        ShowMessage "El legajo no existe"
    End If
End Sub


Private Sub gridProcesos_DblClick()
    If (GetTickCount - lastkeypressMS) > 30000 Then
        LimpiarData
        finalizandoProceso = False
        ShowMessage "Tiempo agotado"
    Else
        If Me.gridProcesos.RowIndex(Me.gridProcesos.row) > 0 And tareasAsignadasSinIniciar.count > 0 Then
            Set tiempoProcesoDet = tareasAsignadasSinIniciar.item(Me.gridProcesos.RowIndex(Me.gridProcesos.row))
            IniciarProceso
        End If
    End If
End Sub

Private Sub gridProcesos_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex >= 1 And RowIndex <= tareasAsignadasSinIniciar.count Then
        Set tiempoProcesoDet = tareasAsignadasSinIniciar.item(RowIndex)
        Values.value(1) = tiempoProcesoDet.PlaneamientoTiempoProceso.Tarea.id & " - " & tiempoProcesoDet.PlaneamientoTiempoProceso.Tarea.Tarea
        Values.value(2) = tiempoProcesoDet.PlaneamientoTiempoProceso.Tarea.Sector.Sector
        Values.value(3) = tiempoProcesoDet.PlaneamientoTiempoProceso.idpedido & "/" & Format(tiempoProcesoDet.PlaneamientoTiempoProceso.item, "000")
    End If
End Sub

Private Sub txtBackend_Change()
    Me.txtBackend.text = vbNullString
End Sub

