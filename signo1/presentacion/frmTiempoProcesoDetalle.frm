VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmTiempoProcesoDetalle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Carga y Descarga de Tareas"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15015
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   15015
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtBackend 
      Height          =   360
      Left            =   3165
      TabIndex        =   28
      Top             =   5955
      Width           =   6270
   End
   Begin XtremeSuiteControls.PushButton cmd1 
      Height          =   900
      Left            =   11895
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   165
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
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   2880
      Left            =   165
      TabIndex        =   5
      Top             =   2535
      Width           =   11340
      _Version        =   786432
      _ExtentX        =   20002
      _ExtentY        =   5080
      _StockProps     =   79
      Caption         =   "Info Tarea"
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
      Appearance      =   6
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
   Begin XtremeSuiteControls.PushButton cmd4 
      Height          =   900
      Left            =   11895
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1140
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
      Left            =   11895
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   2115
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
      Left            =   12900
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   165
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
      Left            =   12900
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1140
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
      Left            =   12900
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   2115
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
      Left            =   13890
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   165
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
      Left            =   13905
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1140
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
      Left            =   13905
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   2115
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
      Left            =   12915
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   3105
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
      Left            =   11910
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   4530
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
      Left            =   12915
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   4530
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
   Begin XtremeSuiteControls.PushButton cmdTeclado 
      Height          =   585
      Left            =   9915
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   960
      Width           =   1545
      _Version        =   786432
      _ExtentX        =   2725
      _ExtentY        =   1032
      _StockProps     =   79
      Caption         =   "Teclado ->"
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
      Left            =   11895
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   3105
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
   Begin VB.Line Line 
      BorderColor     =   &H00FFDBBF&
      DrawMode        =   9  'Not Mask Pen
      Index           =   2
      X1              =   11670
      X2              =   11670
      Y1              =   -15
      Y2              =   5565
   End
   Begin VB.Line Line 
      BorderColor     =   &H00FFDBBF&
      DrawMode        =   9  'Not Mask Pen
      Index           =   1
      X1              =   11670
      X2              =   -30
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
      Left            =   2070
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
      X1              =   11445
      X2              =   195
      Y1              =   1695
      Y2              =   1695
   End
   Begin XtremeSuiteControls.Label lblMensajes 
      Height          =   825
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11655
      _Version        =   786432
      _ExtentX        =   20558
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
      Caption         =   "Nº Tarea:"
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
      Left            =   165
      TabIndex        =   1
      Top             =   1965
      Width           =   1725
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
Private Const CONTECLADO = 15090
Private Const SINTECLADO = 11765
Private Empleado As clsEmpleado
Private proceso As PlaneamientoTiempoProceso
Private det As PlaneamientoTiempoProcesoDetalle
Private scannerBuffer As String
Private fromEmpleado As Boolean

Private finalizandoProceso As Boolean
Private lastkeypressMS As Double


Private Sub MandarTecla(KeyCode As Integer)
'Form_KeyDown keycode, 0
    Form_KeyPress KeyCode
    EnfocarTextBox
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

Private Sub cmdTeclado_Click()
    If Me.Width = CONTECLADO Then
        Me.Width = SINTECLADO
    Else
        Me.Width = CONTECLADO
    End If
    EnfocarTextBox
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

    'Debug.Print KeyAscii, Chr(KeyAscii)
    If KeyAscii = vbKeyReturn Then
        'perform action
        'Debug.Print scannerBuffer

        If finalizandoProceso Then
            FinalizarProcesoPost
        Else
            If LenB(scannerBuffer) > 0 Then
                If StrConv(Left(scannerBuffer, 1), vbUpperCase) = "E" Then
                    GetEmpleado
                Else
                    fromEmpleado = False
                    GetProceso
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

Private Sub GetProceso()

    Dim tmpId As Long
    If Not fromEmpleado Then
        'LimpiarProceso
        'truncamos por overflow

        scannerBuffer = Replace$(scannerBuffer, "e", vbNullString)
        scannerBuffer = Replace$(scannerBuffer, "E", vbNullString)
        scannerBuffer = Replace$(scannerBuffer, ".", vbNullString)

        If Len(scannerBuffer) > 7 Then
            tmpId = Val(Right(scannerBuffer, 7))
        Else
            tmpId = Val(scannerBuffer)
        End If


        Set proceso = DAOTiemposProceso.FindById(tmpId)
    End If

    If IsSomething(proceso) Then

        CargarInfoProceso

        If Not proceso.FINALIZADO Then    'tambien verificar si la ot a la que pertenece el proceso no esta en curso?
            If IsSomething(Empleado) Then
                Dim col As New Collection

                Dim esta As Boolean
                Dim mismoDeta As Boolean
                esta = False
                mismoDeta = False
                Dim deta As PlaneamientoTiempoProcesoDetalle

                Set col = DAOTiemposProcesosDetalles.FindAllWithoutFinishByEmpleado(Empleado.Id)


                For Each deta In col
                    If proceso.Tarea.Id = deta.PlaneamientoTiempoProceso.Tarea.Id Then
                        esta = True
                        Exit For
                    End If
                Next


                For Each deta In col
                    If proceso.Id = deta.IdPlaneamientoTiempoProceso Then
                        mismoDeta = True
                        Set det = deta
                        Exit For
                    End If

                Next

                If esta And Not mismoDeta Then      'proceso.id <> detalle.IdPlaneamientoTiempoProceso Then
                    'si la tarea es la misma y es de otra orden, puede ahcer tareas paralelas
                    'hay q verificar q ya no esté iniciada

                    IniciarProceso

                Else

                    If mismoDeta And esta Then    'si es el mismo deta y la misma tarea, indica q esta iniciada
                        FinalizarProceso

                    Else

                        If col.count > 0 Then
                            If Not esta Then
                                TieneTarea col.item(1)
                            Else
                                IniciarProceso
                            End If
                        Else
                            IniciarProceso
                        End If
                        'If Not esta Then
                        '    'If col.count > 1 Then
                        '    'Else
                        '     If col.count > 0 Then
                        '
                        '        TieneTarea det
                        '     Else
                        '        IniciarProceso
                        '    End If
                        '    'End If


                        'Else
                        '   TieneTarea det
                        'End If
                    End If

                End If
            Else
                LimpiarEmpleado
                ShowMessage "Ahora ingrese legajo"
            End If
        Else
            ShowMessage "La tarea ya esta cerrada o finalizada"
        End If

    Else
        LimpiarProceso
        If tmpId = 0 Then
            ShowMessage "No hay tarea seleccionada"
        Else
            ShowMessage "La tarea Nº " & tmpId & " no existe"
        End If
    End If


End Sub

Private Sub ClearEmpleadoProcesoObject()
    Set Empleado = Nothing
    Set proceso = Nothing
End Sub



Private Sub TieneTarea(det As PlaneamientoTiempoProcesoDetalle)
    Set det.PlaneamientoTiempoProceso = DAOTiemposProceso.FindById(det.IdPlaneamientoTiempoProceso)
    ShowMessage "Ya tiene una tarea iniciada de (" & det.PlaneamientoTiempoProceso.Tarea.Description & ") el " & det.FechaInicioTarea
End Sub

Private Sub FinalizarProceso()
'    det.FechaFinTarea = Now
'    Dim Cant As String
'    Cant = InputBox("Ingrese la cantidad procesada para finalizar la tarea")
'    If LenB(Cant) = 0 Or Not IsNumeric(Cant) Then
'        ShowMessage "Debe ingresar la cantidad para finalizar la tarea"
'    Else
'        det.CantidadProcesada = Val(Cant)
'        If DAOTiemposProcesosDetalles.Save(det) Then
'            ShowMessage "La tarea ha sido finalizada (Duración: " & det.DiferenciaTiempoHorasMinutos & ")"
'        Else
'            ShowMessage "Hubo un error al finalizar la tarea"
'        End If
'    End If





    ShowMessage "Ingrese la cantidad procesada"
    finalizandoProceso = True

End Sub


Private Sub FinalizarProcesoPost()
    det.FechaFinTarea = Now
    Dim Cant As String
    Cant = scannerBuffer
    scannerBuffer = vbNullString
    If LenB(Cant) = 0 Or Not IsNumeric(Cant) Then
        ShowMessage "Debe ingresar la cantidad para finalizar la tarea"
    Else
        If MsgBox("La cantidad ingresada es " & Cant & vbNewLine & "¿Ese valor es correcto?", vbQuestion + vbYesNo) = vbYes Then
            det.CantidadProcesada = Val(Cant)
            If DAOTiemposProcesosDetalles.Save(det) Then
                ClearEmpleadoProcesoObject
                ShowMessage "La tarea ha sido finalizada (Duración: " & det.DiferenciaTiempoHorasMinutos & ")"
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
    If DAOEmpleados.GetTareasIdAsignadasByPersonalId(Empleado.Id).Exists(proceso.Tarea.Id) Then

        Set det = New PlaneamientoTiempoProcesoDetalle
        Set det.Empleado = Empleado
        det.FechaCarga = Now
        det.FechaInicioTarea = Now
        det.IdPlaneamientoTiempoProceso = proceso.Id
        det.legajo = Empleado.legajo
        If DAOTiemposProcesosDetalles.Save(det) Then
            ClearEmpleadoProcesoObject
            ShowMessage "La tarea ha sido iniciada"
        Else
            ShowMessage "Hubo un error al iniciar la tarea"
        End If
        'End If
    Else
        ShowMessage "No puede realizar la tarea (" & proceso.Tarea.Description & ")"
    End If
End Sub
Private Sub CargarInfoProceso()
    If IsSomething(proceso) Then

        If proceso.idDetallePedidoConj = 0 Then
            Set proceso.DetalleOt = DAODetalleOrdenTrabajo.FindById(proceso.idDetallePedido)
            Me.lblItem.caption = proceso.DetalleOt.item
            Me.lblPieza.caption = proceso.DetalleOt.Pieza.nombre
        Else
            Set proceso.DetalleOtConjunto = DAODetalleOrdenTrabajo.FindConjuntoById(proceso.idDetallePedidoConj)    'mal
            Me.lblItem.caption = vbNullString    'proceso.DetalleOtConjunto.
            Me.lblPieza.caption = proceso.DetalleOtConjunto.Pieza.nombre
        End If

        Me.lblTiempoProcesoID.caption = proceso.Id
        Me.lblOT.caption = proceso.idpedido
        Me.lblTarea.caption = proceso.Tarea.Description
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
    If Me.lblMensajes.backColor = vbWhite Then
        Me.lblMensajes.backColor = FormHelper.LetraAzul
        Me.lblMensajes.ForeColor = vbWhite
    Else
        Me.lblMensajes.backColor = vbWhite
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

    Set det = Nothing
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
            'LimpiarProceso
            GetProceso
        Else
            LimpiarProceso
            ShowMessage "Ahora ingrese proceso"
        End If
    Else
        LimpiarEmpleado
        ShowMessage "El legajo no existe"
    End If
End Sub

Private Sub txtBackend_Change()
    Me.txtBackend.text = vbNullString
End Sub
