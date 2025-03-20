VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmPlaneamientoRemitosListaProceso 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remitos en proceso"
   ClientHeight    =   8055
   ClientLeft      =   3600
   ClientTop       =   2055
   ClientWidth     =   12630
   ClipControls    =   0   'False
   Icon            =   "frmListaRemitosProceso.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   12630
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.GroupBox GroupBox 
      Height          =   1935
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   12375
      _Version        =   786432
      _ExtentX        =   21828
      _ExtentY        =   3413
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin VB.TextBox txtDetalle 
         Height          =   285
         Left            =   1680
         TabIndex        =   9
         Top             =   720
         Width           =   4695
      End
      Begin VB.TextBox txtNumero 
         Height          =   285
         Left            =   1680
         TabIndex        =   8
         Top             =   240
         Width           =   1575
      End
      Begin XtremeSuiteControls.ProgressBar ProgressBar 
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   10215
         _Version        =   786432
         _ExtentX        =   18018
         _ExtentY        =   661
         _StockProps     =   93
         Appearance      =   6
      End
      Begin XtremeSuiteControls.PushButton btnLlenarGrid 
         Height          =   495
         Left            =   10440
         TabIndex        =   7
         Top             =   1380
         Width           =   1815
         _Version        =   786432
         _ExtentX        =   3201
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Mostrar"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.GroupBox GroupBox 
         Height          =   1215
         Index           =   0
         Left            =   6840
         TabIndex        =   12
         Top             =   120
         Width           =   5415
         _Version        =   786432
         _ExtentX        =   9551
         _ExtentY        =   2143
         _StockProps     =   79
         Caption         =   "Fecha"
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.DateTimePicker dtpDesde 
            Height          =   315
            Left            =   855
            TabIndex        =   13
            Top             =   735
            Width           =   1470
            _Version        =   786432
            _ExtentX        =   2593
            _ExtentY        =   556
            _StockProps     =   68
            CheckBox        =   -1  'True
            Format          =   1
         End
         Begin XtremeSuiteControls.DateTimePicker dtpHasta 
            Height          =   315
            Left            =   3030
            TabIndex        =   14
            Top             =   735
            Width           =   1470
            _Version        =   786432
            _ExtentX        =   2593
            _ExtentY        =   556
            _StockProps     =   68
            CheckBox        =   -1  'True
            Format          =   1
         End
         Begin XtremeSuiteControls.ComboBox cboRangos 
            Height          =   315
            Left            =   825
            TabIndex        =   15
            Top             =   240
            Width           =   2190
            _Version        =   786432
            _ExtentX        =   3863
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Style           =   2
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.Label Label7 
            Height          =   195
            Left            =   240
            TabIndex        =   18
            Top             =   300
            Width           =   480
            _Version        =   786432
            _ExtentX        =   847
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Rango"
            AutoSize        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label5 
            Height          =   195
            Left            =   285
            TabIndex        =   17
            Top             =   780
            Width           =   465
            _Version        =   786432
            _ExtentX        =   820
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Desde"
            AutoSize        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label6 
            Height          =   195
            Left            =   2460
            TabIndex        =   16
            Top             =   795
            Width           =   420
            _Version        =   786432
            _ExtentX        =   741
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Hasta"
            AutoSize        =   -1  'True
         End
      End
      Begin XtremeSuiteControls.Label Label 
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   11
         Top             =   720
         Width           =   1335
         _Version        =   786432
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Detalle de Remito:"
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label Label 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1455
         _Version        =   786432
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Número de Remito:"
         Alignment       =   1
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   7080
      Width           =   12375
      _Version        =   786432
      _ExtentX        =   21828
      _ExtentY        =   1508
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.PushButton PushButtonAceptar 
         Height          =   495
         Left            =   10440
         TabIndex        =   3
         Top             =   240
         Width           =   1815
         _Version        =   786432
         _ExtentX        =   3201
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Aceptar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButtonCerrar 
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1695
         _Version        =   786432
         _ExtentX        =   2990
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Cerrar"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   495
         Left            =   1920
         TabIndex        =   4
         Top             =   240
         Width           =   8415
         _Version        =   786432
         _ExtentX        =   14843
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Seleccione el Remito y luego Acepte para continuar."
         BackColor       =   -2147483633
         Alignment       =   2
      End
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   4995
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   12345
      _ExtentX        =   21775
      _ExtentY        =   8811
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      PreviewColumn   =   "estado"
      PreviewRowLines =   1
      ColumnAutoResize=   -1  'True
      MethodHoldFields=   -1  'True
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   6
      Column(1)       =   "frmListaRemitosProceso.frx":000C
      Column(2)       =   "frmListaRemitosProceso.frx":0178
      Column(3)       =   "frmListaRemitosProceso.frx":0264
      Column(4)       =   "frmListaRemitosProceso.frx":0358
      Column(5)       =   "frmListaRemitosProceso.frx":044C
      Column(6)       =   "frmListaRemitosProceso.frx":0578
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmListaRemitosProceso.frx":065C
      FormatStyle(2)  =   "frmListaRemitosProceso.frx":0794
      FormatStyle(3)  =   "frmListaRemitosProceso.frx":0844
      FormatStyle(4)  =   "frmListaRemitosProceso.frx":08F8
      FormatStyle(5)  =   "frmListaRemitosProceso.frx":09D0
      FormatStyle(6)  =   "frmListaRemitosProceso.frx":0A88
      ImageCount      =   0
      PrinterProperties=   "frmListaRemitosProceso.frx":0B68
   End
End
Attribute VB_Name = "frmPlaneamientoRemitosListaProceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vrto As Long
Dim strsql As String
Dim claseP As New classPlaneamiento


Dim Remito As Remito
Dim filtro As String
Dim vmostrar As tipoMostrarRemitos
Dim col As New Collection
Dim vIdCliMostrar As Long

Public Enum tipoMostrarRemitos
    MostrarEnProceso = 0
    MostrarAFacturar = 1
    MostrarPendientes = -1
End Enum


Public Property Let idCliMostrar(nIdCliMostrar)
    vIdCliMostrar = nIdCliMostrar
End Property


Public Property Let mostrar(nmostrar As tipoMostrarRemitos)
    vmostrar = nmostrar
End Property


Private Sub btnLlenarGrid_Click()
    llenarGrid
    
        If Me.GridEX1.ItemCount = 0 Then
        MsgBox "No hay remitos en proceso.", vbInformation
        Set Selecciones.RemitoElegido = Nothing
'        Unload Me
    End If
    
End Sub


Private Sub cboRangos_Click()
    funciones.CalculateDateRange Me.cboRangos, Me.dtpDesde, Me.dtpHasta
End Sub


Private Sub Form_Activate()
'    If Me.GridEX1.ItemCount = 0 Then
'        MsgBox "No hay remitos en proceso.", vbInformation
'        Set Selecciones.RemitoElegido = Nothing
'        Unload Me
'    End If
End Sub


Private Sub Form_Load()
    FormHelper.Customize Me
    GridEXHelper.CustomizeGrid Me.GridEX1
    Me.GridEX1.ItemCount = 0
    
    funciones.FillComboBoxDateRanges Me.cboRangos
    For i = 0 To Me.cboRangos.ListCount - 1
        If Me.cboRangos.ItemData(i) = DateRangeValue.DRV_YearCurrent Then Exit For
    Next i
    Me.cboRangos.ListIndex = i

    Me.caption = "Remitos a procesar [Cantidad: " & col.count & "]"

End Sub


Private Sub Form_Terminate()
'Set Selecciones.RemitoElegido = Nothing
'Unload Me
End Sub


Private Sub llenarGrid()
    filtro = ""
    If vmostrar = 0 Then    'en proceso

        filtro = filtro & " and rto.estado=1  and (rto.estadoFacturado=0 or rto.estadoFacturado=1)"
        Me.caption = "Remitos en proceso..."

    ElseIf vmostrar > 0 Then
        If vIdCliMostrar > 0 Then
            filtro = filtro & " and (rto.estadoFacturado=0 or rto.estadoFacturado=1) and rto.estado=2 and rto.idCliente=" & vIdCliMostrar & ""
        Else
            filtro = filtro & " and (rto.estadoFacturado=0 or rto.estadoFacturado=1) and rto.estado=2 "

        End If
        Me.caption = "Remitos a facturar..."

    ElseIf vmostrar = -1 Then

        filtro = filtro & " and rto.estado=2"
        Me.caption = "Remitos finalizados..."
    End If
        
    If Not IsEmpty(Me.txtNumero) And IsNumeric(txtNumero) Then
        filtro = filtro & "  and " & DAORemitoS.TABLA_REMITO & "." & DAORemitoS.CAMPO_NUMERO & "=" & Me.txtNumero
    End If

    If Not IsNull(Me.dtpDesde.value) Then
        filtro = filtro & " and  " & DAORemitoS.TABLA_REMITO & "." & DAORemitoS.CAMPO_FECHA & " >= " & conectar.Escape(Format(Me.dtpDesde.value, "yyyy-mm-dd 00:00:00"))
    End If

    If Not IsNull(Me.dtpHasta.value) Then
        filtro = filtro & " and  " & DAORemitoS.TABLA_REMITO & "." & DAORemitoS.CAMPO_FECHA & " <= " & conectar.Escape(Format(Me.dtpHasta.value, "yyyy-mm-dd 23:59:59"))
    End If
    
        If LenB(Me.txtDetalle.Text) > 0 Then
        filtro = filtro & " and " & DAORemitoS.TABLA_REMITO & "." & DAORemitoS.CAMPO_DETALLE & " like '%" & Trim(Me.txtDetalle) & "%'"
    End If
        
    Set col = New Collection
    Dim tmpCol As New Collection
    
    Set tmpCol = DAORemitoS.FindAll(filtro)
    
    Dim rto As Remito
    
    Me.ProgressBar.min = 0
    
            Me.ProgressBar.Visible = True
    
    Me.ProgressBar.max = tmpCol.count
    Dim d As Long
    d = 0
    For Each rto In tmpCol
 
        col.Add rto, CStr(rto.Id)
        Set remitoDetalle = DAORemitoSDetalle.FindAllByRemito(rto.Id, True, True)
        Dim deta As remitoDetalle
        For Each deta In remitoDetalle
            If deta.idpedido = 0 Or deta.idpedido = -1 Then
                Set idpedido = Nothing
            Else
                If CStr(" " & deta.idpedido) <> rto.OrigenDeConceptos Then
                    rto.OrigenDeConceptos = rto.OrigenDeConceptos & " " & deta.idpedido
                End If
            End If
        Next
        
        d = d + 1
        Me.ProgressBar.value = d

    Next
    
    If col.count > 0 Then

        Me.GridEX1.ItemCount = 0
        Me.GridEX1.ItemCount = col.count

    End If
    
    Me.ProgressBar.value = 0
    
            Me.GridEX1.ItemCount = col.count

End Sub


Private Sub Form_Unload(Cancel As Integer)
'Set Selecciones.RemitoElegido = Nothing
End Sub


Private Sub GridEX1_DblClick()
    If col.count > 0 Then
        GridEX1_SelectionChange
        vrto = Remito.Id
        Set Selecciones.RemitoElegido = Remito
       Unload Me
    End If
End Sub


Private Sub GridEX1_SelectionChange()
If col.count > 0 Then
    Set Remito = col.item(Me.GridEX1.rowIndex(Me.GridEX1.row))
    End If
End Sub


Private Sub GridEX1_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If col.count > 0 And rowIndex > 0 Then
        Set Remito = col.item(rowIndex)
        Values(1) = Remito.numero
        Values(2) = Remito.FEcha
        Values(3) = Remito.cliente.razon
        Values(4) = Remito.detalle
        Values(5) = Remito.VerEstadoFacturado

        'e5fs52- ACA SE VA A MOSTRAR LOS DATOS DE LAS OTS QUE CONTIENE ESE REMITO
        'Values(6) = "OT N°, OT N°, Concepto, Concepto"
        Values(6) = Remito.OrigenDeConceptos

        ' SE TIENE QUE MOSTRAR COMO REMITO.(PUNTO) + ALGO
        'Values(6) = Remito.
        'If IsSomething(Remito.contacto) Then Values(5) = Remito.contacto.nombre

    End If
End Sub


Private Sub PushButtonAceptar_Click()
    If col.count > 0 Then
        GridEX1_SelectionChange
        vrto = Remito.Id
        Set Selecciones.RemitoElegido = Remito
        Unload Me
    End If
End Sub


Private Sub PushButtonCerrar_Click()
    Set Selecciones.RemitoElegido = Nothing
    Unload Me
End Sub

