VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmAdminSubdiariosVentasv2 
   Caption         =   "Subdiario de Ventas"
   ClientHeight    =   8625
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13800
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAdminSubdiariosVentasv2.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8625
   ScaleWidth      =   13800
   Begin XtremeSuiteControls.GroupBox grpTotales 
      Height          =   1680
      Left            =   10545
      TabIndex        =   16
      Top             =   6765
      Width           =   3090
      _Version        =   786432
      _ExtentX        =   5450
      _ExtentY        =   2963
      _StockProps     =   79
      Caption         =   "Totales"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.Label lblTotalTotal 
         Height          =   195
         Left            =   1785
         TabIndex        =   26
         Top             =   1380
         Width           =   1155
         _Version        =   786432
         _ExtentX        =   2037
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   ".-"
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label lblExentoTotal 
         Height          =   195
         Left            =   1785
         TabIndex        =   25
         Top             =   1020
         Width           =   1155
         _Version        =   786432
         _ExtentX        =   2037
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   ".-"
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label lblPercepcionesTotal 
         Height          =   195
         Left            =   1785
         TabIndex        =   24
         Top             =   750
         Width           =   1155
         _Version        =   786432
         _ExtentX        =   2037
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   ".-"
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label lblIVATotal 
         Height          =   195
         Left            =   1785
         TabIndex        =   23
         Top             =   495
         Width           =   1155
         _Version        =   786432
         _ExtentX        =   2037
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   ".-"
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label lblNetoGravadoTotal 
         Height          =   195
         Left            =   1785
         TabIndex        =   22
         Top             =   225
         Width           =   1155
         _Version        =   786432
         _ExtentX        =   2037
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   ".-"
         Alignment       =   1
      End
      Begin VB.Line Line 
         BorderColor     =   &H00FFDBBF&
         DrawMode        =   9  'Not Mask Pen
         X1              =   2955
         X2              =   135
         Y1              =   1305
         Y2              =   1305
      End
      Begin XtremeSuiteControls.Label lblTotal 
         Height          =   195
         Left            =   180
         TabIndex        =   21
         Top             =   1380
         Width           =   420
         _Version        =   786432
         _ExtentX        =   741
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Total:"
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblExento 
         Height          =   195
         Left            =   180
         TabIndex        =   20
         Top             =   1020
         Width           =   570
         _Version        =   786432
         _ExtentX        =   1005
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Exento:"
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblPercepcionesIIBB 
         Height          =   195
         Left            =   180
         TabIndex        =   19
         Top             =   750
         Width           =   1350
         _Version        =   786432
         _ExtentX        =   2381
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Percepciones IIBB:"
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblIVA 
         Height          =   195
         Left            =   180
         TabIndex        =   18
         Top             =   495
         Width           =   315
         _Version        =   786432
         _ExtentX        =   556
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "IVA:"
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblNetoGravado 
         Height          =   195
         Left            =   165
         TabIndex        =   17
         Top             =   225
         Width           =   1065
         _Version        =   786432
         _ExtentX        =   1879
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Neto Gravado:"
         AutoSize        =   -1  'True
      End
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   4890
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   8625
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      ColumnAutoResize=   -1  'True
      MultiSelect     =   -1  'True
      MethodHoldFields=   -1  'True
      RowHeaders      =   -1  'True
      DataMode        =   99
      HeaderFontName  =   "MS Sans Serif"
      FontName        =   "MS Sans Serif"
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   10
      Column(1)       =   "frmAdminSubdiariosVentasv2.frx":000C
      Column(2)       =   "frmAdminSubdiariosVentasv2.frx":011C
      Column(3)       =   "frmAdminSubdiariosVentasv2.frx":0214
      Column(4)       =   "frmAdminSubdiariosVentasv2.frx":0310
      Column(5)       =   "frmAdminSubdiariosVentasv2.frx":03FC
      Column(6)       =   "frmAdminSubdiariosVentasv2.frx":04D0
      Column(7)       =   "frmAdminSubdiariosVentasv2.frx":0624
      Column(8)       =   "frmAdminSubdiariosVentasv2.frx":0750
      Column(9)       =   "frmAdminSubdiariosVentasv2.frx":08A0
      Column(10)      =   "frmAdminSubdiariosVentasv2.frx":09DC
      FormatStylesCount=   7
      FormatStyle(1)  =   "frmAdminSubdiariosVentasv2.frx":0B10
      FormatStyle(2)  =   "frmAdminSubdiariosVentasv2.frx":0C48
      FormatStyle(3)  =   "frmAdminSubdiariosVentasv2.frx":0CF8
      FormatStyle(4)  =   "frmAdminSubdiariosVentasv2.frx":0DAC
      FormatStyle(5)  =   "frmAdminSubdiariosVentasv2.frx":0E84
      FormatStyle(6)  =   "frmAdminSubdiariosVentasv2.frx":0F3C
      FormatStyle(7)  =   "frmAdminSubdiariosVentasv2.frx":101C
      ImageCount      =   0
      PrinterProperties=   "frmAdminSubdiariosVentasv2.frx":10FC
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1680
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   13560
      _Version        =   786432
      _ExtentX        =   23918
      _ExtentY        =   2963
      _StockProps     =   79
      Caption         =   "Parámetros de búsqueda"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.RadioButton rdoRangoFechas 
         Height          =   255
         Left            =   420
         TabIndex        =   5
         Top             =   300
         Width           =   1725
         _Version        =   786432
         _ExtentX        =   3043
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Por rango de fechas"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnMostrar 
         Height          =   360
         Left            =   8850
         TabIndex        =   2
         Top             =   570
         Width           =   2235
         _Version        =   786432
         _ExtentX        =   3942
         _ExtentY        =   635
         _StockProps     =   79
         Caption         =   "Mostrar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton2 
         Height          =   360
         Left            =   11100
         TabIndex        =   3
         Top             =   570
         Width           =   2235
         _Version        =   786432
         _ExtentX        =   3942
         _ExtentY        =   635
         _StockProps     =   79
         Caption         =   "Imprimir"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnExportar 
         Height          =   360
         Left            =   11100
         TabIndex        =   4
         Top             =   945
         Width           =   2235
         _Version        =   786432
         _ExtentX        =   3942
         _ExtentY        =   635
         _StockProps     =   79
         Caption         =   "Exportar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton rdoLiquidacion 
         Height          =   255
         Left            =   3225
         TabIndex        =   6
         Top             =   300
         Width           =   1290
         _Version        =   786432
         _ExtentX        =   2275
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Por liquidacion"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.GroupBox grpRangoFechas 
         Height          =   1140
         Left            =   255
         TabIndex        =   7
         Top             =   330
         Width           =   2595
         _Version        =   786432
         _ExtentX        =   4577
         _ExtentY        =   2011
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.DateTimePicker dtpDesde 
            Height          =   315
            Left            =   915
            TabIndex        =   8
            Top             =   300
            Width           =   1440
            _Version        =   786432
            _ExtentX        =   2540
            _ExtentY        =   556
            _StockProps     =   68
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   1
            CurrentDate     =   40241.6657407407
         End
         Begin XtremeSuiteControls.DateTimePicker dtpHasta 
            Height          =   300
            Left            =   915
            TabIndex        =   9
            Top             =   660
            Width           =   1440
            _Version        =   786432
            _ExtentX        =   2540
            _ExtentY        =   529
            _StockProps     =   68
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   1
            CurrentDate     =   40241.6657407407
         End
         Begin XtremeSuiteControls.Label Label1 
            Height          =   195
            Left            =   300
            TabIndex        =   13
            Top             =   675
            Width           =   480
            _Version        =   786432
            _ExtentX        =   847
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Hasta:"
            AutoSize        =   -1  'True
         End
         Begin XtremeSuiteControls.Label lblDesde 
            Height          =   195
            Left            =   270
            TabIndex        =   12
            Top             =   345
            Width           =   510
            _Version        =   786432
            _ExtentX        =   900
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Desde:"
            AutoSize        =   -1  'True
         End
      End
      Begin XtremeSuiteControls.GroupBox grpLiquidacion 
         Height          =   1140
         Left            =   3045
         TabIndex        =   10
         Top             =   330
         Width           =   5565
         _Version        =   786432
         _ExtentX        =   9816
         _ExtentY        =   2011
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.ComboBox cboLiquidaciones 
            Height          =   315
            Left            =   1170
            TabIndex        =   14
            Top             =   450
            Width           =   4185
            _Version        =   786432
            _ExtentX        =   7382
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Style           =   2
            Text            =   "cboLiquidaciones"
         End
         Begin XtremeSuiteControls.Label lblLiquidacion 
            Height          =   195
            Left            =   225
            TabIndex        =   11
            Top             =   495
            Width           =   840
            _Version        =   786432
            _ExtentX        =   1482
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Liquidacion:"
            AutoSize        =   -1  'True
         End
      End
      Begin XtremeSuiteControls.PushButton btnGuardarLiquidacion 
         Height          =   360
         Left            =   8850
         TabIndex        =   15
         Top             =   945
         Width           =   2235
         _Version        =   786432
         _ExtentX        =   3942
         _ExtentY        =   635
         _StockProps     =   79
         Caption         =   "Generar liquidación"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
      End
   End
End
Attribute VB_Name = "frmAdminSubdiariosVentasv2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim col As New Collection
Dim item As SubdiarioVentasDetalle
Private liqui As LiquidacionSubdiarioVenta
Private liquidaciones As Collection

Private desdeAbsoluto As Date
Private totales As New Dictionary

Private Enum PosicionTotales
    TotNetoGravado = 0
    totIva = 1
    totPercep = 2
    TotExento = 3
    TotTot = 4
End Enum

Private Sub Totalizar()
    Dim sumNeto As Double: sumNeto = 0
    Dim sumIVA As Double: sumIVA = 0
    Dim sumPercep As Double: sumPercep = 0
    Dim sumExento As Double: sumExento = 0
    Dim sumTot As Double: sumTot = 0

    Set totales = New Dictionary

    Dim i As SubdiarioVentasDetalle
    Dim c As Double
    For Each i In col
        If i.estado <> Anulada Then
            sumNeto = sumNeto + i.NetoGravado
            sumIVA = sumIVA + i.Iva
            sumPercep = sumPercep + i.percepciones
            sumExento = sumExento + i.Exento
            sumTot = sumTot + i.Total


        End If
    Next i

    totales.Add PosicionTotales.TotNetoGravado, sumNeto
    totales.Add PosicionTotales.totIva, sumIVA
    totales.Add PosicionTotales.totPercep, sumPercep
    totales.Add PosicionTotales.TotExento, sumExento
    totales.Add PosicionTotales.TotTot, sumTot

    Me.lblNetoGravadoTotal.caption = funciones.FormatearDecimales(totales.item(PosicionTotales.TotNetoGravado))
    Me.lblIVATotal.caption = funciones.FormatearDecimales(totales.item(PosicionTotales.totIva))
    Me.lblPercepcionesTotal.caption = funciones.FormatearDecimales(totales.item(PosicionTotales.totPercep))
    Me.lblExentoTotal.caption = funciones.FormatearDecimales(totales.item(PosicionTotales.TotExento))
    Me.lblTotalTotal.caption = funciones.FormatearDecimales(totales.item(PosicionTotales.TotTot))


End Sub


Private Sub btnExportar_Click()
    ExportaSubDiarioVentas
End Sub

Private Sub btnGuardarLiquidacion_Click()
    If col.count = 0 Then
        MsgBox "No hay detalles para poder guardar la liquidacion", vbExclamation + vbOKOnly
        Exit Sub
    End If


    Dim l As New LiquidacionSubdiarioVenta
    Dim nombre As String
    nombre = InputBox("Ingrese una descripcion para la liquidacion", "Descripcion de liquidacion")
    If LenB(nombre) = 0 Then
        MsgBox "Debe ingresar un nombre para la liquidacion", vbExclamation
    Else
        l.nombre = nombre
        l.desde = Me.dtpDesde.value
        l.hasta = Me.dtpHasta.value
        l.EsDeVenta = True
        Set l.Detalles = col
        If DAOSubdiarios.Guardar(l) Then
            SetearMaxDesde
            MsgBox "La liquidacion se guardó con éxito", vbInformation + vbOKOnly
            CargarLiquidaciones
            If Me.cboLiquidaciones.ListCount > 0 Then
                Me.cboLiquidaciones.ListIndex = Me.cboLiquidaciones.ListCount - 1
            End If
            rdoLiquidacion.value = True
            llenarLista
        Else
            MsgBox "Error al guardar la liquidacion", vbOKOnly + vbCritical
        End If
    End If
End Sub

Private Sub dtpDesde_Change()
    If Me.dtpDesde.value < desdeAbsoluto And CDbl(desdeAbsoluto) <> 0 Then
        Me.dtpDesde.value = desdeAbsoluto
        MsgBox "La fecha desde no puede ser menor que " & desdeAbsoluto & vbNewLine & "Que es la máxima fecha de las liquidaciones hechas.", vbInformation + vbOKOnly
    ElseIf Me.dtpDesde.value > desdeAbsoluto And CDbl(desdeAbsoluto) <> 0 Then
        Me.dtpDesde.value = desdeAbsoluto
        MsgBox "La fecha desde no puede ser mayor que " & desdeAbsoluto & vbNewLine & "Que es la máxima fecha de las liquidaciones hechas.", vbInformation + vbOKOnly
    End If
End Sub

Private Sub dtpHasta_Change()
    If Me.dtpHasta.value < Me.dtpDesde.value Then
        Me.dtpHasta.value = Me.dtpDesde.value
    End If
End Sub

Private Sub Form_Load()

    Customize Me
    GridEXHelper.CustomizeGrid Me.GridEX1, True, False
    GridEXHelper.AutoSizeColumns Me.GridEX1, True


    SetearMaxDesde


    Me.GridEX1.ItemCount = 0

    CargarLiquidaciones

    Me.rdoRangoFechas.value = True
End Sub

Private Sub SetearMaxDesde()
    desdeAbsoluto = DAOSubdiarios.MaxFechaLiqui()

    If CDbl(desdeAbsoluto) <> 0 Then
        Me.dtpDesde.value = desdeAbsoluto
        Me.dtpDesde.MinDate = desdeAbsoluto
        Me.dtpDesde.MaxDate = desdeAbsoluto
        Me.dtpHasta.MinDate = desdeAbsoluto
    Else
        Me.dtpDesde.value = DateSerial(Year(Now), Month(Now), 1)
    End If

    If CLng(desdeAbsoluto) >= CLng(Now) Then
        Me.dtpHasta.value = DateAdd("d", 1, desdeAbsoluto)
    Else
        Me.dtpHasta.value = Now
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Me.GroupBox1.Width = Me.ScaleWidth - 260
    Me.GridEX1.Width = Me.ScaleWidth - 260
    Me.GridEX1.Height = Me.ScaleHeight - 3900

    Me.grpTotales.Left = Me.ScaleWidth - Me.grpTotales.Width - 150
    Me.grpTotales.Top = Me.ScaleHeight - Me.grpTotales.Height - 150
End Sub

Private Sub llenarLista()
    If Me.rdoRangoFechas.value Then
        Set col = DAOSubdiarios.SubDiarioVentas(Me.dtpDesde.value, Me.dtpHasta.value)
    Else
        If Me.cboLiquidaciones.ListIndex <> -1 Then
            Set liqui = liquidaciones.item(CStr(Me.cboLiquidaciones.ItemData(Me.cboLiquidaciones.ListIndex)))
            Set col = liqui.Detalles
        Else
            Set col = New Collection
        End If
    End If

    Me.GridEX1.ItemCount = 0
    Me.GridEX1.ItemCount = col.count

    Me.caption = "Subdiario de Ventas (" & col.count & " comprobantes encontrados)"

    Totalizar
End Sub

Private Sub GridEX1_DblClick()
    If col.count > 0 Then
        Dim f_c3h3 As New frmFacturaEdicion
        f_c3h3.ReadOnly = True
        f_c3h3.idFactura = item.FacturaId
        f_c3h3.Show
    End If
End Sub

Private Sub GridEX1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 67 And Shift = 2 Then    'CTRL + C
        GridEXHelper.Grid2Clipboard Me.GridEX1
        DoEvents
        MsgBox "La lista ha sido copiada al portapapeles.", vbInformation + vbOKOnly
    End If
End Sub



Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)
    If RowBuffer.RowIndex > 0 And col.count > 0 Then
        Set item = col.item(RowBuffer.RowIndex)
        If item.estado = Anulada Then
            RowBuffer.RowStyle = "anulada"
        End If
    End If
End Sub

Private Sub GridEX1_SelectionChange()
    If Me.GridEX1.row <> -1 Then
        Set item = col.item(Me.GridEX1.RowIndex(Me.GridEX1.row))
    End If
End Sub
Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    On Error GoTo err1
    Set item = col.item(RowIndex)
    Values(1) = IIf(item.estado = Anulada, "ANULADO", item.FEcha)
    Values(2) = item.Comprobante
    Values(3) = IIf(item.estado = Anulada, "ANULADO", item.RazonSocial)
    Values(4) = IIf(item.estado = Anulada, "ANULADO", item.Cuit)
    Values(5) = IIf(item.estado = Anulada, "ANULADO", item.CondicionIva)
    Values(6) = funciones.FormatearDecimales(IIf(item.estado = Anulada, 0, item.NetoGravado))
    Values(7) = funciones.FormatearDecimales(IIf(item.estado = Anulada, 0, item.Iva))
    Values(8) = funciones.FormatearDecimales(IIf(item.estado = Anulada, 0, item.percepciones))
    Values(9) = funciones.FormatearDecimales(IIf(item.estado = Anulada, 0, item.Exento))
    Values(10) = funciones.FormatearDecimales(IIf(item.estado = Anulada, 0, (item.Total)))
    Exit Sub
err1:
End Sub
Private Sub btnMostrar_Click()
    llenarLista
End Sub


Private Sub GridEX1_UnboundUpdate(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex > 0 Then
        If item.estado <> Anulada Then
            If MsgBox("¿Desea realmente actualizar los valores del item?", vbYesNo + vbQuestion) = vbYes Then
                Set item = col.item(RowIndex)
                item.NetoGravado = Values(6)
                item.Iva = Values(7)
                item.percepciones = Values(8)
                item.Exento = Values(9)
                item.Total = Values(10)

                DAOSubdiarios.UpdateDetalle item
                Totalizar
            End If
        End If
    End If
End Sub

Private Sub PushButton2_Click()
    With Me.GridEX1.PrinterProperties
        .HeaderDistance = 500
        .FooterDistance = 1550
        .TopMargin = 1800
        .BottomMargin = 2000

        .FitColumns = True
        .DocumentName = "Subdiario Ventas"

        .RepeatHeaders = True
        .Orientation = jgexPPLandscape
        '.HeaderString(jgexHFCenter) = headercenter
        '.HeaderString(jgexHFLeft) = headerLeft'

        '.FooterString(jgexHFLeft) = footerLeft
        .FooterString(jgexHFCenter) = Now



    End With
    Load frmPrintPreview
    frmPrintPreview.Move Me.Left, Me.Top, Me.Width, Me.Height
    Me.GridEX1.PrintPreview frmPrintPreview.GEXPreview1, Me.GridEX1.SelectedItems.count > 1
    frmPrintPreview.Show 1
End Sub

Private Sub rdoLiquidacion_Click()

    ActualizarFrames
    If Me.rdoLiquidacion.value And Me.cboLiquidaciones.ListIndex = -1 And Me.cboLiquidaciones.ListCount > 0 Then
        Me.cboLiquidaciones.ListIndex = 0
    End If
End Sub

Private Sub rdoRangoFechas_Click()
    ActualizarFrames
End Sub

Private Sub ActualizarFrames()
    Me.GridEX1.ItemCount = 0
    Set col = New Collection
    Me.grpRangoFechas.Enabled = Me.rdoRangoFechas.value
    Me.grpLiquidacion.Enabled = Me.rdoLiquidacion.value
    Me.btnGuardarLiquidacion.Enabled = Me.rdoRangoFechas.value

    Me.lblExentoTotal.caption = ".-"
    Me.lblIVATotal.caption = ".-"
    Me.lblNetoGravadoTotal.caption = ".-"
    Me.lblTotalTotal.caption = ".-"
    Me.lblPercepcionesTotal.caption = ".-"

    Me.GridEX1.AllowEdit = Me.rdoLiquidacion.value
    Dim Column As JSColumn
    For Each Column In Me.GridEX1.Columns
        Column.EditType = jgexEditNone
    Next Column

    Me.GridEX1.Columns("neto_gravado").EditType = IIf(Me.rdoRangoFechas.value, jgexEditTypeConstants.jgexEditNone, jgexEditTypeConstants.jgexEditTextBox)
    Me.GridEX1.Columns("iva").EditType = IIf(Me.rdoRangoFechas.value, jgexEditTypeConstants.jgexEditNone, jgexEditTypeConstants.jgexEditTextBox)
    Me.GridEX1.Columns("percepciones").EditType = IIf(Me.rdoRangoFechas.value, jgexEditTypeConstants.jgexEditNone, jgexEditTypeConstants.jgexEditTextBox)
    Me.GridEX1.Columns("exento").EditType = IIf(Me.rdoRangoFechas.value, jgexEditTypeConstants.jgexEditNone, jgexEditTypeConstants.jgexEditTextBox)
    Me.GridEX1.Columns("total").EditType = IIf(Me.rdoRangoFechas.value, jgexEditTypeConstants.jgexEditNone, jgexEditTypeConstants.jgexEditTextBox)
End Sub

Private Sub CargarLiquidaciones()
    Me.cboLiquidaciones.Clear

    Set liquidaciones = DAOSubdiarios.FindAllLiquidacionesVenta()
    For Each liqui In liquidaciones
        Me.cboLiquidaciones.AddItem liqui.nombre & " (" & liqui.desde & " a " & liqui.hasta & ")"
        Me.cboLiquidaciones.ItemData(Me.cboLiquidaciones.NewIndex) = liqui.id
    Next liqui

End Sub



Public Function ExportaSubDiarioVentas() As Boolean
    On Error GoTo errEXCEL
    Dim xlb As New Excel.Workbook
    Dim xla As New Excel.Worksheet
    Dim xls As New Excel.Application

    Dim A As String
    Dim b As String
    Dim offset As Long
    Dim strMsg As String
    Dim CDLGMAIN As CommonDialog
    Dim sFilter As String


    Set xlb = xls.Workbooks.Add
    Set xla = xlb.Worksheets.Add
    xla.Activate


    With xla

        .Range("A1:j1").Merge
        .Range("A2:j2").Merge
        .Range("A1:j3").HorizontalAlignment = xlHAlignCenter
        .Range("A1:j2").Font.Bold = True
        .Range("A3:j2").Font.Bold = True


        .Cells(1, 1).value = "SIGNOPLAST S.A. Subdiario ventas" & IIf(Me.rdoRangoFechas.value, " (NO LIQUIDADO)", vbNullString)

        Dim desde As Date
        Dim hasta As Date
        If Me.rdoRangoFechas.value Then
            desde = Me.dtpDesde.value
            hasta = Me.dtpHasta.value
        Else
            Dim liq As LiquidacionSubdiarioVenta
            Set liq = liquidaciones.item(CStr(Me.cboLiquidaciones.ItemData(Me.cboLiquidaciones.ListIndex)))
            desde = liq.desde
            hasta = liq.hasta
        End If

        .Cells(2, 1).value = "Periodo " & Format(desde, "dd/mm/yyyy") & " - " & Format(hasta, "dd/mm/yyyy")
        .Range("A3:j3").Interior.Color = &HC0C0C0


        Dim Column As JSColumn
        Dim x As Integer

        For Each Column In Me.GridEX1.Columns
            x = x + 1
            .Cells(3, x).value = Column.caption
        Next Column

        .Columns("f").HorizontalAlignment = xlHAlignRight
        .Columns("g").HorizontalAlignment = xlHAlignRight
        .Columns("h").HorizontalAlignment = xlHAlignRight
        .Columns("i").HorizontalAlignment = xlHAlignRight

        .Columns("a").HorizontalAlignment = xlHAlignCenter
        .Columns("b").HorizontalAlignment = xlHAlignCenter
        .Columns("d").HorizontalAlignment = xlHAlignCenter
        .Columns("e").HorizontalAlignment = xlHAlignCenter

        .Columns("j").HorizontalAlignment = xlHAlignRight

        .Columns("a").ColumnWidth = 10
        .Columns("b").ColumnWidth = 8
        .Columns("c").ColumnWidth = 35
        .Columns("d").ColumnWidth = 13
        .Columns("e").ColumnWidth = 15
        .Columns("f").ColumnWidth = 13
        .Columns("g").ColumnWidth = 13
        .Columns("h").ColumnWidth = 13
        .Columns("i").ColumnWidth = 13
        .Columns("j").ColumnWidth = 15



        Dim Total As Double
        Dim totnetog As Double
        Dim totIV As Double
        Dim totperi As Double
        Dim totexen As Double
        Total = 0
        totnetog = 0
        totIV = 0
        totperi = 0
        totexen = 0

        x = 1

        For Each item In col
            If item.estado = Anulada Then
                .Cells(x + 3, 1).value = item.FEcha
                .Cells(x + 3, 2).value = item.Comprobante
                .Cells(x + 3, 3).value = "ANULADO"
                .Cells(x + 3, 4).value = "ANULADO"
                .Cells(x + 3, 5).value = "ANULADO"


                .Cells(x + 3, 6).value = 0
                .Cells(x + 3, 7).value = 0
                .Cells(x + 3, 8).value = 0
                .Cells(x + 3, 9).value = 0
                .Cells(x + 3, 10).value = 0
                .Range(.Cells(x + 3, 1), .Cells(x + 3, 10)).Font.Strikethrough = True
                .Range(.Cells(x + 3, 1), .Cells(x + 3, 10)).Font.Italic = True
            Else


                .Cells(x + 3, 1).value = item.FEcha
                .Cells(x + 3, 2).value = item.Comprobante
                .Cells(x + 3, 3).value = item.RazonSocial
                .Cells(x + 3, 4).value = item.Cuit
                .Cells(x + 3, 5).value = item.CondicionIva


                .Cells(x + 3, 6).value = item.NetoGravado
                .Cells(x + 3, 7).value = item.Iva
                .Cells(x + 3, 8).value = item.percepciones
                .Cells(x + 3, 9).value = item.Exento
                .Cells(x + 3, 10).value = item.Total
            End If

            x = x + 1
        Next item


        A = "j" & x + 2
        offset = x + 3
        b = "j" & offset
        .Range("f1", b).NumberFormat = "0.00"
        .Range("a1", A).Borders.LineStyle = xlContinuous

        .Range("f" & x + 3, b).Interior.Color = &HC0C0C0
        .Range("f" & x + 3, b).Borders.LineStyle = xlContinuous
        .Range("f" & x + 3, b).Font.Bold = True

        .Cells(offset, 10).value = totales.item(PosicionTotales.TotTot)
        .Cells(offset, 9).value = totales.item(PosicionTotales.TotExento)
        .Cells(offset, 8).value = totales.item(PosicionTotales.totPercep)
        .Cells(offset, 7).value = totales.item(PosicionTotales.totIva)
        .Cells(offset, 6).value = totales.item(PosicionTotales.TotNetoGravado)
        .Cells(offset, 5).value = "Totales"





        'xls.Visible = True NO MUESTRO LA HOJA XLS
        strMsg = "Se han transportado los datos correctamente"
        strMsg = strMsg & vbCrLf & "a una hoja de calculo de Excel."
        strMsg = strMsg & vbCrLf & vbCrLf
        strMsg = strMsg & "¿Desea guardar la hoja de calculo de Excel?"
        Set CDLGMAIN = frmPrincipal.cd



        '    If MsgBox(strMsg, vbQuestion + vbYesNo) = vbYes Then
        sFilter = "Hoja de Calculo|*.xls"
        CDLGMAIN.filter = sFilter

        Dim Periodo As String
        Periodo = 1
        Periodo = Format(desde, "ddmmyyyy") & "-" & Format(hasta, "ddmmyyyy")

        Dim archi As String
        archi = "SUBDIARIO_VENTAS_" & Periodo & ".xls"
        frmPrincipal.cd.CancelError = True
        CDLGMAIN.filename = archi
        CDLGMAIN.ShowSave

        If CDLGMAIN.filename <> Empty Then
            xla.SaveAs (CDLGMAIN.filename)
            strMsg = "Los datos del reporte se han guardado en un archivo: " & vbCrLf & vbCrLf
            strMsg = strMsg & CDLGMAIN.filename
            MsgBox strMsg, vbInformation + vbOKOnly, "Hoja de calculo guardada"
            archi = CDLGMAIN.filename
        Else
            ExportaSubDiarioVentas = False
        End If
        xlb.Saved = True
        xlb.Close

        xls.Quit
        Set xls = Nothing
        Set xla = Nothing
        Set xlb = Nothing

        '    End If
        ExportaSubDiarioVentas = True



    End With
    Exit Function
errEXCEL:
    If Err.Number = -2147221080 Then
        ExportaSubDiarioVentas = False
    Else
        ' Resume
        MsgBox "Se produjo un error. No se graban los cambios", vbCritical, "Error"
        ExportaSubDiarioVentas = False
    End If
    xlb.Saved = True
    xlb.Close
    Set xls = Nothing
    Set xla = Nothing
    Set xlb = Nothing

End Function


