VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmAdminExtrasReporteIVACompras 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte IVA Compras"
   ClientHeight    =   8790
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   14085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8790
   ScaleWidth      =   14085
   Begin XtremeSuiteControls.GroupBox GroupBox 
      Height          =   2175
      Index           =   2
      Left            =   4920
      TabIndex        =   3
      Top             =   120
      Width           =   9015
      _Version        =   786432
      _ExtentX        =   15901
      _ExtentY        =   3836
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.CommonDialog CommonDialog1 
         Left            =   240
         Top             =   240
         _Version        =   786432
         _ExtentX        =   423
         _ExtentY        =   423
         _StockProps     =   4
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox 
      Height          =   6135
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Width           =   13815
      _Version        =   786432
      _ExtentX        =   24368
      _ExtentY        =   10821
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.ListBox lstBoxRegistros 
         Height          =   5175
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "lstBoxRegistros"
         Top             =   240
         Width           =   13455
         _Version        =   786432
         _ExtentX        =   23733
         _ExtentY        =   9128
         _StockProps     =   77
         BackColor       =   -2147483643
         Appearance      =   5
         FlatStyle       =   -1  'True
         UseVisualStyle  =   -1  'True
         EnableMarkup    =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnExportarTXT 
         Height          =   495
         Left            =   11040
         TabIndex        =   13
         Top             =   5520
         Width           =   2535
         _Version        =   786432
         _ExtentX        =   4471
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Exportar a TXT"
         UseVisualStyle  =   -1  'True
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox 
      Height          =   2175
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      _Version        =   786432
      _ExtentX        =   8281
      _ExtentY        =   3836
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.RadioButton radioBtnComprobantes 
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   1695
         _Version        =   786432
         _ExtentX        =   2990
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Comprobantes"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton radioBtnAlicuotas 
         Height          =   375
         Left            =   2400
         TabIndex        =   5
         Top             =   240
         Width           =   1095
         _Version        =   786432
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Alícuotas"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.DateTimePicker dtpDesde 
         Height          =   315
         Left            =   855
         TabIndex        =   6
         Top             =   1110
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
         Left            =   3010
         TabIndex        =   7
         Top             =   1110
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
         Left            =   840
         TabIndex        =   8
         Top             =   720
         Width           =   3645
         _Version        =   786432
         _ExtentX        =   6429
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.PushButton btnReportar 
         Height          =   495
         Left            =   2520
         TabIndex        =   12
         Top             =   1560
         Width           =   1935
         _Version        =   786432
         _ExtentX        =   3413
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Reportar"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label7 
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   780
         Width           =   480
         _Version        =   786432
         _ExtentX        =   847
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Rango"
         BackColor       =   12632256
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label5 
         Height          =   195
         Left            =   285
         TabIndex        =   10
         Top             =   1155
         Width           =   465
         _Version        =   786432
         _ExtentX        =   820
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Desde"
         BackColor       =   12632256
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label6 
         Height          =   195
         Left            =   2460
         TabIndex        =   9
         Top             =   1170
         Width           =   420
         _Version        =   786432
         _ExtentX        =   741
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Hasta"
         BackColor       =   12632256
         AutoSize        =   -1  'True
      End
   End
End
Attribute VB_Name = "frmAdminExtrasReporteIVACompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim cliente As clsCliente
    Dim registro As clsRegistroIVACompras
    
    
Private Sub Form_Load()
    FormHelper.Customize Me
    
    Dim i As Integer
    funciones.FillComboBoxDateRanges Me.cboRangos
    For i = 0 To Me.cboRangos.ListCount - 1
        If Me.cboRangos.ItemData(i) = DateRangeValue.DRV_YearCurrent Then Exit For
    Next i
    Me.cboRangos.ListIndex = 6
    
    Me.radioBtnComprobantes.value = True
    
    Me.btnExportarTXT.Enabled = False
    
End Sub


Private Sub btnExportarTXT_Click()
    ' Construye el nombre del archivo con el formato deseado
    Dim nombreArchivo As String
    
    If Me.radioBtnComprobantes.value = True Then
        nombreArchivo = "COMPROBANTES_" & Format(Now, "hhmmss") & ".TXT"
    ElseIf Me.radioBtnAlicuotas.value = True Then
        nombreArchivo = "ALICUOTAS_" & Format(Now, "hhmmss") & ".TXT"
    Else
        MsgBox ("Debe seleccionar un tipo de reporte primero")
        Exit Sub
    End If

    ' Asigna el nombre de archivo personalizado al cuadro de diálogo
    CommonDialog1.filename = nombreArchivo

    ' Abre el cuadro de diálogo para seleccionar la ubicación y el nombre del archivo
    CommonDialog1.filter = "Archivos de texto (*.txt)|*.txt|Todos los archivos (*.*)|*.*"
    CommonDialog1.ShowSave

    ' Verifica si el usuario seleccionó un archivo o canceló
    If CommonDialog1.filename = "" Then
        MsgBox "No se ha seleccionado ningún archivo para guardar.", vbExclamation
        Exit Sub ' Salir si el usuario canceló
    End If

    ' Abre el archivo para escritura
    Open CommonDialog1.filename For Output As #1

    Dim i As Integer

    ' Recorre los elementos del ListBox y escribe cada elemento en una nueva línea
    For i = 0 To Me.lstBoxRegistros.ListCount - 1
        Print #1, Me.lstBoxRegistros.list(i)
    Next i

    ' Cierra el archivo
    Close #1

    MsgBox "Contenido exportado exitosamente al archivo " & CommonDialog1.filename, vbInformation

End Sub



Private Sub btnReportar_Click()

    Me.lstBoxRegistros.Clear
  
    If Me.radioBtnComprobantes.value = True Then
        ReportarComprobantes
   
    ElseIf Me.radioBtnAlicuotas.value = True Then
        ReportarAlicuotas
   
    End If
    
End Sub

Private Sub ReportarComprobantes()

    Dim T As String
    Dim registros As Collection
    Dim filter As String
    filter = "1 = 1"
    
    If Not IsNull(Me.dtpDesde.value) Then
        filter = filter & " AND cp.fecha >= " & conectar.Escape(Me.dtpDesde.value)
    End If

    If Not IsNull(Me.dtpHasta.value) Then
        filter = filter & " AND cp.fecha <= " & conectar.Escape(Me.dtpHasta.value)
    End If
        
    Set registros = DAORegistrosCompras.FindAllComprobantes(filter)
    
    
    
    For Each registro In registros
    
    Dim FEcha As String
    Dim fecha_01 As String
        FEcha = registro.FEcha
        fecha_01 = Format(CDate(FEcha), "yyyyMMdd")
        
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
    Dim tipoComprobante As String
    
    Dim tipodecomprobante_02 As String
    
    If registro.tipodoccontable = 0 Then
        If registro.idconfigfactura = 1 Then
            ' Código para cuando tipo_doc_contable=0 y id_config_factura=1
            tipodecomprobante_02 = "001"
        ElseIf registro.idconfigfactura = 2 Then
            ' Código para cuando tipo_doc_contable=0 y id_config_factura=2
            tipodecomprobante_02 = "006"
        ElseIf registro.idconfigfactura = 3 Then
            ' Código para cuando tipo_doc_contable=0 y id_config_factura=3
            tipodecomprobante_02 = "011"
        ElseIf registro.idconfigfactura = 6 Then
            ' Código para cuando tipo_doc_contable=0 y id_config_factura=6
            tipodecomprobante_02 = "001"
        ElseIf registro.idconfigfactura = 7 Then
            ' Código para cuando tipo_doc_contable=0 y id_config_factura=7
            tipodecomprobante_02 = "006"
        ElseIf registro.idconfigfactura = 9 Then
            ' Código para cuando tipo_doc_contable=0 y id_config_factura=9
            tipodecomprobante_02 = "019"
        ElseIf registro.idconfigfactura = 10 Then
            ' Código para cuando tipo_doc_contable=0 y id_config_factura=10
            tipodecomprobante_02 = "011"
        ElseIf registro.idconfigfactura = 12 Then
            ' Código para cuando tipo_doc_contable=0 y id_config_factura=12
            tipodecomprobante_02 = "051"
        End If
    ElseIf registro.tipodoccontable = 1 Then
        If registro.idconfigfactura = 1 Then
            ' Código para cuando tipo_doc_contable=1 y id_config_factura=1
            tipodecomprobante_02 = "003"
        ElseIf registro.idconfigfactura = 3 Then
            ' Código para cuando tipo_doc_contable=1 y id_config_factura=3
            tipodecomprobante_02 = "013"
        ElseIf registro.idconfigfactura = 6 Then
            ' Código para cuando tipo_doc_contable=1 y id_config_factura=6
            tipodecomprobante_02 = "003"
        ElseIf registro.idconfigfactura = 7 Then
            ' Código para cuando tipo_doc_contable=1 y id_config_factura=7
            tipodecomprobante_02 = "008"
        ElseIf registro.idconfigfactura = 10 Then
            ' Código para cuando tipo_doc_contable=1 y id_config_factura=10
            tipodecomprobante_02 = "013"
        ElseIf registro.idconfigfactura = 12 Then
            ' Código para cuando tipo_doc_contable=1 y id_config_factura=12
            tipodecomprobante_02 = "053"
        End If
    ElseIf registro.tipodoccontable = 2 Then
        If registro.idconfigfactura = 3 Then
            ' Código para cuando tipo_doc_contable=2 y id_config_factura=3
            tipodecomprobante_02 = "012"
        ElseIf registro.idconfigfactura = 6 Then
            ' Código para cuando tipo_doc_contable=2 y id_config_factura=6
            tipodecomprobante_02 = "002"
        ElseIf registro.idconfigfactura = 10 Then
            ' Código para cuando tipo_doc_contable=2 y id_config_factura=10
            tipodecomprobante_02 = "013"
        End If
    ElseIf registro.tipodoccontable = 4 Then
        If registro.idconfigfactura = 6 Then
            ' Código para cuando tipo_doc_contable=4 y id_config_factura=6
            tipodecomprobante_02 = "039"
        End If
    ElseIf registro.tipodoccontable = 5 Then
        If registro.idconfigfactura = 6 Then
            ' Código para cuando tipo_doc_contable=5 y id_config_factura=6
            tipodecomprobante_02 = "030"
        End If
    End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim puntoDeVenta As String
    Dim puntodeventa_03 As String
    
    ' Asigna el valor a la variable
    puntoDeVenta = registro.numerodecomprobante
    
    ' Extrae la parte izquierda hasta el guion del medio
    puntodeventa_03 = Mid(puntoDeVenta, 1, InStr(puntoDeVenta, "-") - 1)
    
    ' Asegúrate de que tenga 5 caracteres completándose con ceros a la izquierda
    puntodeventa_03 = String(5 - Len(puntodeventa_03), "0") & puntodeventa_03
    
    ' Muestra el resultado

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Dim numerodecomprobante As String
    Dim numerodecomprobante_04 As String
    
    ' Asigna el valor a la variable
    numerodecomprobante = registro.numerodecomprobante
    
    ' Extrae la parte derecha después del guion del medio
    numerodecomprobante_04 = Mid(numerodecomprobante, InStr(numerodecomprobante, "-") + 1)
    
    ' Asegúrate de que tenga 20 caracteres completándose con ceros a la izquierda
    numerodecomprobante_04 = String(20 - Len(numerodecomprobante_04), "0") & numerodecomprobante_04
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim espacioenblanco As String
    espacioenblanco = "                "
    
    Dim codigodelvendedor_05 As String
    codigodelvendedor_05 = "80"
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim cuit_06 As String
    
    cuit_06 = registro.Cuit
    cuit_06 = String(20 - Len(cuit_06), "0") & cuit_06
   
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim denominacionvendedor_07 As String
    
    denominacionvendedor_07 = registro.denominacionvendedor
    
    ' Realiza las sustituciones necesarias
    denominacionvendedor_07 = Replace(denominacionvendedor_07, "Ñ", "N")
    denominacionvendedor_07 = Replace(denominacionvendedor_07, "Á", "A")
    denominacionvendedor_07 = Replace(denominacionvendedor_07, "É", "E")
    denominacionvendedor_07 = Replace(denominacionvendedor_07, "Í", "I")
    denominacionvendedor_07 = Replace(denominacionvendedor_07, "Ó", "O")
    denominacionvendedor_07 = Replace(denominacionvendedor_07, "Ú", "U")
    
    ' Recorta a 30 caracteres si es más largo
    denominacionvendedor_07 = Left(denominacionvendedor_07 & Space(30), 30)
    
    ' Convierte la cadena a mayúsculas
    denominacionvendedor_07 = UCase(denominacionvendedor_07)
        
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Dim importettoperacionnuevo As Double
    Dim importettoperacionnuevo_08 As String
    
    importettoperacionnuevo = registro.montoneto + registro.redondeoiva + registro.impuestosinternos + registro.percepcionesSoloIva + registro.percepcionessSinIva + registro.ivavalor
    importettoperacionnuevo_08 = Format(importettoperacionnuevo, "0.00")
    importettoperacionnuevo_08 = Replace(importettoperacionnuevo_08, ".", "")
    importettoperacionnuevo_08 = String(15 - Len(importettoperacionnuevo_08), "0") & importettoperacionnuevo_08
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Dim redondeo_09 As String
    
    redondeo_09 = Format(registro.redondeoiva, "0.00")
    redondeo_09 = Replace(redondeo_09, ".", "")


    If CDbl(redondeo_09) < 0 Then
                redondeo_09 = String(15 - Len(redondeo_09), "0") & redondeo_09
                redondeo_09 = Replace(redondeo_09, "-", "")
        redondeo_09 = "-" & redondeo_09
        

    Else
            redondeo_09 = String(15 - Len(redondeo_09), "0") & redondeo_09
    End If
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    '10_15caracteres
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim percepcionsoloiva As Double
    Dim PercepcionesSoloIva_11 As String
    
    PercepcionesSoloIva_11 = Format(registro.percepcionesSoloIva, "0.00")
    PercepcionesSoloIva_11 = Replace(PercepcionesSoloIva_11, ".", "")
    PercepcionesSoloIva_11 = String(15 - Len(PercepcionesSoloIva_11), "0") & PercepcionesSoloIva_11
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    '12_15caracteres
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    '13_PercepcionessSinIva
    
    Dim percepcionsiniva As Double
    Dim PercepcionesSinIva_13 As String
    
    PercepcionesSinIva_13 = Format(registro.percepcionessSinIva, "0.00")
    PercepcionesSinIva_13 = Replace(PercepcionesSinIva_13, ".", "")
    PercepcionesSinIva_13 = String(15 - Len(PercepcionesSinIva_13), "0") & PercepcionesSinIva_13
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    '14_15caracteres
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim impuestosinternos_15 As String
    
    impuestosinternos_15 = Format(registro.impuestosinternos, "0.00")
    impuestosinternos_15 = Replace(impuestosinternos_15, ".", "")
    impuestosinternos_15 = String(15 - Len(impuestosinternos_15), "0") & impuestosinternos_15

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    '16_moneda
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    '17_10caracteres'
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Dim cantidaddealicuotas_18 As String
    Dim cantidaddealicuotas As Integer
    
    Dim result As String
    
    Select Case True
        Case registro.tipodoccontable = 0 And registro.idconfigfactura = 1 And registro.cantidaddealicuotas = 1
            cantidaddealicuotas_18 = "1"
        Case registro.tipodoccontable = 0 And registro.idconfigfactura = 2 And registro.cantidaddealicuotas = 1
            cantidaddealicuotas_18 = "0"
        Case registro.tipodoccontable = 0 And registro.idconfigfactura = 3 And registro.cantidaddealicuotas = 1
            cantidaddealicuotas_18 = "0"
        Case registro.tipodoccontable = 0 And registro.idconfigfactura = 6 And registro.cantidaddealicuotas = 1
            cantidaddealicuotas_18 = "1"
        Case registro.tipodoccontable = 0 And registro.idconfigfactura = 7 And registro.cantidaddealicuotas = 1
            cantidaddealicuotas_18 = "0"
        Case registro.tipodoccontable = 0 And registro.idconfigfactura = 9 And registro.cantidaddealicuotas = 1
            cantidaddealicuotas_18 = "1"
        Case registro.tipodoccontable = 0 And registro.idconfigfactura = 10 And registro.cantidaddealicuotas = 1
            cantidaddealicuotas_18 = "0"
        Case registro.tipodoccontable = 0 And registro.idconfigfactura = 12 And registro.cantidaddealicuotas = 1
            cantidaddealicuotas_18 = "1"
                        
        Case registro.tipodoccontable = 1 And registro.idconfigfactura = 1 And registro.cantidaddealicuotas = 1
            cantidaddealicuotas_18 = "1"
        Case registro.tipodoccontable = 1 And registro.idconfigfactura = 2 And registro.cantidaddealicuotas = 1
            cantidaddealicuotas_18 = "0"
        Case registro.tipodoccontable = 1 And registro.idconfigfactura = 3 And registro.cantidaddealicuotas = 1
            cantidaddealicuotas_18 = "0"
        Case registro.tipodoccontable = 1 And registro.idconfigfactura = 6 And registro.cantidaddealicuotas = 1
            cantidaddealicuotas_18 = "1"
        Case registro.tipodoccontable = 1 And registro.idconfigfactura = 7 And registro.cantidaddealicuotas = 1
            cantidaddealicuotas_18 = "0"
        Case registro.tipodoccontable = 1 And registro.idconfigfactura = 10 And registro.cantidaddealicuotas = 1
            cantidaddealicuotas_18 = "0"
        Case registro.tipodoccontable = 1 And registro.idconfigfactura = 12 And registro.cantidaddealicuotas = 1
            cantidaddealicuotas_18 = "1"
         
        Case registro.tipodoccontable = 2 And registro.idconfigfactura = 3 And registro.cantidaddealicuotas = 1
            cantidaddealicuotas_18 = "0"
        Case registro.tipodoccontable = 2 And registro.idconfigfactura = 6 And registro.cantidaddealicuotas = 1
            cantidaddealicuotas_18 = "1"
        Case registro.tipodoccontable = 2 And registro.idconfigfactura = 10 And registro.cantidaddealicuotas = 1
            cantidaddealicuotas_18 = "0"
            
        Case registro.tipodoccontable = 4 And registro.idconfigfactura = 6 And registro.cantidaddealicuotas = 1
            cantidaddealicuotas_18 = "1"
            
        Case registro.idIVA = 5
            cantidaddealicuotas_18 = registro.cantidadidIVADistintas
        Case registro.idIVA = 2
            cantidaddealicuotas_18 = registro.cantidadidIVADistintas
        Case registro.idIVA = 10
            cantidaddealicuotas_18 = registro.cantidadidIVADistintas
        Case registro.idIVA = 6
            cantidaddealicuotas_18 = registro.cantidadidIVADistintas
        Case registro.idIVA = 4
            cantidaddealicuotas_18 = registro.cantidadidIVADistintas
        Case Else
            cantidaddealicuotas_18 = "0"
    End Select
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Dim codigodeoperacionXXX_19 As String
    
    Select Case True
        Case registro.idIVA = 8 And registro.cantidaddealicuotas = 1
            codigodeoperacionXXX_19 = "E"
        Case registro.idIVA = 16 And registro.cantidaddealicuotas = 1
            codigodeoperacionXXX_19 = "E"
        Case registro.idIVA = 15 And registro.cantidaddealicuotas = 1
            codigodeoperacionXXX_19 = "E"
        Case registro.idIVA = 12 And registro.cantidaddealicuotas = 1
            codigodeoperacionXXX_19 = "E"
        Case registro.idIVA = 9 And registro.cantidaddealicuotas = 1
            codigodeoperacionXXX_19 = "E"
        Case registro.idIVA = 7 And registro.cantidaddealicuotas = 1
            codigodeoperacionXXX_19 = "E"
        Case registro.idIVA = 17 And registro.cantidaddealicuotas = 1
            codigodeoperacionXXX_19 = "E"
            
        Case registro.idIVA = 5 And registro.cantidaddealicuotas = 1
            codigodeoperacionXXX_19 = "N"
        Case registro.idIVA = 10 And registro.cantidaddealicuotas = 1
            codigodeoperacionXXX_19 = "N"
        Case registro.idIVA = 2 And registro.cantidaddealicuotas = 1
            codigodeoperacionXXX_19 = "N"
        Case registro.idIVA = 4 And registro.cantidaddealicuotas = 1
            codigodeoperacionXXX_19 = "N"
        Case registro.idIVA = 6 And registro.cantidaddealicuotas = 1
            codigodeoperacionXXX_19 = "N"
        Case registro.idIVA = 19 And registro.cantidaddealicuotas = 1
            codigodeoperacionXXX_19 = "N"
        
        Case Else
            codigodeoperacionXXX_19 = "N"
    End Select
    
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    
    
    Dim valoriva_20 As String
    valoriva_20 = Format(registro.ivavalor, "0.00")
    valoriva_20 = Replace(valoriva_20, ".", "")
    valoriva_20 = String(15 - Len(valoriva_20), "0") & valoriva_20

    
    T = 0
        
        T = fecha_01 _
        & tipodecomprobante_02 _
        & puntodeventa_03 _
        & numerodecomprobante_04 _
        & "                " _
        & codigodelvendedor_05 _
        & cuit_06 _
        & denominacionvendedor_07 _
        & importettoperacionnuevo_08 _
        & redondeo_09 _
        & "000000000000000" _
        & PercepcionesSoloIva_11 _
        & "000000000000000" _
        & PercepcionesSinIva_13 _
        & "000000000000000" _
        & impuestosinternos_15 _
        & "PES" _
        & "0001000000" _
        & cantidaddealicuotas_18 _
        & codigodeoperacionXXX_19 _
        & valoriva_20 _
        & "00000000000000000000000000" & " " _
        & "                             " _
        & "000000000000000"
        
        
        Me.lstBoxRegistros.AddItem T
                
    Next

    Me.btnExportarTXT.Enabled = True
    
            MsgBox "Los datos se cargaron éxitosamente!", vbOKOnly, "Confirmación"
    
End Sub

Private Sub ReportarAlicuotas()
    Dim T As String
    Dim registros As Collection
    Dim filter As String
    filter = "1 = 1"
    
    If Not IsNull(Me.dtpDesde.value) Then
        filter = filter & " AND cp.fecha >= " & conectar.Escape(Me.dtpDesde.value)
    End If

    If Not IsNull(Me.dtpHasta.value) Then
        filter = filter & " AND cp.fecha <= " & conectar.Escape(Me.dtpHasta.value)
    End If
        
    Set registros = DAORegistrosCompras.FindAllAlicuotas(filter)
    
    
    
    For Each registro In registros
    
    Dim FEcha As String
    Dim fecha_01 As String
        FEcha = registro.FEcha
        fecha_01 = Format(CDate(FEcha), "yyyyMMdd")
        
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
    Dim tipoComprobante As String
    
    Dim tipodecomprobante_02 As String
    
    If registro.tipodoccontable = 0 Then
        If registro.idconfigfactura = 1 Then
            ' Código para cuando tipo_doc_contable=0 y id_config_factura=1
            tipodecomprobante_02 = "001"
        ElseIf registro.idconfigfactura = 2 Then
            ' Código para cuando tipo_doc_contable=0 y id_config_factura=2
            tipodecomprobante_02 = "006"
        ElseIf registro.idconfigfactura = 3 Then
            ' Código para cuando tipo_doc_contable=0 y id_config_factura=3
            tipodecomprobante_02 = "011"
        ElseIf registro.idconfigfactura = 6 Then
            ' Código para cuando tipo_doc_contable=0 y id_config_factura=6
            tipodecomprobante_02 = "001"
        ElseIf registro.idconfigfactura = 7 Then
            ' Código para cuando tipo_doc_contable=0 y id_config_factura=7
            tipodecomprobante_02 = "006"
        ElseIf registro.idconfigfactura = 9 Then
            ' Código para cuando tipo_doc_contable=0 y id_config_factura=9
            tipodecomprobante_02 = "019"
        ElseIf registro.idconfigfactura = 10 Then
            ' Código para cuando tipo_doc_contable=0 y id_config_factura=10
            tipodecomprobante_02 = "011"
        ElseIf registro.idconfigfactura = 12 Then
            ' Código para cuando tipo_doc_contable=0 y id_config_factura=12
            tipodecomprobante_02 = "051"
        End If
    ElseIf registro.tipodoccontable = 1 Then
        If registro.idconfigfactura = 1 Then
            ' Código para cuando tipo_doc_contable=1 y id_config_factura=1
            tipodecomprobante_02 = "003"
        ElseIf registro.idconfigfactura = 3 Then
            ' Código para cuando tipo_doc_contable=1 y id_config_factura=3
            tipodecomprobante_02 = "013"
        ElseIf registro.idconfigfactura = 6 Then
            ' Código para cuando tipo_doc_contable=1 y id_config_factura=6
            tipodecomprobante_02 = "003"
        ElseIf registro.idconfigfactura = 7 Then
            ' Código para cuando tipo_doc_contable=1 y id_config_factura=7
            tipodecomprobante_02 = "008"
        ElseIf registro.idconfigfactura = 10 Then
            ' Código para cuando tipo_doc_contable=1 y id_config_factura=10
            tipodecomprobante_02 = "013"
        ElseIf registro.idconfigfactura = 12 Then
            ' Código para cuando tipo_doc_contable=1 y id_config_factura=12
            tipodecomprobante_02 = "053"
        End If
    ElseIf registro.tipodoccontable = 2 Then
        If registro.idconfigfactura = 3 Then
            ' Código para cuando tipo_doc_contable=2 y id_config_factura=3
            tipodecomprobante_02 = "012"
        ElseIf registro.idconfigfactura = 6 Then
            ' Código para cuando tipo_doc_contable=2 y id_config_factura=6
            tipodecomprobante_02 = "002"
        ElseIf registro.idconfigfactura = 10 Then
            ' Código para cuando tipo_doc_contable=2 y id_config_factura=10
            tipodecomprobante_02 = "013"
        End If
    ElseIf registro.tipodoccontable = 4 Then
        If registro.idconfigfactura = 6 Then
            ' Código para cuando tipo_doc_contable=4 y id_config_factura=6
            tipodecomprobante_02 = "039"
        End If
    ElseIf registro.tipodoccontable = 5 Then
        If registro.idconfigfactura = 6 Then
            ' Código para cuando tipo_doc_contable=5 y id_config_factura=6
            tipodecomprobante_02 = "030"
        End If
    End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim puntoDeVenta As String
    Dim puntodeventa_03 As String
    
    ' Asigna el valor a la variable
    puntoDeVenta = registro.numerodecomprobante
    
    ' Extrae la parte izquierda hasta el guion del medio
    puntodeventa_03 = Mid(puntoDeVenta, 1, InStr(puntoDeVenta, "-") - 1)
    
    ' Asegúrate de que tenga 5 caracteres completándose con ceros a la izquierda
    puntodeventa_03 = String(5 - Len(puntodeventa_03), "0") & puntodeventa_03
    
    ' Muestra el resultado

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Dim numerodecomprobante As String
    Dim numerodecomprobante_04 As String
    
    ' Asigna el valor a la variable
    numerodecomprobante = registro.numerodecomprobante
    
    ' Extrae la parte derecha después del guion del medio
    numerodecomprobante_04 = Mid(numerodecomprobante, InStr(numerodecomprobante, "-") + 1)
    
    ' Asegúrate de que tenga 20 caracteres completándose con ceros a la izquierda
    numerodecomprobante_04 = String(20 - Len(numerodecomprobante_04), "0") & numerodecomprobante_04
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim espacioenblanco As String
    espacioenblanco = "                "
    
    Dim codigodelvendedor_05 As String
    codigodelvendedor_05 = "80"
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim cuit_06 As String
    
    cuit_06 = registro.Cuit
    cuit_06 = String(20 - Len(cuit_06), "0") & cuit_06

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim importeNetoGravado As String
    importeNetoGravado = Format(registro.valorAlicuota, "0.00")
    importeNetoGravado = Replace(importeNetoGravado, ".", "")
    importeNetoGravado = String(15 - Len(importeNetoGravado), "0") & importeNetoGravado
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim alicuotaDeIVa As String
    
    Select Case True
        Case registro.idIVA = 8
            alicuotaDeIVa = "0003"
        Case registro.idIVA = 16
            alicuotaDeIVa = "0003"
        Case registro.idIVA = 1
            alicuotaDeIVa = "0003"
        Case registro.idIVA = 15
            alicuotaDeIVa = "0003"
        Case registro.idIVA = 12
            alicuotaDeIVa = "0003"
        Case registro.idIVA = 10
            alicuotaDeIVa = "0005"
        Case registro.idIVA = 9
            alicuotaDeIVa = "0003"
        Case registro.idIVA = 7
            alicuotaDeIVa = "0003"
        Case registro.idIVA = 5
            alicuotaDeIVa = "0004"
        Case registro.idIVA = 2
            alicuotaDeIVa = "0005"
        Case registro.idIVA = 17
            alicuotaDeIVa = "0003"
        Case registro.idIVA = 6
            alicuotaDeIVa = "0006"
        Case registro.idIVA = 4
            alicuotaDeIVa = "0005"
        Case registro.idIVA = 19
            alicuotaDeIVa = "0008"
        Case Else
            alicuotaDeIVa = "0000"
    End Select

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim importeNetoDGravado As String
    
    Select Case True
        Case registro.idIVA = 8
            importeNetoDGravado = (0 * registro.valorAlicuota) / 100 + 0.0000000001
        Case registro.idIVA = 16
            importeNetoDGravado = (0 * registro.valorAlicuota) / 100 + 0.0000000001
        Case registro.idIVA = 1
            importeNetoDGravado = (0 * registro.valorAlicuota) / 100 + 0.0000000001
        Case registro.idIVA = 15
            importeNetoDGravado = (0 * registro.valorAlicuota) / 100 + 0.0000000001
        Case registro.idIVA = 12
            importeNetoDGravado = (0 * registro.valorAlicuota) / 100 + 0.0000000001
        Case registro.idIVA = 10
            importeNetoDGravado = (21 * registro.valorAlicuota) / 100 + 0.0000000001
        Case registro.idIVA = 9
            importeNetoDGravado = (0 * registro.valorAlicuota) / 100 + 0.0000000001
        Case registro.idIVA = 7
            importeNetoDGravado = (0 * registro.valorAlicuota) / 100 + 0.0000000001
        Case registro.idIVA = 5
            importeNetoDGravado = (10.5 * registro.valorAlicuota) / 100 + 0.0000000001
        Case registro.idIVA = 2
            importeNetoDGravado = (21 * registro.valorAlicuota) / 100 + 0.0000000001
        Case registro.idIVA = 17
            importeNetoDGravado = (0 * registro.valorAlicuota) / 100 + 0.0000000001
        Case registro.idIVA = 6
            importeNetoDGravado = (27 * registro.valorAlicuota) / 100 + 0.0000000001
        Case registro.idIVA = 4
            importeNetoDGravado = (21 * registro.valorAlicuota) / 100 + 0.0000000001
        Case registro.idIVA = 19
            importeNetoDGravado = (5 * registro.valorAlicuota) / 100 + 0.0000000001
        Case Else
            importeNetoDGravado = "0000"
    End Select
    
    importeNetoDGravado = Format(importeNetoDGravado, "0.00")
    importeNetoDGravado = Replace(importeNetoDGravado, ".", "")
        importeNetoDGravado = String(15 - Len(importeNetoDGravado), "0") & importeNetoDGravado

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        T = 0
        
        T = tipodecomprobante_02 _
        & puntodeventa_03 _
        & numerodecomprobante_04 _
        & codigodelvendedor_05 _
        & cuit_06 _
        & importeNetoGravado _
        & alicuotaDeIVa _
        & importeNetoDGravado
       
        
        Me.lstBoxRegistros.AddItem T
                
    Next
    
        Me.btnExportarTXT.Enabled = True
        
        MsgBox "Los datos se cargaron éxitosamente!", vbOKOnly, "Confirmación"
    
End Sub



Private Sub cboRangos_Click()
    funciones.CalculateDateRange Me.cboRangos, Me.dtpDesde, Me.dtpHasta

End Sub





