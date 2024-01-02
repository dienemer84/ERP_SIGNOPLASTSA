VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmSistemasTests 
   Caption         =   "Tests"
   ClientHeight    =   7530
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2415
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7530
   ScaleWidth      =   2415
   Begin XtremeSuiteControls.PushButton btnImprimirNuevo 
      Height          =   735
      Left            =   240
      TabIndex        =   6
      Top             =   6480
      Width           =   1935
      _Version        =   786432
      _ExtentX        =   3413
      _ExtentY        =   1296
      _StockProps     =   79
      Caption         =   "Imprimir Sistema NUEVO"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton btnImprimir 
      Height          =   735
      Left            =   240
      TabIndex        =   5
      Top             =   5400
      Width           =   1935
      _Version        =   786432
      _ExtentX        =   3413
      _ExtentY        =   1296
      _StockProps     =   79
      Caption         =   "Imprimir Sistema ACTUAL"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton PushButton 
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   4560
      Width           =   1935
      _Version        =   786432
      _ExtentX        =   3413
      _ExtentY        =   1085
      _StockProps     =   79
      Caption         =   "Prueba_05 (Lista Liquidacion Caja)"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton btnPrueba_04 
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   3480
      Width           =   1935
      _Version        =   786432
      _ExtentX        =   3413
      _ExtentY        =   1085
      _StockProps     =   79
      Caption         =   "Prueba_04"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton btnPrueba_03 
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   2400
      Width           =   1935
      _Version        =   786432
      _ExtentX        =   3413
      _ExtentY        =   1085
      _StockProps     =   79
      Caption         =   "Prueba_03"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton btnPrueba2 
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   1935
      _Version        =   786432
      _ExtentX        =   3413
      _ExtentY        =   1085
      _StockProps     =   79
      Caption         =   "Prueba_02"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton btnTest01 
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1935
      _Version        =   786432
      _ExtentX        =   3413
      _ExtentY        =   1085
      _StockProps     =   79
      Caption         =   "Prueba_01"
      UseVisualStyle  =   -1  'True
   End
End
Attribute VB_Name = "frmSistemasTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim colProveedores As New Collection
    Dim colFacturas As New Collection
    Dim colTransferenciasB As New Collection
    
    
Private Sub btnImprimir_Click()

    ' Preguntar al usuario si desea mostrar en PDF o en Word
    Dim respuesta As Integer
    respuesta = MsgBox("¿Desea mostrar en PDF (Sí) o en Word (No)?", vbQuestion + vbYesNo, "Seleccionar formato")

    ' Realizar acciones según la respuesta del usuario
    If respuesta = vbYes Then
        ' Código para mostrar en PDF (por ejemplo, en Word)
        MsgBox "Mostrar en PDF"
        ' Aquí puedes agregar el código específico para mostrar en PDF
        
    Else
        ' Código para mostrar en Word
        MsgBox "Mostrar en Word"
        ' Aquí puedes agregar el código específico para mostrar en Word
        
    End If

    On Error GoTo err2

    Dim Obj As PageSet.PrinterControl
    Set Obj = New PrinterControl
    
    Dim r As Recordset

    Materiales.Sections("cabeza").Controls("lblOT").caption = ": "

    Materiales.Sections("s3").Controls("lblbarcode").caption = "*"

    Set Materiales.DataSource = r
    Obj.ChngOrientationLandscape
    Materiales.Show 1
    Obj.ReSetOrientation    'This resets the printer to portrait.

err2:
    MsgBox Err.Description

    Obj.ReSetOrientation

End Sub

'''''''''''''''Private Sub btnImprimirNuevo_Click()
''''''''''''''''Variable de tipo Aplicación de Word
'''''''''''''''Dim o_Word As New Word.Application
''''''''''''''''Variable de tipo documento de Word
'''''''''''''''Dim documento As Word.Document
'''''''''''''''' variable para hacer referencia al párrafo
'''''''''''''''Dim oSelection As Word.Selection
'''''''''''''''Dim Parrafo As table
''''''''''''''''F es para recorrer la Fila y C para la Columna
'''''''''''''''
'''''''''''''''Dim F, c As Double
''''''''''''''''Nuevo instancia del objeto
'''''''''''''''' Set o_Word = New Word.Application
''''''''''''''''Agrega un Nuevo documento de word
'''''''''''''''Set documento = o_Word.Documents.Add()
'''''''''''''''Set oSelection = o_Word.Selection
'''''''''''''''
'''''''''''''''
''''''''''''''''Creamos una tabla dentro del nuevo documento
'''''''''''''''''Set Parrafo = Documento.Tables.Add(Documento.Range(0, 0), _
'''''''''''''''''grid.rows, grid.Cols)
'''''''''''''''
'''''''''''''''documento.PageSetup.Orientation = wdOrientLandscape
'''''''''''''''
''''''''''''''''Recorremos el Flexgrid para agregar las columnas y filas a nuestra tabla
''''''''''''''''For C = 0 To grid.Cols - 1
''''''''''''''''Agregar columnas
''''''''''''''''Parrafo.Cell(0, C + 1).Range.text = grid.TextMatrix(0, C)
''''''''''''''''Agregar filas
''''''''''''''''For F = 0 To grid.rows - 1
''''''''''''''''Parrafo.Cell(F + 1, C + 1).Range.text = grid.TextMatrix(F, C)
''''''''''''''''Next F
''''''''''''''''Next C
''''''''''''''''Parrafo.Columns.item(1).Width = o_Word.InchesToPoints(2)
''''''''''''''''Parrafo.Columns.item(3).Width = o_Word.InchesToPoints(1.2)
''''''''''''''''Parrafo.al
''''''''''''''''Hacemos visible el word
'''''''''''''''o_Word.Visible = True
'''''''''''''''
'''''''''''''''
''''''''''''''''Documento.PrintPreview
'''''''''''''''
'''''''''''''''
'''''''''''''''documento.SaveAs "c:\word.tmp"
'''''''''''''''o_Word.Quit
'''''''''''''''Set o_Word = Nothing
'''''''''''''''Kill "c:\word.tmp"
'''''''''''''''
'''''''''''''''
'''''''''''''''
''''''''''''''''Eliminamos los objetos
'''''''''''''''Set o_Word = Nothing
'''''''''''''''Set documento = Nothing
'''''''''''''''Set Parrafo = Nothing
'''''''''''''''
''''''''''''''''ps.ReSetOrientation
'''''''''''''''Exit Sub
'''''''''''''''
''''''''''''''''error
'''''''''''''''ErrSub:
'''''''''''''''
'''''''''''''''MsgBox Err.Description
'''''''''''''''
'''''''''''''''On Error Resume Next
'''''''''''''''
'''''''''''''''Set o_Word = Nothing
'''''''''''''''Set documento = Nothing
'''''''''''''''Set Parrafo = Nothing
'''''''''''''''End Sub

Private Sub btnPrueba_03_Click()
    Dim f125 As New frmAdminPagosTransferenciasBancarias
    f125.Show
        
End Sub

'Private Sub btnPrueba_04_Click()
'    Dim f126 As New frmPlaneamientoRemitosListaCompleta
'    f126.Show
'
'End Sub

Private Sub btnPrueba2_Click()
    TraerTransferencias
    
End Sub

Public Sub TraerTransferencias()
    
    Dim transf As operacion
    
    Set colTransferenciasB = DAOTransferenciaBcaria.FindAll(Banco)
    
    For Each transf In colTransferenciasB
'       Debug.Print (transf.Monto & " | " & transf.Comprobante)
    Next
        
'    Debug.Print (colTransferenciasB.count)
    
End Sub

Public Sub CompletarGirdEx()
    MsgBox ("Hola")
End Sub

Private Sub btnTest01_Click()
    TraerDatos
    
End Sub


Public Sub TraerDatos()
    Dim prov As clsProveedor
        frmLoading.ProgressBar.min = 0
        
    Set colProveedores = DAOProveedor.FindAll()
    
        frmLoading.ProgressBar.max = colProveedores.count
    
    Dim i As Integer
    i = 0
    
    For Each prov In colProveedores
    'aca va el iterable
    
    i = i + 1
    
    frmLoading.ProgressBar.value = i
        
    Next
    
    If i = frmLoading.ProgressBar.max Then
        Unload frmLoading
    End If
    
    MostrarFacturas

        

End Sub

Private Sub MostrarFacturas()

        Dim fac As clsFacturaProveedor
            frmLoading.ProgressBar.min = 0
        
        Set colFacturas = DAOFacturaProveedor.FindAll("AdminComprasFacturasProveedores.estado = " & EstadoFacturaProveedor.Aprobada & "", False, "proveedores.razon ASC", False, True)
            frmLoading.ProgressBar.max = colFacturas.count
        
        Dim i As Integer
        i = 0
        
        For Each fac In colFacturas
        i = i + 1
            'aca va otro iterable
            frmLoading.ProgressBar.value = i
        
        Next
        
        If i = frmLoading.ProgressBar.max Then
            Unload frmLoading
        End If

End Sub

Private Sub Form_Load()

    Me.Height = 8040
    Me.Width = 2535

End Sub


Private Sub PushButton_Click()
    Dim f12324 As New frmAdminPagosLiquidaciondeCajaLista
    f12324.Show

End Sub
