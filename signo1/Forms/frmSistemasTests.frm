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
       Debug.Print (transf.Monto & " | " & transf.Comprobante)
    Next
        
    Debug.Print (colTransferenciasB.count)
    
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
