VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmPlaneamientoRemitosListaProceso 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Remitos en proceso"
   ClientHeight    =   6240
   ClientLeft      =   3600
   ClientTop       =   2055
   ClientWidth     =   10125
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmListaRemitosProceso.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   10125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   5160
      Width           =   9855
      _Version        =   786432
      _ExtentX        =   17383
      _ExtentY        =   1508
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.PushButton PushButtonAceptar 
         Height          =   495
         Left            =   7800
         TabIndex        =   4
         Top             =   240
         Width           =   1815
         _Version        =   786432
         _ExtentX        =   3201
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Aceptar"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButtonCerrar 
         Height          =   495
         Left            =   240
         TabIndex        =   3
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
         Left            =   2040
         TabIndex        =   5
         Top             =   240
         Width           =   5655
         _Version        =   786432
         _ExtentX        =   9975
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Seleccione el Remito y luego Acepte para continuar."
         BackColor       =   -2147483633
         Alignment       =   2
      End
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   4755
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   9780
      _ExtentX        =   17251
      _ExtentY        =   8387
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
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Command1"
      Height          =   375
      Left            =   8640
      TabIndex        =   0
      Top             =   1080
      Width           =   855
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

Private Sub Form_Activate()
    If Me.GridEX1.ItemCount = 0 Then
        MsgBox "No hay remitos en proceso.", vbInformation
        Set Selecciones.RemitoElegido = Nothing
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    FormHelper.Customize Me
    GridEXHelper.CustomizeGrid Me.GridEX1
    Me.GridEX1.ItemCount = 0
    LlenarGrid
    
    'Me.caption = caption & " (" & Name & ")"


End Sub

Private Sub Form_Terminate()
    'Set Selecciones.RemitoElegido = Nothing
    'Unload Me
End Sub
Private Sub LlenarGrid()
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


    Set col = New Collection
    
    
    Dim tmpCol As New Collection
    Set tmpCol = DAORemitoS.FindAll(filtro)

    Dim rto As Remito
    
    

    
'    MsgBox (remitos.Detalles)
        
    For Each rto In tmpCol
        '  If rto.estado = RemitoAprobado And (rto.EstadoFacturado = RemitoFacturadoParcial Or rto.EstadoFacturado = RemitoNoFacturado) Then

        col.Add rto, CStr(rto.Id)
        ' If
        
    Dim detallesRemito As New Collection
    Set detallesRemito = DAORemitoSDetalle.FindAllByRemito(rto.Id, False, True)
    
    Next

    If col.count > 0 Then

        Me.GridEX1.ItemCount = 0
        Me.GridEX1.ItemCount = col.count
        
    End If

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
    Set Remito = col.item(Me.GridEX1.RowIndex(Me.GridEX1.row))
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If col.count > 0 And RowIndex > 0 Then
        Set Remito = col.item(RowIndex)
        Values(1) = Remito.numero
        Values(2) = Remito.FEcha
        Values(3) = Remito.cliente.razon
        Values(4) = Remito.detalle
        Values(5) = Remito.VerEstadoFacturado
        
'e5fs52- ACA SE VA A MOSTRAR LOS DATOS DE LAS OTS QUE CONTIENE ESE REMITO, PUEDE SER ID O CONCEPTO

        Values(6) = "OT N°, OT N°, Concepto, Concepto"
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
