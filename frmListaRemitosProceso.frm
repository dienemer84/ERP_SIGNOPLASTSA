VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmPlaneamientoRemitosListaProceso 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Remitos en proceso"
   ClientHeight    =   6330
   ClientLeft      =   3600
   ClientTop       =   2055
   ClientWidth     =   11400
   ClipControls    =   0   'False
   Icon            =   "frmListaRemitosProceso.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   11400
   ShowInTaskbar   =   0   'False
   Begin GridEX20.GridEX GridEX1 
      Height          =   6315
      Left            =   15
      TabIndex        =   1
      Top             =   0
      Width           =   11340
      _ExtentX        =   20003
      _ExtentY        =   11139
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      ColumnAutoResize=   -1  'True
      MethodHoldFields=   -1  'True
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   5
      Column(1)       =   "frmListaRemitosProceso.frx":000C
      Column(2)       =   "frmListaRemitosProceso.frx":0124
      Column(3)       =   "frmListaRemitosProceso.frx":0210
      Column(4)       =   "frmListaRemitosProceso.frx":0304
      Column(5)       =   "frmListaRemitosProceso.frx":03F8
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmListaRemitosProceso.frx":04F4
      FormatStyle(2)  =   "frmListaRemitosProceso.frx":062C
      FormatStyle(3)  =   "frmListaRemitosProceso.frx":06DC
      FormatStyle(4)  =   "frmListaRemitosProceso.frx":0790
      FormatStyle(5)  =   "frmListaRemitosProceso.frx":0868
      FormatStyle(6)  =   "frmListaRemitosProceso.frx":0920
      ImageCount      =   0
      PrinterProperties=   "frmListaRemitosProceso.frx":0A00
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
    '



    Set col = New Collection
    Dim tmpCol As New Collection
    Set tmpCol = DAORemitoS.FindAll(filtro)


    Dim rto As Remito

    For Each rto In tmpCol
        '  If rto.estado = RemitoAprobado And (rto.EstadoFacturado = RemitoFacturadoParcial Or rto.EstadoFacturado = RemitoNoFacturado) Then

        col.Add rto, CStr(rto.id)
        ' If

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
        vrto = Remito.id
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
        Values(3) = Remito.Cliente.razon
        Values(4) = Remito.detalle
        If IsSomething(Remito.contacto) Then Values(5) = Remito.contacto.nombre
    End If
End Sub
