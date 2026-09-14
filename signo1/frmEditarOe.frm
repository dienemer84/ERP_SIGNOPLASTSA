VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmPlaneamientoOEEditar 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Editar OE"
   ClientHeight    =   8760
   ClientLeft      =   3720
   ClientTop       =   1995
   ClientWidth     =   9780
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   9780
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   8300
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Guadar"
      Height          =   375
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   8300
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ Orígen stock ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   9735
      Begin XtremeSuiteControls.ComboBox cboClientes 
         Height          =   315
         Left            =   840
         TabIndex        =   24
         Top             =   360
         Width           =   7455
         _Version        =   786432
         _ExtentX        =   13150
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Text            =   "ComboBox1"
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ver Stock"
         Default         =   -1  'True
         Height          =   375
         Left            =   8520
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   360
         Width           =   1095
      End
      Begin MSComctlLib.ListView lstStockPositivo 
         Height          =   2175
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   3836
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Detalle"
            Object.Width           =   12577
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Cantidad"
            Object.Width           =   2434
         EndProperty
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label idCliente 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Label6"
         Height          =   255
         Left            =   8760
         TabIndex        =   18
         Top             =   600
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ Entrega ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   3240
      Width           =   9735
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "[ Detalle ]"
         Height          =   3375
         Left            =   120
         TabIndex        =   3
         Top             =   1440
         Width           =   9495
         Begin XtremeSuiteControls.ComboBox cboClientesDestino 
            Height          =   315
            Left            =   840
            TabIndex        =   25
            Top             =   3000
            Width           =   4335
            _Version        =   786432
            _ExtentX        =   7646
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Text            =   "ComboBox1"
         End
         Begin VB.CommandButton q 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Quitar"
            Height          =   255
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   2640
            Width           =   735
         End
         Begin VB.ComboBox cboMonedas 
            Height          =   315
            Left            =   6000
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   3000
            Width           =   1215
         End
         Begin VB.TextBox txrRefe 
            Height          =   285
            Left            =   1080
            TabIndex        =   4
            Top             =   360
            Width           =   8295
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   255
            Left            =   8160
            TabIndex        =   5
            Top             =   3000
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            _Version        =   393216
            Format          =   65601537
            CurrentDate     =   38923
         End
         Begin MSComctlLib.ListView lstOE 
            Height          =   1815
            Left            =   120
            TabIndex        =   6
            Top             =   720
            Width           =   9255
            _ExtentX        =   16325
            _ExtentY        =   3201
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   6
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Detalle"
               Object.Width           =   6703
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Cantidad"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Valor"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Cliente"
               Object.Width           =   3792
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "idCliente"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Stock"
               Object.Width           =   1693
            EndProperty
         End
         Begin VB.Label Label8 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Moneda"
            Height          =   255
            Left            =   5160
            TabIndex        =   21
            Top             =   3000
            Width           =   735
         End
         Begin VB.Label Label4 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Cliente"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   3000
            Width           =   735
         End
         Begin VB.Label Label6 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Entrega"
            Height          =   255
            Left            =   7440
            TabIndex        =   8
            Top             =   3000
            Width           =   615
         End
         Begin VB.Label Label7 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Referencia"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.TextBox txtCantidad 
         Height          =   285
         Left            =   4920
         TabIndex        =   2
         Text            =   "0"
         Top             =   600
         Width           =   3375
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Agregar"
         Height          =   375
         Left            =   8520
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Detalle"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblDetalle 
         BackColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   960
         TabIndex        =   13
         Top             =   360
         Width           =   7575
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cantidad Disponible"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label lblCantDispo 
         BackColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   2160
         TabIndex        =   11
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cantidad Requerida"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   10
         Top             =   600
         Width           =   1815
      End
   End
   Begin VB.Label idPieza 
      Height          =   255
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "frmPlaneamientoOEEditar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim clasea As New classAdministracion
Dim grabado As Boolean
Dim mCargando As Boolean
Dim claseC As New classConfigurar
Dim vidOe As Long
Dim rss As Recordset
Dim rs As Recordset
Dim claseSP As New classSignoplast
Dim IdMoneda As Long
Dim claseS As New classStock

Dim claseP As New classPlaneamiento
Dim Cantidad As Double
Dim detalle As String
Dim idStock As Long
Dim vValor As Double
Dim c As Long

Public Property Let idOE(nidoe As Long)
    vidOe = nidoe
End Property

Private Sub llenarLstClientes(rs As Recordset)
    Dim x As ListItem
    lstStockPositivo.ListItems.Clear
    While Not rs.EOF
        Set x = Me.lstStockPositivo.ListItems.Add(, , rs!detalle)
        x.SubItems(1) = rs!Cantidad
        x.SubItems(2) = rs!razon
        x.SubItems(3) = rs!id_cliente
        x.Tag = rs!idPieza
        rs.MoveNext
    Wend
End Sub


Public Sub LlenarListaOE()

    On Error GoTo errHandler

    Set ordenes = DAOOrdenDeEntrega.GetAll()

    Me.gridEntregas.ItemCount = 0

    If Not ordenes Is Nothing Then
        Me.gridEntregas.ItemCount = ordenes.count
    End If

    Exit Sub

errHandler:

    MsgBox "Error al cargar las Ordenes de Entrega." & vbCrLf & _
           "Error " & Err.Number & vbCrLf & _
           Err.Description & vbCrLf & _
           "Origen: " & Err.Source, _
           vbCritical, "Ordenes de Entrega"

End Sub




Private Sub cboMonedas_Click()

    On Error GoTo errHandler
    
    If mCargando Then Exit Sub

    Dim monedaAnterior As Long
    Dim monedaNueva As Long

    If Me.cboMonedas.ListIndex < 0 Then Exit Sub

    monedaNueva = CLng(Me.cboMonedas.ItemData(Me.cboMonedas.ListIndex))
    monedaAnterior = IdMoneda

    'Primera carga del formulario
    If monedaAnterior <= 0 Then
        IdMoneda = monedaNueva
        Exit Sub
    End If

    If monedaAnterior = monedaNueva Then Exit Sub

    cambiarPrecios monedaAnterior, monedaNueva

    IdMoneda = monedaNueva
    grabado = False

    Exit Sub

errHandler:

    MsgBox "Error al cambiar la moneda de la Orden de Entrega." & _
           vbCrLf & _
           "Error " & Err.Number & vbCrLf & _
           Err.Description, _
           vbCritical, _
           "Orden de Entrega"

End Sub


Private Sub Command1_Click()
'If MsgBox("¿Desea crear una nueva O/E?", vbYesNo, "Confirmación") = vbYes Then
'    Me.lstOE.ListItems.Clear
'    Me.lstStockPositivo.ListItems.Clear
    llenarListaStock
    'End If

End Sub


Private Sub Command2_Click()

    On Error GoTo errHandler

    Dim cantpedida As Double
    Dim esta As Boolean
    Dim valorPieza As Double
    Dim idMonedaPieza As Long
    Dim i As Long
    Dim cantidadNueva As Double

    Dim itemOE As MSComctlLib.ListItem
    Dim piezaSeleccionada As Pieza


    '--------------------------------------------------
    ' VALIDAR QUE HAYA UNA PIEZA SELECCIONADA
    '--------------------------------------------------
    If Me.lstStockPositivo.selectedItem Is Nothing Then

        MsgBox "Seleccione una pieza del stock.", _
               vbExclamation, _
               "Orden de Entrega"

        Exit Sub

    End If


    '--------------------------------------------------
    ' VALIDAR CANTIDAD
    '--------------------------------------------------
    If Len(Trim$(Me.txtCantidad.Text)) = 0 Then

        MsgBox "Ingrese la cantidad requerida.", _
               vbExclamation, _
               "Orden de Entrega"

        Me.txtCantidad.SetFocus
        Exit Sub

    End If


    If Not IsNumeric(Me.txtCantidad.Text) Then

        MsgBox "La cantidad ingresada no es válida.", _
               vbExclamation, _
               "Orden de Entrega"

        Me.txtCantidad.SetFocus
        Exit Sub

    End If


    cantpedida = CDbl(Me.txtCantidad.Text)

    If cantpedida <= 0 Then

        MsgBox "La cantidad debe ser mayor a cero.", _
               vbExclamation, _
               "Orden de Entrega"

        Me.txtCantidad.SetFocus
        Exit Sub

    End If


    '--------------------------------------------------
    ' PIEZA SELECCIONADA
    '--------------------------------------------------
    idStock = CLng(Me.lstStockPositivo.selectedItem.Tag)

    detalle = Me.lstStockPositivo.selectedItem.Text

    Cantidad = CDbl( _
        Me.lstStockPositivo.selectedItem.ListSubItems(1).Text _
    )


    '--------------------------------------------------
    ' CONTROL DE STOCK
    '--------------------------------------------------
    If cantpedida > Cantidad Then

        MsgBox "No hay stock suficiente de esta pieza." & vbCrLf & _
               "Disponible: " & funciones.FormatearDecimales(Cantidad, 2) & vbCrLf & _
               "Solicitado: " & funciones.FormatearDecimales(cantpedida, 2), _
               vbExclamation, _
               "Orden de Entrega"

        Exit Sub

    End If


    '--------------------------------------------------
    ' VER SI YA ESTA EN LA OE
    '--------------------------------------------------
    esta = False

    For i = 1 To Me.lstOE.ListItems.count

        If CLng(Me.lstOE.ListItems(i).Tag) = idStock Then

            esta = True

            cantidadNueva = _
                CDbl(Me.lstOE.ListItems(i).ListSubItems(1).Text) + _
                cantpedida


            If cantidadNueva > Cantidad Then

                MsgBox "No hay disponibilidad de stock suficiente." & vbCrLf & _
                       "Disponible: " & funciones.FormatearDecimales(Cantidad, 2) & vbCrLf & _
                       "Cantidad total solicitada: " & _
                       funciones.FormatearDecimales(cantidadNueva, 2), _
                       vbExclamation, _
                       "Orden de Entrega"

                Exit Sub

            End If


            Me.lstOE.ListItems(i).ListSubItems(1).Text = _
                funciones.FormatearDecimales(cantidadNueva, 2)

            Exit For

        End If

    Next i


    '--------------------------------------------------
    ' SI NO EXISTE, AGREGARLA
    '--------------------------------------------------
    If Not esta Then

        valorPieza = claseP.precio_pieza2( _
                        idStock, _
                        idMonedaPieza)


        'Convertir precio si la moneda es distinta
        If IdMoneda > 0 And _
           idMonedaPieza > 0 And _
           IdMoneda <> idMonedaPieza Then

            valorPieza = clasea.realizaCambio( _
                            valorPieza, _
                            idMonedaPieza, _
                            IdMoneda)

        End If


        Set piezaSeleccionada = _
            DAOPieza.FindById(idStock, FL_0)


        If piezaSeleccionada Is Nothing Then

            MsgBox "No se pudo recuperar la pieza seleccionada.", _
                   vbCritical, _
                   "Orden de Entrega"

            Exit Sub

        End If


        If piezaSeleccionada.Cliente Is Nothing Then

            MsgBox "La pieza seleccionada no tiene un cliente asociado.", _
                   vbExclamation, _
                   "Orden de Entrega"

            Exit Sub

        End If


        Set itemOE = _
            Me.lstOE.ListItems.Add( _
                , _
                , _
                detalle)


        itemOE.SubItems(1) = _
            funciones.FormatearDecimales(cantpedida, 2)

        itemOE.SubItems(2) = _
            funciones.FormatearDecimales(valorPieza, 2)

        itemOE.SubItems(3) = _
            piezaSeleccionada.Cliente.razon

        itemOE.SubItems(4) = _
            piezaSeleccionada.Cliente.Id

        itemOE.SubItems(5) = _
            funciones.FormatearDecimales(Cantidad, 2)

        itemOE.Tag = idStock

    End If


    grabado = False

    Me.txtCantidad.Text = "0"

    Exit Sub


errHandler:

    MsgBox "Error al agregar la pieza a la Orden de Entrega." & _
           vbCrLf & _
           "Error " & Err.Number & vbCrLf & _
           Err.Description, _
           vbCritical, _
           "Orden de Entrega"

End Sub


Private Sub Command3_Click()

    On Error GoTo errHandler

    Dim refe As String
    Dim idClienteDestino As Long
    Dim idMonedaSeleccionada As Long


    '--------------------------------------------------
    ' VALIDAR OE
    '--------------------------------------------------
    If vidOe <= 0 Then

        MsgBox "No se pudo determinar la Orden de Entrega a modificar.", _
               vbCritical, _
               "Orden de Entrega"

        Exit Sub

    End If


    '--------------------------------------------------
    ' VALIDAR CLIENTE DESTINO
    '--------------------------------------------------
    If Me.cboClientesDestino.ListIndex < 0 Then

        MsgBox "Seleccione el cliente destino de la Orden de Entrega.", _
               vbExclamation, _
               "Orden de Entrega"

        Me.cboClientesDestino.SetFocus
        Exit Sub

    End If


    idClienteDestino = _
        CLng(Me.cboClientesDestino.ItemData(Me.cboClientesDestino.ListIndex))


    If idClienteDestino <= 0 Then

        MsgBox "Debe seleccionar un cliente destino válido.", _
               vbExclamation, _
               "Orden de Entrega"

        Me.cboClientesDestino.SetFocus
        Exit Sub

    End If


    '--------------------------------------------------
    ' VALIDAR MONEDA
    '--------------------------------------------------
    If Me.cboMonedas.ListIndex < 0 Then

        MsgBox "Seleccione la moneda de la Orden de Entrega.", _
               vbExclamation, _
               "Orden de Entrega"

        Me.cboMonedas.SetFocus
        Exit Sub

    End If


    idMonedaSeleccionada = _
        CLng(Me.cboMonedas.ItemData(Me.cboMonedas.ListIndex))


    If idMonedaSeleccionada <= 0 Then

        MsgBox "La moneda seleccionada no es válida.", _
               vbExclamation, _
               "Orden de Entrega"

        Me.cboMonedas.SetFocus
        Exit Sub

    End If


    '--------------------------------------------------
    ' REFERENCIA
    '--------------------------------------------------
    refe = normaliza(Trim$(Me.txrRefe.Text))


    '--------------------------------------------------
    ' CONFIRMAR
    '--------------------------------------------------
    If MsgBox( _
        "¿Desea guardar los cambios de la Orden de Entrega Nro. " & _
        CStr(vidOe) & "?", _
        vbYesNo + vbQuestion + vbDefaultButton2, _
        "Confirmación") <> vbYes Then

        Exit Sub

    End If


    '--------------------------------------------------
    ' GUARDAR
    '--------------------------------------------------
    If claseP.editarOE( _
            Me.lstOE, _
            Me.DTPicker1.value, _
            refe, _
            idClienteDestino, _
            vidOe, _
            idMonedaSeleccionada) Then
    
    
        IdMoneda = idMonedaSeleccionada
    
        grabado = True
    
    
        '--------------------------------------------------
        ' ACTUALIZAR LISTADO PRINCIPAL
        '--------------------------------------------------
        If FormularioCargado("frmPlaneamientoOELista") Then
    
            frmPlaneamientoOELista.RefrescarListadoActual
    
        End If
    
    
        MsgBox "Los cambios de la Orden de Entrega Nro. " & _
               CStr(vidOe) & _
               " fueron guardados correctamente.", _
               vbInformation, _
               "Orden de Entrega"

    Else

        grabado = False

        MsgBox "No fue posible guardar los cambios de la Orden de Entrega.", _
               vbExclamation, _
               "Orden de Entrega"

    End If


    Exit Sub


errHandler:

    grabado = False

    MsgBox "Error al guardar la Orden de Entrega Nro. " & _
           CStr(vidOe) & "." & vbCrLf & _
           "Error " & Err.Number & vbCrLf & _
           Err.Description, _
           vbCritical, _
           "Orden de Entrega"

End Sub


Private Sub Command4_Click()

    If grabado Then

        Unload Me

    Else

        If MsgBox( _
            "¿Está seguro de salir?" & vbCrLf & _
            "Hay cambios que no fueron guardados.", _
            vbYesNo + vbQuestion + vbDefaultButton2, _
            "Confirmación") = vbYes Then

            Unload Me

        End If

    End If

End Sub


Private Sub Form_Load()

    On Error GoTo errHandler

    mCargando = True
    grabado = True

    FormHelper.Customize Me

    DAOCliente.llenarComboXtremeSuite Me.cboClientes, True
    DAOCliente.llenarComboXtremeSuite Me.cboClientesDestino, True

    DAOMoneda.LlenarCombo Me.cboMonedas

    Me.DTPicker1.value = Now

    llenarDatosOE

    'Terminó la carga inicial
    grabado = True
    mCargando = False

    Exit Sub

errHandler:

    mCargando = False

    MsgBox "Error al cargar la Orden de Entrega." & vbCrLf & _
           "Error " & Err.Number & vbCrLf & _
           Err.Description, _
           vbCritical, _
           "Orden de Entrega"

End Sub

Public Sub llenarDatosOE()
    Dim x As ListItem
    On Error GoTo err551
    Set rs = conectar.RSFactory("select fecha,referencia,idmoneda,idCliente from PedidosEntregas where id=" & vidOe)
    While Not rs.EOF
        Me.DTPicker1 = rs!FEcha
        Me.txrRefe = rs!referencia
        
        Me.cboMonedas.ListIndex = funciones.PosIndexCbo(rs!IdMoneda, Me.cboMonedas)
        
        If Me.cboMonedas.ListIndex >= 0 Then
            IdMoneda = CInt(Me.cboMonedas.ItemData(Me.cboMonedas.ListIndex))
        End If
        
        Me.cboClientesDestino.ListIndex = funciones.PosIndexCbo(rs!idCliente, Me.cboClientesDestino)
        rs.MoveNext
    Wend

    Set rs = conectar.RSFactory("Select s.cantidad as cantStock,s.id as idpieza,s.detalle,dp.cantidad,dp.vale,c.razon,c.id as idcliente from stock s,detallesPedidosEntregas dp, PedidosEntregas p, clientes c  where idPedidoEntrega=" & vidOe & " And s.id_cliente = c.id And p.id = dp.idPedidoEntrega And dp.idPieza = s.id")
    While Not rs.EOF
        Set x = Me.lstOE.ListItems.Add(, , rs!detalle)
        x.SubItems(1) = funciones.FormatearDecimales(rs!Cantidad, 2)
        x.SubItems(2) = funciones.FormatearDecimales(rs!vale, 2)
        x.SubItems(3) = rs!razon
        x.SubItems(4) = rs!idCliente
        x.SubItems(5) = rs!cantStock
        x.Tag = rs!idPieza


        If rs!cantStock < rs!Cantidad Then
            x.ForeColor = vbRed
            x.ListSubItems(1).ForeColor = vbRed
            x.ListSubItems(2).ForeColor = vbRed
            x.ListSubItems(3).ForeColor = vbRed
            x.ListSubItems(4).ForeColor = vbRed
            x.ListSubItems(5).ForeColor = vbRed
        End If

        rs.MoveNext
    Wend

    Exit Sub
err551:
    MsgBox Err.Description


End Sub


Private Sub verMarcado()

    If Me.lstStockPositivo.ListItems.count > 0 Then
        idStock = CLng(Me.lstStockPositivo.selectedItem.Tag)
        detalle = Me.lstStockPositivo.selectedItem
        Cantidad = CDbl(Me.lstStockPositivo.selectedItem.ListSubItems(1).Text)
        Me.lblCantDispo = Cantidad
        Me.lblDetalle = detalle
        Me.idPieza = idStock
    End If
    
End Sub


Private Sub Form_Terminate()
    Set rss = Nothing
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set rss = Nothing
End Sub


Private Sub lstOE_DblClick()
    If Me.lstOE.ListItems.count > 0 Then
        funciones.valorOE = funciones.FormatearDecimales(Me.lstOE.selectedItem.ListSubItems(2), 2)
        funciones.cantOE = funciones.FormatearDecimales(Me.lstOE.selectedItem.ListSubItems(1), 2)
        frmPlaneamientoOEModificarCantidad.Show 1
        Me.lstOE.selectedItem.ListSubItems(2) = FormatearDecimales(funciones.valorOE, 2)
        Me.lstOE.selectedItem.ListSubItems(1) = FormatearDecimales(funciones.cantOE, 2)
        grabado = False
    End If
End Sub


Private Sub lstStockPositivo_ItemClick(ByVal item As MSComctlLib.ListItem)
    verMarcado
End Sub


Private Sub q_Click()

    On Error GoTo errHandler

    Dim i As Long
    Dim cantidadEliminados As Long

    cantidadEliminados = 0

    If Me.lstOE.ListItems.count = 0 Then
        MsgBox "No hay items para quitar.", _
               vbExclamation, _
               "Orden de Entrega"
        Exit Sub
    End If

    If MsgBox( _
        "¿Está seguro de eliminar los items seleccionados?", _
        vbYesNo + vbQuestion, _
        "Confirmación") <> vbYes Then

        Exit Sub

    End If

    For i = Me.lstOE.ListItems.count To 1 Step -1

        If Me.lstOE.ListItems(i).Checked Then

            Me.lstOE.ListItems.remove i

            cantidadEliminados = cantidadEliminados + 1
            grabado = False

        End If

    Next i

    If cantidadEliminados = 0 Then

        MsgBox "No había ningún item marcado para quitar.", _
               vbInformation, _
               "Orden de Entrega"

    End If

    Exit Sub

errHandler:

    MsgBox "Error al quitar items de la Orden de Entrega." & vbCrLf & _
           "Error " & Err.Number & vbCrLf & _
           Err.Description, _
           vbCritical, _
           "Orden de Entrega"

End Sub

Private Sub txtCantidad_GotFocus()
    foco Me.txtCantidad
End Sub


Private Sub txtCantidad_Validate(Cancel As Boolean)
    If Not IsNumeric(Me.txtCantidad) Then Cancel = True
End Sub


Private Sub cambiarPrecios( _
    ByVal monedaOrigen As Long, _
    ByVal MonedaDestino As Long)

    On Error GoTo errHandler

    Dim vale As Double
    Dim x As Long

    If monedaOrigen <= 0 Then Exit Sub
    If MonedaDestino <= 0 Then Exit Sub
    If monedaOrigen = MonedaDestino Then Exit Sub

    For x = 1 To Me.lstOE.ListItems.count

        vale = CDbl(Me.lstOE.ListItems(x).ListSubItems(2).Text)

        vale = clasea.realizaCambio( _
                    vale, _
                    monedaOrigen, _
                    MonedaDestino)

        Me.lstOE.ListItems(x).ListSubItems(2).Text = _
            funciones.FormatearDecimales(vale, 2)

    Next x

    Exit Sub

errHandler:

    MsgBox "Error al convertir los valores de la Orden de Entrega." & _
           vbCrLf & _
           "Error " & Err.Number & vbCrLf & _
           Err.Description, _
           vbCritical, _
           "Orden de Entrega"

End Sub


Private Sub llenarListaStock()

    On Error GoTo errHandler

    Dim rsStock As ADODB.Recordset
    Dim strsql As String
    Dim idClienteSeleccionado As Long
    Dim itemStock As MSComctlLib.ListItem

    '--------------------------------------------
    ' Validar cliente seleccionado
    '--------------------------------------------
    If Me.cboClientes.ListCount = 0 Then

        MsgBox "No hay clientes cargados.", _
               vbExclamation, _
               "Orden de Entrega"

        Exit Sub

    End If

    If Me.cboClientes.ListIndex < 0 Then

        MsgBox "Seleccione un cliente para consultar el stock.", _
               vbExclamation, _
               "Orden de Entrega"

        Exit Sub

    End If


    '--------------------------------------------
    ' Obtener cliente
    '--------------------------------------------
    idClienteSeleccionado = _
        CLng(Me.cboClientes.ItemData(Me.cboClientes.ListIndex))

    Me.idCliente.caption = CStr(idClienteSeleccionado)


    '--------------------------------------------
    ' Armar consulta
    '--------------------------------------------
    If idClienteSeleccionado = -1 Then

        strsql = _
            "SELECT id, detalle, cantidad " & _
            "FROM stock " & _
            "WHERE cantidad > 0 " & _
            "ORDER BY detalle"

    Else

        strsql = _
            "SELECT id, detalle, cantidad " & _
            "FROM stock " & _
            "WHERE cantidad > 0 " & _
            "AND id_cliente = " & CStr(idClienteSeleccionado) & " " & _
            "ORDER BY detalle"

    End If


    '--------------------------------------------
    ' Ejecutar
    '--------------------------------------------
    Set rsStock = conectar.RSFactory(strsql)

    If rsStock Is Nothing Then

        MsgBox "No se pudo consultar el stock.", _
               vbCritical, _
               "Orden de Entrega"

        Exit Sub

    End If


    '--------------------------------------------
    ' Limpiar lista anterior
    '--------------------------------------------
    Me.lstStockPositivo.ListItems.Clear


    '--------------------------------------------
    ' Cargar stock
    '--------------------------------------------
    Do While Not rsStock.EOF

        Set itemStock = _
            Me.lstStockPositivo.ListItems.Add( _
                , _
                , _
                CStr(rsStock!detalle))

        itemStock.SubItems(1) = _
            funciones.FormatearDecimales(CDbl(rsStock!Cantidad), 2)

        itemStock.Tag = CLng(rsStock!Id)

        rsStock.MoveNext

    Loop


    Set rsStock = Nothing

    Exit Sub


errHandler:

    MsgBox "Error al cargar el stock." & vbCrLf & _
           "Error " & Err.Number & vbCrLf & _
           Err.Description & vbCrLf & vbCrLf & _
           "Consulta:" & vbCrLf & strsql, _
           vbCritical, _
           "Orden de Entrega"

    Set rsStock = Nothing

End Sub


Private Sub txrRefe_Change()

    If mCargando Then Exit Sub

    grabado = False

End Sub


Private Sub DTPicker1_Change()

    If mCargando Then Exit Sub

    grabado = False

End Sub


Private Sub cboClientesDestino_Click()

    If mCargando Then Exit Sub

    grabado = False

End Sub


Private Function FormularioCargado(ByVal nombreFormulario As String) As Boolean

    Dim frm As Form

    FormularioCargado = False

    For Each frm In Forms

        If StrComp(frm.Name, _
                   nombreFormulario, _
                   vbTextCompare) = 0 Then

            FormularioCargado = True
            Exit Function

        End If

    Next frm

End Function

