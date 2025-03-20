VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmComprasRequesNuevo 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nuevo requerimiento"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   255
   ClientWidth     =   13140
   ClipControls    =   0   'False
   Icon            =   "frmRequeCompra.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   13140
   Begin XtremeSuiteControls.GroupBox GroupBox4 
      Height          =   3300
      Left            =   75
      TabIndex        =   26
      Top             =   1755
      Width           =   12975
      _Version        =   786432
      _ExtentX        =   22886
      _ExtentY        =   5821
      _StockProps     =   79
      Caption         =   "Materiales"
      UseVisualStyle  =   -1  'True
      Begin GridEX20.GridEX GridEX 
         Height          =   2955
         Left            =   105
         TabIndex        =   27
         Top             =   225
         Width           =   12765
         _ExtentX        =   22516
         _ExtentY        =   5212
         Version         =   "2.0"
         PreviewRowIndent=   200
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         PreviewColumn   =   6
         PreviewRowLines =   1
         ColumnAutoResize=   -1  'True
         MultiSelect     =   -1  'True
         MethodHoldFields=   -1  'True
         BackColorInfoText=   16777215
         AllowDelete     =   -1  'True
         BorderStyle     =   3
         GroupByBoxVisible=   0   'False
         BackColorHeader =   16761024
         RowHeaders      =   -1  'True
         DataMode        =   99
         BackColorBkg    =   16777215
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   10
         Column(1)       =   "frmRequeCompra.frx":000C
         Column(2)       =   "frmRequeCompra.frx":012C
         Column(3)       =   "frmRequeCompra.frx":0220
         Column(4)       =   "frmRequeCompra.frx":0318
         Column(5)       =   "frmRequeCompra.frx":040C
         Column(6)       =   "frmRequeCompra.frx":0500
         Column(7)       =   "frmRequeCompra.frx":061C
         Column(8)       =   "frmRequeCompra.frx":0730
         Column(9)       =   "frmRequeCompra.frx":082C
         Column(10)      =   "frmRequeCompra.frx":092C
         FormatStylesCount=   9
         FormatStyle(1)  =   "frmRequeCompra.frx":0A24
         FormatStyle(2)  =   "frmRequeCompra.frx":0B5C
         FormatStyle(3)  =   "frmRequeCompra.frx":0C0C
         FormatStyle(4)  =   "frmRequeCompra.frx":0CC0
         FormatStyle(5)  =   "frmRequeCompra.frx":0D98
         FormatStyle(6)  =   "frmRequeCompra.frx":0E50
         FormatStyle(7)  =   "frmRequeCompra.frx":0F30
         FormatStyle(8)  =   "frmRequeCompra.frx":0FC0
         FormatStyle(9)  =   "frmRequeCompra.frx":1054
         ImageCount      =   0
         PrinterProperties=   "frmRequeCompra.frx":10E4
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox3 
      Height          =   1890
      Left            =   75
      TabIndex        =   23
      Top             =   5100
      Width           =   4020
      _Version        =   786432
      _ExtentX        =   7091
      _ExtentY        =   3334
      _StockProps     =   79
      Caption         =   "Entregas del Item"
      UseVisualStyle  =   -1  'True
      Begin GridEX20.GridEX grilla_entregas_materiales 
         Height          =   1560
         Left            =   105
         TabIndex        =   24
         Top             =   225
         Width           =   3810
         _ExtentX        =   6720
         _ExtentY        =   2752
         Version         =   "2.0"
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         CalendarTodayText=   "Hoy"
         CalendarNoneText=   "Nada"
         ColumnAutoResize=   -1  'True
         MethodHoldFields=   -1  'True
         AllowDelete     =   -1  'True
         BorderStyle     =   3
         GroupByBoxVisible=   0   'False
         BackColorHeader =   16761024
         RowHeaders      =   -1  'True
         DataMode        =   99
         AllowAddNew     =   -1  'True
         ColumnHeaderHeight=   285
         IntProp1        =   0
         ColumnsCount    =   2
         Column(1)       =   "frmRequeCompra.frx":12BC
         Column(2)       =   "frmRequeCompra.frx":13FC
         FormatStylesCount=   7
         FormatStyle(1)  =   "frmRequeCompra.frx":1530
         FormatStyle(2)  =   "frmRequeCompra.frx":1668
         FormatStyle(3)  =   "frmRequeCompra.frx":1718
         FormatStyle(4)  =   "frmRequeCompra.frx":17CC
         FormatStyle(5)  =   "frmRequeCompra.frx":18A4
         FormatStyle(6)  =   "frmRequeCompra.frx":195C
         FormatStyle(7)  =   "frmRequeCompra.frx":1A3C
         ImageCount      =   0
         PrinterProperties=   "frmRequeCompra.frx":1AC8
      End
   End
   Begin XtremeSuiteControls.PushButton Command5 
      Height          =   375
      Left            =   10695
      TabIndex        =   21
      Top             =   6600
      Width           =   2325
      _Version        =   786432
      _ExtentX        =   4101
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Guardar Requerimiento"
      BackColor       =   16744576
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.GroupBox GroupBox2 
      Height          =   1710
      Left            =   4260
      TabIndex        =   6
      Top             =   30
      Width           =   8775
      _Version        =   786432
      _ExtentX        =   15478
      _ExtentY        =   3016
      _StockProps     =   79
      Caption         =   "Material"
      BackColor       =   16744576
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   285
         Left            =   6255
         TabIndex        =   22
         Top             =   240
         Width           =   300
         _Version        =   786432
         _ExtentX        =   529
         _ExtentY        =   494
         _StockProps     =   79
         Caption         =   "..."
         BackColor       =   16744576
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cboMateriales 
         Height          =   315
         Left            =   1875
         TabIndex        =   20
         Top             =   225
         Width           =   4350
         _Version        =   786432
         _ExtentX        =   7673
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Sorted          =   -1  'True
         Style           =   2
         Text            =   "ComboBox1"
         AutoComplete    =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton Command1 
         Height          =   405
         Left            =   6990
         TabIndex        =   25
         Top             =   180
         Width           =   1620
         _Version        =   786432
         _ExtentX        =   2857
         _ExtentY        =   714
         _StockProps     =   79
         Caption         =   "Agregar a la lista"
         BackColor       =   16744576
         UseVisualStyle  =   -1  'True
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FF8080&
         Caption         =   "Rubro "
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
         TabIndex        =   19
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblRubro 
         BackColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   1560
         TabIndex        =   18
         Top             =   705
         Width           =   4155
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FF8080&
         Caption         =   "Grupo "
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
         TabIndex        =   17
         Top             =   945
         Width           =   1215
      End
      Begin VB.Label lblGrupo 
         BackColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   1560
         TabIndex        =   16
         Top             =   945
         Width           =   4155
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FF8080&
         Caption         =   "Descripción "
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
         TabIndex        =   15
         Top             =   1185
         Width           =   1215
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FF8080&
         Caption         =   "Kg x M2/Ml"
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
         Left            =   5805
         TabIndex        =   14
         Top             =   1185
         Width           =   1095
      End
      Begin VB.Label lblKgM2 
         BackColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   7005
         TabIndex        =   13
         Top             =   1185
         Width           =   1605
      End
      Begin VB.Label lblUnidad 
         BackColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   7005
         TabIndex        =   12
         Top             =   945
         Width           =   1605
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FF8080&
         Caption         =   "Unidad"
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
         Left            =   5925
         TabIndex        =   11
         Top             =   945
         Width           =   975
      End
      Begin VB.Label lblEspesor 
         BackColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   7005
         TabIndex        =   10
         Top             =   705
         Width           =   1605
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FF8080&
         Caption         =   "Espesor"
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
         Left            =   5925
         TabIndex        =   9
         Top             =   705
         Width           =   975
      End
      Begin VB.Label lblDescripcion 
         BackColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   1560
         TabIndex        =   8
         Top             =   1185
         Width           =   4155
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "Nombre Material"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   405
         TabIndex        =   7
         Top             =   300
         Width           =   1395
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   900
      Left            =   165
      TabIndex        =   2
      Top             =   720
      Width           =   1905
      _Version        =   786432
      _ExtentX        =   3360
      _ExtentY        =   1587
      _StockProps     =   79
      Caption         =   "Destino de materiales"
      BackColor       =   16744576
      UseVisualStyle  =   -1  'True
      Begin VB.TextBox lblOt 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   750
         TabIndex        =   5
         Top             =   225
         Width           =   1005
      End
      Begin VB.OptionButton OpOT 
         BackColor       =   &H00FF8080&
         Caption         =   "OT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   225
         Width           =   705
      End
      Begin VB.OptionButton OpStock 
         BackColor       =   &H00FF8080&
         Caption         =   "Stock"
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
         TabIndex        =   3
         Top             =   555
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin XtremeSuiteControls.ComboBox cboSectores 
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      Top             =   225
      Width           =   2340
      _Version        =   786432
      _ExtentX        =   4128
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Style           =   2
      Text            =   "ComboBox1"
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "Sector Solicitante"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   180
      TabIndex        =   0
      Top             =   285
      Width           =   1530
   End
   Begin VB.Menu mnuAcciones 
      Caption         =   "Acciones"
      Visible         =   0   'False
      Begin VB.Menu mnuAprobar 
         Caption         =   "Aprobar"
      End
      Begin VB.Menu mnuProcesarProveedores 
         Caption         =   "Procesar Proveedores"
      End
      Begin VB.Menu mnuFinProcesoProveedores 
         Caption         =   "Fin Proceso Proveedores"
      End
      Begin VB.Menu mnuCrearPO 
         Caption         =   "Crear PO"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAnular 
         Caption         =   "Anular"
      End
   End
End
Attribute VB_Name = "frmComprasRequesNuevo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements ISuscriber


Dim vRow As Long
Dim requeVerificador As clsRequerimiento
Dim Guardado As Boolean

Dim vDetalle As clsRequeMateriales
Dim vmaterial As clsMaterial
Dim tmpEntrega As clsRequeEntregas
Dim vReque As clsRequerimiento    'requerimiento en edicion


Dim vSoloVer As Boolean
Public Property Let SoloVer(nvalue As Boolean)
    vSoloVer = nvalue
End Property
Public Property Let Requerimiento(Id As Long)
    Set vReque = DAORequerimiento.FindById(Id, True, True, True, True)
    cargarDatosReque
End Property


Private Sub cboMateriales_Click()
    buscarMaterial
End Sub


Private Sub Command1_Click()
    If IsSomething(vmaterial) Then

        Set vDetalle = New clsRequeMateriales
        vDetalle.Largo = vmaterial.Largo
        vDetalle.Ancho = vmaterial.Ancho
        vDetalle.Material = vmaterial
        vDetalle.Entregas = New Collection
        vDetalle.ListaProveedores = New Collection

        If vReque.estado = EstadoRequeCompra.Finalizado_ Then
            vDetalle.estado = EstadoRequeCompra.Finalizado_
        ElseIf vReque.estado = EstadoRequeCompra.EnEdición_ Then
            vDetalle.estado = EstadoRequeCompra.EnEdición_
        End If

        vReque.Materiales.Add vDetalle
        Set vDetalle = Nothing
        GridEX.ItemCount = 0
        GridEX.ItemCount = vReque.Materiales.count
        GridEXHelper.AutoSizeColumns Me.GridEX

    Else
        MsgBox "Debe seleccionar un material.", vbOKOnly + vbExclamation
    End If
End Sub

Private Sub Guardar()
    Dim nuevo As Boolean
    Dim destino_ As destino

    If OpOT.value Then
        destino_ = ot_
    ElseIf OpStock.value Then
        destino_ = stock_
    End If

    If destino_ = ot_ And Not IsNumeric(Me.lblOT) Then
        MsgBox "Debe ingresar el nro de OT destino!", vbInformation
        Exit Sub
    End If


    If Me.cboSectores.ListIndex = -1 Then
        MsgBox "Debe ingresar el sector solicitante.", vbExclamation + vbOKOnly
        Exit Sub
    End If


    If Not (vReque.estado = EstadoRequeCompra.EnEdición_ Or vReque.estado = EstadoRequeCompra.Finalizado_) Then
        MsgBox "No puede guardar este requerimiento en el estado actual!", vbCritical
        Exit Sub
    End If


    Dim ot_destino As Long
    If Not IsNumeric(Me.lblOT) Then
        ot_destino = 0
    Else
        ot_destino = lblOT
    End If

    vReque.DestinoOT = ot_destino
    vReque.Tipo = destino_




    If vReque.Id = 0 Then
        vReque.fechaCreado = Date
        vReque.estado = EstadoRequeCompra.EnEdición_
        vReque.Usuario_creador = DAOUsuarios.GetById(funciones.getUser)
    End If



    vReque.Sector = DAOSectores.GetById(Me.cboSectores.ItemData(Me.cboSectores.ListIndex))

    nuevo = True

    Dim aid As Long
    aid = vReque.Id    'para saber si hay q agregarlo o modificar la colection


    If vReque.ValidarEntregas Then

        If DAORequerimiento.Save(vReque) Then
            Dim EVENTO As New clsEventoObserver

            If aid = 0 Then
                EVENTO.EVENTO = agregar_
            Else
                EVENTO.EVENTO = modificar_
            End If
            Set EVENTO.Elemento = vReque
            Set EVENTO.Originador = Me

            Channel.Notificar EVENTO, RequerimientosCompra_

            If aid = 0 Then
                MsgBox "Requerimiento Nº " & vReque.Id & " creado correctamente.", vbInformation + vbOKOnly
            Else
                MsgBox "Requerimiento modificado correctamente.", vbInformation + vbOKOnly
            End If
            Unload Me
        Else
            MsgBox "Se produjo algún error al grabar!", vbCritical, "Error"
        End If

    Else
        MsgBox "Debe definir las entregas correctamente!", vbCritical, "Error"
    End If
End Sub
Private Sub Command5_Click()
    Guardar
End Sub

Private Sub LlenarComboMateriales()
    Me.cboMateriales.Clear
    Dim mattt As clsMaterial
    For Each mattt In DAOMateriales.FindAll()
        Me.cboMateriales.AddItem mattt.codigo & " - " & mattt.descripcion
        Me.cboMateriales.ItemData(Me.cboMateriales.NewIndex) = mattt.Id
    Next mattt

End Sub

Private Sub Form_Load()
    FormHelper.Customize Me

    GridEXHelper.CustomizeGrid Me.GridEX, False, True
    GridEXHelper.CustomizeGrid Me.grilla_entregas_materiales, False, True

    Me.Command5.Enabled = Not vSoloVer

    Guardado = True
    DAOSectores.LlenarComboXtreme Me.cboSectores
    Me.GridEX.ItemCount = 0

    LlenarComboMateriales

    Me.GroupBox1.Enabled = Not vSoloVer
    Me.GroupBox2.Enabled = Not vSoloVer
    Me.cboSectores.Enabled = Not vSoloVer

    'Me.GroupBox4.Enabled = Not vSoloVer 'items
    'Me.GroupBox3.Enabled = Not vSoloVer 'entregas
    Me.GridEX.AllowEdit = Not vSoloVer
    GridEX.AllowDelete = Not vSoloVer
    Me.grilla_entregas_materiales.AllowEdit = Not vSoloVer
    grilla_entregas_materiales.AllowDelete = Not vSoloVer

    Me.grilla_entregas_materiales.ItemCount = 0

    Channel.AgregarSuscriptor Me, Materiales_

    If IsSomething(vReque) Then
        Set requeVerificador = DAORequerimiento.FindById(vReque.Id, True, True, True, True)

        If requeVerificador.Guardado <> vReque.Guardado Then
            MsgBox "El requerimiento ya fue guardado y editado por alguien, se actualizará!"
            Set vReque = requeVerificador
            'actualizo la collection

            Set requeVerificador = Nothing
        End If
        cargarDatosReque
    Else
        Set vReque = New clsRequerimiento
        Set vDetalle = Nothing
    End If
End Sub
Private Sub cargarDatosReque()
    Me.caption = "Requerimiento Nº " & vReque.Id
    'Me.lblFecha = vReque.FechaCreado
    Me.cboSectores.ListIndex = funciones.PosIndexCbo(vReque.Sector.Id, Me.cboSectores)
    Dim Pedido As Long
    Pedido = vReque.DestinoOT
    If vReque.Tipo = ot_ Then
        Me.OpOT.value = True
        Me.lblOT = Format(vReque.DestinoOT, "0000")
    Else
        Me.OpStock.value = True
        Me.lblOT = Empty
    End If
    Me.GridEX.ItemCount = 0
    Me.GridEX.ItemCount = vReque.Materiales.count
    GridEXHelper.AutoSizeColumns Me.GridEX
End Sub

Private Sub Form_Terminate()
    Set vReque = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set vReque = Nothing
    Channel.RemoverSuscripcionTotal Me
End Sub
Private Sub GridEX_AfterDelete()
    seleccionarItem
End Sub
Private Sub GridEX_BeforeDelete(ByVal Cancel As GridEX20.JSRetBoolean)
    If vDetalle.Id = 0 Or vReque.estado = EnEdición_ Then
        Cancel = (MsgBox("¿Está seguro de eliminar el item?", vbYesNo + vbQuestion) = vbNo)
    Else
        Cancel = True
        MsgBox "El item ya fue guardado, no se puede eliminar, anulelo.", vbExclamation
    End If
End Sub
Private Sub seleccionarItem()
    On Error Resume Next
    Dim A As Long
    A = Me.GridEX.rowIndex(Me.GridEX.row)
    Set vDetalle = vReque.Materiales.item(A)
    'Set vMaterial = vDetalle.Material
    'mostrarMaterial vMaterial
End Sub
Private Sub GridEX_ColumnHeaderClick(ByVal Column As GridEX20.JSColumn)
    ordenar_grilla Column, Me.GridEX
End Sub

Private Sub GridEX_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    seleccionarItem
    If Button = 2 And IsSomething(vDetalle) And IsSomething(vReque) Then
        Me.mnuAprobar.Enabled = (vDetalle.estado = EstadoRequeCompra.Finalizado_ And Permisos.ComprasRequesAprobaciones)
        Me.mnuProcesarProveedores.Enabled = (vDetalle.estado = Aprobado_ Or vDetalle.estado = EnProceso_) And Permisos.ComprasRequesProcesar
        Me.mnuFinProcesoProveedores.Enabled = (vDetalle.estado = EnProceso_) And Permisos.ComprasRequesProcesar
        Me.mnuCrearPO.Enabled = (vDetalle.estado = EstadoRequeCompra.Procesado_) And Permisos.ComprasPOCrear
        Me.mnuAnular.Enabled = (vDetalle.estado <> EstadoRequeCompra.Anulado) And Permisos.ComprasRequesAnular
        Me.PopupMenu Me.mnuAcciones
    End If
End Sub

Private Sub GridEX_RowFormat(RowBuffer As GridEX20.JSRowData)
    On Error GoTo E
    Set vDetalle = vReque.Materiales.item(RowBuffer.rowIndex)

    If vDetalle.estado = EstadoRequeCompra.Aprobado_ _
       Or vDetalle.estado = EstadoRequeCompra.EnProceso_ _
       Or vDetalle.estado = EstadoRequeCompra.Procesado_ _
       Or vDetalle.estado = EstadoRequeCompra.AprobadoParcial_ _
       Or vDetalle.estado = EstadoRequeCompra.EnProcesoParcial_ _
       Or vDetalle.estado = EstadoRequeCompra.ProcesadoParcial_ _
       Or vDetalle.estado = EstadoRequeCompra.EnEdición_ Then

        For Each tmpEntrega In vDetalle.Entregas
            If tmpEntrega.FEcha < Date Then RowBuffer.CellStyle(8) = "vencidos"
            If tmpEntrega.FEcha = Date Then RowBuffer.CellStyle(9) = "vencenhoy"
            If tmpEntrega.FEcha > Date And tmpEntrega.FEcha <= DateAdd("d", DAORequeMateriales.CANT_DIAS_AVISO_VENCIMIENTO, Date) Then RowBuffer.CellStyle(10) = "avencer"
        Next tmpEntrega

    End If

E:
End Sub

Private Sub GridEX_SelectionChange()
    seleccionarItem
    MostrarEntregasMateriales
End Sub
Private Sub GridEX_UnboundDelete(ByVal rowIndex As Long, ByVal Bookmark As Variant)
    vReque.Materiales.remove rowIndex
End Sub
Private Sub GridEX_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If vReque.Materiales.count > 0 Then
        Set vDetalle = vReque.Materiales.item(rowIndex)
        With vDetalle
            Values(1) = .Cantidad
            Values(2) = .observaciones
            Values(3) = .Material.codigo
            Values(4) = .Material.Grupo.rubros.rubro
            Values(5) = .Material.Grupo.Grupo
            Values(6) = "Material: " & .Material.descripcion & " | Medidas: " & funciones.JoinCollectionValues(.Material.Atributos, ", ")
            Values(7) = enums.enumEstadoRequeCompra(.estado)
        End With
    End If
End Sub

Private Sub GridEX_UnboundUpdate(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If rowIndex > 0 And vReque.Materiales.count > 0 Then
        Set vDetalle = vReque.Materiales.item(rowIndex)
        vDetalle.Cantidad = Val(Values(1))
        vDetalle.observaciones = Values(2)
    End If
End Sub

Private Sub grilla_entregas_materiales_BeforeDelete(ByVal Cancel As GridEX20.JSRetBoolean)
    Cancel = (MsgBox("¿Está seguro de eliminar la entrega?", vbYesNo + vbQuestion) = vbNo)
End Sub

Private Sub grilla_entregas_materiales_BeforeUpdate(ByVal Cancel As GridEX20.JSRetBoolean)
    Cancel = (Me.GridEX.SelectedItems.count = 0) Or Not IsNumeric(Me.grilla_entregas_materiales.value(1)) Or Not IsDate(Me.grilla_entregas_materiales.value(2))
End Sub

Private Sub grilla_entregas_materiales_RowFormat(RowBuffer As GridEX20.JSRowData)
    If Not vSoloVer Then
        If RowBuffer.value(2) < Date Then
            RowBuffer.RowStyle = "fecha"
        End If
    End If
End Sub

Private Sub grilla_entregas_materiales_UnboundAddNew(ByVal NewRowBookmark As GridEX20.JSRetVariant, ByVal Values As GridEX20.JSRowData)
    If Me.GridEX.SelectedItems.count > 0 Then
        Set tmpEntrega = New clsRequeEntregas
        tmpEntrega.Tipo = material_
        tmpEntrega.Cantidad = Values(1)
        tmpEntrega.FEcha = Values(2)
        vDetalle.Entregas.Add tmpEntrega
    End If
End Sub

Private Sub grilla_entregas_materiales_UnboundDelete(ByVal rowIndex As Long, ByVal Bookmark As Variant)
    On Error Resume Next
    vDetalle.Entregas.remove (rowIndex)
End Sub

Private Sub grilla_entregas_materiales_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    On Error Resume Next
    Set tmpEntrega = vDetalle.Entregas.item(rowIndex)
    With tmpEntrega
        Values(1) = .Cantidad
        Values(2) = .FEcha
    End With

End Sub

Private Sub grilla_entregas_materiales_UnboundUpdate(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Set tmpEntrega = vDetalle.Entregas.item(rowIndex)
    tmpEntrega.Cantidad = Values(1)
    tmpEntrega.FEcha = Values(2)
End Sub


Private Property Get ISuscriber_id() As String
    Static Id As String
    If LenB(Id) = 0 Then Id = funciones.CreateGUID()
    ISuscriber_id = Id
End Property

Private Function ISuscriber_Notificarse(EVENTO As clsEventoObserver) As Variant
    LlenarComboMateriales
End Function

Private Sub lblOt_Click()
    Me.OpOT.value = True
End Sub
Private Sub lblOt_Validate(Cancel As Boolean)
    If IsNumeric(Me.lblOT) Then
        Cancel = (DAOOrdenTrabajo.FindById(Me.lblOT.Text) Is Nothing)
        If Cancel Then
            MsgBox "La OT no existe", vbExclamation
        End If
    Else
        Cancel = (LenB(Trim(Me.lblOT.Text)) <> 0)
    End If
End Sub

Private Sub NotificarCambio()
    Dim ev As New clsEventoObserver
    Set ev.Elemento = vReque
    ev.EVENTO = modificar_
    ev.Tipo = RequerimientosCompra_

    Channel.Notificar ev, RequerimientosCompra_


End Sub




Private Sub mnuAnular_Click()
    If IsSomething(vReque) Then
        If MsgBox("¿Desea anular el item?", vbQuestion + vbYesNo) = vbYes Then
            If DAORequeMateriales.Anular(vDetalle, vReque) Then
                NotificarCambio
                Me.GridEX.ReBind
                Me.GridEX.Refresh
                MsgBox "Item Anulado.", vbInformation
            Else
                MsgBox "Error al anular el item.", vbCritical
            End If
        End If
    End If

End Sub

Private Sub mnuAprobar_Click()

    If IsSomething(vReque) Then
        If MsgBox("¿Desea aprobar el item?", vbQuestion + vbYesNo) = vbYes Then
            If DAORequeMateriales.aprobar(vDetalle, vReque) Then
                NotificarCambio
                Me.GridEX.ReBind
                Me.GridEX.Refresh
                MsgBox "Item Aprobado.", vbInformation
            Else
                MsgBox "Error al aprobar el item.", vbCritical
            End If
        End If
    End If

End Sub

Private Sub mnuCrearPO_Click()
    If IsSomething(vReque) Then
        If MsgBox("¿Desea finalizar el proceso para el item?", vbQuestion + vbYesNo) = vbYes Then
            If DAORequeMateriales.finalizarProcesoProveedores(vDetalle, vReque) Then
                NotificarCambio
                Me.GridEX.ReBind
                Me.GridEX.Refresh
                MsgBox "Proceso finalizado para el item.", vbInformation
            Else
                MsgBox "Error al finalizar el proceso del item.", vbCritical
            End If
        End If
    End If
End Sub

Private Sub mnuFinProcesoProveedores_Click()
    If IsSomething(vReque) Then
        If MsgBox("¿Desea finalizar el proceso para el item?", vbQuestion + vbYesNo) = vbYes Then
            If DAORequeMateriales.finalizarProcesoProveedores(vDetalle, vReque) Then
                NotificarCambio
                Me.GridEX.ReBind
                Me.GridEX.Refresh
                MsgBox "Proceso finalizado para el item.", vbInformation
            Else
                MsgBox "Error al finalizar el proceso del item.", vbCritical
            End If
        End If
    End If
End Sub

Private Sub mnuProcesarProveedores_Click()
    If IsSomething(vReque) Then
        Dim ret As VbMsgBoxResult
        Dim procesoPor1raVez As Boolean
        ret = vbYes
        If vDetalle.estado <> EnProceso_ Then
            ret = MsgBox("¿Desea procesar el item?", vbQuestion + vbYesNo)
            procesoPor1raVez = True
        End If
        If ret = vbNo Then Exit Sub

        If DAORequeMateriales.procesarProveedores(vDetalle, vReque) Then
            NotificarCambio
            Me.GridEX.ReBind
            Me.GridEX.Refresh

            If procesoPor1raVez Then MsgBox "Item listo para procesar proveedores.", vbInformation

            Dim F As New frmComprasRequesProcesar
            F.reque = vReque
            F.Show
        Else
            MsgBox "Error al procesar el item.", vbCritical
        End If
    End If

End Sub

Private Sub OpStock_Click()
    lblOT.Text = vbNullString
End Sub

Private Sub PushButton1_Click()
    On Error GoTo err1
    Dim frm As New frmMaterialesLista2_modal
    frm.Usable = True
    Set Selecciones.Material = Nothing
    frm.Show 1
    If IsSomething(Selecciones.Material) Then
        Set vmaterial = Selecciones.Material
        Set Selecciones.Material = Nothing
        Me.cboMateriales.ListIndex = PosIndexCbo(vmaterial.Id, Me.cboMateriales)
    End If
    Exit Sub
err1:
    Set Selecciones.Material = Nothing
End Sub




Private Function buscarMaterial() As Boolean
    If Me.cboMateriales.ListIndex = -1 Then
        Set vmaterial = Nothing
        Exit Function
    End If

    buscarMaterial = True

    Set vmaterial = DAOMateriales.FindById(Me.cboMateriales.ItemData(Me.cboMateriales.ListIndex))


    If IsSomething(vmaterial) Then
        MostrarMaterial vmaterial
        Me.lblUnidad = enums.enumUnidades(vmaterial.unidad)
        buscarMaterial = True
    Else
        limpiar
        buscarMaterial = False
        MsgBox "No se encontro material con ese codigo", vbExclamation
    End If
End Function
Private Sub limpiar()
    Me.lblDescripcion = Empty
    Me.lblGrupo = Empty
    Me.lblKgM2 = Empty
    Me.lblRubro = Empty
    Me.lblEspesor = Empty
    Me.lblUnidad = Empty
End Sub


Private Sub MostrarMaterial(vmat As clsMaterial)
    Me.lblDescripcion = vmat.descripcion
    Me.lblRubro = vmat.Grupo.rubros.rubro
    Me.lblEspesor = vmat.Espesor
    Me.lblGrupo = vmat.Grupo.Grupo
    Me.lblKgM2 = vmat.PesoXUnidad
    Me.lblUnidad = enums.enumUnidades(vmat.unidad)
End Sub
Private Sub MostrarEntregasMateriales()
    On Error Resume Next
    'Dim a As Long
    'a = Me.GridEX.RowIndex(Me.GridEX.row)
    Me.grilla_entregas_materiales.ItemCount = 0
    Me.grilla_entregas_materiales.ItemCount = vDetalle.Entregas.count
    Me.grilla_entregas_materiales.ReBind
End Sub


