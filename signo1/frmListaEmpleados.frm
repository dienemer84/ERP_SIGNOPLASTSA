VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmListaEmpleados 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Empleados"
   ClientHeight    =   6330
   ClientLeft      =   60
   ClientTop       =   270
   ClientWidth     =   14865
   ClipControls    =   0   'False
   Icon            =   "frmListaEmpleados.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6330
   ScaleWidth      =   14865
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1020
      Left            =   90
      TabIndex        =   2
      Top             =   45
      Width           =   6000
      _Version        =   786432
      _ExtentX        =   10583
      _ExtentY        =   1799
      _StockProps     =   79
      Caption         =   "Filtro"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   330
         Left            =   4725
         TabIndex        =   10
         Top             =   180
         Width           =   1125
         _Version        =   786432
         _ExtentX        =   1984
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "Etiquetas"
         UseVisualStyle  =   -1  'True
      End
      Begin VB.TextBox txtApellido 
         Height          =   285
         Left            =   870
         TabIndex        =   7
         Top             =   585
         Width           =   3015
      End
      Begin VB.TextBox txtNroLeg 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2775
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtNroDoc 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   870
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin XtremeSuiteControls.PushButton cmdBuscar 
         Default         =   -1  'True
         Height          =   345
         Left            =   4710
         TabIndex        =   5
         Top             =   555
         Width           =   1170
         _Version        =   786432
         _ExtentX        =   2064
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Buscar"
         UseVisualStyle  =   -1  'True
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Apellido"
         Height          =   195
         Left            =   210
         TabIndex        =   9
         Top             =   615
         Width           =   615
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nº Leg"
         Height          =   195
         Left            =   2160
         TabIndex        =   8
         Top             =   270
         Width           =   615
      End
      Begin VB.Label lblDocumento 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nº Doc"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   270
         Width           =   615
      End
   End
   Begin GridEX20.GridEX grilla 
      Height          =   5130
      Left            =   15
      TabIndex        =   1
      Top             =   1170
      Width           =   13350
      _ExtentX        =   23548
      _ExtentY        =   9049
      Version         =   "2.0"
      PreviewRowIndent=   200
      AutomaticSort   =   -1  'True
      HoldSortSettings=   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      PreviewColumn   =   "domicilio"
      PreviewRowLines =   1
      ColumnAutoResize=   -1  'True
      MethodHoldFields=   -1  'True
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      ItemCount       =   1
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   10
      Column(1)       =   "frmListaEmpleados.frx":000C
      Column(2)       =   "frmListaEmpleados.frx":0104
      Column(3)       =   "frmListaEmpleados.frx":021C
      Column(4)       =   "frmListaEmpleados.frx":02EC
      Column(5)       =   "frmListaEmpleados.frx":03C0
      Column(6)       =   "frmListaEmpleados.frx":049C
      Column(7)       =   "frmListaEmpleados.frx":057C
      Column(8)       =   "frmListaEmpleados.frx":06D0
      Column(9)       =   "frmListaEmpleados.frx":082C
      Column(10)      =   "frmListaEmpleados.frx":0944
      SortKeysCount   =   1
      SortKey(1)      =   "frmListaEmpleados.frx":0A38
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmListaEmpleados.frx":0AA0
      FormatStyle(2)  =   "frmListaEmpleados.frx":0BD8
      FormatStyle(3)  =   "frmListaEmpleados.frx":0C88
      FormatStyle(4)  =   "frmListaEmpleados.frx":0D3C
      FormatStyle(5)  =   "frmListaEmpleados.frx":0E14
      FormatStyle(6)  =   "frmListaEmpleados.frx":0ECC
      ImageCount      =   0
      PrinterProperties=   "frmListaEmpleados.frx":0FAC
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   12645
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6600
      Width           =   975
   End
   Begin VB.Image imgFoto 
      Height          =   975
      Left            =   6240
      Stretch         =   -1  'True
      Top             =   120
      Width           =   975
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Main"
      Visible         =   0   'False
      Begin VB.Menu mnuEditar 
         Caption         =   "Editar..."
      End
      Begin VB.Menu mnuTareas 
         Caption         =   "Tareas..."
      End
      Begin VB.Menu mnuSectorizar 
         Caption         =   "Sectorizar..."
      End
      Begin VB.Menu mnuArchivosEmpleado 
         Caption         =   "Archivos de Empleado"
      End
      Begin VB.Menu mnuNuevoSiniestro 
         Caption         =   "Nuevo siniestro para empleado"
      End
   End
End
Attribute VB_Name = "frmListaEmpleados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim empleados As Collection
Dim Foto As archivo
Dim tmppath As String
Dim emple As clsEmpleado
'Dim rec As Recordset
Dim CantArchivos As Dictionary
Dim clasea As New classArchivos
Dim CantSiniestros As Dictionary


Private Sub cmdBuscar_Click()
    llenarLista
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub llenarLista()
    Dim filter As String
    filter = " 1 = 1"

    If LenB(Me.txtNroDoc.text) > 0 And IsNumeric(Me.txtNroDoc) Then filter = filter & " AND documento = " & Me.txtNroDoc.text
    If LenB(Me.txtApellido.text) > 0 Then filter = filter & " AND apellido like '%" & Me.txtApellido.text & "%'"
    If LenB(Me.txtNroLeg.text) > 0 And IsNumeric(Me.txtNroLeg) Then filter = filter & " AND legajo = " & Me.txtNroLeg.text


    Set empleados = DAOEmpleados.GetAll(filter)

    Set CantArchivos = DAOArchivo.GetCantidadArchivosPorReferencia(OA_Empleados)
    Set CantSiniestros = DAOSiniestroPersonal.GetCantidadSiniestrosPorEmpleado()

    Me.grilla.ItemCount = 0
    Me.grilla.ItemCount = empleados.count

    Me.grilla.Refresh
    grilla_SelectionChange
End Sub

Private Sub Form_Activate()
    GridEXHelper.AutoSizeColumns Me.grilla
End Sub

Private Sub Form_Initialize()
    llenarLista
End Sub

Private Sub Form_Load()
    FormHelper.Customize Me
    GridEXHelper.CustomizeGrid Me.grilla, False
    Me.grilla.ItemCount = 0

End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Me.grilla.Width = Me.ScaleWidth
    Me.grilla.Height = (Me.ScaleHeight - Me.grilla.Top)
    GridEXHelper.AutoSizeColumns Me.grilla
End Sub

Private Sub grilla_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    On Error Resume Next
    Set emple = empleados.item(Me.grilla.rowIndex(Me.grilla.row))
    ShowImage
    If Button = 2 And Not emple Is Nothing Then
        Me.mnuNuevoSiniestro.Enabled = Permisos.RRHHSiniestros
        Me.mnuEditar.caption = "Editar a " & emple.NombreCompleto
        Me.PopupMenu Me.mnuMain
    End If
End Sub

Private Sub grilla_SelectionChange()
    On Error Resume Next
    Set emple = empleados.item(Me.grilla.rowIndex(Me.grilla.row))
    ShowImage
End Sub
Private Sub ShowImage()
    On Error GoTo err1
    If IsSomething(emple) Then
        Set Foto = DAOArchivo.FindAll(OA_FotoEmpleado, "idPieza=" & emple.Id)(1)
        If IsSomething(Foto) Then
            tmppath = clasea.exportarArchivo(Foto.Id)
            If LenB(tmppath) > 0 Then
                Set Me.imgFoto.Picture = LoadPicture(tmppath)
                Kill tmppath
            End If
        Else
            Set imgFoto.Picture = Nothing
        End If
    Else
        Set imgFoto.Picture = Nothing
    End If
    Exit Sub
err1:
    Set imgFoto.Picture = Nothing
End Sub
Private Sub grilla_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Dim emple2 As clsEmpleado
    Set emple2 = empleados.item(rowIndex)

    With Values
        .value(1) = emple2.Apellido & " " & emple2.nombre & " " & emple2.Nombres
        .value(2) = emple2.legajo
        .value(3) = emple2.documento
        .value(4) = emple2.FechaNacimiento
        .value(5) = emple2.FechaIngreso
        .value(6) = emple2.GrupoSanguineo
        .value(7) = Val(CantArchivos(emple2.Id))
        .value(8) = Val(CantSiniestros(emple2.Id))
        If LenB(emple2.DireccionCompleta) > 0 Then
            .value(9) = "Dirección: " & emple2.DireccionCompleta
        End If
        If IsSomething(emple2.ObraSocial) Then
            .value(10) = emple2.ObraSocial.nombre
        End If

    End With
End Sub


Private Sub mnuArchivosEmpleado_Click()
    If Not emple Is Nothing Then
        Dim frmArchi As New frmArchivos2
        frmArchi.Origen = OrigenArchivos.OA_Empleados
        frmArchi.ObjetoId = emple.Id
        frmArchi.caption = "Empleado - " & emple.NombreCompleto
        frmArchi.Show
    End If
End Sub

Private Sub mnuEditar_Click()


    If Not emple Is Nothing Then
        Dim F As New frmAltaEmpleados
        Load F
        ' grilla_SelectionChange
        Set F.Empleado = emple
        F.Show
    End If
End Sub

Private Sub mnuNuevoSiniestro_Click()
    If Not emple Is Nothing Then
        Dim ffff As New frmSiniestro
        ffff.Show
        ffff.cboAsegurado.ListIndex = funciones.PosIndexCbo(emple.Id, ffff.cboAsegurado)
    End If
End Sub

Private Sub mnuSectorizar_Click()
    If Not emple Is Nothing Then
        Dim F As New frmSectorizar
        Load F
        F.Text1.text = emple.legajo
        F.Command1_Click
        F.Show
    End If
End Sub

Private Sub mnuTareas_Click()
    If Not emple Is Nothing Then
        Dim F As New frmEmpleadosTareas
        Load F
        F.personalId = emple.Id
        F.Show
    End If
End Sub

Private Sub PushButton1_Click()
    If MsgBox("Esta seguro?", vbYesNo) = vbYes Then LabelHelper.PrintEtiquetaLegajos
End Sub
