VERSION 5.00
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#12.0#0"; "CODEJO~3.OCX"
Begin VB.Form frmCategoriasSueldo 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Categorias de sueldo"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11115
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCategoriasSueldo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   11115
   Begin XtremeReportControl.ReportControl ReportControl 
      Height          =   3450
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   5850
      _Version        =   786432
      _ExtentX        =   10319
      _ExtentY        =   6085
      _StockProps     =   64
      BorderStyle     =   3
   End
   Begin XtremeReportControl.ReportControl ReportHistorico 
      Height          =   5175
      Left            =   6240
      TabIndex        =   15
      Top             =   480
      Width           =   4695
      _Version        =   786432
      _ExtentX        =   8281
      _ExtentY        =   9128
      _StockProps     =   64
   End
   Begin VB.CommandButton cmdActualizarValores 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Actualizar valores"
      Height          =   345
      Left            =   105
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5400
      Width           =   5955
   End
   Begin VB.CommandButton cmdCancelar 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   345
      Left            =   4920
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   345
      Left            =   2520
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton cmdGuardar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Guardar"
      Height          =   345
      Left            =   3720
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4920
      Width           =   1100
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   345
      Left            =   1320
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4920
      Width           =   1100
   End
   Begin VB.CommandButton cmdNuevo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Nuevo"
      Height          =   345
      Left            =   120
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4920
      Width           =   1100
   End
   Begin VB.Frame fraDatos 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Datos de la categoria"
      Height          =   1125
      Left            =   90
      TabIndex        =   1
      Top             =   3630
      Width           =   5865
      Begin VB.TextBox txtEspecializacion 
         Height          =   285
         Left            =   3345
         TabIndex        =   7
         Text            =   "0"
         Top             =   675
         Width           =   690
      End
      Begin VB.TextBox txtValor 
         Height          =   285
         Left            =   840
         TabIndex        =   6
         Top             =   660
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
         Height          =   285
         Left            =   840
         TabIndex        =   5
         Top             =   270
         Width           =   4875
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "% Especializacion"
         Height          =   195
         Left            =   1995
         TabIndex        =   4
         Top             =   705
         Width           =   1260
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "Valor"
         Height          =   195
         Left            =   390
         TabIndex        =   3
         Top             =   690
         Width           =   360
      End
      Begin VB.Label lblNombre 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "Nombre"
         Height          =   195
         Left            =   180
         TabIndex        =   2
         Top             =   300
         Width           =   555
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Historial de valores de la categoría"
      Height          =   300
      Left            =   6255
      TabIndex        =   14
      Top             =   150
      Width           =   4680
   End
End
Attribute VB_Name = "frmCategoriasSueldo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private registro As Long
Private categorias As Collection
Private categoria As CategoriaSueldo
Private idle As Boolean

Private Sub ActivarControles()
    Me.ReportControl.Enabled = idle
    Me.fraDatos.Enabled = Not idle


    Me.cmdEliminar.Enabled = Not categoria Is Nothing And idle
    Me.cmdGuardar.Enabled = Not idle
    Me.cmdModificar.Enabled = Not categoria Is Nothing And idle
    Me.cmdNuevo.Enabled = idle
    Me.cmdActualizarValores.Enabled = idle
End Sub

Private Sub cmdActualizarValores_Click()
    On Error GoTo err1:
    Dim Sueldo As Double

    Sueldo = val(InputBox("Ingrese incremental", "Actualización Salarial Global", 0))
    Dim cat As CategoriaSueldo

    Dim categorias As Collection
    conectar.BeginTransaction
    Set categorias = DAOCategoriaSueldo.FindAll()
    Dim porc As Double
    porc = 1 + (Sueldo / 100)
    For Each cat In categorias
        cat.valor = funciones.RedondearDecimales(cat.valor * porc, 2)
        If Not DAOCategoriaSueldo.Save(cat) Then GoTo err1
    Next
    conectar.CommitTransaction
    MsgBox "Actualización correcta de categorias", vbInformation, "Información"
    Exit Sub
err1:
    MsgBox "Se produjo un error al actualizar los valores. Por favor, contacar al dto. de Sistemas", vbCritical, "Error"
    conectar.RollBackTransaction
End Sub

Private Sub cmdCancelar_Click()
    If idle Then
        Unload Me
        Exit Sub
    End If

    idle = True
    LimpiarControles
    Me.ReportControl.SelectedRows.DeleteAll    'deselecciono todas las seleccionadas

    If Me.ReportControl.rows.count > 0 Then
        Me.ReportControl.SelectedRows.Add Me.ReportControl.rows(0)
        'no se como seleccionar un item especifico y que se dispare el evento
    End If

    ActivarControles
End Sub

Private Sub cmdEliminar_Click()
    If Not categoria Is Nothing Then
        If MsgBox("¿Desea eliminar la categoria?", vbQuestion + vbYesNo) = vbYes Then
            If DAOCategoriaSueldo.Delete(categoria) Then
                CargarLista
                LimpiarControles
                ActivarControles
            Else
                MsgBox "Hubo un error al borrar.", vbCritical
            End If
        End If
    End If
End Sub

Private Sub cmdGuardar_Click()
    On Error GoTo E

    If categoria Is Nothing Then Set categoria = New CategoriaSueldo

    categoria.nombre = Me.txtNombre.Text
    categoria.PorcentajeEspecializacion = CDbl(Me.txtEspecializacion.Text)
    categoria.valor = CDbl(Me.txtValor.Text)

    If DAOCategoriaSueldo.Save(categoria) Then
        Dim EVENTO As New clsEventoObserver
        Set EVENTO.Elemento = categoria
        EVENTO.EVENTO = agregar_     'es lo mismo cualquiera
        Set EVENTO.Originador = Me

        idle = True
        CargarLista
        LimpiarControles
        ActivarControles
        Channel.Notificar EVENTO, Tareas_
    Else
        MsgBox "Hubo un error al guardar.", vbCritical
    End If

    Exit Sub
E:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub cmdModificar_Click()
    idle = False
    ActivarControles

    Me.txtNombre.SetFocus
End Sub

Private Sub cmdNuevo_Click()
    idle = False
    ActivarControles
    LimpiarControles
    Me.txtEspecializacion = 0
    Me.txtNombre.SetFocus
End Sub

Private Sub LimpiarControles()
    Me.txtEspecializacion.Text = vbNullString
    Me.txtNombre.Text = vbNullString
    Me.txtValor.Text = vbNullString

    Set categoria = Nothing
End Sub

Private Sub Form_Load()
    FormHelper.Customize Me
    Dim Column As ReportColumn

    Set Column = Me.ReportControl.Columns.Add(0, "Nombre", 10, True)
    Column.Icon = 0
    Column.Sortable = True
    Column.AllowDrag = False
    Column.AllowRemove = False
    Column.Alignment = xtpAlignmentIconLeft

    Set Column = Me.ReportControl.Columns.Add(1, "Valor", 10, True)
    Column.Icon = 0
    Column.Sortable = True
    Column.AllowDrag = False
    Column.AllowRemove = False
    Column.Alignment = xtpAlignmentRight

    Set Column = Me.ReportControl.Columns.Add(2, "% Especializacion", 10, True)
    Column.Icon = 0
    Column.Sortable = True
    Column.AllowDrag = False
    Column.AllowRemove = False
    Column.Alignment = xtpAlignmentRight

    Me.ReportControl.PaintManager.HorizontalGridStyle = xtpGridSmallDots
    Me.ReportControl.PaintManager.VerticalGridStyle = xtpGridSmallDots

    
    '----------------------------------------
    ' Columnas del historial de sueldos
    '----------------------------------------
    
    Set Column = Me.ReportHistorico.Columns.Add(0, "Fecha", 100, True)
    Column.Sortable = True
    Column.AllowDrag = False
    Column.AllowRemove = False
    
    Set Column = Me.ReportHistorico.Columns.Add(1, "Valor", 100, True)
    Column.Sortable = True
    Column.AllowDrag = False
    Column.AllowRemove = False
    Column.Alignment = xtpAlignmentRight
    
    Set Column = Me.ReportHistorico.Columns.Add(2, "Usuario", 100, True)
    Column.Sortable = True
    Column.AllowDrag = False
    Column.AllowRemove = False
    
    Me.ReportHistorico.PaintManager.HorizontalGridStyle = xtpGridSmallDots
    Me.ReportHistorico.PaintManager.VerticalGridStyle = xtpGridSmallDots

    CargarLista

    idle = True
    ActivarControles
End Sub

Private Sub CargarLista()
    Me.ReportControl.Records.DeleteAll
    Dim rec As ReportRecord

    Set categorias = DAOCategoriaSueldo.FindAll()
    For Each categoria In categorias
        Set rec = Me.ReportControl.Records.Add
        rec.AddItem categoria.nombre
        rec.AddItem funciones.FormatearDecimales(categoria.valor)
        rec.AddItem categoria.PorcentajeEspecializacion
        rec.Tag = categoria.Id
    Next categoria

    Set categoria = Nothing

    Me.ReportControl.Populate
    Me.ReportControl.FocusedRow = Me.ReportControl.rows(registro)
End Sub

Private Sub ReportControl_SelectionChanged()

    If Me.ReportControl.SelectedRows.count = 0 Then

        Set categoria = Nothing

        LimpiarControles

        Me.ReportHistorico.Records.DeleteAll
        Me.ReportHistorico.Populate

        ActivarControles

        Exit Sub

    End If

    Set categoria = categorias.item( _
        CStr(Me.ReportControl.SelectedRows(0).record.Tag))

    If Not categoria Is Nothing Then

        Me.txtEspecializacion.Text = categoria.PorcentajeEspecializacion
        Me.txtValor.Text = categoria.valor
        Me.txtNombre.Text = categoria.nombre

        'Cargar historial de la categoría seleccionada
        CargarHistorico categoria.Id

    Else

        LimpiarControles

        Me.ReportHistorico.Records.DeleteAll
        Me.ReportHistorico.Populate

    End If

    registro = Me.ReportControl.SelectedRows(0).Index

    ActivarControles

End Sub

Private Sub txtEspecializacion_Validate(Cancel As Boolean)
    funciones.ValidarTextBox Me.txtEspecializacion, Cancel
End Sub

Private Sub CargarHistorico(ByVal idCategoriaSueldo As Long)

    On Error GoTo E

    Dim rs As ADODB.Recordset
    Dim rec As ReportRecord

    Me.ReportHistorico.Records.DeleteAll

    If idCategoriaSueldo <= 0 Then
        Me.ReportHistorico.Populate
        Exit Sub
    End If

    Set rs = DAOCategoriaSueldo.FindHistorico(idCategoriaSueldo)

    While Not rs.EOF

        Set rec = Me.ReportHistorico.Records.Add

        'Fecha
        If IsNull(rs!FEcha) Then
            rec.AddItem vbNullString
        Else
            rec.AddItem Format$(rs!FEcha, "dd/mm/yyyy")
        End If

        'Valor
        If IsNull(rs!valor) Then
            rec.AddItem vbNullString
        Else
            rec.AddItem funciones.FormatearDecimales(CDbl(rs!valor))
        End If

        'Usuario
        If IsNull(rs!Usuario) Then
            rec.AddItem vbNullString
        Else
            rec.AddItem CStr(rs!Usuario)
        End If

        rs.MoveNext
    Wend

    Me.ReportHistorico.Populate

    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
    End If

    Set rs = Nothing

    Exit Sub

E:
    MsgBox "Error al cargar el historial de la categoría." & vbCrLf & _
           Err.Description, vbCritical, "Historial de sueldos"

End Sub
