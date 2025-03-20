VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmMaterialesNuevo 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Material"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7740
   ClipControls    =   0   'False
   Icon            =   "frmNuevoMaterial.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   7740
   Begin VB.TextBox txtPtoReposicion 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3915
      TabIndex        =   16
      Text            =   "0"
      Top             =   4875
      Width           =   885
   End
   Begin VB.TextBox txtStockMinimo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1515
      TabIndex        =   15
      Text            =   "0"
      Top             =   4860
      Width           =   885
   End
   Begin VB.TextBox txtAltura 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   6675
      TabIndex        =   8
      Text            =   "0"
      Top             =   2610
      Width           =   885
   End
   Begin XtremeSuiteControls.ComboBox cboRubros 
      Height          =   315
      Left            =   1545
      TabIndex        =   0
      Top             =   765
      Width           =   6015
      _Version        =   786432
      _ExtentX        =   10610
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.PushButton Command1 
      Height          =   375
      Left            =   6300
      TabIndex        =   17
      Top             =   4785
      Width           =   1245
      _Version        =   786432
      _ExtentX        =   2196
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Guardar"
      BackColor       =   16744576
      UseVisualStyle  =   -1  'True
   End
   Begin VB.TextBox txtDescripcion 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1545
      TabIndex        =   2
      Top             =   1500
      Width           =   6015
   End
   Begin VB.TextBox txtEspesor 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   5070
      TabIndex        =   7
      Text            =   "0"
      Top             =   2610
      Width           =   885
   End
   Begin VB.TextBox txtKgXM2Ml 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1530
      TabIndex        =   9
      Top             =   2955
      Width           =   900
   End
   Begin VB.TextBox txtValor 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3780
      TabIndex        =   10
      Top             =   2970
      Width           =   855
   End
   Begin VB.TextBox txtFecha 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1515
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   4455
      Width           =   1380
   End
   Begin VB.TextBox txtCantidad 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3900
      TabIndex        =   14
      Top             =   4470
      Width           =   1335
   End
   Begin VB.TextBox txtLargo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1545
      TabIndex        =   5
      Text            =   "0"
      Top             =   2610
      Width           =   885
   End
   Begin VB.TextBox txtAncho 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3300
      TabIndex        =   6
      Text            =   "0"
      Top             =   2610
      Width           =   885
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ Altas ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6135
      Left            =   8640
      TabIndex        =   18
      Top             =   4680
      Width           =   7455
   End
   Begin XtremeSuiteControls.ComboBox cboGrupos 
      Height          =   315
      Left            =   1545
      TabIndex        =   1
      Top             =   1125
      Width           =   6015
      _Version        =   786432
      _ExtentX        =   10610
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboMonedas 
      Height          =   315
      Left            =   5595
      TabIndex        =   11
      Top             =   2955
      Width           =   1980
      _Version        =   786432
      _ExtentX        =   3492
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboAlmacenes 
      Height          =   315
      Left            =   1515
      TabIndex        =   12
      Top             =   4080
      Width           =   6015
      _Version        =   786432
      _ExtentX        =   10610
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Appearance      =   5
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboUnidades 
      Height          =   315
      Left            =   1545
      TabIndex        =   4
      Top             =   2235
      Width           =   6015
      _Version        =   786432
      _ExtentX        =   10610
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Style           =   2
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboTipoMaterial 
      Height          =   315
      Left            =   1545
      TabIndex        =   3
      Top             =   1845
      Width           =   6015
      _Version        =   786432
      _ExtentX        =   10610
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Style           =   2
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboUnidadPedido 
      Height          =   315
      Left            =   1530
      TabIndex        =   40
      Top             =   3330
      Width           =   6015
      _Version        =   786432
      _ExtentX        =   10610
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Style           =   2
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboUnidadCompra 
      Height          =   315
      Left            =   1530
      TabIndex        =   42
      Top             =   3705
      Width           =   6015
      _Version        =   786432
      _ExtentX        =   10610
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Style           =   2
      Text            =   "ComboBox1"
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "Unidad Compra"
      Height          =   195
      Left            =   390
      TabIndex        =   43
      Top             =   3750
      Width           =   1095
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "Unidad Pedido"
      Height          =   195
      Left            =   450
      TabIndex        =   41
      Top             =   3375
      Width           =   1050
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "Pto Reposición"
      Height          =   195
      Left            =   2790
      TabIndex        =   39
      Top             =   4905
      Width           =   1080
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "Stock Minimo"
      Height          =   195
      Left            =   525
      TabIndex        =   38
      Top             =   4920
      Width           =   960
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "Código Nuevo"
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
      Left            =   270
      TabIndex        =   37
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblCodigoNuevo 
      AutoSize        =   -1  'True
      Caption         =   "lblCodigo"
      Height          =   195
      Left            =   1545
      TabIndex        =   36
      Top             =   120
      Width           =   645
   End
   Begin VB.Label lblCodigoAnterior 
      AutoSize        =   -1  'True
      Caption         =   "lblCodigo"
      Height          =   195
      Left            =   1545
      TabIndex        =   35
      Top             =   435
      Width           =   645
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "Altura"
      Height          =   195
      Left            =   6195
      TabIndex        =   34
      Top             =   2655
      Width           =   405
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "Tipo"
      Height          =   195
      Left            =   1155
      TabIndex        =   33
      Top             =   1875
      Width           =   315
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "Un Trabajo/Costos"
      Height          =   195
      Left            =   165
      TabIndex        =   32
      Top             =   2280
      Width           =   1350
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "Kg x M2/Ml "
      Height          =   195
      Left            =   630
      TabIndex        =   31
      Top             =   3000
      Width           =   870
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "Espesor "
      Height          =   195
      Left            =   4440
      TabIndex        =   30
      Top             =   2655
      Width           =   615
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "Descripción "
      Height          =   195
      Left            =   630
      TabIndex        =   29
      Top             =   1515
      Width           =   885
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "Grupo "
      Height          =   195
      Left            =   1050
      TabIndex        =   28
      Top             =   1170
      Width           =   480
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "Rubro "
      Height          =   195
      Left            =   1035
      TabIndex        =   27
      Top             =   810
      Width           =   480
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "Código Anterior"
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
      Left            =   165
      TabIndex        =   26
      Top             =   420
      Width           =   1320
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      Caption         =   "Valor Unitario "
      Height          =   195
      Left            =   2760
      TabIndex        =   25
      Top             =   3000
      Width           =   990
   End
   Begin VB.Label FEcha 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "Fecha "
      Height          =   195
      Left            =   1005
      TabIndex        =   24
      Top             =   4470
      Width           =   495
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      Caption         =   "Moneda "
      Height          =   195
      Left            =   4920
      TabIndex        =   23
      Top             =   3015
      Width           =   630
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "Almacen "
      Height          =   195
      Left            =   825
      TabIndex        =   22
      Top             =   4110
      Width           =   660
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      Caption         =   "Cantidad "
      Height          =   195
      Left            =   3180
      TabIndex        =   21
      Top             =   4500
      Width           =   675
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "Largo "
      Height          =   195
      Left            =   1065
      TabIndex        =   20
      Top             =   2640
      Width           =   450
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "Ancho "
      Height          =   195
      Left            =   2760
      TabIndex        =   19
      Top             =   2655
      Width           =   510
   End
End
Attribute VB_Name = "frmMaterialesNuevo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements ISuscriber
Dim rubroElegido As clsRubros
Dim vmaterial As clsMaterial
Dim vId As String

'Dim base As New classNuevoMaterial
Dim claseSP As New classSignoplast
Dim idModificar As Long
Dim strsql As String
Dim clsConf As New classConfigurar

Dim codigoViejo As String
Dim valorViejo As Double


Public Property Let Material(nvalue As clsMaterial)
    Set vmaterial = nvalue
End Property

Private Sub cboGrupos_Click()
    BuildCodigoMaterial
End Sub


Private Sub cboRubros_Click()
    Set rubroElegido = DAORubros.FindById(cboRubros.ItemData(cboRubros.ListIndex))
    DAOGrupos.llenarComboXtremeSuite Me.cboGrupos, rubroElegido
    BuildCodigoMaterial
End Sub


Public Sub BuildCodigoMaterial()
    Dim cod As String

    If Me.cboRubros.ListIndex > -1 Then
        Dim r As clsRubros
        Set r = DAORubros.FindById(Me.cboRubros.ItemData(Me.cboRubros.ListIndex))
        cod = r.iniciales
    End If

    If Me.cboGrupos.ListIndex > -1 Then
        Dim g As clsGrupo
        Set g = DAOGrupos.GetById(Me.cboGrupos.ItemData(Me.cboGrupos.ListIndex))
        cod = cod & Format(g.Id, "000")
    End If

    If IsSomething(vmaterial) Then
        cod = cod & "-" & vmaterial.Id
    End If

    Me.lblCodigoNuevo.caption = cod
End Sub

Private Sub cboTipoMaterial_Click()
    If Me.cboTipoMaterial.ListIndex > -1 Then
        Dim cboUnidadIdx As Unidades: cboUnidadIdx = -1
        Dim tipoMat As TipoMaterial
        tipoMat = Me.cboTipoMaterial.ItemData(Me.cboTipoMaterial.ListIndex)

        Select Case tipoMat
        Case TipoMaterial.TM_PerfilEspecial, TipoMaterial.TM_PerfilCuadrado, TipoMaterial.TM_PerfilRectangular, TipoMaterial.TM_PerfilCuadrado, TipoMaterial.TM_PerfilTubo, TipoMaterial.TM_PerfilELE
            cboUnidadIdx = Unidades.Ml_
        Case TipoMaterial.TM_HojaPlancha
            cboUnidadIdx = Unidades.m2_
        Case TipoMaterial.TM_UnidadKilo
            cboUnidadIdx = Unidades.un_
        End Select

        Me.txtLargo.Enabled = Not (tipoMat = TM_UnidadKilo)
        Me.txtAncho.Enabled = (tipoMat = TM_HojaPlancha Or tipoMat = TM_PerfilRectangular Or tipoMat)
        Me.txtAltura.Enabled = (tipoMat = TM_PerfilRectangular)
        Me.txtEspesor.Enabled = Not (tipoMat = TM_UnidadKilo)

        If cboUnidadIdx <> -1 Then
            Me.cboUnidades.ListIndex = funciones.PosIndexCbo(cboUnidadIdx, Me.cboUnidades)
        Else
            Me.cboUnidades.ListIndex = -1
        End If
    Else
        Me.cboUnidades.ListIndex = -1
    End If

End Sub

Private Sub cboUnidades_Click()
    ValidarUnidades
End Sub
Private Function accion() As Boolean
    On Error GoTo err2
    accion = True
    If cboGrupos.ListCount = 0 Then
        MsgBox "Debe tener un grupo seleccionado!", vbCritical, "Error"
        Exit Function
    End If
    Dim ErrorCode As Integer
    ErrorCode = 0
    If cboRubros.ListIndex = -1 Or cboGrupos.ListIndex = -1 Or cboUnidades.ListIndex = -1 Or cboUnidadPedido.ListIndex = -1 Or cboUnidadCompra.ListIndex = -1 Then
        ErrorCode = 1
    End If

    If cboMonedas.ListIndex = -1 Then ErrorCode = 3

    If Not IsNumeric(Me.txtLargo.Text) Or Not IsNumeric(Me.txtEspesor.Text) Or Not IsNumeric(Me.txtAltura.Text) Or Not IsNumeric(Me.txtAncho.Text) Or Not IsNumeric(txtValor) Then
        ErrorCode = 2
    End If
    'If DAOMateriales.existeCodigo(Trim(UCase(Me.txtCodigo.text))) = 1 And Me.txtCodigo.text <> codigoViejo Then     '1=codigo existente.
    '     ErrorCode = 4
    'End If
    Select Case ErrorCode
    Case 1: MsgBox "Debe seleccionar Rubros/Grupos", vbCritical, "Error"
    Case 2: MsgBox "Debe introducir datos válidos para espesor/Peso/Valor", vbCritical, "Error"
    Case 3: MsgBox "Debe seleccionar Moneda", vbCritical, "Error"
    Case 4: MsgBox "El código existe", vbCritical, "Error"
    End Select
    If ErrorCode = 0 Then
        If IsSomething(vmaterial) Then
            'se modifica
            armarObjeto

            If DAOMateriales.modificar(vmaterial, True) Then
                MsgBox "El material se ha modificado.", vbInformation
                DaoHistorico.Save "materiales_historial", "Material desaprobado", vmaterial.Id
                If CDbl(valorViejo <> vmaterial.Valor) Then
                    MsgBox "Hubo una variacion. Se almacena en históricos", vbInformation, "Información"
                    DAOMaterialHistorico.crear vmaterial
                    Unload Me
                End If
            Else
                GoTo err2
                Exit Function
            End If
        Else
            'es nuevo material
            Set vmaterial = New clsMaterial
            armarObjeto
            If DAOMateriales.crear(vmaterial) Then
                BuildCodigoMaterial
                MsgBox "Material creado. Código Material: " & Me.lblCodigoNuevo.caption, vbInformation
                Set vmaterial = Nothing
                Unload Me
            Else
                Set vmaterial = Nothing
                MsgBox "Se produjo algún error, no se guardará el nuevo material.", vbCritical
            End If
        End If
    End If


    Exit Function
err2:
    accion = False
    MsgBox "Se produjo un error: " & Err.Description, vbCritical, "Error"
    Set vmaterial = Nothing
End Function
Private Sub armarObjeto()
    vmaterial.Grupo = DAOGrupos.GetById(Me.cboGrupos.ItemData(cboGrupos.ListIndex))
    vmaterial.Grupo.rubros = DAORubros.FindById(Me.cboRubros.ItemData(cboRubros.ListIndex))
    vmaterial.unidad = CInt(Me.cboUnidades.ItemData(cboUnidades.ListIndex))
    vmaterial.UnidadCompra = CInt(Me.cboUnidadCompra.ItemData(cboUnidadCompra.ListIndex))
    vmaterial.UnidadPedido = CInt(Me.cboUnidadCompra.ItemData(cboUnidadPedido.ListIndex))
    vmaterial.almacen = DAOAlmacenes.GetById(Me.cboAlmacenes.ItemData(Me.cboAlmacenes.ListIndex))
    vmaterial.Largo = Val(Me.txtLargo)
    vmaterial.Ancho = Val(Me.txtAncho)
    vmaterial.Cantidad = Val(Me.txtCantidad)

    If vmaterial.Id = 0 Then
        vmaterial.codigo = UCase(Me.lblCodigoNuevo.caption)
    Else
        If Me.lblCodigoAnterior.caption = Me.lblCodigoNuevo.caption Then
            vmaterial.codigo = UCase(Me.lblCodigoNuevo.caption)
        Else
            If MsgBox("¿Desea utilizar el codigo nuevo para el material?" & vbNewLine & "[Si] : Actualiza con codigo nuevo" & vbNewLine & "[No] : Preserva codigo anterior", vbQuestion + vbYesNo) = vbYes Then
                vmaterial.codigo = UCase(Me.lblCodigoNuevo.caption)
            Else
                vmaterial.codigo = UCase(Me.lblCodigoAnterior.caption)
            End If
        End If
    End If
    vmaterial.StockMinimo = Val(Me.txtStockMinimo.Text)
    vmaterial.PuntoReposicion = Val(Me.txtPtoReposicion.Text)
    vmaterial.descripcion = UCase(Me.txtDescripcion)
    vmaterial.Espesor = Val(Me.txtEspesor)
    vmaterial.PesoXUnidad = Val(Me.txtKgXM2Ml)
    vmaterial.Valor = Val(txtValor)
    vmaterial.moneda = DAOMoneda.GetById(Me.cboMonedas.ItemData(Me.cboMonedas.ListIndex))
    vmaterial.almacen = DAOAlmacenes.GetById(Me.cboAlmacenes.ItemData(Me.cboAlmacenes.ListIndex))
    vmaterial.estado = 1
    vmaterial.Cantidad = Val(Me.txtCantidad)
    vmaterial.FechaValor = Date

    vmaterial.Altura = Val(Me.txtAltura.Text)
    vmaterial.Tipo = Me.cboTipoMaterial.ItemData(Me.cboTipoMaterial.ListIndex)
End Sub
Private Sub Command1_Click()
    accion
End Sub

Private Sub llenarCombosRubrosGrupos()
    DAORubros.LlenarComboExtremeSuite Me.cboRubros
    Set rubroElegido = DAORubros.FindById(cboRubros.ItemData(cboRubros.ListIndex))
    DAOGrupos.llenarComboXtremeSuite Me.cboGrupos, rubroElegido
End Sub

Private Sub Form_Load()
    FormHelper.Customize Me
    limpiar
    llenarCombosRubrosGrupos

    DAOAlmacenes.llenarComboXtremeSuite Me.cboAlmacenes
    DAOMoneda.llenarComboXtremeSuite Me.cboMonedas
    llenarComboPerfiles
    llenarComboUnidades
    Me.cboUnidades.ListIndex = -1
    Me.cboUnidadCompra.ListIndex = -1
    Me.cboUnidadPedido.ListIndex = -1
    Me.cboTipoMaterial.ListIndex = -1
    vId = funciones.CreateGUID
    Channel.AgregarSuscriptor Me, RubrosGrupos_
    If Not vmaterial Is Nothing Then mostrarForm

    ''Me.caption = caption & " (" & Name & ")"

End Sub

Private Sub llenarComboPerfiles()
    Me.cboTipoMaterial.Clear
    Dim key As Variant

    For Each key In tipoMateriales.Keys()
        Me.cboTipoMaterial.AddItem tipoMateriales.item(key)
        Me.cboTipoMaterial.ItemData(Me.cboTipoMaterial.NewIndex) = key
    Next key

End Sub

Private Sub Form_Terminate()
    Set vmaterial = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set vmaterial = Nothing
End Sub
Private Sub llenarComboUnidades()
    Me.cboUnidades.AddItem enums.enumUnidades(Unidades.kg_)
    Me.cboUnidades.ItemData(Me.cboUnidades.NewIndex) = Unidades.kg_
    Me.cboUnidades.AddItem enums.enumUnidades(Unidades.m2_)
    Me.cboUnidades.ItemData(Me.cboUnidades.NewIndex) = Unidades.m2_
    Me.cboUnidades.AddItem enums.enumUnidades(Unidades.Ml_)
    Me.cboUnidades.ItemData(Me.cboUnidades.NewIndex) = Unidades.Ml_
    Me.cboUnidades.AddItem enums.enumUnidades(Unidades.un_)
    Me.cboUnidades.ItemData(Me.cboUnidades.NewIndex) = Unidades.un_
    Me.cboUnidades.AddItem enums.enumUnidades(Unidades.litro_)
    Me.cboUnidades.ItemData(Me.cboUnidades.NewIndex) = Unidades.litro_

    Me.cboUnidadPedido.AddItem enums.enumUnidades(Unidades.kg_)
    Me.cboUnidadPedido.ItemData(Me.cboUnidadPedido.NewIndex) = Unidades.kg_
    Me.cboUnidadPedido.AddItem enums.enumUnidades(Unidades.m2_)
    Me.cboUnidadPedido.ItemData(Me.cboUnidadPedido.NewIndex) = Unidades.m2_
    Me.cboUnidadPedido.AddItem enums.enumUnidades(Unidades.Ml_)
    Me.cboUnidadPedido.ItemData(Me.cboUnidadPedido.NewIndex) = Unidades.Ml_
    Me.cboUnidadPedido.AddItem enums.enumUnidades(Unidades.un_)
    Me.cboUnidadPedido.ItemData(Me.cboUnidadPedido.NewIndex) = Unidades.un_
    Me.cboUnidadPedido.AddItem enums.enumUnidades(Unidades.litro_)
    Me.cboUnidadPedido.ItemData(Me.cboUnidadPedido.NewIndex) = Unidades.litro_

    Me.cboUnidadCompra.AddItem enums.enumUnidades(Unidades.kg_)
    Me.cboUnidadCompra.ItemData(Me.cboUnidadCompra.NewIndex) = Unidades.kg_
    Me.cboUnidadCompra.AddItem enums.enumUnidades(Unidades.m2_)
    Me.cboUnidadCompra.ItemData(Me.cboUnidadCompra.NewIndex) = Unidades.m2_
    Me.cboUnidadCompra.AddItem enums.enumUnidades(Unidades.Ml_)
    Me.cboUnidadPedido.ItemData(Me.cboUnidadCompra.NewIndex) = Unidades.Ml_
    Me.cboUnidadCompra.AddItem enums.enumUnidades(Unidades.un_)
    Me.cboUnidadCompra.ItemData(Me.cboUnidadCompra.NewIndex) = Unidades.un_
    Me.cboUnidadCompra.AddItem enums.enumUnidades(Unidades.litro_)
    Me.cboUnidadCompra.ItemData(Me.cboUnidadCompra.NewIndex) = Unidades.litro_

End Sub

Private Property Get ISuscriber_id() As String
    ISuscriber_id = vId
End Property

Private Function ISuscriber_Notificarse(EVENTO As clsEventoObserver) As Variant
    If EVENTO.Tipo = RubrosGrupos_ Then
        llenarCombosRubrosGrupos

    End If

End Function


Private Sub txtAncho_GotFocus()
    foco Me.txtAncho
End Sub
Private Sub txtAncho_Validate(Cancel As Boolean)
    If Not IsNumeric(Me.txtAncho) Then Cancel = True Else Cancel = False
End Sub
Private Sub txtLargo_GotFocus()
    foco Me.txtLargo
End Sub
Private Sub txtLargo_Validate(Cancel As Boolean)
    If Not IsNumeric(Me.txtLargo) Then Cancel = True Else Cancel = False
End Sub
Private Sub mostrarForm()


    If vmaterial.Tipo = 0 Then
        Me.cboTipoMaterial.ListIndex = -1
    Else
        Me.cboTipoMaterial.ListIndex = funciones.PosIndexCbo(vmaterial.Tipo, Me.cboTipoMaterial)
    End If

    Me.cboUnidades.ListIndex = funciones.PosIndexCbo(vmaterial.unidad, Me.cboUnidades)
    Me.cboUnidadCompra.ListIndex = funciones.PosIndexCbo(vmaterial.UnidadCompra, Me.cboUnidadCompra)
    Me.cboUnidadPedido.ListIndex = funciones.PosIndexCbo(vmaterial.UnidadPedido, Me.cboUnidadPedido)

    Me.cboRubros.ListIndex = funciones.PosIndexCbo(vmaterial.Grupo.rubros.Id, Me.cboRubros)
    Me.cboGrupos.ListIndex = funciones.PosIndexCbo(vmaterial.Grupo.Id, Me.cboGrupos)
    Me.cboAlmacenes.ListIndex = funciones.PosIndexCbo(vmaterial.almacen.Id, Me.cboAlmacenes)
    'Me.txtCodigo = vMaterial.codigo
    Me.lblCodigoAnterior.caption = vmaterial.codigo
    Me.txtDescripcion = vmaterial.descripcion
    Me.txtEspesor = Format(vmaterial.Espesor, "0.00")
    Me.txtKgXM2Ml = vmaterial.PesoXUnidad
    Me.txtCantidad = vmaterial.Cantidad
    Me.txtValor = vmaterial.Valor
    valorViejo = vmaterial.Valor
    Me.cboMonedas.ListIndex = funciones.PosIndexCbo(vmaterial.moneda.Id, Me.cboMonedas)
    Me.txtStockMinimo.Text = vmaterial.StockMinimo
    Me.txtPtoReposicion.Text = vmaterial.PuntoReposicion


    Me.txtLargo = vmaterial.Largo
    Me.txtAncho = vmaterial.Ancho

    BuildCodigoMaterial
End Sub


Private Sub limpiar()
'Me.txtCodigo.text = vbNullString
    Me.lblCodigoAnterior.caption = vbNullString
    Me.lblCodigoNuevo.caption = vbNullString
    Me.txtDescripcion.Text = vbNullString
    txtValor = Empty
    txtFecha = Date
End Sub
Private Sub ValidarUnidades()
    If Me.cboTipoMaterial.ListIndex = -1 Then Exit Sub
    Me.cboUnidades.Enabled = (Me.cboTipoMaterial.ItemData(Me.cboTipoMaterial.ListIndex) = TipoMaterial.TM_UnidadKilo)

    Dim A As Integer
    A = cboUnidades.ItemData(cboUnidades.ListIndex)

    LimpiarDimensiones

    txtKgXM2Ml.Locked = (A = Unidades.un_ Or A = Unidades.kg_)
    If A = Unidades.un_ Or A = Unidades.kg_ Then
        Me.txtAncho.Enabled = (A = Unidades.m2_)
        Me.txtLargo.Enabled = (A = Unidades.m2_ Or A = Unidades.Ml_)
        Me.txtEspesor.Enabled = False
        Me.txtAltura.Enabled = False
        Me.txtKgXM2Ml = 1
    End If
End Sub

Private Sub LimpiarDimensiones()
    Me.txtAncho.Text = 0
    Me.txtLargo.Text = 0
    Me.txtEspesor.Text = 0
    Me.txtAltura.Text = 0
    Me.txtKgXM2Ml.Text = 0
End Sub
