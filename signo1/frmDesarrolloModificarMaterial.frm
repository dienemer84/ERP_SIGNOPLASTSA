VERSION 5.00
Begin VB.Form frmDesarrolloModificarMaterial 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Modificar material..."
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4365
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   4365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ Detalle ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   4335
      Begin VB.TextBox txtDetalle 
         Height          =   285
         Left            =   960
         TabIndex        =   4
         Top             =   1440
         Width           =   3255
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Aceptar"
         Default         =   -1  'True
         Height          =   375
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   3600
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   3600
         Width           =   975
      End
      Begin VB.TextBox txtYPieza 
         Height          =   285
         Left            =   3120
         TabIndex        =   8
         Top             =   3000
         Width           =   1095
      End
      Begin VB.TextBox txtXPieza 
         Height          =   285
         Left            =   960
         TabIndex        =   7
         Top             =   3000
         Width           =   1095
      End
      Begin VB.TextBox txtYHoja 
         Height          =   285
         Left            =   3120
         TabIndex        =   6
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox txtXHoja 
         Height          =   285
         Left            =   960
         TabIndex        =   5
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox txtScrap 
         Height          =   285
         Left            =   3120
         TabIndex        =   3
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtCantidad 
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Left            =   960
         TabIndex        =   0
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Detalle "
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
         TabIndex        =   22
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Pieza"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Hoja"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1800
         Width           =   735
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   4200
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   4200
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Label lblMaterial 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   960
         TabIndex        =   1
         Top             =   720
         Width           =   3255
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Scrap "
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
         Left            =   2280
         TabIndex        =   19
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ancho "
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
         Left            =   2280
         TabIndex        =   18
         Top             =   3000
         Width           =   855
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Largo "
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
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Largo "
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
         TabIndex        =   16
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ancho "
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
         Left            =   2400
         TabIndex        =   15
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cantidad "
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
         TabIndex        =   14
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Material "
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
         TabIndex        =   13
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Código "
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
         TabIndex        =   12
         Top             =   360
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmDesarrolloModificarMaterial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim idUnidad As Integer
Dim baseS As New classStock
Dim baseM As New classConfigurar
Dim vidcodigo As Long
Dim formu As frmNuevoElemento

Dim Modificado As Boolean
Public Property Let idcodigo(nidcodigo As Long)
    vidcodidgo = nidcodigo
End Property

Public Property Let nuevo_form(frm As Variant)
    Set formu = frm

End Property


Private Sub Command1_Click()
    Dim costo As Double
    If Not Modificado Then
        If MsgBox("¿Está seguro de actualizar?", vbYesNo, "Confirmación") = vbYes Then



            codigo = Me.txtCodigo
            Scrap = CDbl(Me.txtScrap)
            Cantidad = CDbl(Me.txtCantidad)
            xhoja = CDbl(Me.txtXHoja)
            yhoja = CDbl(Me.txtYHoja)
            xpieza = CDbl(Me.txtXPieza)
            ypieza = CDbl(Me.txtYPieza)
            Id = baseM.QueIdMaterial(Trim(Me.txtCodigo))
            baseS.calcularM2MLKGMaterial ypieza, xpieza, Id, Scrap, yhoja, xhoja, Cantidad, Kg, m2ml, Pieza, costo, 0

            'x1 es hoja
            'x es pieza
            baseS.ejecutar "select m.id_unidad,m.espesor,m.descripcion, g.grupo, r.rubro from materiales m,grupos g,rubros r where m.id_grupo=g.id and m.id_rubro=r.id and  m.id=" & Id
            descripcion = baseS.descripcion
            Espesor = baseS.Espesor
            Grupo = baseS.Grupo
            rubro = baseS.rubro
            descripcionPieza = rubro & " " & Grupo & " " & descripcion


            'hay que modificar el peso los m2ml el detalle.. etc etc etc
            '    frmNuevoElemento.ListView1.SelectedItem = codigo
            '    frmNuevoElemento.ListView1.SelectedItem.ListSubItems(1) = id
            '    frmNuevoElemento.ListView1.SelectedItem.ListSubItems(2) = descripcionPieza
            '    frmNuevoElemento.ListView1.SelectedItem.ListSubItems(3) = espesor
            '    frmNuevoElemento.ListView1.SelectedItem.ListSubItems(4) = pieza
            '    frmNuevoElemento.ListView1.SelectedItem.ListSubItems(9) = scrap
            '    frmNuevoElemento.ListView1.SelectedItem.ListSubItems(5) = xhoja 'XHOHA
            '    frmNuevoElemento.ListView1.SelectedItem.ListSubItems(6) = yhoja
            '    frmNuevoElemento.ListView1.SelectedItem.ListSubItems(7) = xpieza 'XPIEZA
            '    frmNuevoElemento.ListView1.SelectedItem.ListSubItems(8) = ypieza
            '    frmNuevoElemento.ListView1.SelectedItem.ListSubItems(10) = kg
            '    frmNuevoElemento.ListView1.SelectedItem.ListSubItems(11) = m2ml
            '    frmNuevoElemento.ListView1.SelectedItem.ListSubItems(12) = funciones.formatearDecimales(costo, 2)
            '    frmNuevoElemento.ListView1.SelectedItem.ListSubItems(14) = cantidad
            '    frmNuevoElemento.ListView1.SelectedItem.Tag = UCase(Trim(Me.txtDetalle))

            formu.ListView1.selectedItem = codigo
            formu.ListView1.selectedItem.ListSubItems(1) = Id
            formu.ListView1.selectedItem.ListSubItems(2) = descripcionPieza
            formu.ListView1.selectedItem.ListSubItems(3) = Espesor
            formu.ListView1.selectedItem.ListSubItems(4) = Pieza
            formu.ListView1.selectedItem.ListSubItems(9) = Scrap
            formu.ListView1.selectedItem.ListSubItems(5) = xhoja    'XHOHA
            formu.ListView1.selectedItem.ListSubItems(6) = yhoja
            formu.ListView1.selectedItem.ListSubItems(7) = xpieza    'XPIEZA
            formu.ListView1.selectedItem.ListSubItems(8) = ypieza
            formu.ListView1.selectedItem.ListSubItems(10) = Kg
            formu.ListView1.selectedItem.ListSubItems(11) = m2ml
            formu.ListView1.selectedItem.ListSubItems(12) = funciones.FormatearDecimales(costo, 2)
            formu.ListView1.selectedItem.ListSubItems(14) = Cantidad
            formu.ListView1.selectedItem.Tag = UCase(Trim(Me.txtDetalle))

            Unload Me
        End If
    Else

        Unload Me
    End If

End Sub

Private Sub Command2_Click()
    If Modificado Then
        If MsgBox("¿Está seguro de volver?", vbYesNo, "Confirmación") = vbYes Then
            Unload Me
        End If
    Else
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    FormHelper.Customize Me
    Modificado = False


End Sub

Private Sub txtCodigo_Change()
'ver estado. si es=0 no se puede usar para cotizar... (valor 0)
    If baseM.QueIdMaterial(Trim(Me.txtCodigo)) > 1 Then    'existe codigo
        IdMaterial = baseM.QueIdMaterial(Trim(Me.txtCodigo))
        verDetalleMateriales IdMaterial
        Dim r As Recordset
        Set r = conectar.RSFactory("select id_unidad from materiales where id=" & IdMaterial)
        If Not r.EOF And Not r.BOF Then
            idUnidad = r!id_Unidad
            If idUnidad = 1 Then
                Me.txtXHoja.Enabled = False
                Me.txtXPieza.Enabled = False
                Me.txtYHoja.Enabled = False
                Me.txtYPieza.Enabled = False
                Me.txtScrap.Enabled = True
            ElseIf idUnidad = 2 Then    'm2
                Me.txtXHoja.Enabled = True
                Me.txtXPieza.Enabled = True
                Me.txtYHoja.Enabled = True
                Me.txtYPieza.Enabled = True
                Me.txtScrap.Enabled = True
            ElseIf idUnidad = 3 Then    'ml
                Me.txtXHoja.Enabled = True
                Me.txtXPieza.Enabled = True
                Me.txtYHoja.Enabled = False
                Me.txtYPieza.Enabled = False
                Me.txtScrap.Enabled = True

            ElseIf idUnidad = 4 Then    'un
                Me.txtXHoja.Enabled = False
                Me.txtXPieza.Enabled = False
                Me.txtYHoja.Enabled = False
                Me.txtYPieza.Enabled = False
                Me.txtScrap.Enabled = False
            End If
        End If
    Else
        Me.lblMaterial = "Código inexistente"
    End If
End Sub
Private Function verDetalleMateriales(Id)
    Dim Kg As Double, m2ml As Double
    Dim descripcion As String
    Dim costo As Double
    cxh = funciones.cantxhoja(x, y, x1, y1)
    baseS.ejecutar "select m.espesor,m.descripcion, g.grupo, r.rubro from materiales m,grupos g, rubros r where m.id_grupo=g.id and m.id_rubro=r.id and  m.id=" & Id
    descripcion = baseS.descripcion
    Espesor = baseS.Espesor
    Grupo = baseS.Grupo
    rubro = baseS.rubro

    Me.lblMaterial = truncar(rubro, 40) & " " & truncar(Grupo, 40) & " " & truncar(MAT, 40)
    If Espesor > 0 Then
        Me.lblMaterial = Me.lblMaterial & " " & Espesor & "mm"
    End If
End Function
Private Sub txtXHoja_GotFocus()
    foco Me.txtXHoja
End Sub
Private Sub txtXPieza_GotFocus()
    foco Me.txtXPieza
End Sub

Private Sub txtYHoja_GotFocus()
    foco Me.txtYHoja
End Sub
Private Sub txtYPieza_GotFocus()
    foco Me.txtYPieza
End Sub
