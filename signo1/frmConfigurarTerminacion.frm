VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmConfigurarTerminacion 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Calculo de terminación-> pintura"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   9450
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   9450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Salir"
      Height          =   255
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Recalcular"
      Height          =   255
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Agregar"
      Height          =   255
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   1680
      Width           =   975
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ Configuración ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   0
      TabIndex        =   3
      Top             =   1920
      Width           =   3615
      Begin VB.TextBox txtOperarios 
         Height          =   285
         Left            =   2640
         TabIndex        =   13
         Text            =   "Text7"
         Top             =   3480
         Width           =   735
      End
      Begin VB.CommandButton Command23 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Recalcular"
         Height          =   375
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   3960
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   2640
         TabIndex        =   12
         Text            =   "Text6"
         Top             =   3120
         Width           =   735
      End
      Begin VB.TextBox Textf 
         Height          =   285
         Left            =   2640
         TabIndex        =   10
         Text            =   "Text6"
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox Textg 
         Height          =   285
         Left            =   2640
         TabIndex        =   11
         Text            =   "Text7"
         Top             =   2760
         Width           =   735
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   2640
         TabIndex        =   9
         Text            =   "Text5"
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   2640
         TabIndex        =   8
         Text            =   "Text4"
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   2640
         TabIndex        =   7
         Text            =   "Text3"
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2640
         TabIndex        =   6
         Text            =   "Text2"
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2640
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cantidad de operarios"
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
         TabIndex        =   46
         Top             =   3480
         Width           =   2415
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Espesor de la pintura"
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
         TabIndex        =   20
         Top             =   3120
         Width           =   2415
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Indice de aumento MDO"
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
         TabIndex        =   19
         Top             =   2400
         Width           =   2415
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Indice de aumento"
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
         TabIndex        =   18
         Top             =   2760
         Width           =   2415
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tpo. Horneado por M2"
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
         TabIndex        =   17
         Top             =   1920
         Width           =   2415
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tpo. Pintura por M2"
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
         Top             =   1560
         Width           =   2415
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tpo. Preparacion por M2"
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
         TabIndex        =   15
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fosfatos"
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
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cantidad de Pintura por M2"
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
         TabIndex        =   4
         Top             =   480
         Width           =   2415
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Quitar"
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1680
      Width           =   615
   End
   Begin MSComctlLib.ListView lstPiezas 
      Height          =   1455
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   2566
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Código"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Material"
         Object.Width           =   5821
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Kg"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "M2"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Medidas"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Caras"
         Object.Width           =   1279
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Capas"
         Object.Width           =   1279
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "cantidad"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "largo"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "ancho"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ Calculo general ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   3720
      TabIndex        =   0
      Top             =   1920
      Width           =   5655
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "[ Materiales ]"
         Height          =   1695
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   5415
         Begin VB.CheckBox Check6 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Todos"
            Height          =   195
            Left            =   4440
            TabIndex        =   44
            Top             =   0
            Width           =   855
         End
         Begin VB.CheckBox Check2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Check2"
            Height          =   255
            Left            =   4920
            TabIndex        =   40
            Top             =   1200
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Check1"
            Height          =   255
            Left            =   4920
            TabIndex        =   39
            Top             =   600
            Width           =   255
         End
         Begin VB.ComboBox cboFosf 
            Height          =   315
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   35
            Top             =   1200
            Width           =   3495
         End
         Begin VB.ComboBox cboCant 
            Height          =   315
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   600
            Width           =   3495
         End
         Begin VB.Label lblfosfatosReal 
            BackColor       =   &H00C0C0C0&
            Height          =   135
            Left            =   4800
            TabIndex        =   50
            Top             =   720
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label lblcantPintReal 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Height          =   255
            Left            =   4800
            TabIndex        =   49
            Top             =   360
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label8 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Cantidad de pintura"
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
            TabIndex        =   33
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label lblCantPint 
            BackColor       =   &H00E0E0E0&
            Height          =   255
            Left            =   3840
            TabIndex        =   32
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label10 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Fosfatos"
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
            TabIndex        =   31
            Top             =   960
            Width           =   855
         End
         Begin VB.Label lblfosfatos 
            BackColor       =   &H00E0E0E0&
            Height          =   255
            Left            =   3840
            TabIndex        =   30
            Top             =   1200
            Width           =   975
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "[ Mano de obra ]"
         Height          =   2295
         Left            =   120
         TabIndex        =   22
         Top             =   2040
         Width           =   5415
         Begin VB.CheckBox Check5 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Check5"
            Height          =   255
            Left            =   4920
            TabIndex        =   43
            Top             =   1800
            Width           =   255
         End
         Begin VB.CheckBox Check4 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Check4"
            Height          =   255
            Left            =   4920
            TabIndex        =   42
            Top             =   1200
            Width           =   255
         End
         Begin VB.CheckBox Check3 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Check3"
            Height          =   255
            Left            =   4920
            TabIndex        =   41
            Top             =   600
            Width           =   255
         End
         Begin VB.ComboBox cboHorno 
            Height          =   315
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   38
            Top             =   1800
            Width           =   3495
         End
         Begin VB.ComboBox cboApp 
            Height          =   315
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   37
            Top             =   1200
            Width           =   3495
         End
         Begin VB.ComboBox cboSup 
            Height          =   315
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   600
            Width           =   3495
         End
         Begin VB.Label lblappReal 
            BackColor       =   &H00C0C0C0&
            Height          =   255
            Left            =   3120
            TabIndex        =   52
            Top             =   960
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label lblSuperficieReal 
            BackColor       =   &H00C0C0C0&
            Height          =   255
            Left            =   3600
            TabIndex        =   51
            Top             =   240
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label Label11 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Superficies"
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
            TabIndex        =   28
            Top             =   360
            Width           =   975
         End
         Begin VB.Label lblSuperficie 
            BackColor       =   &H00E0E0E0&
            Height          =   255
            Left            =   3840
            TabIndex        =   27
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label12 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Aplicación"
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
            TabIndex        =   26
            Top             =   960
            Width           =   975
         End
         Begin VB.Label lblAplicacion 
            BackColor       =   &H00E0E0E0&
            Height          =   255
            Left            =   3840
            TabIndex        =   25
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label Label13 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Horneado"
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
            TabIndex        =   24
            Top             =   1560
            Width           =   975
         End
         Begin VB.Label lblHorneado 
            BackColor       =   &H00E0E0E0&
            Height          =   255
            Left            =   3840
            TabIndex        =   23
            Top             =   1800
            Width           =   975
         End
         Begin VB.Label lblhornoReal 
            BackColor       =   &H00C0C0C0&
            Height          =   375
            Left            =   3240
            TabIndex        =   53
            Top             =   1440
            Visible         =   0   'False
            Width           =   735
         End
      End
   End
   Begin VB.Label Ancho_total 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Label15"
      Height          =   255
      Left            =   8160
      TabIndex        =   48
      Top             =   1680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label largo_total 
      BackColor       =   &H00C0C0C0&
      DataMember      =   "               "
      Height          =   255
      Left            =   6360
      TabIndex        =   47
      Top             =   1680
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "frmConfigurarTerminacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim baseNMDO As New classNuevaMDO
Dim dto As DTOCuentasTerminacion

Dim idcodigoMDO As Integer
Dim baseC As New classConfigurar
Dim baseM As New classNuevaMDO
Dim baseS As New classStock
'Dim baseR As New classRubrosGrupos
Dim baseSP As New classSignoplast
Dim formu As frmNuevoElemento

Public Property Let nuevo_form(frm As Variant)
    Set formu = frm

End Property

Private Sub Check6_Click()
    If Check6.value = 1 Then
        Check1.value = 1
        Check2.value = 1
        Check3.value = 1
        Check4.value = 1
        Check5.value = 1
    Else
        Check1.value = 0
        Check2.value = 0
        Check3.value = 0
        Check4.value = 0
        Check5.value = 0
    End If
End Sub
Private Sub Command1_Click()
    quitar_de_lista Me.lstPiezas
End Sub
Private Sub Command2_Click()
    calcular
End Sub
Private Sub Command4_Click()
    Unload Me
End Sub
Private Sub Command3_Click()
    On Error Resume Next
    Cant = formu.ListView1.ListItems.count
    Dim costo As Double

    'agrego cuenta cantidad de pintura
    If Me.Check1.value = 1 Then
        Dim idcodigo As Integer
        idcodigo = Me.cboCant.ItemData(Me.cboCant.ListIndex)
        baseC.ejecutar_consulta "select id,codigo,descripcion from materiales where id=" & idcodigo
        codigo = baseC.codigoMaterial
        descripcion = baseC.descripcionMaterial
        esta = False
        For i = 1 To Cant
            If formu.ListView1.ListItems(i).ListSubItems(10) = idcodigo Then
                esta = True
                pos = i
            End If
        Next i
        baseS.calcularM2MLKGMaterial 0, 0, idcodigo, 0, 0, 0, CDbl(lblcantPintReal), Kg, m2ml, Pieza, costo, 0
        If Not esta Then
            Set x = formu.ListView1.ListItems.Add(, , codigo)
            x.SubItems(1) = idcodigo
            x.SubItems(2) = descripcion
            x.SubItems(3) = 0
            x.SubItems(4) = Me.lblCantPint    '& "x" & Me.largo_total & "x" & Me.Ancho_total
            x.SubItems(5) = 0
            x.SubItems(6) = 0
            x.SubItems(7) = 0
            x.SubItems(8) = 0
            x.SubItems(9) = 0
            x.SubItems(10) = Me.lblcantPintReal
            x.SubItems(11) = 0
            x.SubItems(12) = Math.Round(costo, 2)
            x.SubItems(13) = 0
            x.SubItems(14) = Me.lblcantPintReal

        Else
            hay = True
        End If
    End If

    'agrego cuenta fosfatos
    If Me.Check2.value = 1 Then
        idcodigo = Me.cboCant.ItemData(Me.cboFosf.ListIndex)
        baseC.ejecutar_consulta "select codigo,descripcion from materiales where id=" & idcodigo
        codigo = baseC.codigoMaterial
        descripcion = baseC.descripcionMaterial
        esta = False
        For i = 1 To Cant
            If formu.ListView1.ListItems(i).ListSubItems(10) = idcodigo Then
                esta = True
                pos = i
            End If
        Next i
        baseS.calcularM2MLKGMaterial 0, 0, idcodigo, 0, 0, 0, CDbl(lblfosfatosReal), Kg, m2ml, Pieza, costo, 0
        If Not esta Then
            Set x = formu.ListView1.ListItems.Add(, , codigo)
            x.SubItems(1) = idcodigo
            x.SubItems(2) = descripcion
            x.SubItems(3) = 0
            x.SubItems(4) = Me.lblfosfatos
            x.SubItems(5) = 0
            x.SubItems(6) = 0
            x.SubItems(7) = 0
            x.SubItems(8) = 0
            x.SubItems(9) = 0
            x.SubItems(10) = Me.lblfosfatosReal
            x.SubItems(11) = 0
            x.SubItems(12) = funciones.FormatearDecimales(costo, 2)
            x.SubItems(13) = 0
            x.SubItems(14) = Me.lblfosfatosReal
        Else
            hay = True
        End If
    End If

    'agrego superficie
    If Check3.value = 1 Then
        idcodigoMDO = Me.cboSup.ItemData(Me.cboSup.ListIndex)
        Tiempo = Me.lblSuperficieReal
        operarios = Me.txtOperarios
        Me.agrego operarios, Tiempo, Sector, idcodigoMDO, formu.ListView2
    End If
    If Check4.value = 1 Then
        idcodigoMDO = Me.cboApp.ItemData(Me.cboApp.ListIndex)
        Tiempo = Me.lblappReal
        operarios = Me.txtOperarios
        Me.agrego operarios, Tiempo, Sector, idcodigoMDO, formu.ListView2
    End If
    If Check5.value = 1 Then
        idcodigoMDO = Me.cboHorno.ItemData(Me.cboHorno.ListIndex)
        Tiempo = Me.lblhornoReal
        operarios = Me.txtOperarios
        Me.agrego operarios, Tiempo, Sector, idcodigoMDO, formu.ListView2
    End If
    Unload Me
End Sub
Public Function agrego(operarios, Tiempo, Sector, codigo As Integer, lst As ListView, Optional Valor)
    Dim totmin As Double
    Dim totplata As Double
    esta = False
    Cant = formu.ListView2.ListItems.count
    For i = 1 To Cant
        If formu.ListView2.ListItems(i).ListSubItems(1) = idcodigoMDO Then
            esta = True
            pos = i
        End If
    Next


    If Not esta Then
        baseM.VERMDO idcodigoMDO, cantxproc, Sector, Tarea, descrip
        If cantxproc = -1 Then cantxproc2 = "Cambio"
        If cantxproc = 0 Then cantxproc2 = "Fijo"
        Set x = lst.ListItems.Add(, , idcodigoMDO)
        x.SubItems(1) = idcodigoMDO
        x.SubItems(2) = operarios
        x.SubItems(3) = Tiempo
        x.SubItems(4) = Sector
        x.SubItems(5) = cantxproc2
        x.SubItems(7) = cantxproc
        x.SubItems(6) = Tarea
        x.SubItems(8) = descripcion
        baseSP.ejecutar "select valor from valores_MDO where id_tarea=" & codigo
        Valor = baseSP.valorMDO
        Tiempo = Tiempo
        cpp = cantxproc
        cantop = operarios
        If cpp > 0 Then    '(cpp variable)
            totmin = cantop * Tiempo / cpp
            totplata = totmin * Valor
        Else
            totmin = cantop * Tiempo
            totplata = totmin * Valor
        End If

        x.SubItems(9) = funciones.FormatearDecimales(totmin, 2)
        x.SubItems(10) = funciones.FormatearDecimales(totplata, 2)


    Else  'si esta
        baseM.VERMDO idcodigoMDO, cantxproc, Sector, Tarea, descrip
        Tiempo = Tiempo + CDbl(lst.ListItems(pos).ListSubItems(3))
        lst.ListItems(pos).SubItems(3) = Tiempo

        baseSP.ejecutar "select valor from valores_MDO where id_tarea=" & codigo
        Valor = baseSP.valorMDO
        Tiempo = Tiempo
        cpp = cantxproc
        cantop = operarios
        If cpp > 0 Then    '(cpp variable)
            totmin = cantop * Tiempo / cpp
            totplata = totmin * Valor
        Else
            totmin = cantop * Tiempo
            totplata = totmin * Valor
        End If

        lst.ListItems(pos).SubItems(9) = funciones.FormatearDecimales(totmin, 2)
        lst.ListItems(pos).SubItems(10) = funciones.FormatearDecimales(totplata, 2)

    End If
End Function
Private Sub Form_Activate()
    Dim i As Integer
    Me.txtOperarios = 1
    Me.Check6.value = 1
    calcular
    i = baseC.sector_terminacion
    Set dto = DAOdtoCuentasTerminacion.GetConfigTerminacion
    LlenarCuentasMDO
    LlenarCuentasMAT
    Dim r As Integer
    Me.Refresh
End Sub
Private Sub LlenarCuentasMDO()
    On Error Resume Next
    DAOTareas.LlenarComboPorSector Me.cboHorno, dto.Sector
    DAOTareas.LlenarComboPorSector Me.cboApp, dto.Sector
    DAOTareas.LlenarComboPorSector Me.cboSup, dto.Sector
    Me.cboHorno.ListIndex = funciones.PosIndexCbo(dto.Horneado.id, Me.cboHorno)
    Me.cboApp.ListIndex = funciones.PosIndexCbo(dto.Aplicacion.id, Me.cboApp)
    Me.cboSup.ListIndex = funciones.PosIndexCbo(dto.Limpieza.id, Me.cboSup)

End Sub

Private Sub LlenarCuentasMAT()
    On Error Resume Next
    DAOMateriales.LlenarComboPorRubro Me.cboCant, dto.rubro
    DAOMateriales.LlenarComboPorRubro Me.cboFosf, dto.rubro
    Me.cboCant.ListIndex = funciones.PosIndexCbo(dto.CantidadPintura.id, cboCant)
    Me.cboFosf.ListIndex = funciones.PosIndexCbo(dto.CantidadFosfatos.id, cboFosf)
End Sub

Private Sub Form_Load()
    FormHelper.Customize Me
    Dim baseS As New classStock
    baseC.ver_datos_pintura A, B, c, d, E, F, g, h, i
    Me.Text6 = i
    Me.Text1 = A  'cantpintm2
    Me.Text2 = B    'cantfosfatos
    Me.Text3 = c    'tpo prrp sup
    Me.Text4 = d    'tpo pint m2
    Me.Text5 = E    'tpo horno
    Me.Textg = h    'factor mdo
    Me.Textf = g    'factor mat
    
        ''Me.caption = caption & " (" & Name & ")"
        
        
End Sub

Private Sub calcular()
    Dim cPint As Double, cFosf As Double, cSup As Double, ctiempo As Double, chorno As Double, largoT As Double, anchoT As Double
    baseC.Calcular_terminacion Me.lstPiezas, CDbl(Me.Textg), CDbl(Me.Text1), CDbl(Me.Text6), cPint, cFosf, CDbl(Me.Text2), CDbl(Me.Text3), cSup, ctiempo, CDbl(Me.Text4), chorno, CDbl(Me.Text5), largoT, anchoT
    Me.Ancho_total = Math.Round(anchoT, 4)
    Me.largo_total = Math.Round(largoT, 4)
    Me.lblCantPint = Math.Round(cPint, 4) & " Kg"
    Me.lblcantPintReal = Math.Round(cPint, 4)
    Me.lblfosfatos = Math.Round(cFosf, 4) & " Kg"
    Me.lblfosfatosReal = Math.Round(cFosf, 4)
    Me.lblSuperficie = Math.Round(cSup, 4) & " Min"
    Me.lblSuperficieReal = Math.Round(cSup, 4)
    Me.lblAplicacion = Math.Round(ctiempo, 4) & " Min"
    Me.lblappReal = Math.Round(ctiempo, 4)
    Me.lblHorneado = Math.Round(chorno, 4) & " Min"
    Me.lblhornoReal = Math.Round(chorno, 4)
End Sub

Private Sub lstPiezas_DblClick()
    If Me.lstPiezas.ListItems.count > 0 Then
        frmCapaCaras.txtCaras = Me.lstPiezas.selectedItem.ListSubItems(5)
        frmCapaCaras.txtCapa = Me.lstPiezas.selectedItem.ListSubItems(6)
        frmCapaCaras.Show 1

    End If
End Sub
