VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDesarrolloConjuntosTiempos 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Definir tiempos de conjunto"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   12885
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ Mano de obra ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   4815
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   12855
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Salir"
         Height          =   375
         Left            =   11640
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CommandButton btnModificar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Modificar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10440
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txtCodigoMDO 
         Height          =   285
         Left            =   840
         TabIndex        =   0
         Text            =   "0"
         Top             =   480
         Width           =   3375
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00C0C0C0&
         Caption         =   "[ Detalle ]"
         ForeColor       =   &H00000000&
         Height          =   1095
         Left            =   4320
         TabIndex        =   9
         Top             =   360
         Width           =   8415
         Begin VB.Label va 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Caption         =   "Valor "
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
            Left            =   5160
            TabIndex        =   19
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label T 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Caption         =   "Tarea "
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
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Caption         =   "Descripcion "
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
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Caption         =   "Cant x Proceso "
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
            Left            =   5160
            TabIndex        =   16
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label lblTarea 
            BackColor       =   &H00C0C0C0&
            Height          =   255
            Left            =   1200
            TabIndex        =   15
            Top             =   240
            Width           =   3975
         End
         Begin VB.Label lblDescripcion 
            BackColor       =   &H00C0C0C0&
            Height          =   255
            Left            =   1200
            TabIndex        =   14
            Top             =   480
            Width           =   7095
         End
         Begin VB.Label lblCPP 
            BackColor       =   &H00C0C0C0&
            Height          =   255
            Left            =   6720
            TabIndex        =   13
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Caption         =   "Sector "
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
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label lblSector 
            BackColor       =   &H00C0C0C0&
            Height          =   255
            Left            =   1200
            TabIndex        =   11
            Top             =   720
            Width           =   3975
         End
         Begin VB.Label lblValor 
            BackColor       =   &H00C0C0C0&
            Height          =   255
            Left            =   6720
            TabIndex        =   10
            Top             =   720
            Width           =   1575
         End
      End
      Begin VB.CommandButton btnAgregarMDO 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Agregar"
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox txtCantOp 
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Text            =   "0"
         Top             =   840
         Width           =   2775
      End
      Begin VB.TextBox txtTiempo 
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Text            =   "0"
         ToolTipText     =   "Tiempo en minutos"
         Top             =   1200
         Width           =   3255
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Quitar"
         Height          =   255
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   4440
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Invertir"
         Default         =   -1  'True
         Height          =   255
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   4440
         Width           =   735
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   2175
         Left            =   120
         TabIndex        =   4
         Top             =   2160
         Width           =   12615
         _ExtentX        =   22251
         _ExtentY        =   3836
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Cant OP"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Tiempo"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Sector"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "CPP"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Tarea"
            Object.Width           =   4762
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Descripcion"
            Object.Width           =   4057
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Text            =   "T.Total"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   8
            Text            =   "Costo $"
            Object.Width           =   1587
         EndProperty
      End
      Begin VB.Label Label19 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Código"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label23 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cant Operarios"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label24 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tiempo"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   1200
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmDesarrolloConjuntosTiempos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vconjunto As Long
Dim grabado As Boolean
Dim rss As Recordset
Dim vIdMDO As Long
Dim vCant As Long
Dim vtiempo As Long
Dim vidCPP As Integer
Dim baseS As New classStock
Dim baseM As New classConfigurar
Dim base As New classNuevoElemento
Dim baseSP As New classSignoplast

Public Property Let idConjunto(nconjunto As Long)
    vconjunto = nconjunto
End Property
Private Sub btnAgregarMDO_Click()
    If Not IsNumeric(Me.txtCantOp) Or Not IsNumeric(Me.txtTiempo) Then
        MsgBox "Ingrese datos válidos por favor", vbCritical, "Error"
    Else


        base.ver_detalle_mdo CInt(Me.txtCodigoMDO), idcpp, cantxproc, mdoDescrip, Tarea, Sector, Valor
        vIdMDO = Me.txtCodigoMDO
        Me.lblCPP = cantxproc
        vidCPP = idcpp
        lblTarea = Tarea
        lblDescripcion = mdoDescrip
        lblSector = Sector
        Me.lblValor = Valor




        Set x = Me.ListView2.ListItems.Add(, , vIdMDO)
        x.SubItems(1) = CDbl(Me.txtCantOp)
        x.SubItems(2) = CDbl(Me.txtTiempo)
        x.SubItems(3) = Me.lblSector
        x.SubItems(4) = Me.lblCPP
        x.ListSubItems(4).Tag = vidCPP
        x.SubItems(5) = Me.lblTarea
        x.SubItems(6) = Me.lblDescripcion

        Valor = CDbl(Me.lblValor)
        Tiempo = CDbl(Me.txtTiempo)
        cpp = CInt(vidCPP)
        cantop = CDbl(Me.txtCantOp)
        If cpp > 0 Then    '(cpp variable)
            totmin = cantop * Tiempo / cpp
            totplata = totmin * Valor
        Else
            totmin = cantop * Tiempo
            totplata = totmin * Valor

        End If
        x.SubItems(7) = Math.Round(totmin, 2)
        x.SubItems(8) = Math.Round(totplata, 2)

        Me.txtCodigoMDO.SetFocus
    End If
    grabado = False

End Sub


Private Sub btnModificar_Click()
    On Error GoTo errb
    If MsgBox("¿Desea grabar los cambios?", vbYesNo, "Confirmacion") = vbYes Then
        If Not baseS.GrabarTiemposConjunto(vconjunto, Me.ListView2) Then
            MsgBox "Se produjo algún error!", vbCritical, "Error"
            grabado = False
        Else
            grabado = True
        End If
    End If


    Exit Sub
errb:
    MsgBox Err.Description

End Sub





Private Sub Command1_Click()
    If Not grabado Then
        If MsgBox("¿Está seguro de salir?", vbYesNo, "Confirmación") = vbYes Then
            Unload Me
        End If
    Else
        Unload Me
    End If
End Sub

Private Sub Command2_Click()
'base.ver_detalle_elemento Trim(Me.txtCodigoMaterial), Me, 0
    For i = 1 To Me.ListView2.ListItems.count
        If Me.ListView2.ListItems(i).Checked Then
            Me.ListView2.ListItems(i).Checked = False
        Else
            Me.ListView2.ListItems(i).Checked = True
        End If

    Next i

End Sub


Private Sub Command4_Click()
    If MsgBox("¿Está seguro de eliminar los items seleecionados?", vbYesNo, "Confirmacion") = vbYes Then
        For i = Me.ListView2.ListItems.count To 1 Step -1
            If Me.ListView2.ListItems(i).Checked = True Then
                Me.ListView2.ListItems.remove (i)
                grabado = False
            End If
        Next i
    End If

End Sub






Private Sub Form_Load()
    FormHelper.Customize Me
    Dim rs As Recordset
    Dim x As ListItem
    Set rs = conectar.RSFactory("select vm.valor,t.cantxproc,vm.descripcion,t.id as codigo,d.id,d.cantidad,t.tarea,d.tiempo,s.sector,t.cantxproc from valores_MDO vm,sp.sectores s,desarrollo_mdo d, tareas t  where d.id_pieza=" & vconjunto & " and t.id=d.codigo and s.id=t.id_sector and t.id=vm.id_tarea")
    While Not rs.EOF

        cpa = rs!cantxproc
        cpp = rs!cantxproc

        If cpa = -1 Then cpa = "Cambio"
        If cpa = 0 Then cpa = "Fijo"

        Set x = Me.ListView2.ListItems.Add(, , rs!codigo)
        x.SubItems(1) = rs!Cantidad
        x.SubItems(2) = rs!Tiempo
        x.SubItems(3) = rs!Sector
        'X.SubItems(4) = cpa
        x.SubItems(5) = rs!Tarea
        x.SubItems(4) = cpp
        x.ListSubItems(4).Tag = cpa
        x.SubItems(6) = rs!descripcion
        cantop = rs!Cantidad
        Valor = rs!Valor
        Tiempo = rs!Tiempo
        If cpp > 0 Then    '(cpp variable)
            totmin = cantop * Tiempo / cpp
            totplata = totmin * Valor
        Else
            totmin = cantop * Tiempo
            totplata = totmin * Valor

        End If
        x.SubItems(7) = Math.Round(totmin, 2)
        x.SubItems(8) = Math.Round(totplata, 2)



        rs.MoveNext
    Wend

    Me.limpiar_txt



    grabado = True
End Sub
Function limpiar_txt()

    Me.txtCodigoMDO = 1
    Me.txtTiempo = 1
    Me.txtCantOp = 1

End Function

Private Sub Form_Unload(Cancel As Integer)
    If Not grabado Then
        If MsgBox("¿Desea descartar los cambios?", vbYesNo, "Confirmación") = vbYes Then
            Unload Me
        Else
            Cancel = 1
        End If
    End If
End Sub


Private Sub ListView2_DblClick()
    frmModificarMDO.nuevo_form = Me
    frmModificarMDO.lblCPP = Me.ListView2.selectedItem.ListSubItems(4)
    frmModificarMDO.lblSector = Me.ListView2.selectedItem.ListSubItems(3)
    frmModificarMDO.lblTarea = Me.ListView2.selectedItem.ListSubItems(1) & " - " & Me.ListView2.selectedItem.ListSubItems(5)
    frmModificarMDO.lblDescripcion = Me.ListView2.selectedItem.ListSubItems(8)
    frmModificarMDO.idDesMDO = Me.ListView2.selectedItem
    frmModificarMDO.txtCantOp = Me.ListView2.selectedItem.ListSubItems(1)
    frmModificarMDO.txtTiempo = Me.ListView2.selectedItem.ListSubItems(2)
    frmModificarMDO.Show 1
End Sub
Private Sub txtCantOp_Change()
    If Trim(Me.txtCodigoMDO) = Empty Or Trim(Me.txtCantOp) = Empty Or Trim(Me.txtTiempo) = Empty Then

        Me.btnAgregarMDO.Enabled = False
    Else
        Me.btnAgregarMDO.Enabled = True

    End If
    grabado = False
End Sub
Private Sub txtCantOp_GotFocus()
    foco txtCantOp
End Sub

Private Sub txtCantOp_Validate(Cancel As Boolean)
    If Not IsNumeric(Me.txtCantOp) Then Cancel = True
End Sub
Private Sub txtCodigoMDO_Change()
    If Trim(Me.txtCodigoMDO) = Empty Or Trim(Me.txtCantOp) = Empty Or Trim(Me.txtTiempo) = Empty Then

        Me.btnAgregarMDO.Enabled = False
    Else
        Me.btnAgregarMDO.Enabled = True

    End If
    grabado = False
End Sub

Private Sub txtCodigoMDO_GotFocus()
    foco Me.txtCodigoMDO
End Sub

Private Sub txtCodigoMDO_KeyPress(KeyAscii As Integer)
    Set base = New classNuevoElemento
    If KeyAscii = 13 Then
        base.ver_detalle_mdo CInt(Me.txtCodigoMDO), idcpp, cantxproc, mdoDescrip, Tarea, Sector, Valor
        vIdMDO = Me.txtCodigoMDO
        Me.lblCPP = cantxproc
        vidCPP = idcpp


        lblTarea = Tarea
        lblDescripcion = mdoDescrip
        lblSector = Sector
        Me.lblValor = Valor


    End If
End Sub

Private Sub txtCodigoMDO_LostFocus()
    Set base = New classNuevoElemento
    If Not Trim(Me.txtCodigoMDO) = Empty Then
        base.ver_detalle_mdo CInt(Me.txtCodigoMDO), idcpp, cantxproc, mdoDescrip, Tarea, Sector, Valor
        vIdMDO = Me.txtCodigoMDO
        Me.lblCPP = cantxproc
        vlblidCPP = idcpp
        lblTarea = Tarea
        lblDescripcion = mdoDescrip
        lblSector = Sector
        Me.lblValor = Valor

    End If
End Sub

Private Sub txtCodigoMDO_Validate(Cancel As Boolean)
'If Not IsNumeric(Me.txtCodigoMDO) Then Cancel = True
End Sub




Private Sub txtTiempo_Change()
    If Trim(Me.txtCodigoMDO) = Empty Or Trim(Me.txtCantOp) = Empty Or Trim(Me.txtTiempo) = Empty Then

        Me.btnAgregarMDO.Enabled = False
    Else
        Me.btnAgregarMDO.Enabled = True

    End If
    grabado = False
End Sub

Private Sub txtTiempo_GotFocus()
    foco txtTiempo

End Sub

Private Sub txtTiempo_Validate(Cancel As Boolean)
    If Not IsNumeric(Me.txtTiempo) Then Cancel = True
End Sub

