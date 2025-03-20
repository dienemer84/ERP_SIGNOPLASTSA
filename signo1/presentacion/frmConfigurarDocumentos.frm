VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmConfigurarDocumentos 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configurar Documentos"
   ClientHeight    =   14715
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   16245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   14715
   ScaleWidth      =   16245
   Begin XtremeSuiteControls.PushButton cmdGuardar 
      Height          =   375
      Left            =   120
      TabIndex        =   27
      Top             =   8520
      Width           =   2640
      _Version        =   786432
      _ExtentX        =   4657
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Guardar"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.GroupBox GroupBox4 
      Height          =   2460
      Left            =   90
      TabIndex        =   26
      Top             =   6000
      Width           =   2655
      _Version        =   786432
      _ExtentX        =   4683
      _ExtentY        =   4339
      _StockProps     =   79
      Caption         =   "Nuevo"
      BackColor       =   12632256
      UseVisualStyle  =   -1  'True
      Begin MSComctlLib.ListView ListView1 
         Height          =   1455
         Left            =   120
         TabIndex        =   34
         Top             =   495
         Width           =   2430
         _ExtentX        =   4286
         _ExtentY        =   2566
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Campo"
            Object.Width           =   3705
         EndProperty
      End
      Begin XtremeSuiteControls.PushButton Command3 
         Height          =   315
         Left            =   600
         TabIndex        =   30
         Top             =   2085
         Width           =   1485
         _Version        =   786432
         _ExtentX        =   2619
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Agregar"
         UseVisualStyle  =   -1  'True
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Campo"
         Height          =   195
         Left            =   135
         TabIndex        =   33
         Top             =   225
         Width           =   495
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox2 
      Height          =   1830
      Left            =   90
      TabIndex        =   8
      Top             =   2385
      Width           =   2655
      _Version        =   786432
      _ExtentX        =   4683
      _ExtentY        =   3228
      _StockProps     =   79
      Caption         =   "Posición"
      BackColor       =   -2147483633
      UseVisualStyle  =   -1  'True
      Begin VB.TextBox txtTop 
         Height          =   285
         Left            =   795
         TabIndex        =   14
         Top             =   1425
         Width           =   1200
      End
      Begin VB.TextBox txtLeft 
         Height          =   285
         Left            =   795
         TabIndex        =   13
         Top             =   1035
         Width           =   1200
      End
      Begin VB.TextBox txtHeight 
         Height          =   285
         Left            =   795
         TabIndex        =   10
         Top             =   645
         Width           =   1200
      End
      Begin VB.TextBox txtWidth 
         Height          =   285
         Left            =   795
         TabIndex        =   9
         Top             =   315
         Width           =   1200
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Y"
         Height          =   195
         Left            =   615
         TabIndex        =   16
         Top             =   1470
         Width           =   105
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "X"
         Height          =   195
         Left            =   600
         TabIndex        =   15
         Top             =   1080
         Width           =   105
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Alto"
         Height          =   195
         Left            =   450
         TabIndex        =   12
         Top             =   690
         Width           =   270
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ancho"
         Height          =   195
         Left            =   255
         TabIndex        =   11
         Top             =   360
         Width           =   465
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   2235
      Left            =   60
      TabIndex        =   2
      Top             =   105
      Width           =   2655
      _Version        =   786432
      _ExtentX        =   4683
      _ExtentY        =   3942
      _StockProps     =   79
      Caption         =   "Documento"
      BackColor       =   -2147483633
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.ComboBox cboTipoDocs 
         Height          =   315
         Left            =   810
         TabIndex        =   32
         Top             =   300
         Width           =   1695
         _Version        =   786432
         _ExtentX        =   2990
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Text            =   "ComboBox1"
      End
      Begin VB.TextBox txtNombre 
         Height          =   285
         Left            =   810
         TabIndex        =   28
         Top             =   660
         Width           =   1710
      End
      Begin XtremeSuiteControls.PushButton Command1 
         Height          =   525
         Left            =   2190
         TabIndex        =   7
         Top             =   1095
         Width           =   285
         _Version        =   786432
         _ExtentX        =   503
         _ExtentY        =   926
         _StockProps     =   79
         Caption         =   "SET"
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
      End
      Begin VB.TextBox txtWidthContenedor 
         Height          =   285
         Left            =   795
         TabIndex        =   4
         Top             =   1020
         Width           =   1155
      End
      Begin VB.TextBox txtHeightContenedor 
         Height          =   285
         Left            =   795
         TabIndex        =   3
         Top             =   1410
         Width           =   1155
      End
      Begin XtremeSuiteControls.PushButton Command2 
         Height          =   300
         Left            =   570
         TabIndex        =   17
         Top             =   1800
         Width           =   1440
         _Version        =   786432
         _ExtentX        =   2540
         _ExtentY        =   529
         _StockProps     =   79
         Caption         =   "Background"
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   300
         Left            =   2040
         TabIndex        =   18
         Top             =   1800
         Width           =   390
         _Version        =   786432
         _ExtentX        =   688
         _ExtentY        =   529
         _StockProps     =   79
         Caption         =   "X"
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Nombre"
         Height          =   195
         Left            =   150
         TabIndex        =   31
         Top             =   330
         Width           =   555
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Nombre"
         Height          =   195
         Left            =   150
         TabIndex        =   29
         Top             =   690
         Width           =   555
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Alto"
         Height          =   195
         Left            =   390
         TabIndex        =   6
         Top             =   1395
         Width           =   270
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ancho"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   1050
         Width           =   465
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   15
      Left            =   390
      Top             =   9480
   End
   Begin VB.PictureBox PicContainer 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Height          =   9150
      Left            =   2820
      ScaleHeight     =   9090
      ScaleWidth      =   13185
      TabIndex        =   1
      Top             =   90
      Width           =   13245
   End
   Begin VB.PictureBox trashContainer 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FF00FF&
      Height          =   780
      Left            =   15285
      OLEDropMode     =   1  'Manual
      Picture         =   "frmConfigurarDocumentos.frx":0000
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   0
      Top             =   9390
      Width           =   780
   End
   Begin XtremeSuiteControls.GroupBox GroupBox3 
      Height          =   1725
      Left            =   90
      TabIndex        =   19
      Top             =   4260
      Width           =   2655
      _Version        =   786432
      _ExtentX        =   4683
      _ExtentY        =   3043
      _StockProps     =   79
      Caption         =   "Formato"
      BackColor       =   -2147483633
      UseVisualStyle  =   -1  'True
      Begin VB.CheckBox chkTachado 
         Caption         =   "Tachado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   -1  'True
         EndProperty
         Height          =   195
         Left            =   1275
         TabIndex        =   24
         Top             =   585
         Width           =   1245
      End
      Begin VB.CheckBox chkSubrayado 
         Caption         =   "Subrayado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1275
         TabIndex        =   23
         Top             =   885
         Width           =   1245
      End
      Begin VB.CheckBox chkCursiva 
         Caption         =   "Cursiva"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   22
         Top             =   585
         Width           =   1020
      End
      Begin VB.CheckBox chkNegrita 
         Caption         =   "Negrita"
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
         TabIndex        =   21
         Top             =   885
         Width           =   1020
      End
      Begin XtremeSuiteControls.PushButton PushButton2 
         Height          =   300
         Left            =   570
         TabIndex        =   20
         Top             =   1260
         Width           =   1440
         _Version        =   786432
         _ExtentX        =   2540
         _ExtentY        =   529
         _StockProps     =   79
         Caption         =   "Fuente"
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
      End
      Begin VB.Label lblFuente 
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   315
         Width           =   2295
      End
   End
   Begin XtremeSuiteControls.PushButton cmdPrintDemo 
      Height          =   375
      Left            =   120
      TabIndex        =   35
      Top             =   8895
      Width           =   2640
      _Version        =   786432
      _ExtentX        =   4657
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Imprimir Demo"
      UseVisualStyle  =   -1  'True
   End
End
Attribute VB_Name = "frmConfigurarDocumentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents DragEvents As ClsEvents
Attribute DragEvents.VB_VarHelpID = -1
Private WithEvents DragEventsTrash As ClsEvents
Attribute DragEventsTrash.VB_VarHelpID = -1
Dim WithEvents txt As TextBox
Attribute txt.VB_VarHelpID = -1
Dim ctrl As Control
Dim ultimoHwnd As Long
Dim Doc As documento

Dim ultimo_id As String
Dim ultimoWidht As Double
Dim ultimoHeight As Double

Public Function Document(nvalue As documento)
    Set Doc = nvalue
End Function



Private Sub cboTipoDocs_Click()
    FillFields
End Sub

Private Sub FillFields()
    Dim dic As New Collection
    Set dic = DAODocumentos.GetFieldsByTipo(Me.cboTipoDocs.ItemData(Me.cboTipoDocs.ListIndex))


    Dim x As ListItem

    Me.ListView1.ListItems.Clear

    Dim A As DTOCampoBD
    For Each A In dic
        Set x = Me.ListView1.ListItems.Add(, , A.NombreCampo)
        x.Tag = A.CampoEnBD
    Next



End Sub


Private Sub chkCursiva_Click()
    SetearFormato
End Sub

Private Sub chkNegrita_Click()
    SetearFormato
End Sub

Private Sub chkSubrayado_Click()
    SetearFormato
End Sub

Private Sub chkTachado_Click()
    SetearFormato
End Sub

Private Sub cmdGuardar_Click()

    If Doc Is Nothing Then
        Set Doc = New documento
    End If
    Doc.Alto = Val(Me.txtHeightContenedor)
    Doc.Ancho = Val(Me.txtWidthContenedor)
    Doc.estado = True
    Doc.nombre = UCase(Me.txtNombre)

    Dim det As DocumentoDetalle
    Dim col As New Collection
    Dim ctrl As Control

    For Each ctrl In Me.Controls
        If TypeOf ctrl Is Timer Then GoTo prox
        If ctrl.Container.Name = "PicContainer" Then
            Set det = New DocumentoDetalle
            det.Alineacion = vbCenter
            det.Alto = ctrl.Height
            det.Ancho = ctrl.Width
            det.Cursiva = ctrl.FontItalic
            Set det.documento = Doc
            det.Fijo = True
            det.Negrita = ctrl.FontBold
            det.nombreFuente = ctrl.FontName
            det.PosX = ctrl.Left
            det.PosY = ctrl.Top
            det.Subrayado = ctrl.FontUnderline
            det.Tachado = ctrl.FontStrikethru
            det.Tag = ctrl.Tag
            det.Tamano = ctrl.FontSize
            col.Add det
        End If
prox:
    Next

    Set Doc.detalles = col

    If Not DAODocumentos.SaveDocumento(Doc, True) Then GoTo err1 Else MsgBox "Guardado correctamente!"

    Exit Sub
err1:

End Sub

Private Sub ResizePicContainter()
    Me.PicContainer.Width = FormHelper.ConvertCmToTwip(Val(Me.txtWidthContenedor))
    Me.PicContainer.Height = FormHelper.ConvertCmToTwip(Val(Me.txtHeightContenedor))
    PosicionarTrash
End Sub

Private Sub cmdPrintDemo_Click()
    Printer.Orientation = 2
    Printer.Height = FormHelper.ConvertCmToTwip(Doc.Alto)
    Printer.Width = FormHelper.ConvertCmToTwip(Doc.Ancho)


    Dim deta As DocumentoDetalle
    Dim xx As Long


    Printer.Line (0, 0)-(Printer.Width, 0)
    Printer.Line (0, 0)-(0, Printer.Width)

    Printer.Line (Printer.Width, 0)-(0, Printer.Height)

    For Each deta In Doc.detalles
        '    Printer.CurrentX = deta.PosX
        '    Printer.CurrentY = 0
        '
        '    Printer.Print "punta " & xx
        '    xx = xx + 1
    Next


    Printer.EndDoc
End Sub

Private Sub Command1_Click()
    ResizePicContainter
    PosicionarTrash
End Sub

Private Sub RetrieveFields()

    Me.txtHeightContenedor = Doc.Alto
    Me.txtWidthContenedor = Doc.Ancho
    Me.txtNombre = Doc.nombre
    ResizePicContainter
    Dim deta As DocumentoDetalle

    For Each deta In Doc.detalles
        Dim objid As String
        objid = funciones.CreateGUID
        Set txt = Controls.Add("vb.textbox", "txt" & objid)
        Set txt.Container = Me.PicContainer
        txt.Width = deta.Ancho
        txt.Height = deta.Alto
        txt.Top = deta.PosY
        txt.Left = deta.PosX
        txt.BorderStyle = 1
        txt.Appearance = 0
        txt.Visible = True
        txt.Text = deta.Tag
        txt.Tag = deta.Tag
        txt.FontBold = deta.Negrita
        txt.FontItalic = deta.Cursiva
        txt.FontName = deta.nombreFuente
        txt.FontSize = deta.Tamano
        txt.FontStrikethru = deta.Tachado
        txt.FontUnderline = deta.Subrayado
        txt.Alignment = deta.Alineacion
        txt.Text = deta.Tag

        ReiniciarContainer
        ultimoHwnd = txt.hWnd
        ultimoWidht = RedondearDecimales(FormHelper.ConvertTwipToCm(txt.Width))
        ultimoHeight = RedondearDecimales(FormHelper.ConvertTwipToCm(txt.Height))
        MostrarValores




    Next









End Sub


Private Sub Command2_Click()
    On Error GoTo err1
    Dim archivo As String
    frmPrincipal.CD.ShowOpen
    archivo = frmPrincipal.CD.filename
    Me.PicContainer.Picture = LoadPicture(archivo)
    ShowContainerSize
    Exit Sub
err1:

End Sub

Private Sub limpiar()
    Me.txtLeft = 0
    Me.txtTop = 0
    Me.txtWidth = 0
    Me.txtHeight = 0
End Sub

Private Sub Command3_Click()

    Me.chkCursiva = False
    Me.chkNegrita = False
    Me.chkSubrayado = False
    Me.chkTachado = False
    Dim objid As String
    objid = funciones.CreateGUID
    Set txt = Controls.Add("vb.textbox", "txt" & objid)
    Set txt.Container = Me.PicContainer
    txt.Width = 3500
    txt.Height = 100
    txt.Top = 0
    txt.Left = 0
    txt.BorderStyle = 1
    txt.Appearance = 0
    txt.Visible = True
    txt.Text = Me.ListView1.selectedItem.Tag
    'txt.MultiLine = True

    ReiniciarContainer
    ultimoHwnd = txt.hWnd
    ultimoWidht = RedondearDecimales(FormHelper.ConvertTwipToCm(txt.Width))
    ultimoHeight = RedondearDecimales(FormHelper.ConvertTwipToCm(txt.Height))
    MostrarValores

End Sub
Private Sub Command4_Click()
'Debug.Print DAODocumentos.FindAll(True).count
End Sub

Private Sub DragEventsTrash_ObjectDrop(hWndContainerSource As Long, hWndContainer As Long, hWndObject As Long, Reject As Boolean)
    Me.trashContainer.Appearance = 0
    For Each ctrl In Me.Controls
        If TypeOf ctrl Is TextBox Then


            If ctrl.hWnd = hWndObject Then
                Controls.remove ctrl.Name
                limpiar

            End If
        End If
    Next
    Me.trashContainer.Appearance = 1
End Sub

Private Sub Form_Load()
    Customize Me
    limpiar
    Set DragEvents = New ClsEvents
    Set DragEventsTrash = New ClsEvents
    ShowContainerSize

    LlenarTiposDocs
    If IsSomething(Doc) Then RetrieveFields

End Sub

Private Sub LlenarTiposDocs()
    Me.cboTipoDocs.AddItem "Cheque"
    Me.cboTipoDocs.ItemData(cboTipoDocs.NewIndex) = TipoDocumentoImpresion.TDI_Cheque

    Me.cboTipoDocs.ListIndex = 0
End Sub

Private Sub PosicionarTrash()
    Me.trashContainer.Top = Me.PicContainer.Height + Me.PicContainer.Top + 50
    Me.trashContainer.Left = Me.PicContainer.Width + Me.PicContainer.Left - Me.trashContainer.Width
End Sub
Private Sub ShowContainerSize()
    Me.txtHeightContenedor = funciones.RedondearDecimales(FormHelper.ConvertTwipToCm(Me.PicContainer.Height))
    Me.txtWidthContenedor = funciones.RedondearDecimales(FormHelper.ConvertTwipToCm(Me.PicContainer.Width))

    PosicionarTrash
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call UnInitializeAllContainer
End Sub
Private Sub DragEvents_StopDrag(hWndContainer As Long, hWndObject As Long, x As Long, y As Long, Width As Long, Height As Long, Cancel As Boolean)
    ultimoHwnd = hWndObject
    ultimoWidht = RedondearDecimales(FormHelper.ConvertPixelToCm(Width))
    ultimoHeight = RedondearDecimales(FormHelper.ConvertPixelToCm(Height))
    Me.Timer1.Enabled = True
End Sub

Private Sub PushButton1_Click()
    Me.PicContainer.Picture = Nothing
End Sub

Private Sub PushButton2_Click()
    Dim stdfont As New stdfont

    On Error GoTo err1
    Set ctrl = LocateLastControlInContainer

    frmPrincipal.CD.Flags = cdlCFPrinterFonts

    If IsSomething(ctrl) Then
        frmPrincipal.CD.FontBold = ctrl.FontBold
        frmPrincipal.CD.FontItalic = ctrl.FontItalic
        frmPrincipal.CD.FontName = ctrl.FontName
        frmPrincipal.CD.FontStrikethru = ctrl.FontStrikethru

        frmPrincipal.CD.FontUnderline = ctrl.FontUnderline
        frmPrincipal.CD.FontSize = ctrl.FontSize

        frmPrincipal.CD.ShowFont


        stdfont.Bold = frmPrincipal.CD.FontBold
        stdfont.Italic = frmPrincipal.CD.FontItalic
        stdfont.Underline = frmPrincipal.CD.FontUnderline
        stdfont.Strikethrough = frmPrincipal.CD.FontStrikethru
        stdfont.Size = frmPrincipal.CD.FontSize
        stdfont.Name = frmPrincipal.CD.FontName

        SetearFormato stdfont
        Set ctrl.Font = stdfont

        MostrarValores ctrl
    End If


    Exit Sub
err1:
End Sub

Private Sub radCentrado_Click()
    SetearFormato
End Sub

Private Sub radDerecha_Click()
    SetearFormato
End Sub

Private Sub radIzquierda_Click()
    SetearFormato
End Sub

Private Sub PushButton3_Click()

End Sub

Private Sub Timer1_Timer()
    If Not Me.Timer1.Enabled Then Exit Sub
    Me.Timer1.Enabled = False
    MostrarValores


End Sub


Private Function LocateLastControlInContainer() As Control
    Dim ctrl As Control

    For Each ctrl In Me.Controls
        If TypeOf ctrl Is Timer Then GoTo prox
        If ctrl.Container.Name = "PicContainer" Then


            'Debug.Print ctrl.hwnd, ultimoHwnd

            If ctrl.hWnd = ultimoHwnd Then

                Set LocateLastControlInContainer = ctrl
                Exit For
            End If
        End If
prox:
    Next

    Set LocateLastControlInContainer = ctrl
End Function

Private Sub SetearFormato(Optional Font As stdfont)

    Set ctrl = LocateLastControlInContainer()

    If IsSomething(ctrl) Then

        If IsSomething(Font) Then
            ctrl.Font.Size = Font.Size
            ctrl.Font.Bold = Font.Bold
            ctrl.Font.Italic = Font.Italic
            ctrl.Font.Underline = Font.Underline
            ctrl.Font.Strikethrough = Font.Strikethrough
            ctrl.Font.Name = Font.Name
        Else
            ctrl.Font.Bold = Me.chkNegrita.value
            ctrl.Font.Italic = Me.chkCursiva.value
            ctrl.Font.Underline = Me.chkSubrayado.value
            ctrl.Font.Strikethrough = Me.chkTachado.value
            ctrl.Font.Name = Me.lblFuente

        End If


        '        If Me.radCentrado.value Then ctrl.Alignment = vbCenter
        '        If Me.radDerecha.value Then ctrl.Alignment = vbRightJustify
        '        If Me.radIzquierda.value Then ctrl.Alignment = vbLeftJustify
        '
        '        ReiniciarContainer
        '        ultimoHwnd = ctrl.hwnd

    End If
End Sub


Private Sub MostrarValores(Optional ctrl As Object = Nothing)

    If ctrl Is Nothing Then Set ctrl = LocateLastControlInContainer()

    If IsSomething(ctrl) Then
        Me.txtLeft = RedondearDecimales(FormHelper.ConvertTwipToCm(ctrl.Left))
        Me.txtTop = RedondearDecimales(FormHelper.ConvertTwipToCm(ctrl.Top))
        Me.txtHeight = ultimoHeight
        Me.txtWidth = ultimoWidht
        Me.chkNegrita = Abs(ctrl.FontBold)
        Me.chkCursiva = Abs(ctrl.FontItalic)
        Me.chkSubrayado = Abs(ctrl.FontUnderline)
        Me.chkTachado = Abs(ctrl.FontStrikethru)
        Me.lblFuente = ctrl.FontName & "," & ctrl.FontSize



        'If ctrl.Alignment = vbRightJustify Then Me.radDerecha.value = True '
        ' If ctrl.Alignment = vbLeftJustify Then Me.radIzquierda.value = True
        '  If ctrl.Alignment = vbCenter Then Me.radCentrado.value = True


    End If
End Sub









Private Sub txtHeight_KeyUp(KeyCode As Integer, Shift As Integer)
    Set ctrl = LocateLastControlInContainer()
    ctrl.Height = FormHelper.ConvertCmToTwip(Val(Me.txtHeight))

End Sub

Private Sub txtLeft_KeyUp(KeyCode As Integer, Shift As Integer)
    Set ctrl = LocateLastControlInContainer()
    ctrl.Left = FormHelper.ConvertCmToTwip(Val(Me.txtLeft))
End Sub

Private Sub txtTop_KeyUp(KeyCode As Integer, Shift As Integer)
    Set ctrl = LocateLastControlInContainer()
    ctrl.Top = FormHelper.ConvertCmToTwip(Val(Me.txtTop))

End Sub

Private Sub txtWidth_KeyUp(KeyCode As Integer, Shift As Integer)
    Set ctrl = LocateLastControlInContainer()
    ctrl.Width = FormHelper.ConvertCmToTwip(Val(Me.txtWidth))

End Sub

Private Sub ReiniciarContainer()
    Call UnInitializeAllContainer
    Call InitializeContainer(PicContainer, True, True, True, DragEvents)
    Call InitializeContainer(trashContainer, True, True, True, DragEventsTrash)
End Sub
