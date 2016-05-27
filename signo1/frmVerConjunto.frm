VERSION 5.00
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#13.2#0"; "Codejock.ReportControl.v13.2.1.ocx"
Begin VB.Form frmDesarrolloConjuntosVer 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ver Conjunto..."
   ClientHeight    =   5190
   ClientLeft      =   2130
   ClientTop       =   1230
   ClientWidth     =   6330
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmVerConjunto.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   6330
   Begin XtremeReportControl.ReportControl reportControl 
      Height          =   5010
      Left            =   75
      TabIndex        =   1
      Top             =   90
      Width           =   6165
      _Version        =   851970
      _ExtentX        =   10874
      _ExtentY        =   8837
      _StockProps     =   64
      BorderStyle     =   3
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4005
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6390
      Width           =   990
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Main"
      Visible         =   0   'False
      Begin VB.Menu mnuPieza 
         Caption         =   "NOMBREPIEZA"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSeparador 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDesarrollo 
         Caption         =   "Desarrollo"
      End
      Begin VB.Menu mnuCambiar 
         Caption         =   "Cambiar"
      End
      Begin VB.Menu mnuCopiar 
         Caption         =   "Copiar"
      End
      Begin VB.Menu mnuMateriales 
         Caption         =   "Materiales"
      End
      Begin VB.Menu mnuTiempos 
         Caption         =   "Tiempos"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmDesarrolloConjuntosVer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim claseS As New classStock
Dim claseP As New classPlaneamiento
'Dim vconjunto As Long
Private m_pieza As pieza
Private PiezaMenuEmergente As pieza

Public Property Let idConjunto(conjunto As Long)
    'vconjunto = conjunto
    Set m_pieza = DAOPieza.FindById(conjunto, FL_3, , True)
End Property

Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Form_Load()
FormHelper.Customize Me
    Me.ReportControl.Columns.DeleteAll

    Dim repCol As ReportColumn
    Set repCol = Me.ReportControl.Columns.Add(0, "Pieza", 285, True)
        repCol.Alignment = xtpAlignmentIconLeft
        repCol.Sortable = False
        repCol.AllowDrag = False
        repCol.AllowRemove = False
        repCol.Icon = 0
        repCol.TreeColumn = True
    Set repCol = Me.ReportControl.Columns.Add(1, "Cantidad", 60, True)
        repCol.Alignment = xtpAlignmentRight
        repCol.Sortable = False
        repCol.AllowDrag = False
        repCol.AllowRemove = False
        repCol.Icon = 0
    Set repCol = Me.ReportControl.Columns.Add(2, "Kg", 60, True)
        repCol.Alignment = xtpAlignmentRight
        repCol.Sortable = False
        repCol.AllowDrag = False
        repCol.AllowRemove = False
        repCol.Icon = 0
    
    Me.ReportControl.PaintManager.HorizontalGridStyle = xtpGridSmallDots
    Me.ReportControl.PaintManager.VerticalGridStyle = xtpGridSmallDots
    
    '----------------------
    
    Me.caption = "Conjunto [" & m_pieza.nombre & "]"
    
    Me.ReportControl.Records.DeleteAll
    AddRecord m_pieza
    Me.ReportControl.Populate
    
End Sub

Private Sub AddRecord(ByVal P As pieza, Optional ByVal parent As ReportRecord = Nothing)
    Dim rec As ReportRecord
  
    If parent Is Nothing Then
        Set rec = Me.ReportControl.Records.Add
    Else
        Set rec = parent.Childs.Add
    End If

    rec.AddItem P.nombre
    rec.AddItem P.cantidad
    rec.AddItem P.Kilage
    rec.Tag = P.id
    rec.Expanded = True
    
    Dim tmpPieza As pieza
    For Each tmpPieza In P.PiezasHijas
        AddRecord tmpPieza, rec
    Next tmpPieza
    
End Sub


Private Sub mnuCambiar_Click()
    If Not PiezaMenuEmergente Is Nothing Then
        frmDefinirConjunto.accion = 1
        frmDefinirConjunto.idPiezaMadre = PiezaMenuEmergente.id
        frmDefinirConjunto.Show
    End If
End Sub

Private Sub mnuCopiar_Click()
    If Not PiezaMenuEmergente Is Nothing Then
        Dim nuevoNombre As String
        If MsgBox("¿Está seguro de copiar el conjunto?", vbYesNo + vbQuestion, "Confirmación") = vbYes Then
            nuevoNombre = funciones.ingreso()
            If Len(Trim(nuevoNombre)) > 0 Then
                If claseS.copiarConjuntoV2(PiezaMenuEmergente.id, nuevoNombre, 0) Then
                    MsgBox "Conjunto copiado satisfactoriamente!", vbInformation, "Información"
                Else
                    MsgBox "Error en la copia de conjuntos!", vbCritical, "Error"
                End If
            End If
        End If
    End If
End Sub

Private Sub mnuDesarrollo_Click()

    If Not PiezaMenuEmergente Is Nothing Then
        Dim F As New frmDesarrollo
        Load F
        F.CargarPieza PiezaMenuEmergente.id
        F.Show
    End If
End Sub
Private Sub mnuMateriales_Click()
    If Not PiezaMenuEmergente Is Nothing Then
        DAOOrdenTrabajo.informePiezaMateriales PiezaMenuEmergente.id, 3, True
    End If
End Sub

Private Sub mnuTiempos_Click()
    If Not PiezaMenuEmergente Is Nothing Then
        frmDesarrolloConjuntosTiempos.idConjunto = PiezaMenuEmergente.id
        frmDesarrolloConjuntosTiempos.Show
    End If
End Sub

'Private Sub lista()
'    Dim rs As recordset
'    Dim x As ListItem
'    Set rs = conectar.RSFactory("select detalle from stock where id=" & vconjunto)
'    While Not rs.EOF
'        rs.MoveNext
'    Wend
'    Set rs = conectar.RSFactory("select s.id,s.detalle,sc.cantidad,sc.idPiezaHija from stockConjuntos sc inner join stock s on s.id=sc.idPiezaHija where sc.idPiezaPadre=" & vconjunto)
'    While Not rs.EOF
'        claseS.CalcularValorePieza rs!idPiezaHija, Kg, m2
'        Set x = Me.lstConjunto.ListItems.Add(, , rs!Cantidad)
'        x.SubItems(1) = rs!detalle
'        x.SubItems(2) = Kg * rs!Cantidad & " Kg"
'        x.Tag = rs!Id
'
'        rs.MoveNext
'    Wend
'
'End Sub



Private Sub reportControl_MouseUp(Button As Integer, Shift As Integer, x As Long, y As Long)
    If Button = 2 Then
        Dim hitinfo As ReportHitTestInfo
        Set hitinfo = Me.ReportControl.HitTest(x, y)
        
        If Not hitinfo.item Is Nothing Then
            Dim record As ReportRecord
            Set record = hitinfo.item.record
            
            Set PiezaMenuEmergente = m_pieza.LocatePiezaInPiezasHijas(record.Tag)
            If Not PiezaMenuEmergente Is Nothing Then
                Me.mnuPieza.caption = PiezaMenuEmergente.nombre
                Me.mnuCopiar.Enabled = PiezaMenuEmergente.EsConjunto
                Me.mnuCambiar.Enabled = PiezaMenuEmergente.EsConjunto
                Me.mnuTiempos.Enabled = PiezaMenuEmergente.EsConjunto
            
                Me.PopupMenu Me.mnuMain
            End If
        End If
        
    End If
End Sub
