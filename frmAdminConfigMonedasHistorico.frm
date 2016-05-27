VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmAdminConfigCambioHistorico 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Fructuación moneda extranjera"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   12360
   ShowInTaskbar   =   0   'False
   Begin MSChart20Lib.MSChart grafica 
      Height          =   3735
      Left            =   5280
      OleObjectBlob   =   "frmAdminConfigMonedasHistorico.frx":0000
      TabIndex        =   3
      Top             =   720
      Width           =   7095
   End
   Begin MSComctlLib.ListView lstValores 
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   5220
      _ExtentX        =   9208
      _ExtentY        =   7011
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Moneda"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Valor"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Fecha"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Usuario"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Moneda"
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
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frmAdminConfigCambioHistorico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim claseSP As New classSignoplast
Dim rs As Recordset
'Dim claseA As New classAdministracion
Dim vIdMoneda As Long

Public Property Let IdMoneda(nIdMoneda As Long)
    vIdMoneda = nIdMoneda
End Property


Private Sub Form_Load()
    FormHelper.Customize Me


    Dim x As ListItem
    Set rs = conectar.RSFactory("select m.nombre_corto,cm.FechaActualizacion, cm.IdUsuarioActualizacion,cm.valor from AdminConfigMonedasHistorial cm inner join AdminConfigMonedas m on cm.idMoneda=m.id where idMoneda=" & vIdMoneda)
    While Not rs.EOF
        Set x = Me.lstValores.ListItems.Add(, , rs!Nombre_corto)
        x.SubItems(1) = rs!Valor
        x.SubItems(2) = rs!fechaActualizacion
        x.SubItems(3) = claseSP.queUsuario(rs!idUsuarioActualizacion)
        rs.MoveNext
    Wend
    grafico


End Sub


Function grafico()
    grafica.ColumnCount = Me.lstValores.ListItems.count
    grafica.rowcount = 1




    For x = 1 To grafica.ColumnCount
        grafica.Column = x
        grafica.data = Me.lstValores.ListItems(x).ListSubItems(1)
        grafica.RowLabel = lstValores.ListItems(x).ListSubItems(1)
        grafica.ColumnLabel = Format(CDate(lstValores.ListItems(x).ListSubItems(2)), "dd-mm-yyyy")
    Next
End Function

