VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmMaterializacion 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Materializacion"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   11325
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Exportar"
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5280
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1800
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Imprimir"
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4920
      Width           =   1095
   End
   Begin MSComctlLib.ListView lstMaterialiazacion 
      Height          =   4815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11265
      _ExtentX        =   19870
      _ExtentY        =   8493
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Pieza"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Material"
         Object.Width           =   13053
      EndProperty
   End
End
Attribute VB_Name = "frmMaterializacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim claseP As New classPlaneamiento
Dim claseS As New classStock
Dim vId As Long
Dim vot As Boolean
Dim vpresu As Boolean
Dim votro As Boolean

Public Property Let Id(nId As Long)
    vId = nId
End Property

Public Property Let Ot(nuot As Boolean)
    vot = nuot
End Property

Public Property Let presu(npresu As Boolean)
    vpresu = npresu
End Property
Public Property Let otro(notro As Boolean)
    votro = notro
End Property


Private Sub Command1_Click()
    claseS.exportaMaterializacion vId, vpresu, vot, votro
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub
Private Sub Command4_Click()
    MsgBox "no implementado"
    Exit Sub

    On Error GoTo err4
    Me.CommonDialog1.ShowPrinter
    For x = 1 To Me.CommonDialog1.Copies
        Imprimir
    Next
    Exit Sub
err4:
End Sub
Private Function Imprimir()
    Dim rs As Recordset
    Dim rs2 As Recordset
    Printer.Font.Size = 10
    Espacio = 0
    Printer.Font.Bold = True
    Printer.Print "MATERIALIZACION"
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)
    Printer.Print Tab(75);
    Printer.Print rs2!Cantidad;
    Printer.Print Tab(85);
    Printer.Print rs2!Remito;
    Printer.Print Tab(95);
    Printer.Print Format(rs2!FEcha, "dd-mm-yyyy")

    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)
    Printer.Print
    Printer.Print
    ''Imprime la línea de final de impresión
    'Texto del pie>
    Printer.Print "Fecha emisión " & Format(Date, "dd-mm-yyyy")
    'Comenzamos la impresión
    'Printer.EndDoc
End Function

Private Sub Form_Load()
    FormHelper.Customize Me
    Dim lista As Recordset
    Dim r As Recordset
    Dim x As ListItem
    Set lista = claseS.ListaPiezas(vId, vot, vpresu, votro)

    'lista.Sort

    While Not lista.EOF
        A = claseS.detalle_pieza(lista!idPieza)
        Set x = Me.lstMaterialiazacion.ListItems.Add(, , A)
        Set r = claseS.materializacion(lista!idPieza)
        While Not r.EOF
            x.SubItems(1) = r!codigo & " " & r!rubro & " " & r!Grupo & " " & r!descripcion & " " & Math.Round(r!Espesor, 2) & "mm"
            Set x = Me.lstMaterialiazacion.ListItems.Add
            r.MoveNext
        Wend
        lista.MoveNext
    Wend
End Sub
