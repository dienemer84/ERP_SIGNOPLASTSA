VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCostosIncidencia 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Costos"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5130
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   5130
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog CD 
      Left            =   4440
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ Incidencias]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5055
      Begin VB.CommandButton Command3 
         BackColor       =   &H00E0E0E0&
         Cancel          =   -1  'True
         Caption         =   "Salir"
         Height          =   375
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   4680
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Exportar "
         Height          =   375
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   4680
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Imprimir"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   4680
         Width           =   975
      End
      Begin MSComctlLib.ListView lstCostos 
         Height          =   3135
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   5530
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   2011
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label MAT 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   1320
         TabIndex        =   12
         Top             =   1080
         Width           =   3495
      End
      Begin VB.Label MDO 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   1320
         TabIndex        =   11
         Top             =   840
         Width           =   3495
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "% Materiales"
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
         TabIndex        =   10
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "% M.D.O."
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
         TabIndex        =   9
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   1320
         TabIndex        =   4
         Top             =   600
         Width           =   3495
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
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
         TabIndex        =   3
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   1320
         TabIndex        =   2
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Elemento"
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
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmCostosIncidencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As Recordset
Dim classStock As New classStock
Dim idPieza As Long
Dim strsql As String
Dim cli As Long
Public Property Let cliente(ncli As Long)
    cli = ncli
End Property
Public Property Let idp(Id As Long)
    idPieza = Id
End Property

Private Sub Command1_Click()
    If MsgBox("¿Seguro de imprimir?", vbYesNo, "Confirmación") = vbYes Then

        AnchoCol = 0

        Set rs = conectar.RSFactory("select detalle from stock where id=" & idPieza)
        Pieza = rs!detalle

        For i = 1 To Me.lstCostos.ColumnHeaders.count
            AnchoCol = AnchoCol + lstCostos.ColumnHeaders(i).Width
        Next
        Espacio = 0
        Printer.Font.Bold = True
        Printer.Print "Incidencia de materiales"
        Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)
        Printer.Print
        Printer.Print "Cliente: " & classStock.cLiente_pieza(CInt(idPieza))
        Printer.Print "Pieza: " & Pieza
        Printer.Print

        Printer.Print "% MDO: " & Me.MDO
        Printer.Print "% Materiales: " & Me.MAT
        'Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)


        With Me.lstCostos
            Printer.Font.Bold = True
            For i = 1 To .ColumnHeaders.count
                Espacio = Espacio + CInt(.ColumnHeaders(i).Width)    '* Printer.ScaleWidth / AnchoCol)
                Printer.Print lstCostos.ColumnHeaders(i).text;
                Printer.CurrentX = Espacio
            Next
            Printer.Font.Bold = False
            Printer.Print

            'Imprime una línea
            'Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)
            Printer.Print

            For i = 1 To .ListItems.count
                Espacio = 0

                Set lItem = .ListItems(i)
                Printer.Print lItem.text;
                'Recorremos las columnas
                For x = 1 To .ColumnHeaders.count - 1
                    Espacio = Espacio + CInt(.ColumnHeaders(x).Width)    ' * Printer.ScaleWidth / AnchoCol)
                    Printer.CurrentX = Espacio

                    Printer.Print lItem.SubItems(x);


                Next

                'Otro espacio en blanco
                Printer.Print


            Next

        End With

        Printer.Print

        ''Imprime la línea de final de impresión

        'Texto del pie>

        Printer.Print Format(Date, "dd-mm-yyyy")


        'Comenzamos la impresión
        Printer.EndDoc
    End If


End Sub

Private Sub Command2_Click()
    A = classStock.exporta(idPieza, Me.lstCostos)
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    FormHelper.Customize Me
    Dim rs2 As Recordset
    Dim rs1 As Recordset
    Set rs1 = conectar.RSFactory("SELECT r.rubro, if(m.id_unidad=2,sum(((dm.largoTerm*dm.anchoTerm)/1000000*dm.cantidad*m.pesoxunidad)*valor_unitario),if(m.id_unidad=3,sum(dm.cantidad*dm.largoTerm*pesoxunidad)/1000*vm.valor_unitario, if(m.id_unidad=4,0,sum(dm.cantidad)*vm.valor_unitario))) as valor from rubros r,grupos g, materiales m, desarrollo_material dm, stock s, valores_MATERIALES vm where s.id=" & idPieza & " and dm.id_pieza=s.id and dm.id_material=m.id and m.id_rubro=r.id and m.id_grupo=g.id and vm.id_material=m.id  group by r.id")
    c = 0
    While Not rs1.EOF
        c = c + rs1!Valor
        rs1.MoveNext
    Wend


    Set rs2 = conectar.RSFactory("select  t.tarea,sum(dm.cantidad*dm.tiempo*vm.valor) as mdo from tareas t, desarrollo_mdo dm, valores_MDO vm where dm.codigo=t.id and vm.id_tarea=t.id and dm.id_pieza=" & idPieza & " group by t.id")
    md = 0
    While Not rs2.EOF
        md = rs2!MDO + md
        rs2.MoveNext
    Wend


    totp = c + md

    Me.MAT = funciones.FormatearDecimales((c / totp) * 100, 2) & "%"

    Me.MDO = funciones.FormatearDecimales((md / totp) * 100, 2) & "%"

    Me.Label2 = classStock.detalle_pieza(CLng(idPieza))
    Me.Label4 = classStock.cLiente_pieza(CLng(idPieza))
    Dim x As ListItem

    tot = 0



    'While Not rs.EOF
    'tot = tot + rs!valor
    'c = c + 1
    'rs.MoveNext
    'Wend
    If c > 0 Then
        rs1.MoveFirst

        Set x = Me.lstCostos.ListItems.Add(, , "Materiales")
        While Not rs1.EOF
            Set x = Me.lstCostos.ListItems.Add(, , "")
            x.SubItems(1) = rs1!rubro
            x.SubItems(2) = Math.Round((rs1!Valor / c) * 100, 0) & "%"
            rs1.MoveNext
        Wend

    End If


    'Set X = Me.lstCostos.ListItems.Add(, , "Mano de obra")
    'rs2.MoveFirst
    'While Not rs2.EOF
    '        Set X = Me.lstCostos.ListItems.Add(, , "")
    '        X.SubItems(1) = rs2!tarea
    '            X.SubItems(2) = Math.Round((rs2!MDO / md) * 100, 0) & "%"
    'rs2.MoveNext
    'Wend

End Sub




