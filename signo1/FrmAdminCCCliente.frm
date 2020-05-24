VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmAdminCCCliente 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cuenta Corriente"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   300
      Left            =   6570
      TabIndex        =   19
      Top             =   330
      Width           =   270
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6360
      Top             =   6480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ Rango de fecha ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   0
      TabIndex        =   6
      Top             =   5760
      Width           =   2295
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Habilitar"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   975
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   255
         Left            =   720
         TabIndex        =   10
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   52494337
         CurrentDate     =   39196
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   255
         Left            =   720
         TabIndex        =   8
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   52494337
         CurrentDate     =   39196
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Hasta"
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
         Width           =   615
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Desde"
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
         TabIndex        =   7
         Top             =   600
         Width           =   615
      End
   End
   Begin VB.CommandButton Im 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Imprimir"
      Height          =   255
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6120
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Default         =   -1  'True
      Height          =   255
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5760
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ Cuenta Corriente ]"
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
      TabIndex        =   3
      Top             =   480
      Width           =   6855
      Begin MSComctlLib.ListView lstCC 
         Height          =   4815
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   8493
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Fecha"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Comprob"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Nro"
            Object.Width           =   1094
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Debe"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Haber"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Saldo"
            Object.Width           =   2293
         EndProperty
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ver CC"
      Height          =   255
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.ComboBox cboClientes 
      Height          =   315
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   4695
   End
   Begin VB.Label lblSaldo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   3720
      TabIndex        =   18
      Top             =   6240
      Width           =   1335
   End
   Begin VB.Label lblTotalHaber 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
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
      Left            =   3720
      TabIndex        =   17
      Top             =   6000
      Width           =   1335
   End
   Begin VB.Label lblTotalDebe 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
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
      Left            =   3720
      TabIndex        =   16
      Top             =   5760
      Width           =   1335
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Total Haber "
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
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Total Debe "
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
      TabIndex        =   14
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Saldo "
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
      TabIndex        =   13
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Label Label1 
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
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "FrmAdminCCCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim claseA As New classAdministracion
Dim claseS As New classStock

Private Sub Check1_Click()
    estadoRango
End Sub

Private Sub Command1_Click()
    Dim idCliente As Long
    idCliente = Me.cboClientes.ItemData(Me.cboClientes.ListIndex)
    llenarLST (idCliente)
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
Dim f332 As New frmCtaCte
f332.Show
End Sub

Private Sub Form_Load()
FormHelper.Customize Me
    Me.DTPicker1 = Now
    Me.DTPicker2 = Now
    claseS.llenar_combo_clientes Me.cboClientes, 9999
    estadoRango
End Sub

Private Sub llenarLST(idCliente As Long)
    Dim saldo As Double
    Dim totaldebe As Double, totalhaber As Double
    Dim tip As Integer
    Dim strsql As String
    Dim rs As recordset
    Dim x As ListItem
    Me.lstCC.ListItems.Clear
    If Me.Check1 Then    'esta maraco el check de filtro x rango de fecha
        Me.DTPicker1.Enabled = True
        Me.DTPicker2.Enabled = True
        FechaDesde = Format(CDate(Me.DTPicker1), "yyyy-mm-dd")
        FechaHasta = Format(CDate(Me.DTPicker2), "yyyy-mm-dd")
        strsql = "select  tipo,identificador,fecha,if (tipo=0,'Factura',if(tipo=1,'N/C',if(tipo=2,'Recibo',if(tipo=3,'Ret',if(tipo=5,'N/D',0))))) as comprobante,if(operacion=0,debehaber,0) as debe, if(operacion=1,debehaber,0) as haber  from AdminClientesCC where fecha >= '" & FechaDesde & "' and fecha<='" & FechaHasta & "' and idcliente=" & idCliente
    Else
        strsql = "select  tipo,identificador,fecha,if (tipo=0,'Factura',if(tipo=1,'N/C',if(tipo=2,'Recibo',if(tipo=3,'Ret',if(tipo=5,'N/D',0))))) as comprobante,if(operacion=0,debehaber,0) as debe, if(operacion=1,debehaber,0) as haber  from AdminClientesCC where idcliente=" & idCliente
        Me.DTPicker1.Enabled = False
        Me.DTPicker2.Enabled = False
    End If

    Set rs = conectar.RSFactory(strsql)
    saldo = 0
    totaldebe = 0
    totalhaber = 0
    While Not rs.EOF
        Set x = Me.lstCC.ListItems.Add(, , Format(rs!Fecha, "DD-MM-YYYY"))
        x.Tag = identificador
        x.SubItems(1) = rs!Comprobante

        tip = rs!Tipo
        If tip = 2 Then
            IDENT = rs!identificador
        ElseIf tip = 3 Then
            IDENT = rs!identificador
        ElseIf tip = 4 Then
            IDENT = rs!identificador

        Else
            IDENT = claseA.queFactura(rs!identificador)

        End If
        x.SubItems(2) = IDENT



        x.SubItems(3) = funciones.FormatearDecimales(rs!Debe, 2)
        totaldebe = totaldebe + rs!Debe
        x.SubItems(4) = FormatearDecimales(rs!Haber, 2)
        totalhaber = totalhaber + rs!Haber
        saldo = saldo + (rs!Debe - rs!Haber)
        x.SubItems(5) = FormatearDecimales(saldo, 2)
        If saldo < 0 Then x.ListSubItems(4).ForeColor = vbRed
        rs.MoveNext
    Wend
    Me.lblTotalDebe = funciones.FormatearDecimales(totaldebe, 2)
    Me.lblTotalHaber = funciones.FormatearDecimales(totalhaber, 2)
    Me.lblSaldo = funciones.FormatearDecimales(totaldebe - totalhaber, 2)


End Sub


Private Sub estadoRango()
    If Me.Check1 Then    'esta maraco el check de filtrox rango de fecha
        Me.DTPicker1.Enabled = True
        Me.DTPicker2.Enabled = True
    Else
        Me.DTPicker1.Enabled = False
        Me.DTPicker2.Enabled = False
    End If
End Sub

Private Sub Im_Click()
    On Error GoTo err41
    Me.CommonDialog1.Copies = 1
    Me.CommonDialog1.ShowPrinter

    For x = 1 To Me.CommonDialog1.Copies
        Imprimir
    Next x
    Exit Sub
err41:
End Sub


Private Function Imprimir()
    On Error GoTo err31:
    Cliente = Me.cboClientes
    Printer.Font.Size = 10
    AnchoCol = 0

    For i = 1 To Me.lstCC.ColumnHeaders.count
        AnchoCol = AnchoCol + lstCC.ColumnHeaders(i).Width
    Next
    Espacio = 0
    Printer.Font.Bold = True
    a = Printer.Font.Size
    Printer.Font.Size = 12
    Printer.Print "MOVIMIENTOS DE CUENTA CORRIENTE"
    Printer.Font.Bold = False
    Printer.Print Cliente
    Printer.Font.Size = a
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)





    With Me.lstCC

        'Acá se imprimen los encabezados del ListView
        Printer.Font.Bold = True
        For i = 1 To .ColumnHeaders.count

            Espacio = Espacio + CInt(.ColumnHeaders(i).Width * Printer.ScaleWidth / AnchoCol)
            If lstCC.ColumnHeaders(i).Width > 1 Then
                'Printer.Print i
                Printer.Print lstCC.ColumnHeaders(i).text;
            End If
            Printer.CurrentX = Espacio
        Next
        Printer.Font.Bold = False
        Printer.Print

        'Imprime una línea
        Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)
        Printer.Print

        'Este bucle recorre los items y subitems del ListView  y los imprime
        For i = 1 To .ListItems.count
            Espacio = 0

            Set lItem = .ListItems(i)
            Printer.Print lItem.text;
            'Recorremos las columnas
            For x = 1 To .ColumnHeaders.count - 1
                Espacio = Espacio + CInt(.ColumnHeaders(x).Width * Printer.ScaleWidth / AnchoCol)
                Printer.CurrentX = Espacio
                If Me.lstCC.ColumnHeaders(x + 1).Width > 1 Then
                    'If X = 8 Then 'si es la col 11, trunco a 30 caracteres
                    ' Printer.Print truncar(litem.SubItems(X), 30)
                    ' Else
                    Printer.Print lItem.SubItems(x);
                    'End If


                End If
            Next

            'Otro espacio en blanco
            Printer.Print


        Next

    End With

    Printer.Print

    ''Imprime la línea de final de impresión

    'Texto del pie>


    Printer.Print "Fecha emisión: " & Format(Date, "dd-mm-yyyy")
    'Comenzamos la impresión
    Printer.EndDoc

    Exit Function
err31:
    MsgBox Err.Description

End Function
