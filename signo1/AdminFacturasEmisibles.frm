VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmAdminFacturasEmisibles 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Facturas emitibles"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7470
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ Comprobantes ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   3480
      TabIndex        =   3
      Top             =   0
      Width           =   3975
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   1575
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   2778
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Facturas emisibles"
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "TipoFactura"
            Caption         =   "Tipo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "alicuota"
            Caption         =   "I.V.A."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "discrimina"
            Caption         =   "Disc"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "Sí"
               FalseValue      =   "No"
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   7
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
               Alignment       =   2
               ColumnAllowSizing=   0   'False
               Locked          =   -1  'True
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column01 
               ColumnAllowSizing=   0   'False
               Locked          =   -1  'True
               ColumnWidth     =   1544.882
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               ColumnAllowSizing=   0   'False
               Locked          =   -1  'True
               ColumnWidth     =   705.26
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ Manejo de facturas ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3375
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Cancel          =   -1  'True
         Caption         =   "Salir"
         Height          =   255
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1560
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Agregar"
         Default         =   -1  'True
         Height          =   255
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1560
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Discrimina el I.V.A."
         Height          =   255
         Left            =   960
         TabIndex        =   6
         Top             =   1080
         Width           =   1695
      End
      Begin VB.ComboBox cboIVA 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Alicuota "
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
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo "
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
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmAdminFacturasEmisibles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim clasea As New classAdministracion

Private Sub Command1_Click()
    Unload Me
End Sub


Private Sub DataGrid1_DblClick()
    Dim estado As Boolean
    If MsgBox("¿Desea cambiar el estado?", vbYesNo, "Confirmación") = vbYes Then
        estado = rs!valido
        estado = Not estado

        rs!valido = estado
        rs.Update

        'verRecordset
        LlenarRS

    End If

End Sub

Private Sub Form_Load()
    FormHelper.Customize Me
    LlenarCbo
    LlenarRS
End Sub


Private Sub Text1_GotFocus()
    foco Me.Text1
End Sub


Private Sub LlenarRS()


    Set rs = conectar.RSFactoryCliente("select ft.TipoFactura,f.discriminaIVA as discrimina,f.discriminaIVA,concat(i.detalle,' ',i.Alicuota,'%') as alicuota from AdminConfigFacturas f inner join AdminConfigIVA i on i.idIVA=f.idIva inner join AdminConfigFacturasTipos ft on f.tipoFactura=ft.id")

    Set Me.DataGrid1.DataSource = rs
End Sub

Private Sub LlenarCbo()
    Dim rs As Recordset, strsql As String
    cboIVA.Clear
    strsql = "select idIVA,detalle,alicuota from AdminConfigIVA where idIVA not in (select idIVA from AdminConfigFacturas)"
    Set rs = conectar.RSFactory(strsql)
    While Not rs.EOF
        muestra = rs!detalle & " " & rs!alicuota & "%"
        Me.cboIVA.AddItem muestra
        Me.cboIVA.ItemData(cboIVA.NewIndex) = rs!idIVA
        rs.MoveNext
    Wend
    If cboIVA.ListCount > 0 Then
        cboIVA.ListIndex = 0
    End If

End Sub
