VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmAdminIVA 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configurar IVA"
   ClientHeight    =   2655
   ClientLeft      =   540
   ClientTop       =   690
   ClientWidth     =   5220
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   5220
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Command1"
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   3000
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5175
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2295
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   4048
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         AllowAddNew     =   -1  'True
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
         Caption         =   "Configuración de IVA"
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "idIVA"
            Caption         =   "Pos"
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
            DataField       =   "detalle"
            Caption         =   "Detalle"
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
            DataField       =   "alicuota"
            Caption         =   "Alicuota"
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
         BeginProperty Column03 
            DataField       =   "valid"
            Caption         =   "Valido"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "0,000E+00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            AllowRowSizing  =   0   'False
            BeginProperty Column00 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   404.787
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2324.977
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               ColumnAllowSizing=   0   'False
               Locked          =   -1  'True
               WrapText        =   -1  'True
               ColumnWidth     =   705.26
            EndProperty
         EndProperty
      End
   End
   Begin VB.Menu valid 
      Caption         =   "Validar"
      Visible         =   0   'False
      Begin VB.Menu validar 
         Caption         =   "validar"
      End
   End
End
Attribute VB_Name = "frmAdminIVA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim clasea As New classAdministracion
Dim rs As Recordset

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub DataGrid1_DblClick()
    Dim valid As Boolean
    If MsgBox("¿Está seguro de cambiar el estado?", vbYesNo, "Confirmación") = vbYes Then
        With rs


            valid = rs!valido
            valid = Not valid
            'id = rs!idIVA
            'If valido = 0 Then
            '    valiNew = 1
            'ElseIf valido = 1 Then
            '    valiNew = 0
            'End If
            rs!valido = valid
            rs.Update
        End With
        'claseA.ejecutarComando "update AdminConfigIVA set valido=" & valiNew & " where idIVA=" & id
        Form_Load
    End If
End Sub



Private Sub Form_Load()
    FormHelper.Customize Me
    Set rs = conectar.RSFactoryCliente("select idIVA,detalle,alicuota, valido  from AdminConfigIVA ")
    Set Me.DataGrid1.DataSource = rs
End Sub
