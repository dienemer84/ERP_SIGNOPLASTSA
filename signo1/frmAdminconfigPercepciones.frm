VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmAdminconfigPercepciones 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Percepciones"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtValor 
      Height          =   285
      Left            =   840
      TabIndex        =   4
      Top             =   840
      Width           =   3375
   End
   Begin VB.TextBox txtDescripcion 
      Height          =   285
      Left            =   840
      TabIndex        =   3
      Top             =   480
      Width           =   3375
   End
   Begin VB.TextBox txtCodigo 
      Height          =   285
      Left            =   840
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Agregar"
      Height          =   255
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin MSDataGridLib.DataGrid dataGrid1 
      Height          =   1815
      Left            =   80
      TabIndex        =   0
      Top             =   1320
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   3201
      _Version        =   393216
      AllowUpdate     =   -1  'True
      ColumnHeaders   =   -1  'True
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
      Caption         =   "Lista de percepciones"
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "codigo"
         Caption         =   "Cod"
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
         DataField       =   "percepcion"
         Caption         =   "Percepción"
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
         DataField       =   "porcentaje"
         Caption         =   "Valor"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Valido"
         Caption         =   "Disp"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   5
            Format          =   "0,00%"
            HaveTrueFalseNull=   1
            TrueValue       =   "Si"
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
         RecordSelectors =   0   'False
         BeginProperty Column00 
            Alignment       =   2
            ColumnAllowSizing=   0   'False
            Locked          =   -1  'True
            ColumnWidth     =   494.929
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   2145.26
         EndProperty
         BeginProperty Column02 
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnAllowSizing=   0   'False
            Locked          =   -1  'True
            ColumnWidth     =   645.165
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Valor"
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
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Detalle"
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
      TabIndex        =   6
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Código"
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
      TabIndex        =   5
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmAdminconfigPercepciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As Recordset
Dim clasea As New classAdministracion
Private Sub Command1_Click()
    Unload Me
End Sub
Private Sub Command2_Click()
    On Error GoTo err55
    Dim strsql As String
    If MsgBox("¿Desea agregar esta percepción?", vbYesNo, "Confirmación") Then
        cod = UCase(Trim(Me.txtCodigo))
        descripcion = normaliza(Me.txtDescripcion)
        Valor = CDbl(Me.txtValor)
        strsql = "Insert into AdminConfigPercepciones (codigo,percepcion,porcentaje) values ('" & cod & "','" & UCase(descripcion) & "'," & Valor & ")"
        If clasea.ejecutarComando(strsql) Then
            MsgBox "Percepción ingresada con éxito!", vbInformation, "Información"
        Else
            MsgBox "Se produjo un error. No se guardarán los cambios!", vbCritical, "Error"
        End If
    End If
    Exit Sub
err55:
    If Err.Number = 13 Then
        MsgBox "Ingrese datos válidos!", vbCritical, "Error"
    Else
        MsgBox "Error " & Err.numbres & ": " & Err.Description
    End If
    verRecordset
End Sub
Private Sub verRecordset()
    Set rs = conectar.RSFactoryCliente("select * from AdminConfigPercepciones")


    Set Me.DataGrid1.DataSource = rs
End Sub
Private Sub DataGrid1_DblClick()
    Dim estado As Boolean
    If MsgBox("¿Desea cambiar el estado de esta perecpción?", vbYesNo, "Confirmación") = vbYes Then
        estado = rs!valido
        estado = Not estado
        rs!valido = estado
        rs.Update
        verRecordset
    End If
End Sub
Private Sub Form_Load()
    FormHelper.Customize Me
    verRecordset
End Sub

