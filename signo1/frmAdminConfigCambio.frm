VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmAdminConfigCambio 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Centro de cambio"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   7845
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command4 
      Cancel          =   -1  'True
      Caption         =   "Command4"
      Height          =   375
      Left            =   480
      TabIndex        =   21
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ Cambio ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   0
      TabIndex        =   12
      Top             =   3960
      Width           =   7815
      Begin VB.ComboBox cboOrigen 
         Height          =   315
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   360
         Width           =   735
      End
      Begin VB.ComboBox cboDestino 
         Height          =   315
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtOrigen 
         Height          =   285
         Left            =   1440
         TabIndex        =   15
         Top             =   360
         Width           =   4815
      End
      Begin VB.TextBox txtDestino 
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   840
         Width           =   4815
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Aceptar"
         Default         =   -1  'True
         Height          =   255
         Left            =   6600
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Origen"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Destino"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ MONEDA ]"
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
      TabIndex        =   1
      Top             =   1920
      Width           =   7815
      Begin VB.CommandButton Command3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Historico"
         Height          =   255
         Left            =   6600
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Actualizar"
         Height          =   255
         Left            =   6600
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1560
         Width           =   1095
      End
      Begin VB.ComboBox cboMonedaCambio 
         Height          =   315
         Left            =   2640
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Left            =   840
         TabIndex        =   9
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox txtCambio 
         Height          =   285
         Left            =   840
         TabIndex        =   8
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox txtDetalle 
         Height          =   285
         Left            =   840
         TabIndex        =   7
         Top             =   720
         Width           =   6495
      End
      Begin VB.TextBox txtMoneda 
         Height          =   285
         Left            =   840
         TabIndex        =   6
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cambio"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Detalle"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Moneda"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   615
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   3201
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
      Caption         =   "Centro de cambio"
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "nombre_corto"
         Caption         =   "Moneda"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "nombre_largo"
         Caption         =   "Detalle"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "cambio"
         Caption         =   "Cambio"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "destino"
         Caption         =   "Moneda"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "patron"
         Caption         =   "Patrón"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   5
            Format          =   ""
            HaveTrueFalseNull=   1
            TrueValue       =   "Sí"
            FalseValue      =   "No"
            NullValue       =   ""
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   7
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "FechaActual"
         Caption         =   "Fecha"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   3
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         RecordSelectors =   0   'False
         BeginProperty Column00 
            Locked          =   -1  'True
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   2505.26
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column03 
            Locked          =   -1  'True
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column04 
            Locked          =   -1  'True
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column05 
            Locked          =   -1  'True
            ColumnWidth     =   1395.213
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmAdminConfigCambio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Id As Long
Dim rs As Recordset
Dim clasea As New classAdministracion
Dim cambioOriginal As Double
Dim claseC As New classConfigurar
Private Sub Command1_Click()
    Dim ori As Double
    If IsNumeric(Me.txtOrigen) Then
        ori = CDbl(Me.txtOrigen)
        Me.txtDestino = clasea.realizaCambio(ori, Me.cboOrigen.ListIndex, Me.cboDestino.ListIndex)

    End If
End Sub

Private Sub Command2_Click()
    If MsgBox("¿Está seguro de actualizar?", vbYesNo, "Confirmación") = vbYes Then

        If MsgBox("La moneda que está actualizando es: " & vbCrLf & Me.txtMoneda & "-" & Me.txtDetalle & vbCrLf & "Desea continuar con la actualización?", vbYesNo, "Confirmación") = vbYes Then

            Id = rs!Id
            idMonedaCambio = Me.cboMonedaCambio.ListIndex
            Cambio = CDbl(Me.txtCambio)
            hoy = Format(Now, "yyyy-mm-dd")
            detalle = Me.txtDetalle
            moneda = Me.txtMoneda

            'comentado el 15-07-2014 por solicitud de sabrina scaldafferro
            'If Cambio = cambioOriginal Then
            '    clasea.ejecutarComando "update AdminConfigMonedas set idMonedaCambio=" & idMonedaCambio & ", cambio= " & Cambio & ", nombre_largo='" & detalle & "',nombre_corto='" & Moneda & "' where id=" & Id
            ' Else
            clasea.ejecutarComando "update AdminConfigMonedas set idMonedaCambio=" & idMonedaCambio & ", cambio= " & Cambio & ", nombre_largo='" & detalle & "',nombre_corto='" & moneda & "',fechaActual='" & hoy & "' where id=" & Id
            ' End If

            clasea.ejecutarComando "insert into AdminConfigMonedasHistorial (IdMoneda, FechaActualizacion, idUsuarioActualizacion,Valor) values (" & Id & ",'" & funciones.datetimeFormateada(Now) & "'," & funciones.getUser & "," & Cambio & ")"
            Me.mostrarRS
        End If

    End If
End Sub

Private Sub Command3_Click()
    frmAdminConfigCambioHistorico.IdMoneda = Id
    frmAdminConfigCambioHistorico.Show

End Sub

Private Sub Command4_Click()
    Unload Me
End Sub



Public Sub mostrarRS()
    Set rs = conectar.RSFactoryCliente("select a.id,a.nombre_largo,a.fechaActual, a.nombre_corto,a.cambio,a.patron,a.IdMonedaCambio,b.nombre_corto as destino from AdminConfigMonedas a inner join AdminConfigMonedas b on b.id=a.idMonedaCambio")
    Set Me.DataGrid1.DataSource = rs
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Me.txtCambio = rs!Cambio
    'Me.txtDestino = rs!destino
    Id = rs!Id
    Me.txtDetalle = rs!nombre_largo
    Me.txtMoneda = rs!Nombre_corto
    Me.txtFecha = rs!FechaActual
    cambioOriginal = rs!Cambio

    Me.cboMonedaCambio.ListIndex = funciones.PosIndexCbo(rs!idMonedaCambio, Me.cboMonedaCambio)
End Sub

Private Sub Form_Load()
    FormHelper.Customize Me
    mostrarRS


    DAOMoneda.LlenarCombo Me.cboOrigen
    DAOMoneda.LlenarCombo Me.cboDestino
    DAOMoneda.LlenarCombo Me.cboMonedaCambio
    DataGrid1_RowColChange 0, 0
End Sub

