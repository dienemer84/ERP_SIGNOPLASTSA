VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmRequeCompraDefinir 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Definir Entregas"
   ClientHeight    =   5970
   ClientLeft      =   2190
   ClientTop       =   1965
   ClientWidth     =   10950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   10950
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   5655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10935
      Begin VB.Frame Frame2 
         Caption         =   "Estado"
         Height          =   735
         Left            =   6240
         TabIndex        =   8
         Top             =   2880
         Width           =   4575
         Begin VB.Label lblEstado 
            Alignment       =   2  'Center
            Caption         =   "Label3"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   4335
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Agregar"
         Height          =   255
         Left            =   4920
         TabIndex        =   6
         Top             =   3000
         Width           =   1095
      End
      Begin VB.TextBox txtCantidad 
         Height          =   285
         Left            =   3240
         TabIndex        =   5
         Text            =   "0"
         Top             =   3000
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   255
         Left            =   720
         TabIndex        =   4
         Top             =   3000
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   20905985
         CurrentDate     =   39094
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2535
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   4471
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
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
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "id"
            Caption         =   "id"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "material"
            Caption         =   "Material"
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
         BeginProperty Column03 
            DataField       =   "cantidad"
            Caption         =   "Cant"
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
         BeginProperty Column04 
            DataField       =   "x"
            Caption         =   "Largo"
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
         BeginProperty Column05 
            DataField       =   "y"
            Caption         =   "Ancho"
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
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            RecordSelectors =   0   'False
            BeginProperty Column00 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   14.74
            EndProperty
            BeginProperty Column01 
               ColumnAllowSizing=   -1  'True
               ColumnWidth     =   4110.236
            EndProperty
            BeginProperty Column02 
               ColumnAllowSizing=   -1  'True
               ColumnWidth     =   3465.071
            EndProperty
            BeginProperty Column03 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   629.858
            EndProperty
            BeginProperty Column04 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   1094.74
            EndProperty
            BeginProperty Column05 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   1094.74
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Height          =   1455
         Left            =   120
         TabIndex        =   7
         Top             =   3600
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   2566
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
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
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "id"
            Caption         =   "id"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "cantidad"
            Caption         =   "Cantidad"
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
            DataField       =   "fecha"
            Caption         =   "Fecha"
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
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            RecordSelectors =   0   'False
            BeginProperty Column00 
               Alignment       =   2
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column01 
               ColumnAllowSizing=   -1  'True
               ColumnWidth     =   2654.929
            EndProperty
            BeginProperty Column02 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   1094.74
            EndProperty
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "Cantidad"
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
         TabIndex        =   3
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha"
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
         TabIndex        =   2
         Top             =   3000
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmRequeCompraDefinir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As Recordset
Dim rs2 As Recordset
Dim classStock As New classStock
Dim idDetalleReque As Long
Dim idReque As Long

Public Property Let idR(newId)
idReque = newId
End Property

Private Sub Command1_Click()
Dim deta As String
Dim idDetalleReque As Long

idDetalleReque = rs2!id
valo = comprobarEstado(idDetalleReque, deta)
Me.lblEstado = deta
End Sub

Private Sub Form_Load()
Me.Frame1.Caption = "[ Requerimiento Nro. " & Format(idReque, "0000") & " ]"
llenarLst

End Sub


Private Function llenarLst()
Set rs2 = classStock.CrearRS("select rd.id,concat(m.codigo,' ',m.descripcion) as material, rd.detalle, rd.cantidad, rd.x, rd.y  from requerimientosDetalles rd inner join sp.materiales m on m.id=rd.idMaterial where idReque=" & idReque)
Set Me.DataGrid1.DataSource = rs2
End Function


Private Function comprobarEstado(idDetalleReque As Long, ByRef DetalleEstado As String) As Boolean
Set rs = classStock.CrearRS("Select cantidad from requerimientosDetalles where id=" & idDetalleReque)
cantPedido = rs!cantidad
Set rs = classStock.CrearRS("Select sum(cantidad) as cantidad from requerimientosDetallesEntregas where idDetalleReque=" & idDetalleReque & " group by idDetalleReque")
If rs.EOF And rs.BOF Then
cantDefinido = 0
Else
cantDefinido = rs!cantidad
End If

If cantDefinido = cantPedido Then
 comprobarEstado = True
 detallePedido = "Totalmente definida"
ElseIf cantDefinido < cantPedido Then
comprobarEstado = False
 detallePedido = "Parcialmente definida"
Else
 detallePedido = "SDF"
End If

Me.lblEstado = detallePedido
End Function
