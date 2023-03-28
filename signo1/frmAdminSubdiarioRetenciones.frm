VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmAdminSubdiarioRetenciones 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Subdiario de retenciones..."
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   15165
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   15165
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ Seleccione período de cobranzas ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15135
      Begin XtremeSuiteControls.ProgressBar progreso 
         Height          =   420
         Left            =   4440
         TabIndex        =   10
         Top             =   1180
         Visible         =   0   'False
         Width           =   4215
         _Version        =   786432
         _ExtentX        =   7435
         _ExtentY        =   741
         _StockProps     =   93
         Appearance      =   6
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00E0E0E0&
         Cancel          =   -1  'True
         Caption         =   "Salir"
         Height          =   255
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Exportar"
         Height          =   255
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Generar"
         Height          =   255
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Imprimir"
         Height          =   255
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   600
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker DTHasta 
         Height          =   255
         Left            =   840
         TabIndex        =   5
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         _Version        =   393216
         Format          =   62783489
         CurrentDate     =   39660
      End
      Begin MSComCtl2.DTPicker DTDesde 
         Height          =   255
         Left            =   840
         TabIndex        =   6
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         _Version        =   393216
         Format          =   62783489
         CurrentDate     =   39660
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
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
         TabIndex        =   8
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
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
         Top             =   360
         Width           =   615
      End
   End
   Begin MSComctlLib.ListView lstSubdiarioRetenciones 
      Height          =   4575
      Left            =   0
      TabIndex        =   9
      Top             =   1800
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   8070
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
      NumItems        =   0
   End
End
Attribute VB_Name = "frmAdminSubdiarioRetenciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim clase As New classAdministracion
Dim rs As Recordset
Dim desde As Date
Dim hasta As Date

Public Sub Command2_Click()

    Dim x As ListItem
    Me.lstSubdiarioRetenciones.ListItems.Clear


    If Me.DTDesde > Me.DTHasta Then
        MsgBox "Error en la seleccion de fechas!", vbCritical, "Error"
        Exit Sub
    Else
        desde = Me.DTDesde
        hasta = Me.DTHasta
    End If

    Set rs = clase.subdiario_retenciones(Me.lstSubdiarioRetenciones, desde, hasta)
End Sub

Private Sub Command3_Click()
    clase.exportaSubDiarioRetenciones Me.lstSubdiarioRetenciones, desde, hasta, progreso

End Sub

Private Sub Command4_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    FormHelper.Customize Me
    desde = CDate(1 & "-" & Month(Now) & "-" & Year(Now))
    Me.DTDesde = desde
    Me.DTHasta = Now


End Sub



