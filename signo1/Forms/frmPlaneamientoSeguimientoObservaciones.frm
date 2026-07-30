VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmPlaneamientoSeguimientoObservaciones 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ver Observaciones"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7275
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   7275
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.PushButton btnCerrar 
      Height          =   495
      Left            =   2670
      TabIndex        =   1
      Top             =   4320
      Width           =   1935
      _Version        =   786432
      _ExtentX        =   3413
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Cerrar"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
      _Version        =   786432
      _ExtentX        =   12303
      _ExtentY        =   7223
      _StockProps     =   79
      Caption         =   "Observación"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.FlatEdit txtObservacion 
         Height          =   3495
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   6495
         _Version        =   786432
         _ExtentX        =   11456
         _ExtentY        =   6165
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Text            =   "-"
      End
   End
End
Attribute VB_Name = "frmPlaneamientoSeguimientoObservaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub CargarObservacion(Observacion As String)
    Me.txtObservacion.Text = Observacion
    
End Sub

Private Sub btnCerrar_Click()
    Unload Me
End Sub


Private Sub Form_Load()
    FormHelper.Customize Me
End Sub
