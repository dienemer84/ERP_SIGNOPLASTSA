VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmPrintPreview 
   BackColor       =   &H00F0E1D1&
   Caption         =   "Previa de la Impresión"
   ClientHeight    =   8925
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   12525
   Icon            =   "frmPrintPreview.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8925
   ScaleWidth      =   12525
   StartUpPosition =   3  'Windows Default
   Begin GridEX20.GEXPreview GEXPreview1 
      Height          =   6735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   11880
      BackColor       =   15786449
      BeginProperty ToolbarFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PageSetupText   =   "Configurar Pagina..."
      PrintText       =   "Imprimir..."
      CloseButtonText =   "Cerrar"
   End
End
Attribute VB_Name = "frmPrintPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    FormHelper.Customize Me
End Sub

Private Sub Form_Resize()
    Me.GEXPreview1.Move 0, 0, ScaleWidth, ScaleHeight
End Sub

Private Sub GEXPreview1_OnCloseClick()
    Unload Me
End Sub

Private Sub GEXPreview1_OnPrintClick(ByVal UsePrintSetupDlg As GridEX20.JSRetBoolean)
    MsgBox "El listado será enviado a la impresora.", vbInformation + vbOKOnly
End Sub
