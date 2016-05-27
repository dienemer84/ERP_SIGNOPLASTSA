VERSION 5.00
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#12.0#0"; "CODEJO~1.OCX"
Begin VB.Form frmComprasOrdenEdit 
   Caption         =   "Edición de Orden de Compra Nº "
   ClientHeight    =   7515
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10035
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmComprasOrdenEdit.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7515
   ScaleWidth      =   10035
   StartUpPosition =   3  'Windows Default
   Tag             =   "Edición de Orden de Compra Nº "
   Begin VB.Frame Frame1 
      Caption         =   "Detalles disponibles de peticiones de oferta"
      Height          =   6930
      Left            =   4215
      TabIndex        =   5
      Top             =   465
      Width           =   5715
      Begin XtremeReportControl.ReportControl rptctrlDetalleOrden 
         Height          =   6600
         Left            =   90
         TabIndex        =   6
         Top             =   240
         Width           =   5520
         _Version        =   786432
         _ExtentX        =   9737
         _ExtentY        =   11642
         _StockProps     =   64
         BorderStyle     =   3
      End
   End
   Begin VB.Frame fraDetallesDisponibles 
      Caption         =   "Detalles disponibles de peticiones de oferta"
      Height          =   6930
      Left            =   60
      TabIndex        =   3
      Top             =   450
      Width           =   3615
      Begin XtremeReportControl.ReportControl rptctrlDetallesDisponibles 
         Height          =   6600
         Left            =   90
         TabIndex        =   4
         Top             =   240
         Width           =   3435
         _Version        =   786432
         _ExtentX        =   6059
         _ExtentY        =   11642
         _StockProps     =   64
         BorderStyle     =   3
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<-"
      Height          =   615
      Left            =   3780
      TabIndex        =   2
      Top             =   4080
      Width           =   345
   End
   Begin VB.CommandButton Command 
      Caption         =   "->"
      Height          =   615
      Left            =   3780
      TabIndex        =   1
      Top             =   3000
      Width           =   330
   End
   Begin VB.Label lblProveedor 
      AutoSize        =   -1  'True
      Caption         =   "Proveedor: "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   135
      TabIndex        =   0
      Tag             =   "Proveedor: "
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmComprasOrdenEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    FormHelper.Customize Me
End Sub
