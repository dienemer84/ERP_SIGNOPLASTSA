VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmAdminPagosTransferenciasBancariasEditar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Modificar Datos de Transferencia"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   6600
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   6375
      Begin VB.TextBox txtNumeroTransferencia 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Text            =   "Text"
         Top             =   1080
         Width           =   4935
      End
      Begin XtremeSuiteControls.PushButton btnBorrarContenido 
         Height          =   375
         Index           =   6
         Left            =   5400
         TabIndex        =   4
         Top             =   1080
         Width           =   495
         _Version        =   786432
         _ExtentX        =   873
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "X"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblNumeroCbte 
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   4995
         _Version        =   786432
         _ExtentX        =   8819
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Datos de la Transferencia"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Line Line 
         X1              =   5880
         X2              =   240
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label lblTextoAdicional 
         Caption         =   "Número de Transferencia"
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
         Index           =   1
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Width           =   2775
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   6375
      Begin XtremeSuiteControls.PushButton btnGuardar 
         Height          =   495
         Index           =   0
         Left            =   4560
         TabIndex        =   1
         Top             =   240
         Width           =   1575
         _Version        =   786432
         _ExtentX        =   2778
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Guardar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnRestablecer 
         Height          =   495
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   1575
         _Version        =   786432
         _ExtentX        =   2778
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Reestablecer"
         UseVisualStyle  =   -1  'True
      End
   End
End
Attribute VB_Name = "frmAdminPagosTransferenciasBancariasEditar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private TransfBancaria As clsTransferenciaBcaria

Public Property Let idTransfBancaria(value As Long)
    Set TransfBancaria = DAOTransferenciaBcaria.FindById(value)

End Property


Private Sub btnBorrarContenido_Click(Index As Integer)
    Me.txtNumeroTransferencia(1) = ""
End Sub




Private Sub btnGuardar_Click(Index As Integer)
    If MsgBox("Está segur@ de los cambios realizados?", vbYesNo, "Confirmación") = vbYes Then
        

        
        TransfBancaria.Comprobante = Me.txtNumeroTransferencia(1)

        If DAOTransferenciaBcaria.ActualizarNroComprobante(TransfBancaria) Then
            MsgBox "Los datos del comprobante han sido actualizados.", vbOKOnly + vbInformation
   
        Else
            Err.Raise "9999", "Guardando transferencia", Err.Description
        End If
    End If
End Sub


Private Sub btnRestablecer_Click(Index As Integer)
    Me.txtNumeroTransferencia(1) = TransfBancaria.Comprobante

End Sub


Private Sub Form_Load()
    Customize Me
    
    Me.lblNumeroCbte.caption = TransfBancaria.FechaOperacion & "- " & FormatCurrency(funciones.FormatearDecimales(TransfBancaria.Monto))
     
     
    Me.txtNumeroTransferencia(1) = TransfBancaria.Comprobante

End Sub


