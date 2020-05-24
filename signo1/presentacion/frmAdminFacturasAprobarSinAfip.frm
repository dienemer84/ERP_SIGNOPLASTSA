VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~3.OCX"
Begin VB.Form frmAdminFacturasAprobarSinAfip 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Aprobar comprobante sin envio a AFIP"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   5655
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Carga de datos"
      Height          =   1695
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5415
      Begin VB.TextBox txtCAE 
         Height          =   315
         Left            =   1200
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   420
         Width           =   3615
      End
      Begin MSComCtl2.DTPicker dtVtoCAE 
         Height          =   375
         Left            =   1200
         TabIndex        =   3
         Top             =   960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   58785793
         CurrentDate     =   43960
      End
      Begin VB.Label Label1 
         Caption         =   "CAE:"
         Height          =   255
         Left            =   600
         TabIndex        =   6
         Top             =   450
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Vto.CAE:"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   1020
         Width           =   735
      End
   End
   Begin XtremeSuiteControls.PushButton cmdAceptar 
      Height          =   375
      Left            =   3960
      TabIndex        =   0
      Top             =   1920
      Width           =   1575
      _Version        =   786432
      _ExtentX        =   2778
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Aceptar"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton cmdCancelar 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   1575
      _Version        =   786432
      _ExtentX        =   2778
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Cancelar"
      UseVisualStyle  =   -1  'True
   End
End
Attribute VB_Name = "frmAdminFacturasAprobarSinAfip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Factura As Factura

Private Sub cmdAceptar_Click()
On Error GoTo errCae

If MsgBox("¿Está seguro de actualizar los datos?", vbYesNo, "Confirmación") = vbYes Then

    Factura.CAE = Me.txtCAE
    Factura.CAEVto = Me.dtVtoCAE
    DAOFactura.ActualizarCAE Factura
    MsgBox "Datos de CAE actualizados correctamente", vbInformation, "Proceso correcto"
   Unload Me
End If
Exit Sub

errCae:
 MsgBox Err.Description, vbCritical, "Se produjo un error"
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Customize Me
    Me.txtCAE = Factura.CAE
    Me.dtVtoCAE = Factura.CAEVto
    

End Sub
